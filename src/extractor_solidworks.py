from __future__ import annotations

"""
extractor_solidworks.py
Extrae metadatos forenses de archivos SolidWorks (.sldprt / .sldasm).

Fuentes verificadas con diagnostico_raw.py en archivos reales:
  SummaryInfo[5]  = username del creador ('lupe', 'gladi', etc.)
  SummaryInfo[6]  = fecha de CREACIÓN — queda FIJA aunque guardes de nuevo
  SummaryInfo[7]  = fecha del ÚLTIMO GUARDADO — cambia cada vez que guardas
  Shell col[10]   = Propietario Windows — cambia al copiar (NO usar para autor)

Para detectar copias cuando alguien "abre y guarda":
  - SW_Created_Date (idx 6) sigue igual → mismo origen
  - SW_Saved_Date  (idx 7) cambia     → diferente persona lo guardó
  - Feature tree idéntico              → misma pieza
"""

import os
import re
from typing import Any, Dict, List, Optional, Tuple

import pythoncom
import win32com.client

from config import (
    CAD_EXTENSIONS, IGNORED_FEATURE_TYPES,
    PROBE_PROPERTY_NAMES, SW_SUMMARY_INFO_FIELDS, SW_SUMMARY_INFO_EXTRA,
)
from utils import (
    coerce_tuple_first, fast_file_hash, format_datetime, normalize_text,
)


# ─────────────────────────────────────────────────────────────────────────────
# Sesión
# ─────────────────────────────────────────────────────────────────────────────

class SolidWorksSession:
    def __init__(self) -> None:
        self.app = None
        self._co_initialized = False

    def connect(self):
        if self.app is not None:
            # Verificar que la sesión sigue viva
            try:
                _ = self.app.Visible
                return self.app
            except Exception:
                # Sesión muerta — reconectar
                self.app = None

        if not self._co_initialized:
            pythoncom.CoInitialize()
            self._co_initialized = True

        # Intentar conectar a una instancia existente de SW primero
        try:
            self.app = win32com.client.GetActiveObject("SldWorks.Application")
            return self.app
        except Exception:
            pass

        # Si no hay instancia activa, crear una nueva
        self.app = win32com.client.Dispatch("SldWorks.Application")
        try:
            self.app.Visible = False
        except Exception:
            pass
        return self.app

    def close(self) -> None:
        self.app = None
        if self._co_initialized:
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass
            self._co_initialized = False


# ─────────────────────────────────────────────────────────────────────────────
# Apertura de documentos
# ─────────────────────────────────────────────────────────────────────────────

def _doc_type(path: str) -> Optional[int]:
    return {".sldprt": 1, ".sldasm": 2}.get(os.path.splitext(path.lower())[1])


def _set_safe(obj, attr, val):
    try:
        setattr(obj, attr, val)
    except Exception:
        pass


def _unwrap(result) -> Any:
    if result is None:
        return None
    if isinstance(result, tuple):
        for item in result:
            if item is not None and not isinstance(item, (int, bool)):
                return item
        return None
    return result


def _open_document(sw_app, path: str) -> Tuple[Any, str]:
    doc_type = _doc_type(path)
    if doc_type is None:
        return None, "unsupported"

    # Si el archivo ya está abierto en SolidWorks, usarlo directamente
    try:
        existing = sw_app.GetOpenDocumentByName(path)
        if existing is not None:
            return existing, "AlreadyOpen"
    except Exception:
        pass

    # Intentar por nombre de archivo también
    try:
        basename = os.path.basename(path)
        existing = sw_app.GetOpenDocumentByName(basename)
        if existing is not None:
            return existing, "AlreadyOpen"
    except Exception:
        pass

    # Abrir normalmente
    try:
        spec = sw_app.GetOpenDocSpec(path)
        if spec is not None:
            for a, v in (("Silent", True), ("ReadOnly", False),
                         ("AddToRecentDocumentList", False),
                         ("LightWeight", False), ("LoadModel", True)):
                _set_safe(spec, a, v)
            m = _unwrap(sw_app.OpenDoc7(spec))
            if m is not None:
                return m, "OpenDoc7"
    except Exception:
        pass
    try:
        m = _unwrap(sw_app.OpenDoc6(path, doc_type, 0, "", 0, 0))
        if m is not None:
            return m, "OpenDoc6"
    except Exception:
        pass
    return None, "failed"


def _activate_and_rebuild(sw_app, path: str, model) -> Any:
    """Activa el documento y hace rebuild para cargar el feature tree."""
    if model is None:
        return model
    try:
        # ActivateDoc2 en win32com late-binding a veces es propiedad, no método.
        # Intentar múltiples formas.
        title = None
        try:
            title = model.GetTitle()
        except Exception:
            pass
        title = title or os.path.basename(path)

        activated = False
        for candidate in (title, os.path.basename(path)):
            if activated:
                break
            # Forma 1: llamada directa
            try:
                sw_app.ActivateDoc2(candidate, False, 0)
                activated = True
                continue
            except TypeError:
                pass
            except Exception:
                pass
            # Forma 2: via __getattr__ explícito para evitar el proxy de win32com
            try:
                method = sw_app.__getattr__("ActivateDoc2")
                if callable(method):
                    method(candidate, False, 0)
                    activated = True
            except Exception:
                pass

        # Obtener modelo activo (puede ser diferente al abierto)
        try:
            active = sw_app.ActiveDoc
            if active is not None:
                model = active
        except Exception:
            pass

        # Rebuild — necesario para que el feature tree se cargue completamente
        try:
            model.ForceRebuild3(False)
        except Exception:
            try:
                model.EditRebuild3()
            except Exception:
                pass
    except Exception:
        pass
    return model


def _close_doc(sw_app, model) -> None:
    if model is None:
        return
    try:
        t = model.GetTitle()
        if t:
            sw_app.CloseDoc(t)
    except Exception:
        pass


# ─────────────────────────────────────────────────────────────────────────────
# Helpers
# ─────────────────────────────────────────────────────────────────────────────

def _is_date(text: str) -> bool:
    if not text:
        return False
    patterns = [
        r"\d{1,2}/\d{1,2}/\d{2,4}",
        r"\d{4}-\d{2}-\d{2}",
        r"\d{1,2}:\d{2}:\d{2}",
        r"\b(a\.\s*m\.|p\.\s*m\.|am|pm)\b",
        r"\b(lunes|martes|mi[eé]rcoles|jueves|viernes|s[aá]bado|domingo|"
        r"january|february|march|april|may|june|july|august|"
        r"september|october|november|december|"
        r"enero|febrero|marzo|abril|mayo|junio|julio|agosto|"
        r"septiembre|octubre|noviembre|diciembre)\b",
    ]
    for p in patterns:
        if re.search(p, text, re.IGNORECASE):
            return True
    return False


def _clean_name(value: Any) -> Optional[str]:
    if value is None:
        return None
    s = str(value).strip()
    if not s or s.lower() in ("desconocido", "unknown", "", "nan"):
        return None
    if _is_date(s):
        return None
    return s


def _strip_domain(u: str) -> str:
    return u.split("\\", 1)[1] if "\\" in u else u


def _safe_str(value: Any) -> str:
    """Convierte a string limpio, devuelve '' si es NaN/None/vacío."""
    if value is None:
        return ""
    s = str(value).strip()
    if s.lower() in ("nan", "none", "desconocido", ""):
        return ""
    return s


# ─────────────────────────────────────────────────────────────────────────────
# Shell metadata
# ─────────────────────────────────────────────────────────────────────────────

def _read_shell_metadata(path: str) -> Dict[str, str]:
    meta: Dict[str, str] = {}
    try:
        shell  = win32com.client.Dispatch("Shell.Application")
        folder = shell.Namespace(os.path.dirname(os.path.abspath(path)))
        if folder is None:
            return meta
        item = folder.ParseName(os.path.basename(path))
        if item is None:
            return meta

        col_map: Dict[str, int] = {}
        for i in range(300):
            try:
                h = folder.GetDetailsOf(None, i)
            except Exception:
                break
            if h:
                col_map[normalize_text(h)] = i

        def get_col(*kws) -> Optional[str]:
            for kw in kws:
                for h_n, idx in col_map.items():
                    if normalize_text(kw) in h_n:
                        try:
                            v = folder.GetDetailsOf(item, idx)
                            if v and v.strip():
                                return v.strip()
                        except Exception:
                            pass
            return None

        owner = get_col("propietario", "owner", "dueño", "dueno")
        if owner:
            meta["owner"]       = owner
            meta["owner_short"] = _strip_domain(owner)

        computer = get_col("equipo", "computer", "maquina")
        if computer:
            c = re.sub(r"\s*\(.*?\)\s*", "", computer).strip()
            if c:
                meta["computer_name"] = c
    except Exception:
        pass
    return meta


# ─────────────────────────────────────────────────────────────────────────────
# SummaryInfo — índices verificados
# ─────────────────────────────────────────────────────────────────────────────

def _read_summary_info(model) -> Dict[str, str]:
    """
    Lee SummaryInfo. Índices verificados en archivos reales:
      [5] = author username (puede cambiar si SW lo sobreescribe con usuario actual)
      [6] = fecha de CREACIÓN — FIJA, no cambia al guardar de nuevo
      [7] = fecha del ÚLTIMO GUARDADO — cambia cada vez que se guarda
    """
    meta: Dict[str, str] = {}
    all_fields = {**SW_SUMMARY_INFO_FIELDS, **SW_SUMMARY_INFO_EXTRA}
    for idx in range(16):
        try:
            val = model.SummaryInfo(idx)
            if val is None:
                continue
            v = str(val).strip()
            if not v:
                continue
            key = all_fields.get(idx, f"sw_raw_{idx}")
            meta[key] = v
        except Exception:
            continue
    return meta


# ─────────────────────────────────────────────────────────────────────────────
# Autor desde binario (fallback)
# ─────────────────────────────────────────────────────────────────────────────

def _read_author_from_binary(path: str, exclude: str = "") -> Optional[str]:
    try:
        with open(path, "rb") as f:
            data = f.read(80000)

        pattern = rb"(?:[\x20-\x7e]\x00){3,25}"
        matches = re.findall(pattern, data)

        skip = {
            "solidworks", "version", "part", "assembly", "drawing",
            "default", "pieza", "ensamble", "plano", "true", "false",
            "none", "null", "yes", "no", "si", "ok", "cancel",
            "alumno", "student", "usuario", "user", "admin",
        }
        exclude_n = normalize_text(exclude)

        candidates = []
        seen: set = set()
        for m in matches:
            try:
                s = m.decode("utf-16-le").strip()
            except Exception:
                continue
            if not s or s in seen:
                continue
            seen.add(s)
            if any(c in s for c in r'/\:*?"<>|.@#$%^&()[]{}0123456789'):
                continue
            if _is_date(s):
                continue
            if not re.match(r'^[a-zA-Z_áéíóúüñÁÉÍÓÚÜÑ][\w\s\-]*$', s):
                continue
            if s.lower() in skip:
                continue
            if exclude_n and normalize_text(s) == exclude_n:
                continue
            if 2 <= len(s) <= 20:
                candidates.append(s)

        if not candidates:
            return None
        from collections import Counter as _C
        return _C(candidates).most_common(1)[0][0]
    except Exception:
        return None


# ─────────────────────────────────────────────────────────────────────────────
# Custom Properties
# ─────────────────────────────────────────────────────────────────────────────

def _get_ext(model) -> Any:
    for fn in (lambda: model.Extension, lambda: model.IExtension):
        try:
            e = fn()
            if e is not None:
                return e
        except Exception:
            pass
    return None


def _props_from_mgr(mgr) -> Dict[str, str]:
    if mgr is None:
        return {}
    props: Dict[str, str] = {}
    names: List[str] = []
    try:
        raw = mgr.GetNames()
        if raw is not None:
            names = list(raw) if isinstance(raw, (list, tuple)) else ([str(raw)] if raw else [])
            names = [str(n) for n in names if n]
    except Exception:
        pass
    name_lower = {n.lower() for n in names}
    for p in PROBE_PROPERTY_NAMES:
        if p.lower() not in name_lower:
            names.append(p)
    for name in names:
        value = ""
        for getter, args in (
            ("Get5", (name, True, "", "")), ("Get4", (name, True, "", "")),
            ("Get3", (name, True)),         ("Get2", (name,)),
            ("Get",  (name,)),
        ):
            try:
                r = getattr(mgr, getter)(*args)
                c = coerce_tuple_first(r, "")
                if c:
                    value = str(c).strip()
                    break
            except Exception:
                pass
        if value:
            props[normalize_text(name)] = value
    return props


def _read_custom_properties(model) -> Dict[str, str]:
    props: Dict[str, str] = {}
    ext = _get_ext(model)
    if ext is None:
        return props
    try:
        mgr = ext.CustomPropertyManager("")
        if mgr:
            props.update(_props_from_mgr(mgr))
    except Exception:
        pass
    try:
        cr = model.GetConfigurationNames()
        cfgs: List[str] = []
        if isinstance(cr, tuple):
            for item in cr:
                if isinstance(item, (list, tuple)):
                    cfgs = [str(x) for x in item if x]
                    break
        elif isinstance(cr, (list, tuple)):
            cfgs = [str(x) for x in cr if x]
        for cfg in cfgs:
            try:
                mgr = ext.CustomPropertyManager(str(cfg))
                if mgr:
                    for k, v in _props_from_mgr(mgr).items():
                        props[f"cfg:{cfg}:{k}"] = v
            except Exception:
                pass
    except Exception:
        pass
    return props


# ─────────────────────────────────────────────────────────────────────────────
# Feature Tree
# ─────────────────────────────────────────────────────────────────────────────

def _feat_type(feat) -> str:
    """
    GetTypeName2 en win32com late-binding devuelve el string DIRECTAMENTE
    (es una propiedad, no un método). El diagnóstico lo confirmó:
    feat.GetTypeName2 = 'ProfileFeature' (string), NO llamable con ().
    """
    for attr in ("GetTypeName2", "GetTypeName"):
        try:
            val = getattr(feat, attr)
            if isinstance(val, str):
                return val.strip()      # propiedad → ya es el tipo
            if callable(val):
                r = val()
                if isinstance(r, str):
                    return r.strip()    # método → llamar
        except Exception:
            continue
    return ""


def _feat_name(feat) -> str:
    try:
        val = feat.Name
        if isinstance(val, str):
            return val.strip()
        if callable(val):
            return str(val()).strip()
    except Exception:
        pass
    return ""


def _feat_next(feat) -> Any:
    """
    GetNextFeature también es propiedad en late-binding.
    Devuelve el objeto COM del siguiente feature, o None al final.
    """
    for attr in ("GetNextFeature", "IGetNextFeature"):
        try:
            val = getattr(feat, attr)
            if val is None:
                return None
            if isinstance(val, str):
                return None     # string = no hay siguiente
            if callable(val):
                return val()    # método → llamar
            return val          # objeto COM → es el siguiente
        except Exception:
            continue
    return None


def _feat_sub(feat) -> Any:
    for attr in ("GetFirstSubFeature", "IGetFirstSubFeature"):
        try:
            val = getattr(feat, attr)
            if val is None:
                return None
            if isinstance(val, str):
                return None
            if callable(val):
                return val()
            return val
        except Exception:
            continue
    return None


def _walk(first, depth: int = 0) -> List[Dict]:
    rows: List[Dict] = []
    feat = first
    while feat is not None:
        try:
            ftype = _feat_type(feat)
            fname = _feat_name(feat)
            if ftype and ftype not in IGNORED_FEATURE_TYPES:
                rows.append({
                    "depth": depth, "type": ftype, "name": fname,
                    "token_type": normalize_text(ftype),
                    "token_name": normalize_text(fname),
                })
            sub = _feat_sub(feat)
            if sub is not None:
                rows.extend(_walk(sub, depth + 1))
        except Exception:
            pass
        feat = _feat_next(feat)
    return rows


def _read_features(model) -> List[Dict]:
    """
    Lee el feature tree. Basado en diagnóstico real:
    - GetTypeName2/GetNextFeature son PROPIEDADES (string/objeto), no métodos
    - FeatureManager.GetFeatures() devuelve tuple donde TODOS los elementos
      son features (no hay que hacer unwrap del primero — iterar todos)
    """
    # Estrategia 1: FirstFeature como propiedad (más fiable según diagnóstico)
    try:
        first = model.FirstFeature
        if first is not None and not isinstance(first, str):
            rows = _walk(first)
            if rows:
                return rows
    except Exception:
        pass

    # Estrategia 2: FeatureManager.GetFeatures — el diagnóstico mostró
    # que devuelve un TUPLE de 23 CDispatch objects (todos son features)
    for include_suppressed in (True, False):
        try:
            raw = model.FeatureManager.GetFeatures(include_suppressed)
            if raw is None:
                continue

            # El diagnóstico mostró: tuple de N CDispatch = TODOS son features
            # NO usar _unwrap() — iterar el tuple completo
            if isinstance(raw, tuple):
                feats = list(raw)
            elif isinstance(raw, list):
                feats = raw
            else:
                feats = [raw]

            rows = []
            for feat in feats:
                if feat is None:
                    continue
                try:
                    ft  = _feat_type(feat)
                    fn2 = _feat_name(feat)
                    if ft and ft not in IGNORED_FEATURE_TYPES:
                        rows.append({
                            "depth": 0, "type": ft, "name": fn2,
                            "token_type": normalize_text(ft),
                            "token_name": normalize_text(fn2),
                        })
                except Exception:
                    pass
            if rows:
                return rows
        except Exception:
            pass

    return []


def _sig(rows: List[Dict]) -> str:
    t = " > ".join(r["token_type"] for r in rows if r.get("token_type"))
    n = " > ".join(r["token_name"] for r in rows if r.get("token_name"))
    return f"{t} | {n}".strip(" |")



# ─────────────────────────────────────────────────────────────────────────────
# Extractor principal
# ─────────────────────────────────────────────────────────────────────────────

def _mlabel(a: str, l: str) -> str:
    an, ln = normalize_text(a), normalize_text(l)
    if not an and not ln:
        return "sin_metadata"
    if an and ln and an != ln:
        return "inconsistente"
    return "aparentemente_consistente"


def extract_solidworks_document(path: str, session: SolidWorksSession) -> Dict[str, Any]:
    _base: Dict[str, Any] = {
        "Archivo": os.path.basename(path), "Ruta_Completa": path,
        "Modo": "solidworks_api", "Open_Method": "",
        "Autor_Original": "Desconocido", "Ultimo_Guardado": "Desconocido",
        "Propietario_Windows": "", "Nombre_Maquina": "",
        "SW_Created_Date": "", "SW_Saved_Date": "",
        "Fecha_Creacion_SW": "Desconocido", "Fecha_Ultimo_Guardado_SW": "Desconocido",
        "SW_Author_Raw": "",
        # Hash y tamaño siempre disponibles (no requieren SW)
        "Hash_Corto":   fast_file_hash(path) if os.path.isfile(path) else "",
        "Tamano_Bytes": os.path.getsize(path) if os.path.isfile(path) else 0,
        "Fecha_Modificacion": format_datetime(os.path.getmtime(path)) if os.path.isfile(path) else "Desconocido",
        "Extension": os.path.splitext(path)[1].lower(),
        "Feature_Count": 0, "Feature_Types": "", "Feature_Names": "",
        "Feature_Signature": "", "Custom_Props": {}, "Summary_Info": {},
        "Metadata_Status": "sin_metadata", "Confidence": 0, "Error": "",
    }

    if not os.path.isfile(path):
        return {**_base, "Error": "No existe el archivo"}
    if not path.lower().endswith(CAD_EXTENSIONS):
        return {**_base, "Error": "Extensión no compatible"}

    sw_app = session.connect()
    shell_meta  = _read_shell_metadata(path)
    owner_full  = shell_meta.get("owner", "")
    owner_short = shell_meta.get("owner_short", "")
    computer    = shell_meta.get("computer_name", "")

    file_info = {
        "Extension":          os.path.splitext(path)[1].lower(),
        "Tamano_Bytes":       os.path.getsize(path),
        "Hash_Corto":         fast_file_hash(path),
        "Fecha_Modificacion": format_datetime(os.path.getmtime(path)),
    }

    model, open_method = _open_document(sw_app, path)

    if model is None:
        return {**_base, **file_info,
                "Open_Method": open_method,
                "Error": "No se pudo abrir con SolidWorks",
                "Propietario_Windows": owner_full,
                "Nombre_Maquina": owner_full or computer,
                "Custom_Props": shell_meta, "Confidence": 5}

    try:
        # PASO 1: Leer SummaryInfo ANTES del rebuild
        # SW puede sobreescribir SummaryInfo[5] (autor) con el usuario actual
        # durante ForceRebuild. Leerlo antes da el valor original.
        summary_info = _read_summary_info(model)

        # PASO 2: Activar y reconstruir (necesario para cargar el feature tree)
        model = _activate_and_rebuild(sw_app, path, model)

        # PASO 3: Leer el resto después del rebuild
        custom_props = _read_custom_properties(model)
        try:
            t = model.GetTitle()
            if t and str(t).strip():
                custom_props["runtime:title"] = str(t).strip()
        except Exception:
            pass

        # ── Autor Original ─────────────────────────────────────────────────
        # SummaryInfo[5] leído ANTES del rebuild — es el usuario que guardó el archivo en SW.
        # Para piezas propias: SummaryInfo[5] = tu username = Propietario Windows → coinciden.
        # Para piezas recibidas: SummaryInfo[5] = autor original ≠ Propietario Windows.
        # NO descartar SummaryInfo[5] solo porque coincide con owner_short:
        # si alguien creó su propia pieza, SU username es el autor correcto.
        sw_author = (
            _clean_name(summary_info.get("author")) or       # SummaryInfo idx 5
            _clean_name(summary_info.get("author_alt")) or   # SummaryInfo idx 3
            _clean_name(custom_props.get("sw-author")) or
            _clean_name(custom_props.get("author")) or
            _clean_name(custom_props.get("creator"))
        )

        # Usar binario SOLO si la API no devolvió nada (no para reemplazar)
        if not sw_author:
            sw_author = _read_author_from_binary(path, exclude="")

        author_original = sw_author or "Desconocido"

        # ── Último guardado (solo custom props, informativo) ───────────────
        ultimo_guardado = (
            _clean_name(custom_props.get("sw-last saved by")) or
            _clean_name(custom_props.get("last saved by")) or
            _clean_name(custom_props.get("lastsavedby")) or
            "Desconocido"
        )

        # ── Fechas SW ─────────────────────────────────────────────────────
        fc_sw = _safe_str(summary_info.get("created_date_sw"))  # FIJA
        fs_sw = _safe_str(summary_info.get("saved_date_sw"))    # cambia al guardar

        # ── Feature Tree ──────────────────────────────────────────────────
        feat_rows  = _read_features(model)
        feat_types = " > ".join(r["token_type"] for r in feat_rows if r.get("token_type"))
        feat_names = " > ".join(r["token_name"] for r in feat_rows if r.get("token_name"))
        feat_sig   = _sig(feat_rows)

        # ── Confianza ─────────────────────────────────────────────────────
        conf = 0
        if feat_rows:                         conf += 40
        if feat_sig:                          conf += 10
        if author_original != "Desconocido":  conf += 25
        if fc_sw:                             conf += 15
        if fs_sw:                             conf += 10

        return {
            **file_info,
            "Archivo":        os.path.basename(path),
            "Ruta_Completa":  path,
            "Modo":           "solidworks_api",
            "Open_Method":    open_method,
            "Autor_Original": author_original,
            "Ultimo_Guardado": ultimo_guardado,
            "Propietario_Windows":      owner_full,
            "Nombre_Maquina":           owner_full or computer,
            # Campos planos (nunca NaN, no dict anidado)
            "SW_Created_Date":          fc_sw,   # FIJA — clave para detectar copias
            "SW_Saved_Date":            fs_sw,   # cambia al guardar
            "SW_Author_Raw":            author_original if author_original != "Desconocido" else "",
            # Aliases para el UI
            "Fecha_Creacion_SW":        fc_sw if fc_sw else "Desconocido",
            "Fecha_Ultimo_Guardado_SW": fs_sw if fs_sw else "Desconocido",
            "Feature_Count":     len(feat_rows),
            "Feature_Types":     feat_types,
            "Feature_Names":     feat_names,
            "Feature_Signature": feat_sig,
            "Custom_Props":      custom_props,
            "Summary_Info":      summary_info,
            "Metadata_Status":   _mlabel(author_original, ultimo_guardado),
            "Confidence":        min(conf, 100),
            "Error":             "",
        }

    except Exception as exc:
        return {**_base, **file_info,
                "Open_Method": open_method,
                "Error": str(exc),
                "Propietario_Windows": owner_full,
                "Nombre_Maquina": owner_full or computer,
                "Custom_Props": shell_meta, "Confidence": 5}
    finally:
        _close_doc(sw_app, model)
