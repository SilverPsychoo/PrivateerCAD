from __future__ import annotations

"""
analizador.py — Motor de detección de plagio para SolidWorks.

LÓGICA DE DETECCIÓN:
──────────────────────────────────────────────────────────────
El maestro NO necesita hacer nada manual. La app da el veredicto.

Cuando alguien "abre el archivo y lo guarda":
  - SW_Saved_Date  CAMBIA  → diferente, no sirve para igualdad
  - SW_Created_Date FIJA   → igual en original y copia → MISMO ORIGEN
  - Feature Tree   IGUAL   → misma secuencia de operaciones

Cuando alguien copia el archivo sin abrirlo:
  - SW_Saved_Date  FIJA    → idéntica → COPIA EXACTA
  - Hash           IGUAL   → idéntico

Combinando ambos casos se detecta plagio en cualquier variante.
"""

from collections import Counter, defaultdict
from difflib import SequenceMatcher
from itertools import combinations
from typing import Any, Dict, List, Optional, Tuple

import pandas as pd
import networkx as nx
import matplotlib
matplotlib.use("TkAgg")
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches

from config import (
    GRAPH_EDGE_THRESHOLD, HIGH_RISK_THRESHOLD, SUSPECT_THRESHOLD,
    SW_DATE_COLLISION_WINDOW_SEC,
    SCORE_HASH_IDENTICO, SCORE_FEATURE_TREE_ALTO,
    SCORE_FEATURE_TREE_MEDIO, SCORE_FEATURE_TREE_BAJO,
    SCORE_MISMO_FEATURE_COUNT, SCORE_MISMO_TAMANO,
    GENERIC_USERNAMES,
)
from utils import normalize_text, parse_datetime_any


# ─────────────────────────────────────────────────────────────────────────────
# Helpers
# ─────────────────────────────────────────────────────────────────────────────

def _is_generic(username: str) -> bool:
    return normalize_text(str(username or "")) in GENERIC_USERNAMES


def _valid_author(username: str) -> bool:
    n = normalize_text(str(username or ""))
    return bool(n) and n not in ("desconocido", "unknown", "") and not _is_generic(n)


def _clean_sw_date(value: Any) -> str:
    """
    Limpia un campo de fecha SW. Devuelve '' si es NaN, 'nan', None o vacío.
    Esto es crítico: pandas convierte columnas con NaN mezclado a 'nan' string.
    """
    if value is None:
        return ""
    s = str(value).strip()
    if s.lower() in ("nan", "none", "desconocido", "unknown", ""):
        return ""
    return s


def _sw_created_delta(a: Dict, b: Dict) -> Optional[float]:
    """
    Diferencia entre fechas de CREACIÓN SW (idx 6).
    Fija aunque el archivo se guarde de nuevo.
    Si dos archivos tienen la misma fecha de creación → mismo origen.
    """
    da = _clean_sw_date(a.get("SW_Created_Date") or a.get("Fecha_Creacion_SW"))
    db = _clean_sw_date(b.get("SW_Created_Date") or b.get("Fecha_Creacion_SW"))
    if not da or not db:
        return None
    dta = parse_datetime_any(da)
    dtb = parse_datetime_any(db)
    if dta and dtb:
        return abs((dta - dtb).total_seconds())
    return None


def _sw_saved_delta(a: Dict, b: Dict) -> Optional[float]:
    """
    Diferencia entre fechas de GUARDADO SW (idx 7).
    Igual cuando se copia sin abrir. Cambia cuando alguien abre y guarda.
    """
    da = _clean_sw_date(a.get("SW_Saved_Date") or a.get("Fecha_Ultimo_Guardado_SW"))
    db = _clean_sw_date(b.get("SW_Saved_Date") or b.get("Fecha_Ultimo_Guardado_SW"))
    if not da or not db:
        return None
    dta = parse_datetime_any(da)
    dtb = parse_datetime_any(db)
    if dta and dtb:
        return abs((dta - dtb).total_seconds())
    return None


def _feature_similarity(a: Dict, b: Dict) -> float:
    ca = int(a.get("Feature_Count") or 0)
    cb = int(b.get("Feature_Count") or 0)
    if ca == 0 or cb == 0:
        return 0.0

    type_a = str(a.get("Feature_Types") or "")
    type_b = str(b.get("Feature_Types") or "")
    name_a = str(a.get("Feature_Names") or "")
    name_b = str(b.get("Feature_Names") or "")

    sim = (0.50 * SequenceMatcher(None, f"{type_a}||{name_a}", f"{type_b}||{name_b}").ratio() +
           0.30 * SequenceMatcher(None, type_a, type_b).ratio() +
           0.20 * SequenceMatcher(None, name_a, name_b).ratio())

    if min(ca, cb) < 4:
        sim *= 0.75
    return round(sim, 4)


def _choose_source(a: Dict, b: Dict) -> Tuple[Dict, Dict]:
    """
    Determina cuál archivo es el original y cuál la copia.
    Si la fecha de creación es igual (mismo origen), el guardado más antiguo es el original.
    """
    # Fecha de creación SW — si difieren, el más antiguo es el original
    for field in ("SW_Created_Date", "Fecha_Creacion_SW"):
        da = parse_datetime_any(_clean_sw_date(a.get(field)) or a.get(field))
        db = parse_datetime_any(_clean_sw_date(b.get(field)) or b.get(field))
        if da and db and abs((da - db).total_seconds()) > 5:
            return (a, b) if da < db else (b, a)

    # Fecha de guardado SW — si creación es igual, el guardado más antiguo es el original
    for field in ("SW_Saved_Date", "Fecha_Ultimo_Guardado_SW"):
        da = parse_datetime_any(_clean_sw_date(a.get(field)) or a.get(field))
        db = parse_datetime_any(_clean_sw_date(b.get(field)) or b.get(field))
        if da and db and abs((da - db).total_seconds()) > 5:
            return (a, b) if da < db else (b, a)

    # Fecha Windows como último recurso
    da = parse_datetime_any(a.get("Fecha_Modificacion"))
    db = parse_datetime_any(b.get("Fecha_Modificacion"))
    if da and db:
        return (a, b) if da <= db else (b, a)
    return a, b


# ─────────────────────────────────────────────────────────────────────────────
# Score de plagio entre un par
# ─────────────────────────────────────────────────────────────────────────────

def _pair_score(a: Dict, b: Dict) -> Dict[str, Any]:
    score = 0
    reasons: List[str] = []

    # ── 1. Fecha de CREACIÓN SW idéntica ──────────────────────────────────
    # SummaryInfo[6] queda FIJA aunque el archivo se guarde de nuevo.
    # Si dos archivos distintos tienen la misma fecha de creación → mismo origen.
    created_delta = _sw_created_delta(a, b)
    if created_delta is not None:
        if created_delta <= SW_DATE_COLLISION_WINDOW_SEC:
            score += 50
            fc = _clean_sw_date(a.get("SW_Created_Date") or a.get("Fecha_Creacion_SW"))
            reasons.append(f"misma fecha de creación SW ({fc}) → mismo origen")
        elif created_delta <= 120:
            score += 20
            reasons.append(f"fecha de creación SW casi idéntica (Δ={created_delta:.0f}s)")

    # ── 2. Fecha de GUARDADO SW idéntica ──────────────────────────────────
    # Si además la fecha de guardado coincide → no solo mismo origen sino
    # copia directa (nunca se volvió a guardar).
    saved_delta = _sw_saved_delta(a, b)
    if saved_delta is not None and saved_delta <= SW_DATE_COLLISION_WINDOW_SEC:
        score += 15  # refuerzo adicional
        reasons.append("fecha de guardado SW también idéntica (copia directa)")

    # ── 3. Hash SHA idéntico ──────────────────────────────────────────────
    ha = str(a.get("Hash_Corto") or "").strip()
    hb = str(b.get("Hash_Corto") or "").strip()
    if (ha and hb
            and ha not in ("nan", "none", "")
            and hb not in ("nan", "none", "")
            and not ha.startswith("ERROR")
            and ha == hb):
        score += SCORE_HASH_IDENTICO
        reasons.append("hash idéntico (copia exacta byte a byte)")

    # ── 4. Similitud del Feature Tree ─────────────────────────────────────
    feat_sim = _feature_similarity(a, b)
    if feat_sim >= 0.98:
        score += SCORE_FEATURE_TREE_ALTO
        reasons.append(f"árbol de operaciones casi idéntico ({feat_sim:.0%})")
    elif feat_sim >= 0.90:
        score += SCORE_FEATURE_TREE_MEDIO
        reasons.append(f"árbol de operaciones muy similar ({feat_sim:.0%})")
    elif feat_sim >= 0.82:
        score += SCORE_FEATURE_TREE_BAJO
        reasons.append(f"árbol de operaciones similar ({feat_sim:.0%})")

    # ── 5. Mismo número de operaciones ────────────────────────────────────
    ca = int(a.get("Feature_Count") or 0)
    cb = int(b.get("Feature_Count") or 0)
    if ca > 0 and ca == cb:
        score += SCORE_MISMO_FEATURE_COUNT
        reasons.append(f"mismo número de operaciones ({ca})")

    # ── 6. Mismo tamaño ───────────────────────────────────────────────────
    sa = int(a.get("Tamano_Bytes") or 0)
    sb = int(b.get("Tamano_Bytes") or 0)
    if sa > 0 and sa == sb:
        score += SCORE_MISMO_TAMANO
        reasons.append("mismo tamaño de archivo")

    # ── 7. Mismo autor real no genérico ───────────────────────────────────
    aut_a = _clean_sw_date(a.get("SW_Author_Raw") or a.get("Autor_Original"))
    aut_b = _clean_sw_date(b.get("SW_Author_Raw") or b.get("Autor_Original"))
    if _valid_author(aut_a) and _valid_author(aut_b) and normalize_text(aut_a) == normalize_text(aut_b):
        score += 8
        reasons.append(f"mismo autor SW real '{aut_a}'")

    score = min(score, 100)
    source, target = _choose_source(a, b)

    return {
        "score":              score,
        "feature_similarity": feat_sim,
        "created_delta":      created_delta,
        "saved_delta":        saved_delta,
        "reasons":            reasons,
        "source_file":  source.get("Archivo", ""),
        "target_file":  target.get("Archivo", ""),
        "source_path":  source.get("Ruta_Completa", source.get("Archivo", "")),
        "target_path":  target.get("Ruta_Completa", target.get("Archivo", "")),
        "source_data":  source,
        "target_data":  target,
    }


# ─────────────────────────────────────────────────────────────────────────────
# Diagnóstico individual
# ─────────────────────────────────────────────────────────────────────────────

def diagnostico_unico(data: Dict[str, Any]) -> Tuple[str, str]:
    if not data:
        return "No se pudo extraer información.", "SIN_DATOS"

    lines = []
    autor       = data.get("Autor_Original", "Desconocido")
    propietario = data.get("Propietario_Windows", "")
    feats       = int(data.get("Feature_Count") or 0)
    conf        = int(data.get("Confidence") or 0)
    fc_sw       = _clean_sw_date(data.get("SW_Created_Date") or data.get("Fecha_Creacion_SW"))
    fs_sw       = _clean_sw_date(data.get("SW_Saved_Date") or data.get("Fecha_Ultimo_Guardado_SW"))

    autor_real    = _valid_author(autor)
    autor_generic = _is_generic(autor)

    lines.append(f"📄 Archivo              : {data.get('Archivo', '')}")
    lines.append(f"📁 Ruta                 : {data.get('Ruta_Completa', '')}")
    lines.append(f"⚙️  Modo                 : {data.get('Modo', '')} | {data.get('Open_Method', '')}")
    lines.append("")

    if autor_real:
        lines.append(f"✍️  Autor SW (del archivo): {autor}")
    elif autor_generic:
        lines.append(f"✍️  Autor SW              : {autor}  ← PC genérica (lab/uni)")
    else:
        lines.append(f"✍️  Autor SW              : No disponible")

    if propietario:
        prop_short = propietario.split("\\")[-1] if "\\" in propietario else propietario
        lines.append(f"🖥️  Propietario Windows   : {propietario}")
        lines.append(f"   (Quién tiene el archivo AHORA — cambia al copiarlo para revisar)")
    else:
        lines.append(f"🖥️  Propietario Windows   : No disponible")

    lines.append(f"📅 Fecha modificación    : {data.get('Fecha_Modificacion', '—')}")
    lines.append("")
    lines.append("── Fechas internas SW ─────────────────────────────────────────")
    if fc_sw:
        lines.append(f"🗓️  Creado                : {fc_sw}  ← FIJA aunque guardes de nuevo")
    else:
        lines.append(f"🗓️  Creado                : No disponible")
    if fs_sw:
        lines.append(f"🗓️  Último guardado SW    : {fs_sw}  ← cambia al volver a guardar")
    else:
        lines.append(f"🗓️  Último guardado SW    : No disponible")

    lines.append(f"🔧 Operaciones           : {feats}")
    lines.append(f"🔍 Confianza lectura     : {conf}/100")

    if data.get("Error"):
        lines.append(f"\n❌ Error: {data['Error']}")
        return "\n".join(lines), "ERROR"

    lines.append("")
    lines.append("── VEREDICTO INDIVIDUAL ──────────────────────────────────────")
    lines.append("ℹ️  Un solo archivo no puede probar plagio.")
    lines.append("   Usa 'Analizar carpeta completa' con todos los trabajos del grupo.")

    if autor_real and fc_sw and feats > 0:
        estado = "LIMPIO"
        lines.append(f"\n🟢 Datos completos: autor='{autor}', {feats} operaciones,")
        lines.append(f"   fecha creación='{fc_sw}'")
        lines.append(f"   El lote comparará estos datos contra todos los archivos del grupo.")
    elif autor_real or fc_sw or feats > 0:
        estado = "EVIDENCIA_PARCIAL"
        lines.append(f"\n🟠 Datos parciales — suficientes para comparar en lote.")
    else:
        estado = "BAJA_CONFIANZA"
        lines.append(f"\n⚪ Datos insuficientes. El lote intentará comparar con el grupo.")

    return "\n".join(lines), estado


# ─────────────────────────────────────────────────────────────────────────────
# Análisis por lote
# ─────────────────────────────────────────────────────────────────────────────

def analizar_lote(datos: List[Dict[str, Any]]) -> Tuple[Any, str, List[Dict]]:
    if not datos:
        return None, "No se encontraron archivos CAD válidos.", []

    # Normalizar campos planos antes de crear el DataFrame
    for d in datos:
        si = d.get("Summary_Info")
        if isinstance(si, dict):
            if not _clean_sw_date(d.get("SW_Created_Date")):
                d["SW_Created_Date"] = si.get("created_date_sw", "")
            if not _clean_sw_date(d.get("SW_Saved_Date")):
                d["SW_Saved_Date"]   = si.get("saved_date_sw", "")
            if not d.get("SW_Author_Raw"):
                d["SW_Author_Raw"]   = si.get("author", "")
        # Fallbacks desde aliases
        if not _clean_sw_date(d.get("SW_Created_Date")):
            d["SW_Created_Date"] = _clean_sw_date(d.get("Fecha_Creacion_SW"))
        if not _clean_sw_date(d.get("SW_Saved_Date")):
            d["SW_Saved_Date"] = _clean_sw_date(d.get("Fecha_Ultimo_Guardado_SW"))

    df = pd.DataFrame(datos).copy()

    # Rellenar columnas faltantes con defaults seguros
    for col, default in {
        "Autor_Original": "Desconocido", "Ultimo_Guardado": "Desconocido",
        "Propietario_Windows": "", "Nombre_Maquina": "",
        "Feature_Signature": "", "Feature_Types": "", "Feature_Names": "",
        "Hash_Corto": "", "Fecha_Modificacion": "Desconocido",
        "Tamano_Bytes": 0, "Feature_Count": 0,
        "SW_Created_Date": "", "SW_Saved_Date": "", "SW_Author_Raw": "",
        "Fecha_Creacion_SW": "", "Fecha_Ultimo_Guardado_SW": "",
    }.items():
        if col not in df.columns:
            df[col] = default

    df["Feature_Count"] = pd.to_numeric(df["Feature_Count"], errors="coerce").fillna(0).astype(int)
    df["Tamano_Bytes"]  = pd.to_numeric(df["Tamano_Bytes"],  errors="coerce").fillna(0).astype(int)
    df["Confidence"]    = pd.to_numeric(df.get("Confidence", pd.Series()), errors="coerce").fillna(0).astype(int)

    # CRÍTICO: limpiar strings NaN que pandas genera al mezclar tipos
    str_cols = ("SW_Created_Date", "SW_Saved_Date", "SW_Author_Raw",
                "Hash_Corto", "Feature_Types", "Feature_Names", "Feature_Signature",
                "Autor_Original", "Fecha_Modificacion",
                "Fecha_Creacion_SW", "Fecha_Ultimo_Guardado_SW")
    for col in str_cols:
        if col in df.columns:
            df[col] = df[col].fillna("").astype(str).apply(
                lambda x: "" if x.lower() in ("nan", "none") else x
            )

    # Recuperar SW_Created_Date y SW_Saved_Date desde aliases si quedaron vacíos
    for i in range(len(df)):
        if not df.at[i, "SW_Created_Date"] and df.at[i, "Fecha_Creacion_SW"]:
            df.at[i, "SW_Created_Date"] = df.at[i, "Fecha_Creacion_SW"]
        if not df.at[i, "SW_Saved_Date"] and df.at[i, "Fecha_Ultimo_Guardado_SW"]:
            df.at[i, "SW_Saved_Date"] = df.at[i, "Fecha_Ultimo_Guardado_SW"]

    registros = df.to_dict("records")

    # ── Comparar TODOS los pares ──────────────────────────────────────────
    pares: List[Dict]      = []
    relaciones: List[Dict] = []

    for i, j in combinations(range(len(registros)), 2):
        pair = _pair_score(registros[i], registros[j])
        pares.append(pair)
        if pair["score"] >= GRAPH_EDGE_THRESHOLD:
            relaciones.append({
                "source":      pair["source_path"],
                "target":      pair["target_path"],
                "source_file": pair["source_file"],
                "target_file": pair["target_file"],
                "score":       pair["score"],
                "feature_similarity": pair["feature_similarity"],
                "reason_str":  "; ".join(pair["reasons"]),
                "reasons":     pair["reasons"],
            })

    # ── Mejor score por archivo ───────────────────────────────────────────
    path_to_idx = {(r.get("Ruta_Completa") or r.get("Archivo", "")): i
                   for i, r in enumerate(registros)}

    best_score: Dict[int, int]            = {i: 0    for i in range(len(registros))}
    best_match: Dict[int, Optional[Dict]] = {i: None for i in range(len(registros))}

    for pair in pares:
        for pk in ("source_path", "target_path"):
            idx = path_to_idx.get(pair[pk])
            if idx is not None and pair["score"] > best_score[idx]:
                best_score[idx] = pair["score"]
                best_match[idx] = pair

    # ── Estado por archivo ────────────────────────────────────────────────
    estados, puntajes, detalles, fuentes = [], [], [], []
    for i, row in enumerate(registros):
        match  = best_match[i]
        score  = int(match["score"]) if match else 0
        reason = "; ".join(match["reasons"]) if match else ""

        if score >= HIGH_RISK_THRESHOLD:
            estado = "ALTO RIESGO"
        elif score >= SUSPECT_THRESHOLD:
            estado = "SOSPECHOSO"
        elif int(row.get("Feature_Count") or 0) == 0 and not _clean_sw_date(row.get("SW_Created_Date")):
            estado = "BAJA CONFIANZA"
        else:
            estado = "SIN ANOMALÍAS"

        # "Posible origen" solo se muestra para la COPIA (target), no para el original (source)
        fuente = ""
        if match:
            rp = row.get("Ruta_Completa", row.get("Archivo", ""))
            es_fuente = match.get("source_path") == rp
            if not es_fuente:
                # Este archivo ES la copia → mostrar de dónde vino
                fuente = match["source_file"]
            # Si es la fuente original → no mostrar "posible origen"

        estados.append(estado)
        puntajes.append(score)
        detalles.append(reason)
        fuentes.append(fuente)

    df["Puntaje_Sospecha"] = puntajes
    df["Estado"]           = estados
    df["Detalle_Sospecha"] = detalles
    df["Posible_Fuente"]   = fuentes

    # ── Detecciones especiales ────────────────────────────────────────────
    col_created  = _detectar_colisiones_fecha_creacion(registros)
    col_saved    = _detectar_colisiones_fecha_guardado(registros)
    grupos_fc    = _agrupar_por_feature_count(registros)
    grupos_autor = _agrupar_por_autor(registros)
    paciente     = _detectar_paciente_cero(registros, relaciones)

    # ── Reporte ───────────────────────────────────────────────────────────
    total         = len(df)
    n_alto        = int((df["Puntaje_Sospecha"] >= HIGH_RISK_THRESHOLD).sum())
    n_sospechosos = int((df["Puntaje_Sospecha"] >= SUSPECT_THRESHOLD).sum())

    sep   = "─" * 62
    lines = []
    lines.append(sep)
    lines.append(f"  REPORTE DE ANÁLISIS  —  {total} archivos")
    lines.append(sep)
    lines.append(f"  🔴 ALTO RIESGO (plagio muy probable) : {n_alto}")
    lines.append(f"  🟠 SOSPECHOSOS                       : {n_sospechosos}")
    lines.append("")

    # Copias con misma fecha de CREACIÓN (caso "abrió y guardó")
    if col_created:
        lines.append("🚨 MISMO ORIGEN DETECTADO (fecha de creación SW idéntica):")
        lines.append("   La fecha de creación SW no cambia aunque el archivo se vuelva a guardar.")
        lines.append("   Si dos archivos la tienen igual → se crearon en la misma sesión.")
        lines.append("")
        for col in col_created:
            lines.append(f"   ↔  {col['a']}")
            lines.append(f"      {col['b']}")
            lines.append(f"      Fecha creación SW: {col['fecha']}  (Δ={col['delta']:.1f}s)")
        lines.append("")

    # Copias exactas (misma fecha guardado)
    if col_saved:
        lines.append("🚨 COPIAS EXACTAS (fecha de guardado SW idéntica):")
        lines.append("   Estos archivos son idénticos — se copiaron sin abrir.")
        lines.append("")
        for col in col_saved:
            lines.append(f"   ↔  {col['a']}  ==  {col['b']}")
            lines.append(f"      Fecha guardado SW: {col['fecha']}  (Δ={col['delta']:.1f}s)")
        lines.append("")

    # Mismo número de operaciones
    if grupos_fc:
        lines.append("⚠️  MISMO NÚMERO DE OPERACIONES (posible copia):")
        for fc, archivos in grupos_fc.items():
            lines.append(f"   {fc} operaciones: {', '.join(archivos)}")
        lines.append("")

    # Paciente cero
    if paciente and _valid_author(paciente["nombre"]):
        lines.append(f"🦠 DISTRIBUIDOR ORIGINAL (paciente cero): '{paciente['nombre']}'")
        lines.append(f"   Certeza: {paciente.get('certeza', '—')}")
        if paciente.get("fecha_sw"):
            lines.append(f"   Fecha SW más antigua: {paciente['fecha_sw']}")
        lines.append(f"   Su archivo es origen en {paciente['salidas']} relación(es).")
        lines.append("")

    # Grupos por autor
    if grupos_autor:
        lines.append("👤 ARCHIVOS DEL MISMO AUTOR SW:")
        for autor, archivos in grupos_autor.items():
            lines.append(f"   {autor}: {', '.join(archivos)}")
        lines.append("")

    # Detalle por archivo
    lines.append("DETALLE POR ARCHIVO:")
    lines.append(sep)
    for _, row in df.iterrows():
        icon  = {"ALTO RIESGO": "🔴", "SOSPECHOSO": "🟠",
                 "BAJA CONFIANZA": "⚪", "SIN ANOMALÍAS": "🟢"}.get(row["Estado"], "🔵")

        autor_d  = _clean_sw_date(row.get("SW_Author_Raw") or row.get("Autor_Original")) or "Desconocido"
        fc_d     = _clean_sw_date(row.get("SW_Created_Date") or row.get("Fecha_Creacion_SW"))
        fs_d     = _clean_sw_date(row.get("SW_Saved_Date") or row.get("Fecha_Ultimo_Guardado_SW"))
        feats_d  = int(row.get("Feature_Count") or 0)

        lines.append(f"{icon} {row['Archivo']}")
        lines.append(f"   Estado       : {row['Estado']}  |  Score: {row['Puntaje_Sospecha']}/100")
        lines.append(f"   Autor SW     : {autor_d}")
        if fc_d:
            lines.append(f"   Creado SW    : {fc_d}")
        if fs_d:
            lines.append(f"   Guardado SW  : {fs_d}")
        if feats_d > 0:
            lines.append(f"   Operaciones  : {feats_d}")
        if row["Detalle_Sospecha"]:
            lines.append(f"   Indicadores  : {row['Detalle_Sospecha']}")
        if row["Posible_Fuente"]:
            lines.append(f"   ← Posible origen: {row['Posible_Fuente']}")
        lines.append("")

    return df, "\n".join(lines), relaciones


# ─────────────────────────────────────────────────────────────────────────────
# Detectores de patrones
# ─────────────────────────────────────────────────────────────────────────────

def _detectar_colisiones_fecha_creacion(registros: List[Dict]) -> List[Dict]:
    """Pares con la misma fecha de CREACIÓN SW (±5s). Detecta 'mismo origen'."""
    cols = []
    for i, j in combinations(range(len(registros)), 2):
        a, b = registros[i], registros[j]
        delta = _sw_created_delta(a, b)
        if delta is not None and delta <= SW_DATE_COLLISION_WINDOW_SEC:
            cols.append({
                "a":     a.get("Archivo", ""),
                "b":     b.get("Archivo", ""),
                "delta": delta,
                "fecha": _clean_sw_date(a.get("SW_Created_Date") or a.get("Fecha_Creacion_SW")),
            })
    return cols


def _detectar_colisiones_fecha_guardado(registros: List[Dict]) -> List[Dict]:
    """Pares con la misma fecha de GUARDADO SW (±5s). Detecta copias exactas."""
    cols = []
    for i, j in combinations(range(len(registros)), 2):
        a, b = registros[i], registros[j]
        delta = _sw_saved_delta(a, b)
        if delta is not None and delta <= SW_DATE_COLLISION_WINDOW_SEC:
            cols.append({
                "a":     a.get("Archivo", ""),
                "b":     b.get("Archivo", ""),
                "delta": delta,
                "fecha": _clean_sw_date(a.get("SW_Saved_Date") or a.get("Fecha_Ultimo_Guardado_SW")),
            })
    return cols


def _agrupar_por_feature_count(registros: List[Dict]) -> Dict[int, List[str]]:
    grupos: Dict[int, List[str]] = defaultdict(list)
    for r in registros:
        fc = int(r.get("Feature_Count") or 0)
        if fc >= 3:
            grupos[fc].append(r.get("Archivo", ""))
    return {k: v for k, v in grupos.items() if len(v) >= 2}


def _agrupar_por_autor(registros: List[Dict]) -> Dict[str, List[str]]:
    grupos: Dict[str, List[str]] = defaultdict(list)
    for r in registros:
        a = _clean_sw_date(r.get("SW_Author_Raw") or r.get("Autor_Original"))
        if _valid_author(a):
            grupos[a].append(r.get("Archivo", ""))
    return {k: v for k, v in grupos.items() if len(v) >= 2}


def _detectar_paciente_cero(registros: List[Dict],
                             relaciones: List[Dict]) -> Optional[Dict]:
    """
    Determina quién es el distribuidor original (paciente cero) con certeza.

    Estrategia:
    1. Si hay relaciones en el grafo: el nodo con más aristas SALIENTES es
       el que distribuyó. Entre empates, el que tiene fecha SW de CREACIÓN
       más antigua es el original.
    2. Si no hay grafo pero sí colisiones de fecha SW: el archivo con la
       fecha SW más antigua es el original (creó el archivo antes).
    3. Fallback: autor SW no genérico que aparece en más archivos.
    """
    if relaciones:
        out_deg = Counter(r["source_file"] for r in relaciones)
        in_deg  = Counter(r["target_file"]  for r in relaciones)

        if not out_deg:
            return None

        # Entre los que tienen más salidas, el más antiguo es el original
        max_out = out_deg.most_common(1)[0][1]
        top_sources = [f for f, c in out_deg.items() if c == max_out]

        # Resolver empate por fecha SW de creación más antigua
        best = None
        best_dt = None
        for fname in top_sources:
            rec = next((r for r in registros if r.get("Archivo") == fname), None)
            if rec:
                dt = parse_datetime_any(_clean_sw_date(
                    rec.get("SW_Created_Date") or rec.get("Fecha_Creacion_SW")
                ))
                if best_dt is None or (dt and dt < best_dt):
                    best_dt  = dt
                    best     = fname

        nombre = best or top_sources[0]
        return {
            "nombre":    nombre,
            "salidas":   out_deg[nombre],
            "entradas":  in_deg.get(nombre, 0),
            "certeza":   "ALTA — aparece como fuente en el grafo de relaciones",
            "fecha_sw":  best_dt.strftime("%Y-%m-%d %H:%M") if best_dt else "",
        }

    # Sin grafo: buscar el más antiguo entre archivos con la misma fecha de creación
    # Agrupar por SW_Created_Date
    grupos_creacion: Dict[str, List[Dict]] = defaultdict(list)
    for r in registros:
        fc = _clean_sw_date(r.get("SW_Created_Date") or r.get("Fecha_Creacion_SW"))
        if fc:
            # Redondear a minuto para agrupar variantes del mismo archivo
            dt = parse_datetime_any(fc)
            if dt:
                key = dt.strftime("%Y-%m-%d %H:%M")
                grupos_creacion[key].append(r)

    for key, grupo in grupos_creacion.items():
        if len(grupo) >= 2:
            # El más antiguo por fecha de modificación Windows
            mas_antiguo = min(
                grupo,
                key=lambda r: parse_datetime_any(r.get("Fecha_Modificacion")) or
                              __import__("datetime").datetime.max
            )
            return {
                "nombre":   mas_antiguo.get("Archivo", ""),
                "salidas":  len(grupo) - 1,
                "entradas": 0,
                "certeza":  "MEDIA — mismo origen SW, archivo Windows más antiguo",
                "fecha_sw": key,
            }

    # Fallback: autor SW real más repetido
    c: Counter = Counter()
    for r in registros:
        a = _clean_sw_date(r.get("SW_Author_Raw") or r.get("Autor_Original"))
        if _valid_author(a):
            c[a] += 1
    if c:
        nombre, count = c.most_common(1)[0]
        if count >= 2:
            return {"nombre": nombre, "salidas": count, "entradas": 0,
                    "certeza": "BAJA — mismo autor en múltiples archivos", "fecha_sw": ""}
    return None


# ─────────────────────────────────────────────────────────────────────────────
# Grafo de distribución
# ─────────────────────────────────────────────────────────────────────────────

def mostrar_grafo(df, relaciones: List[Dict]) -> None:
    if df is None or df.empty or not relaciones:
        return

    G = nx.DiGraph()
    color_map = {
        "ALTO RIESGO":    "#e74c3c",
        "SOSPECHOSO":     "#e67e22",
        "BAJA CONFIANZA": "#95a5a6",
        "SIN ANOMALÍAS":  "#27ae60",
    }

    for _, row in df.iterrows():
        node_id = row.get("Ruta_Completa", row.get("Archivo", ""))
        G.add_node(node_id,
                   label=row.get("Archivo", ""),
                   estado=row.get("Estado", "SIN ANOMALÍAS"),
                   autor=_clean_sw_date(row.get("SW_Author_Raw") or row.get("Autor_Original")),
                   score=int(row.get("Puntaje_Sospecha", 0)))

    for rel in relaciones:
        G.add_edge(rel["source"], rel["target"],
                   weight=rel["score"],
                   reasons=rel.get("reason_str", ""))

    fig, ax = plt.subplots(figsize=(15, 10))
    fig.patch.set_facecolor("#1a1a2e")
    ax.set_facecolor("#1a1a2e")

    try:
        pos = nx.kamada_kawai_layout(G)
    except Exception:
        pos = nx.spring_layout(G, seed=42, k=2.0)

    node_colors = [color_map.get(G.nodes[n].get("estado", "SIN ANOMALÍAS"), "#7f8c8d")
                   for n in G.nodes()]
    node_sizes  = [1800 + int(G.nodes[n].get("score", 0)) * 22 for n in G.nodes()]

    nx.draw_networkx_nodes(G, pos, ax=ax, node_color=node_colors,
                           node_size=node_sizes, alpha=0.92)

    labels = {}
    for n in G.nodes():
        nd    = G.nodes[n]
        label = nd.get("label", n)
        autor = nd.get("autor", "")
        labels[n] = f"{label}\n✍ {autor}" if autor and autor != "Desconocido" else label

    nx.draw_networkx_labels(G, pos, labels=labels, ax=ax,
                            font_size=8, font_color="white", font_weight="bold")

    edge_weights = [G[u][v]["weight"] for u, v in G.edges()]
    nx.draw_networkx_edges(G, pos, ax=ax, edge_color="#f39c12", arrows=True,
                           arrowstyle="->", arrowsize=18,
                           width=[max(0.5, w / 20) for w in edge_weights],
                           alpha=0.85, connectionstyle="arc3,rad=0.08")

    edge_labels = {(u, v): f"{G[u][v]['weight']}" for u, v in G.edges()}
    nx.draw_networkx_edge_labels(G, pos, edge_labels=edge_labels, ax=ax,
                                 font_size=8, font_color="#f1c40f")

    legend_patches = [mpatches.Patch(color=c, label=l) for l, c in color_map.items()]
    ax.legend(handles=legend_patches, loc="upper left",
              facecolor="#2c3e50", labelcolor="white", fontsize=9)

    ax.set_title("PRIVATEERCAD — Red de distribución",
                 color="white", fontsize=13, fontweight="bold", pad=14)
    ax.axis("off")
    plt.tight_layout()
    plt.show()
