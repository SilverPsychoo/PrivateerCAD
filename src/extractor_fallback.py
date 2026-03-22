from __future__ import annotations

import os
from typing import Any, Dict

import win32com.client

from config import MAX_SHELL_COLUMNS
from utils import fast_file_hash, format_datetime, normalize_text, safe_str, first_non_empty


def _read_shell_metadata(path: str) -> Dict[str, str]:
    meta: Dict[str, str] = {}
    try:
        shell = win32com.client.Dispatch("Shell.Application")
        folder = shell.Namespace(os.path.dirname(path))
        if folder is None:
            return meta
        item = folder.ParseName(os.path.basename(path))
        if item is None:
            return meta

        column_map: Dict[str, int] = {}
        for i in range(MAX_SHELL_COLUMNS):
            try:
                h = folder.GetDetailsOf(None, i)
            except Exception:
                break
            if h:
                column_map[normalize_text(h)] = i

        candidates = {
            "author":        ("author", "autor", "created by", "creado por"),
            "last_saved_by": ("last saved", "modificado por", "ultimo guardado", "último guardado"),
            "title":         ("title", "titulo", "título"),
            "computer_name": ("computer", "computadora", "machine", "maquina", "hostname"),
        }
        for key, words in candidates.items():
            idx = None
            for w in words:
                for header, col in column_map.items():
                    if normalize_text(w) in header:
                        idx = col
                        break
                if idx is not None:
                    break
            if idx is not None:
                try:
                    v = safe_str(folder.GetDetailsOf(item, idx), "")
                    if v:
                        meta[key] = v
                except Exception:
                    pass
    except Exception:
        pass
    return meta


def _read_ole_summary(path: str) -> Dict[str, str]:
    meta: Dict[str, str] = {}
    try:
        import olefile
        if not olefile.isOleFile(path):
            return meta
        ole = olefile.OleFileIO(path)
        if ole.exists("\x05SummaryInformation"):
            props = ole.getproperties("\x05SummaryInformation")
            for prop_id, key in ((4, "author"), (8, "last_saved_by"),
                                  (5, "last_saved_date"), (12, "created_date")):
                if prop_id in props:
                    val = props[prop_id]
                    if isinstance(val, bytes):
                        val = val.decode("utf-8", "ignore")
                    v = safe_str(val, "")
                    if v:
                        meta[key] = v
        if ole.exists("\x05DocumentSummaryInformation"):
            props = ole.getproperties("\x05DocumentSummaryInformation")
            if 15 in props:
                val = props[15]
                if isinstance(val, bytes):
                    val = val.decode("utf-8", "ignore")
                meta["company"] = safe_str(val, "")
        ole.close()
    except Exception:
        pass
    return meta


def extract_fallback_document(path: str) -> Dict[str, Any]:
    if not os.path.isfile(path):
        return {
            "Archivo": os.path.basename(path), "Ruta_Completa": path,
            "Modo": "fallback", "Open_Method": "Windows/OLE",
            "Error": "No existe el archivo",
        }

    shell_meta = _read_shell_metadata(path)
    ole_meta   = _read_ole_summary(path)

    author = first_non_empty(shell_meta.get("author"), ole_meta.get("author"), default="Desconocido")
    last   = first_non_empty(shell_meta.get("last_saved_by"), ole_meta.get("last_saved_by"),
                              author, default="Desconocido")
    machine = first_non_empty(shell_meta.get("computer_name"), ole_meta.get("computer_name"), default="")

    from utils import normalize_text as _nt
    def _ml(a, l):
        an, ln = _nt(a), _nt(l)
        if not an and not ln:
            return "sin_metadata"
        if an and ln and an != ln:
            return "inconsistente"
        return "aparentemente_consistente"

    conf = 0
    if shell_meta or ole_meta:
        conf = 20

    custom_props: Dict[str, str] = {}
    custom_props.update(ole_meta)
    custom_props.update(shell_meta)

    return {
        "Archivo":            os.path.basename(path),
        "Ruta_Completa":      path,
        "Modo":               "fallback",
        "Open_Method":        "Windows/OLE",
        "Extension":          os.path.splitext(path)[1].lower(),
        "Tamano_Bytes":       os.path.getsize(path),
        "Hash_Corto":         fast_file_hash(path),
        "Fecha_Modificacion": format_datetime(os.path.getmtime(path)),
        "Autor_Original":     safe_str(author, "Desconocido"),
        "Ultimo_Guardado":    safe_str(last, "Desconocido"),
        "Nombre_Maquina":     machine,
        "Feature_Count":      0,
        "Feature_Types":      "",
        "Feature_Names":      "",
        "Feature_Signature":  "",
        "Custom_Props":       custom_props,
        "Metadata_Status":    _ml(author, last),
        "Confidence":         conf,
        "Error":              "",
    }
