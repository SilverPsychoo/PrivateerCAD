from __future__ import annotations

import hashlib
import os
import re
from datetime import datetime
from difflib import SequenceMatcher
from typing import Any, Optional


def normalize_text(value: Any) -> str:
    if value is None:
        return ""
    text = str(value).strip().lower()
    text = re.sub(r"\s+", " ", text)
    text = re.sub(r"[^\w\s\-_/.:()]+", "", text, flags=re.UNICODE)
    return text


def safe_str(value: Any, default: str = "Desconocido") -> str:
    if value is None:
        return default
    text = str(value).strip()
    return text if text else default


def first_non_empty(*values: Any, default: str = "") -> str:
    """Retorna el primer valor no vacío, ignorando 'Desconocido' y None."""
    for value in values:
        if value is None:
            continue
        text = str(value).strip()
        if text and text.lower() not in ("desconocido", "unknown", ""):
            return text
    return default


def similarity_ratio(a: str, b: str) -> float:
    return SequenceMatcher(None, a or "", b or "").ratio()


def coerce_tuple_first(result: Any, default: Any = None) -> Any:
    """Extrae el primer valor útil de un resultado que puede ser tuple (COM win32)."""
    if result is None:
        return default
    if isinstance(result, tuple):
        if not result:
            return default
        for item in result:
            if isinstance(item, str) and item.strip():
                return item.strip()
        for item in result:
            if item is None or isinstance(item, bool):
                continue
            if isinstance(item, (int, float)) and item == 0:
                continue
            text = str(item).strip()
            if text:
                return text
        return default
    if isinstance(result, str):
        return result.strip() or default
    text = str(result).strip()
    return text if text else default


def fast_file_hash(path: str, block_size: int = 1024 * 1024) -> str:
    """Hash rápido: tamaño + primer MB + último MB."""
    h = hashlib.sha256()
    try:
        size = os.path.getsize(path)
        h.update(str(size).encode("utf-8", "ignore"))
        with open(path, "rb") as f:
            first = f.read(block_size)
            h.update(first)
            if size > block_size * 2:
                f.seek(max(0, size - block_size))
                h.update(f.read(block_size))
            elif size > block_size:
                f.seek(block_size)
                h.update(f.read())
    except Exception as exc:
        return f"ERROR_HASH:{exc}"
    return h.hexdigest()


def parse_datetime_any(value: Any) -> Optional[datetime]:
    """Convierte cualquier representación de fecha/hora a datetime."""
    if value in (None, "", "Desconocido"):
        return None
    if isinstance(value, datetime):
        return value
    if isinstance(value, (int, float)):
        try:
            return datetime.fromtimestamp(float(value))
        except Exception:
            return None

    text = str(value).strip()
    if not text or text.lower() == "desconocido":
        return None

    # Normalizar formato español de AM/PM: "p. m." → "PM", "a. m." → "AM"
    text_norm = re.sub(r'p\.\s*m\.', 'PM', text, flags=re.IGNORECASE)
    text_norm = re.sub(r'a\.\s*m\.', 'AM', text_norm, flags=re.IGNORECASE)
    text_norm = text_norm.strip()

    # Intentar como número (Unix timestamp en texto)
    try:
        return datetime.fromtimestamp(float(text_norm))
    except Exception:
        pass

    for fmt in (
        "%d/%m/%Y %I:%M:%S %p",   # 10/03/2026 03:07:23 PM  ← formato SW español
        "%d/%m/%Y %I:%M %p",      # 10/03/2026 03:07 PM
        "%Y-%m-%d %H:%M:%S", "%Y-%m-%d %H:%M",
        "%d/%m/%Y %H:%M:%S", "%d/%m/%Y %H:%M",
        "%m/%d/%Y %H:%M:%S", "%m/%d/%Y %H:%M",
        "%Y-%m-%dT%H:%M:%S", "%Y-%m-%dT%H:%M",
        "%d-%m-%Y %H:%M:%S", "%d-%m-%Y %H:%M",
    ):
        try:
            return datetime.strptime(text_norm, fmt)
        except Exception:
            pass

    try:
        return datetime.fromisoformat(text_norm)
    except Exception:
        return None


def format_datetime(value: Any) -> str:
    """Convierte cualquier valor de tiempo a string legible."""
    if isinstance(value, (int, float)):
        try:
            return datetime.fromtimestamp(float(value)).strftime("%Y-%m-%d %H:%M:%S")
        except Exception:
            return "Desconocido"
    dt = parse_datetime_any(value)
    if dt is None:
        return "Desconocido"
    return dt.strftime("%Y-%m-%d %H:%M:%S")


def safe_getattr_or_call(obj: Any, names: tuple, *args, default: Any = None) -> Any:
    for name in names:
        try:
            attr = getattr(obj, name)
        except Exception:
            continue
        try:
            if callable(attr):
                return attr(*args) if args else attr()
            return attr
        except Exception:
            continue
    return default


def ensure_dir(path: str) -> None:
    os.makedirs(path, exist_ok=True)
