from __future__ import annotations

"""Huellas binarias no destructivas para archivos CAD.

El módulo no depende de SolidWorks. Combina un SHA-256 completo, fragmentos
definidos por contenido y, cuando el archivo es OLE, hashes por flujo interno.
Las huellas permiten reconocer copias con cambios localizados de metadatos sin
confundir archivos únicamente porque pesan lo mismo.
"""

import hashlib
import json
import mmap
import os
from typing import Any, Dict, List, Tuple


_CONTENT_ANCHORS = (
    b"\x00\x00",
    b"\x3d\xa7",
    b"\x91\xc4",
    b"\xe7\x2b",
    b"\x6a\xf1",
)

_VOLATILE_OLE_MARKERS = (
    "summaryinformation",
    "documentsummaryinformation",
    "preview",
    "thumbnail",
)


def hash_and_chunk_file(
    path: str,
    *,
    min_chunk: int = 16 * 1024,
    average_mask: int = 0xFFFF,
    max_chunk: int = 256 * 1024,
    max_fingerprints: int = 4096,
) -> Tuple[str, List[str]]:
    """Calcula SHA-256 completo y huellas de fragmentos en una sola lectura.

    Los límites de fragmento dependen del contenido, no de offsets fijos. Así,
    insertar unos bytes no desplaza todas las comparaciones posteriores.
    """
    # ``average_mask`` permanece en la firma para no romper llamadas antiguas.
    del average_mask
    size = os.path.getsize(path)
    if size == 0:
        return hashlib.sha256(b"").hexdigest(), []

    chunks: List[str] = []
    with open(path, "rb") as handle:
        with mmap.mmap(handle.fileno(), 0, access=mmap.ACCESS_READ) as mapped:
            sha = hashlib.sha256(mapped).hexdigest()
            start = 0
            while start < size and len(chunks) < max_fingerprints:
                search_from = min(size, start + min_chunk)
                hard_end = min(size, start + max_chunk)
                boundary = hard_end
                if search_from < hard_end:
                    positions = [
                        mapped.find(anchor, search_from, hard_end)
                        for anchor in _CONTENT_ANCHORS
                    ]
                    valid = [position for position in positions if position >= 0]
                    if valid:
                        boundary = min(hard_end, min(valid) + 2)
                if boundary <= start:
                    boundary = min(size, start + max_chunk)
                chunks.append(
                    hashlib.blake2b(mapped[start:boundary], digest_size=8).hexdigest()
                )
                start = boundary
    return sha, chunks


def _ole_stream_fingerprints(path: str) -> Dict[str, Dict[str, Any]]:
    try:
        import olefile
    except Exception:
        return {}

    try:
        if not olefile.isOleFile(path):
            return {}
    except Exception:
        return {}

    result: Dict[str, Dict[str, Any]] = {}
    ole = None
    try:
        ole = olefile.OleFileIO(path)
        for parts in ole.listdir(streams=True, storages=False):
            name = "/".join(str(part) for part in parts)
            normalized = name.lower().replace("\x05", "")
            if any(marker in normalized for marker in _VOLATILE_OLE_MARKERS):
                continue
            try:
                payload = ole.openstream(parts).read()
            except Exception:
                continue
            if not payload:
                continue
            result[name] = {
                "hash": hashlib.blake2b(payload, digest_size=12).hexdigest(),
                "size": len(payload),
            }
    except Exception:
        return {}
    finally:
        if ole is not None:
            try:
                ole.close()
            except Exception:
                pass
    return result


def collect_binary_evidence(path: str) -> Dict[str, Any]:
    """Devuelve campos planos listos para agregarse al registro extraído."""
    empty = {
        "Hash_Corto": "",
        "SHA256_Completo": "",
        "Binary_Chunk_Hashes": "",
        "Binary_Chunk_Count": 0,
        "OLE_Stream_Hashes": "{}",
        "OLE_Stream_Count": 0,
    }
    if not os.path.isfile(path):
        return empty

    try:
        sha256, chunks = hash_and_chunk_file(path)
    except Exception as exc:
        return {**empty, "Hash_Corto": f"ERROR_HASH:{exc}"}

    streams = _ole_stream_fingerprints(path)
    return {
        "Hash_Corto": sha256,
        "SHA256_Completo": sha256,
        "Binary_Chunk_Hashes": "|".join(chunks),
        "Binary_Chunk_Count": len(chunks),
        "OLE_Stream_Hashes": json.dumps(
            streams, ensure_ascii=False, sort_keys=True, separators=(",", ":")
        ),
        "OLE_Stream_Count": len(streams),
    }


def parse_chunk_hashes(value: Any) -> List[str]:
    if isinstance(value, (list, tuple, set)):
        return [str(item) for item in value if str(item)]
    text = str(value or "").strip()
    return [item for item in text.split("|") if item] if text else []


def parse_ole_streams(value: Any) -> Dict[str, Dict[str, Any]]:
    if isinstance(value, dict):
        return value
    text = str(value or "").strip()
    if not text:
        return {}
    try:
        parsed = json.loads(text)
        return parsed if isinstance(parsed, dict) else {}
    except Exception:
        return {}
