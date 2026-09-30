from __future__ import annotations

"""Motor de similitud multiseñal de PrivateerCAD.

No decide culpabilidad. Ordena coincidencias para que una persona revise la
evidencia. Las señales débiles tienen límites explícitos y el contexto del lote
reduce el peso de árboles comunes a toda una práctica.
"""

from collections import Counter
from difflib import SequenceMatcher
import json
import math
import os
from typing import Any, Dict, List, Optional, Sequence, Tuple

from forensics import parse_chunk_hashes, parse_ole_streams
from utils import normalize_text, parse_datetime_any


UNKNOWN_VALUES = {"", "nan", "none", "desconocido", "unknown"}


def _text(value: Any) -> str:
    text = str(value or "").strip()
    return "" if text.lower() in UNKNOWN_VALUES else text


def _number(value: Any, default: float = 0.0) -> float:
    try:
        result = float(value)
    except (TypeError, ValueError):
        return default
    return result if math.isfinite(result) else default


def _json(value: Any, default: Any) -> Any:
    if isinstance(value, type(default)):
        return value
    text = _text(value)
    if not text:
        return default
    try:
        parsed = json.loads(text)
        return parsed if isinstance(parsed, type(default)) else default
    except Exception:
        return default


def _date(record: Dict[str, Any], primary: str, alias: str) -> Any:
    return parse_datetime_any(_text(record.get(primary)) or _text(record.get(alias)))


def _date_delta(a: Dict[str, Any], b: Dict[str, Any], primary: str, alias: str) -> Optional[float]:
    da = _date(a, primary, alias)
    db = _date(b, primary, alias)
    return abs((da - db).total_seconds()) if da and db else None


def _split_types(record: Dict[str, Any]) -> List[str]:
    value = record.get("Feature_Types", "")
    if isinstance(value, (list, tuple)):
        raw = value
    else:
        raw = str(value or "").split(">")
    return [normalize_text(item) for item in raw if normalize_text(item)]


def _feature_structure(record: Dict[str, Any]) -> List[Dict[str, Any]]:
    rows = _json(record.get("Feature_Structure"), [])
    output: List[Dict[str, Any]] = []
    for row in rows:
        if not isinstance(row, dict):
            continue
        token = normalize_text(row.get("type"))
        if not token:
            continue
        params = []
        for raw in row.get("parameters", []):
            value = _number(raw, math.nan)
            if math.isfinite(value):
                params.append(value)
        sketch = row.get("sketch") if isinstance(row.get("sketch"), dict) else {}
        output.append({
            "type": token,
            "depth": int(_number(row.get("depth"), 0)),
            "parameters": params,
            "sketch": sketch,
        })
    if output:
        return output
    return [
        {"type": token, "depth": 0, "parameters": [], "sketch": {}}
        for token in _split_types(record)
    ]


def _ngrams(tokens: Sequence[str], size: int) -> Counter:
    if len(tokens) < size:
        return Counter()
    return Counter(tuple(tokens[i:i + size]) for i in range(len(tokens) - size + 1))


def _counter_jaccard(a: Counter, b: Counter) -> float:
    if not a or not b:
        return 0.0
    keys = set(a) | set(b)
    intersection = sum(min(a[key], b[key]) for key in keys)
    union = sum(max(a[key], b[key]) for key in keys)
    return intersection / union if union else 0.0


def _cosine(a: Counter, b: Counter) -> float:
    if not a or not b:
        return 0.0
    dot = sum(a[key] * b.get(key, 0) for key in a)
    norm_a = math.sqrt(sum(value * value for value in a.values()))
    norm_b = math.sqrt(sum(value * value for value in b.values()))
    return dot / (norm_a * norm_b) if norm_a and norm_b else 0.0


def _relative_similarity(a: float, b: float, scale: float = 0.025) -> float:
    denominator = max(abs(a), abs(b), 1e-12)
    relative_error = abs(a - b) / denominator
    return math.exp(-relative_error / scale)


def _numeric_list_similarity(a: Sequence[float], b: Sequence[float]) -> float:
    if not a or not b:
        return 0.0
    count = min(len(a), len(b))
    pair_score = sum(_relative_similarity(a[i], b[i], 0.015) for i in range(count)) / count
    length_score = min(len(a), len(b)) / max(len(a), len(b))
    return 0.85 * pair_score + 0.15 * length_score


def feature_similarity(a: Dict[str, Any], b: Dict[str, Any]) -> Dict[str, Any]:
    rows_a = _feature_structure(a)
    rows_b = _feature_structure(b)
    types_a = [row["type"] for row in rows_a]
    types_b = [row["type"] for row in rows_b]
    if not types_a or not types_b:
        return {"score": 0.0, "complexity": 0.0, "parameter_score": 0.0,
                "parameter_coverage": 0, "count": min(len(types_a), len(types_b))}

    sequence = SequenceMatcher(None, types_a, types_b, autojunk=False).ratio()
    bigrams = _counter_jaccard(_ngrams(types_a, 2), _ngrams(types_b, 2))
    trigrams = _counter_jaccard(_ngrams(types_a, 3), _ngrams(types_b, 3))
    ngram = 0.65 * bigrams + 0.35 * trigrams if min(len(types_a), len(types_b)) >= 3 else bigrams
    multiset = _cosine(Counter(types_a), Counter(types_b))
    count_ratio = min(len(types_a), len(types_b)) / max(len(types_a), len(types_b))

    depths_a = [str(row["depth"]) for row in rows_a]
    depths_b = [str(row["depth"]) for row in rows_b]
    depth = SequenceMatcher(None, depths_a, depths_b, autojunk=False).ratio()

    parameter_scores: List[float] = []
    sketch_scores: List[float] = []
    matcher = SequenceMatcher(None, types_a, types_b, autojunk=False)
    for block in matcher.get_matching_blocks():
        for offset in range(block.size):
            ra = rows_a[block.a + offset]
            rb = rows_b[block.b + offset]
            if ra["parameters"] and rb["parameters"]:
                parameter_scores.append(_numeric_list_similarity(ra["parameters"], rb["parameters"]))
            if ra["sketch"] and rb["sketch"]:
                keys = set(ra["sketch"]) | set(rb["sketch"])
                left = Counter({key: int(_number(ra["sketch"].get(key), 0)) for key in keys})
                right = Counter({key: int(_number(rb["sketch"].get(key), 0)) for key in keys})
                sketch_scores.append(_counter_jaccard(left, right))

    param_score = sum(parameter_scores + sketch_scores) / len(parameter_scores + sketch_scores) \
        if parameter_scores or sketch_scores else 0.0
    base = 0.35 * sequence + 0.25 * ngram + 0.20 * multiset + 0.12 * count_ratio + 0.08 * depth
    if parameter_scores or sketch_scores:
        base = 0.88 * base + 0.12 * param_score

    min_count = min(len(types_a), len(types_b))
    complexity = min(1.0, max(0.0, (min_count - 2) / 8.0))
    return {
        "score": round(max(0.0, min(base, 1.0)), 4),
        "complexity": round(complexity, 4),
        "parameter_score": round(param_score, 4),
        "parameter_coverage": len(parameter_scores) + len(sketch_scores),
        "count": min_count,
    }


def _geometry_value_similarity(a: Any, b: Any, *, count: bool = False) -> Optional[float]:
    if isinstance(a, (list, tuple)) and isinstance(b, (list, tuple)):
        left = [_number(value, math.nan) for value in a]
        right = [_number(value, math.nan) for value in b]
        left = [value for value in left if math.isfinite(value)]
        right = [value for value in right if math.isfinite(value)]
        return _numeric_list_similarity(sorted(left), sorted(right)) if left and right else None
    left = _number(a, math.nan)
    right = _number(b, math.nan)
    if not math.isfinite(left) or not math.isfinite(right):
        return None
    if count:
        denominator = max(abs(left), abs(right), 1.0)
        return max(0.0, 1.0 - abs(left - right) / denominator)
    return _relative_similarity(left, right)


def geometry_similarity(a: Dict[str, Any], b: Dict[str, Any]) -> Dict[str, Any]:
    ga = _json(a.get("Geometry_Data"), {})
    gb = _json(b.get("Geometry_Data"), {})
    if not ga or not gb:
        return {"score": 0.0, "coverage": 0}

    weights = {
        "volume": 1.8,
        "surface_area": 1.6,
        "mass": 0.8,
        "bbox_dimensions": 1.6,
        "principal_moments": 1.0,
        "body_count": 0.7,
        "face_count": 1.0,
        "edge_count": 1.0,
    }
    weighted = 0.0
    total = 0.0
    coverage = 0
    for key, weight in weights.items():
        if key == "mass" and (ga.get("user_assigned_mass") or gb.get("user_assigned_mass")):
            continue
        if key not in ga or key not in gb:
            continue
        similarity = _geometry_value_similarity(
            ga[key], gb[key], count=key.endswith("_count")
        )
        if similarity is None:
            continue
        weighted += weight * similarity
        total += weight
        coverage += 1
    return {"score": round(weighted / total, 4) if total else 0.0, "coverage": coverage}


def component_similarity(a: Dict[str, Any], b: Dict[str, Any]) -> Dict[str, Any]:
    ca = _json(a.get("Component_Structure"), [])
    cb = _json(b.get("Component_Structure"), [])
    if not ca or not cb:
        return {"score": 0.0, "count": 0}

    def token(row: Dict[str, Any]) -> str:
        return "|".join((
            normalize_text(row.get("file")),
            normalize_text(row.get("configuration")),
            normalize_text(row.get("suppression")),
        ))

    left = Counter(token(row) for row in ca if isinstance(row, dict))
    right = Counter(token(row) for row in cb if isinstance(row, dict))
    if not left or not right:
        return {"score": 0.0, "count": 0}
    score = 0.55 * _counter_jaccard(left, right) + 0.45 * _cosine(left, right)
    return {"score": round(score, 4), "count": min(sum(left.values()), sum(right.values()))}


def binary_similarity(a: Dict[str, Any], b: Dict[str, Any]) -> Dict[str, Any]:
    chunks_a = Counter(parse_chunk_hashes(a.get("Binary_Chunk_Hashes")))
    chunks_b = Counter(parse_chunk_hashes(b.get("Binary_Chunk_Hashes")))
    chunk_score = 0.0
    chunk_matches = 0
    if chunks_a and chunks_b:
        keys = set(chunks_a) | set(chunks_b)
        chunk_matches = sum(min(chunks_a[key], chunks_b[key]) for key in keys)
        minimum = min(sum(chunks_a.values()), sum(chunks_b.values()))
        union = sum(max(chunks_a[key], chunks_b[key]) for key in keys)
        containment = chunk_matches / minimum if minimum else 0.0
        jaccard = chunk_matches / union if union else 0.0
        chunk_score = 0.7 * containment + 0.3 * jaccard

    streams_a = parse_ole_streams(a.get("OLE_Stream_Hashes"))
    streams_b = parse_ole_streams(b.get("OLE_Stream_Hashes"))
    matching_bytes = 0
    matching_streams = 0
    total_a = sum(int(_number(item.get("size"), 0)) for item in streams_a.values() if isinstance(item, dict))
    total_b = sum(int(_number(item.get("size"), 0)) for item in streams_b.values() if isinstance(item, dict))
    for name in set(streams_a) & set(streams_b):
        left = streams_a.get(name, {})
        right = streams_b.get(name, {})
        if not isinstance(left, dict) or not isinstance(right, dict):
            continue
        if _text(left.get("hash")) and _text(left.get("hash")) == _text(right.get("hash")):
            matching_streams += 1
            matching_bytes += min(int(_number(left.get("size"), 0)), int(_number(right.get("size"), 0)))
    ole_score = matching_bytes / min(total_a, total_b) if min(total_a, total_b) else 0.0
    return {
        "chunk_score": round(chunk_score, 4),
        "chunk_matches": chunk_matches,
        "chunk_count": min(sum(chunks_a.values()), sum(chunks_b.values())),
        "ole_score": round(ole_score, 4),
        "ole_matches": matching_streams,
        "ole_count": min(len(streams_a), len(streams_b)),
    }


def _created_key(record: Dict[str, Any]) -> str:
    value = _date(record, "SW_Created_Date", "Fecha_Creacion_SW")
    return value.strftime("%Y-%m-%dT%H:%M:%S") if value else ""


def _feature_key(record: Dict[str, Any]) -> str:
    return ">".join(_split_types(record))


def build_cohort_context(records: Sequence[Dict[str, Any]]) -> Dict[str, Any]:
    return {
        "size": len(records),
        "created_frequency": Counter(key for key in map(_created_key, records) if key),
        "feature_frequency": Counter(key for key in map(_feature_key, records) if key),
    }


def _valid_hash(record: Dict[str, Any]) -> str:
    value = _text(record.get("SHA256_Completo") or record.get("Hash_Corto"))
    return "" if value.upper().startswith("ERROR") else value.lower()


def _path(record: Dict[str, Any]) -> str:
    return _text(record.get("Ruta_Completa")) or _text(record.get("Archivo"))


def _direction(a: Dict[str, Any], b: Dict[str, Any]) -> Tuple[Dict[str, Any], Dict[str, Any], float, str]:
    created_a = _date(a, "SW_Created_Date", "Fecha_Creacion_SW")
    created_b = _date(b, "SW_Created_Date", "Fecha_Creacion_SW")
    if created_a and created_b and abs((created_a - created_b).total_seconds()) > 5:
        return (a, b, 0.65, "fecha de creación") if created_a < created_b \
            else (b, a, 0.65, "fecha de creación")

    saved_a = _date(a, "SW_Saved_Date", "Fecha_Ultimo_Guardado_SW")
    saved_b = _date(b, "SW_Saved_Date", "Fecha_Ultimo_Guardado_SW")
    if saved_a and saved_b and abs((saved_a - saved_b).total_seconds()) > 5:
        return (a, b, 0.72, "fecha de guardado") if saved_a < saved_b \
            else (b, a, 0.72, "fecha de guardado")

    # Un orden estable permite renderizar la arista, pero no afirma el origen.
    ordered = sorted((a, b), key=lambda row: (_path(row).lower(), _text(row.get("Archivo")).lower()))
    return ordered[0], ordered[1], 0.0, "dirección indeterminada"


def score_pair(
    a: Dict[str, Any],
    b: Dict[str, Any],
    context: Optional[Dict[str, Any]] = None,
) -> Dict[str, Any]:
    context = context or {"size": 2, "created_frequency": Counter(), "feature_frequency": Counter()}
    reasons: List[str] = []

    hash_a = _valid_hash(a)
    hash_b = _valid_hash(b)
    exact_hash = bool(hash_a and hash_b and hash_a == hash_b)
    created_delta = _date_delta(a, b, "SW_Created_Date", "Fecha_Creacion_SW")
    saved_delta = _date_delta(a, b, "SW_Saved_Date", "Fecha_Ultimo_Guardado_SW")

    features = feature_similarity(a, b)
    geometry = geometry_similarity(a, b)
    components = component_similarity(a, b)
    binary = binary_similarity(a, b)

    source, target, direction_confidence, direction_basis = _direction(a, b)
    if exact_hash:
        reasons.append("SHA-256 completo idéntico (duplicado byte a byte)")
        score = 100
        decision = "DUPLICADO_EXACTO"
        confidence = "ALTA"
    else:
        feature_key = _feature_key(a) if _feature_key(a) == _feature_key(b) else ""
        feature_frequency = int(context.get("feature_frequency", {}).get(feature_key, 0)) if feature_key else 0
        cohort_size = max(2, int(context.get("size", 2)))
        common_ratio = feature_frequency / cohort_size if feature_frequency else 0.0
        common_penalty = 1.0
        if feature_frequency >= 3:
            common_penalty = max(0.50, 1.0 - 0.65 * common_ratio)

        feature_reliability = 0.35 + 0.65 * features["complexity"]
        feature_points = 42.0 * features["score"] * feature_reliability * common_penalty
        component_points = 22.0 * components["score"] * min(1.0, components["count"] / 5.0)
        structure_points = max(feature_points, component_points)

        geometry_reliability = min(1.0, geometry["coverage"] / 4.0)
        geometry_points = 27.0 * geometry["score"] * geometry_reliability
        if feature_frequency >= 3 and features["score"] >= 0.98:
            geometry_points *= max(0.65, 1.0 - 0.35 * common_ratio)

        chunk_reliability = min(1.0, binary["chunk_count"] / 6.0)
        chunk_points = 25.0 * binary["chunk_score"] * chunk_reliability
        ole_reliability = min(1.0, binary["ole_count"] / 5.0)
        ole_points = 18.0 * binary["ole_score"] * ole_reliability
        content_points = max(chunk_points, ole_points)

        created_points = 0.0
        if created_delta is not None and created_delta <= 5:
            key = _created_key(a)
            frequency = int(context.get("created_frequency", {}).get(key, 2))
            rarity = 1.0 / max(1.0, frequency - 1.0)
            created_points = 12.0 * max(0.25, rarity)
        elif created_delta is not None and created_delta <= 120:
            created_points = 3.0
        saved_points = 4.0 if saved_delta is not None and saved_delta <= 5 else 0.0

        score_value = structure_points + geometry_points + content_points + created_points + saved_points

        feature_strong = features["score"] >= 0.88 and features["count"] >= 5
        component_strong = components["score"] >= 0.90 and components["count"] >= 3
        geometry_strong = geometry["score"] >= 0.96 and geometry["coverage"] >= 3
        content_strong = (
            (binary["chunk_score"] >= 0.35 and binary["chunk_matches"] >= 2)
            or (binary["ole_score"] >= 0.45 and binary["ole_matches"] >= 2)
        )
        core_families = sum((feature_strong or component_strong, geometry_strong, content_strong))

        support_count = sum((
            created_points >= 6,
            saved_points > 0,
            features["parameter_score"] >= 0.95 and features["parameter_coverage"] >= 2,
        ))
        if (feature_strong or component_strong) and geometry_strong:
            score_value += 10
        if (feature_strong or component_strong) and content_strong:
            score_value += 10
        if created_points >= 6 and core_families:
            score_value += 6

        if core_families == 0:
            score_value = min(score_value, 29)
        elif core_families == 1 and support_count == 0:
            score_value = min(score_value, 59)
        if (feature_frequency >= 3 and common_ratio >= 0.50
                and not content_strong and created_points < 6):
            # Un patrón mayoritario suele describir la consigna o plantilla.
            score_value = min(score_value, 39)
        elif feature_frequency >= 3 and not content_strong and created_points < 6:
            score_value = min(score_value, 64)

        ext_a = os.path.splitext(_text(a.get("Archivo")))[1].lower()
        ext_b = os.path.splitext(_text(b.get("Archivo")))[1].lower()
        if ext_a and ext_b and ext_a != ext_b:
            score_value = min(score_value, 20)

        score = int(round(max(0.0, min(score_value, 100.0))))
        decision = "COINCIDENCIA_FUERTE" if score >= 75 else (
            "REVISAR" if score >= 45 else "SIN_COINCIDENCIA_RELEVANTE"
        )
        confidence = "ALTA" if core_families >= 2 else ("MEDIA" if core_families == 1 else "BAJA")

        if features["score"] >= 0.65:
            detail = f"estructura {features['score']:.0%} ({features['count']} operaciones)"
            if feature_frequency >= 3:
                detail += f", patrón compartido por {feature_frequency}/{cohort_size} archivos"
            reasons.append(detail)
        if components["score"] >= 0.65:
            reasons.append(f"estructura de ensamble {components['score']:.0%}")
        if geometry["score"] >= 0.80 and geometry["coverage"]:
            reasons.append(f"geometría {geometry['score']:.0%} ({geometry['coverage']} métricas)")
        if binary["chunk_score"] >= 0.20 and binary["chunk_matches"]:
            reasons.append(
                f"contenido binario {binary['chunk_score']:.0%} "
                f"({binary['chunk_matches']} fragmentos coincidentes)"
            )
        if binary["ole_score"] >= 0.20 and binary["ole_matches"]:
            reasons.append(
                f"flujos internos OLE {binary['ole_score']:.0%} "
                f"({binary['ole_matches']} coincidentes)"
            )
        if created_delta is not None and created_delta <= 5:
            reasons.append("misma fecha interna de creación SW (señal de procedencia)")
        elif created_delta is not None and created_delta <= 120:
            reasons.append(f"fechas internas de creación próximas (Δ={created_delta:.0f}s)")
        if saved_delta is not None and saved_delta <= 5:
            reasons.append("misma fecha interna de guardado SW (señal auxiliar)")

    return {
        "score": score,
        "decision": decision,
        "comparison_confidence": confidence,
        "feature_similarity": features["score"],
        "feature_parameter_similarity": features["parameter_score"],
        "geometry_similarity": geometry["score"],
        "geometry_coverage": geometry["coverage"],
        "component_similarity": components["score"],
        "binary_similarity": binary["chunk_score"],
        "ole_similarity": binary["ole_score"],
        "created_delta": created_delta,
        "saved_delta": saved_delta,
        "exact_hash": exact_hash,
        "reasons": reasons,
        "direction_confidence": direction_confidence,
        "direction_basis": direction_basis,
        "source_file": _text(source.get("Archivo")),
        "target_file": _text(target.get("Archivo")),
        "source_path": _path(source),
        "target_path": _path(target),
        "source_data": source,
        "target_data": target,
    }
