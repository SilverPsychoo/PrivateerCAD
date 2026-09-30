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
    GENERIC_USERNAMES,
)
from detection_engine import build_cohort_context, score_pair
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


# ─────────────────────────────────────────────────────────────────────────────
# Score de plagio entre un par
# ─────────────────────────────────────────────────────────────────────────────

def _pair_score(a: Dict, b: Dict, context: Optional[Dict] = None) -> Dict[str, Any]:
    """Compatibilidad interna para el nuevo motor multiseñal."""
    return score_pair(a, b, context)


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
        lines.append("✍️  Autor SW              : No disponible")

    if propietario:
        lines.append(f"🖥️  Propietario Windows   : {propietario}")
        lines.append("   (Quién tiene el archivo AHORA — cambia al copiarlo para revisar)")
    else:
        lines.append("🖥️  Propietario Windows   : No disponible")

    lines.append(f"📅 Fecha modificación    : {data.get('Fecha_Modificacion', '—')}")
    lines.append("")
    lines.append("── Fechas internas SW ─────────────────────────────────────────")
    if fc_sw:
        lines.append(f"🗓️  Creado                : {fc_sw}  ← FIJA aunque guardes de nuevo")
    else:
        lines.append("🗓️  Creado                : No disponible")
    if fs_sw:
        lines.append(f"🗓️  Último guardado SW    : {fs_sw}  ← cambia al volver a guardar")
    else:
        lines.append("🗓️  Último guardado SW    : No disponible")

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
        lines.append("   El lote comparará estos datos contra todos los archivos del grupo.")
    elif autor_real or fc_sw or feats > 0:
        estado = "EVIDENCIA_PARCIAL"
        lines.append("\n🟠 Datos parciales — suficientes para comparar en lote.")
    else:
        estado = "BAJA_CONFIANZA"
        lines.append("\n⚪ Datos insuficientes. El lote intentará comparar con el grupo.")

    return "\n".join(lines), estado


# ─────────────────────────────────────────────────────────────────────────────
# Análisis por lote
# ─────────────────────────────────────────────────────────────────────────────

def analizar_lote(datos: List[Dict[str, Any]]) -> Tuple[Any, str, List[Dict]]:
    if not datos:
        return None, "No se encontraron archivos CAD válidos.", []

    # No modificar los diccionarios que conserva la interfaz/extractor.
    datos = [dict(item) for item in datos]

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
        "Feature_Structure": "[]", "Geometry_Data": "{}",
        "Component_Count": 0, "Component_Structure": "[]",
        "Hash_Corto": "", "Fecha_Modificacion": "Desconocido",
        "SHA256_Completo": "", "Binary_Chunk_Hashes": "",
        "Binary_Chunk_Count": 0, "OLE_Stream_Hashes": "{}",
        "OLE_Stream_Count": 0,
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
                "Feature_Structure", "Geometry_Data", "Component_Structure",
                "SHA256_Completo", "Binary_Chunk_Hashes", "OLE_Stream_Hashes",
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
    cohort_context = build_cohort_context(registros)

    # ── Comparar TODOS los pares ──────────────────────────────────────────
    pares: List[Dict]      = []
    relaciones: List[Dict] = []

    for i, j in combinations(range(len(registros)), 2):
        pair = _pair_score(registros[i], registros[j], cohort_context)
        pares.append(pair)
        if pair["score"] >= GRAPH_EDGE_THRESHOLD:
            relaciones.append({
                "source":      pair["source_path"],
                "target":      pair["target_path"],
                "source_file": pair["source_file"],
                "target_file": pair["target_file"],
                "score":       pair["score"],
                "feature_similarity": pair["feature_similarity"],
                "geometry_similarity": pair["geometry_similarity"],
                "binary_similarity": pair["binary_similarity"],
                "ole_similarity": pair["ole_similarity"],
                "comparison_confidence": pair["comparison_confidence"],
                "direction_confidence": pair["direction_confidence"],
                "direction_basis": pair["direction_basis"],
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
    confianzas, similitudes_estructura, similitudes_geometria = [], [], []
    similitudes_binarias, decisiones = [], []
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
        if (match and score >= SUSPECT_THRESHOLD
                and float(match.get("direction_confidence", 0)) >= 0.60):
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
        confianzas.append(match.get("comparison_confidence", "BAJA") if match else "BAJA")
        similitudes_estructura.append(match.get("feature_similarity", 0) if match else 0)
        similitudes_geometria.append(match.get("geometry_similarity", 0) if match else 0)
        similitudes_binarias.append(
            max(match.get("binary_similarity", 0), match.get("ole_similarity", 0)) if match else 0
        )
        decisiones.append(match.get("decision", "SIN_COINCIDENCIA_RELEVANTE") if match else "SIN_COINCIDENCIA_RELEVANTE")

    df["Puntaje_Sospecha"] = puntajes
    df["Estado"]           = estados
    df["Detalle_Sospecha"] = detalles
    df["Posible_Fuente"]   = fuentes
    df["Confianza_Comparacion"] = confianzas
    df["Similitud_Estructura"] = similitudes_estructura
    df["Similitud_Geometria"] = similitudes_geometria
    df["Similitud_Binaria"] = similitudes_binarias
    df["Decision_Par"] = decisiones

    # ── Detecciones especiales ────────────────────────────────────────────
    col_created  = _detectar_colisiones_fecha_creacion(registros, pares)
    col_saved    = _detectar_colisiones_fecha_guardado(registros, pares)
    grupos_autor = _agrupar_por_autor(registros)
    paciente     = _detectar_paciente_cero(registros, relaciones)
    duplicados_exactos = [pair for pair in pares if pair.get("exact_hash")]

    # ── Reporte ───────────────────────────────────────────────────────────
    total         = len(df)
    n_alto        = int((df["Puntaje_Sospecha"] >= HIGH_RISK_THRESHOLD).sum())
    n_sospechosos = int(((df["Puntaje_Sospecha"] >= SUSPECT_THRESHOLD) &
                         (df["Puntaje_Sospecha"] < HIGH_RISK_THRESHOLD)).sum())

    sep   = "─" * 62
    lines = []
    lines.append(sep)
    lines.append(f"  REPORTE DE ANÁLISIS  —  {total} archivos")
    lines.append(sep)
    lines.append(f"  🔴 COINCIDENCIA FUERTE — revisar : {n_alto}")
    lines.append(f"  🟠 COINCIDENCIA SOSPECHOSA        : {n_sospechosos}")
    lines.append("")

    if duplicados_exactos:
        lines.append("🚨 DUPLICADOS EXACTOS (SHA-256 completo):")
        for pair in duplicados_exactos:
            lines.append(f"   ↔  {pair['source_file']}  ==  {pair['target_file']}")
        lines.append("")

    if col_created:
        lines.append("⚠️  FECHA DE CREACIÓN SW COINCIDENTE, CON OTRA EVIDENCIA:")
        lines.append("   La fecha se usa como procedencia; nunca decide el resultado por sí sola.")
        lines.append("")
        for col in col_created:
            lines.append(f"   ↔  {col['a']}")
            lines.append(f"      {col['b']}")
            lines.append(
                f"      Fecha: {col['fecha']}  (Δ={col['delta']:.1f}s) · "
                f"Score {col['score']}/100"
            )
        lines.append("")

    if col_saved:
        lines.append("⚠️  FECHA DE GUARDADO SW COINCIDENTE, CON OTRA EVIDENCIA:")
        lines.append("   Es una señal auxiliar; la igualdad exacta solo la confirma SHA-256.")
        lines.append("")
        for col in col_saved:
            lines.append(f"   ↔  {col['a']}  /  {col['b']}")
            lines.append(
                f"      Fecha: {col['fecha']}  (Δ={col['delta']:.1f}s) · "
                f"Score {col['score']}/100"
            )
        lines.append("")

    pares_relevantes = sorted(
        (pair for pair in pares if pair["score"] >= SUSPECT_THRESHOLD),
        key=lambda pair: pair["score"], reverse=True,
    )
    if pares_relevantes:
        lines.append("COINCIDENCIAS ENTRE ARCHIVOS:")
        for pair in pares_relevantes:
            lines.append(
                f"   {pair['source_file']}  ↔  {pair['target_file']}  ·  "
                f"{pair['score']}/100  ·  confianza {pair['comparison_confidence'].lower()}"
            )
            if pair["reasons"]:
                lines.append(f"      {'; '.join(pair['reasons'])}")
            if pair.get("direction_confidence", 0) >= 0.60:
                lines.append(
                    f"      Posible dirección: {pair['source_file']} → {pair['target_file']} "
                    f"({pair['direction_basis']})"
                )
        lines.append("")

    if paciente:
        lines.append(f"🦠 POSIBLE ARCHIVO DE ORIGEN: '{paciente['nombre']}'")
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
        lines.append(f"   Confianza    : {row['Confianza_Comparacion']}")
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

def _detectar_colisiones_fecha_creacion(
    registros: List[Dict], pares: Optional[List[Dict]] = None
) -> List[Dict]:
    """Fechas de creación coincidentes que además superan el umbral de revisión."""
    cols = []
    for pair_index, (i, j) in enumerate(combinations(range(len(registros)), 2)):
        a, b = registros[i], registros[j]
        delta = _sw_created_delta(a, b)
        pair = pares[pair_index] if pares and pair_index < len(pares) else _pair_score(a, b)
        if (delta is not None and delta <= SW_DATE_COLLISION_WINDOW_SEC
                and pair.get("score", 0) >= SUSPECT_THRESHOLD):
            cols.append({
                "a":     a.get("Archivo", ""),
                "b":     b.get("Archivo", ""),
                "delta": delta,
                "fecha": _clean_sw_date(a.get("SW_Created_Date") or a.get("Fecha_Creacion_SW")),
                "score": pair.get("score", 0),
            })
    return cols


def _detectar_colisiones_fecha_guardado(
    registros: List[Dict], pares: Optional[List[Dict]] = None
) -> List[Dict]:
    """Fechas de guardado coincidentes que además tienen evidencia corroborante."""
    cols = []
    for pair_index, (i, j) in enumerate(combinations(range(len(registros)), 2)):
        a, b = registros[i], registros[j]
        delta = _sw_saved_delta(a, b)
        pair = pares[pair_index] if pares and pair_index < len(pares) else _pair_score(a, b)
        if (delta is not None and delta <= SW_DATE_COLLISION_WINDOW_SEC
                and pair.get("score", 0) >= SUSPECT_THRESHOLD):
            cols.append({
                "a":     a.get("Archivo", ""),
                "b":     b.get("Archivo", ""),
                "delta": delta,
                "fecha": _clean_sw_date(a.get("SW_Saved_Date") or a.get("Fecha_Ultimo_Guardado_SW")),
                "score": pair.get("score", 0),
            })
    return cols


def _agrupar_por_autor(registros: List[Dict]) -> Dict[str, List[str]]:
    grupos: Dict[str, List[str]] = defaultdict(list)
    for r in registros:
        a = _clean_sw_date(r.get("SW_Author_Raw") or r.get("Autor_Original"))
        if _valid_author(a):
            grupos[a].append(r.get("Archivo", ""))
    return {k: v for k, v in grupos.items() if len(v) >= 2}


def _detectar_paciente_cero(registros: List[Dict],
                             relaciones: List[Dict]) -> Optional[Dict]:
    """Propone un origen solo con varias aristas fuertes y dirección sustentada."""
    firmes = [
        rel for rel in relaciones
        if rel.get("score", 0) >= HIGH_RISK_THRESHOLD
        and float(rel.get("direction_confidence", 0)) >= 0.60
    ]
    if not firmes:
        return None

    out_deg = Counter(rel["source"] for rel in firmes)
    in_deg = Counter(rel["target"] for rel in firmes)
    source_path, salidas = out_deg.most_common(1)[0]
    # Una sola pareja no permite inferir quién distribuyó el archivo.
    if salidas < 2:
        return None

    record = next(
        (row for row in registros
         if (row.get("Ruta_Completa") or row.get("Archivo", "")) == source_path),
        None,
    )
    if not record:
        return None
    date = parse_datetime_any(_clean_sw_date(
        record.get("SW_Created_Date") or record.get("Fecha_Creacion_SW")
    ))
    related = [rel for rel in firmes if rel["source"] == source_path]
    avg_direction = sum(float(rel.get("direction_confidence", 0)) for rel in related) / len(related)
    certainty = "ALTA" if avg_direction >= 0.70 and salidas >= 3 else "MEDIA"
    return {
        "nombre": record.get("Archivo", ""),
        "salidas": salidas,
        "entradas": in_deg.get(source_path, 0),
        "certeza": f"{certainty} — varias coincidencias fuertes apuntan al mismo archivo",
        "fecha_sw": date.strftime("%Y-%m-%d %H:%M") if date else "",
    }


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
                   reasons=rel.get("reason_str", ""),
                   direction_confidence=rel.get("direction_confidence", 0))

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

    directed_edges = [
        (u, v) for u, v in G.edges()
        if float(G[u][v].get("direction_confidence", 0)) >= 0.60
    ]
    ambiguous_edges = [edge for edge in G.edges() if edge not in directed_edges]
    if directed_edges:
        nx.draw_networkx_edges(
            G, pos, ax=ax, edgelist=directed_edges, edge_color="#f39c12", arrows=True,
            arrowstyle="->", arrowsize=18,
            width=[max(0.5, G[u][v]["weight"] / 20) for u, v in directed_edges],
            alpha=0.85, connectionstyle="arc3,rad=0.08",
        )
    if ambiguous_edges:
        nx.draw_networkx_edges(
            G, pos, ax=ax, edgelist=ambiguous_edges, edge_color="#95a5a6", arrows=False,
            style="dashed",
            width=[max(0.5, G[u][v]["weight"] / 20) for u, v in ambiguous_edges],
            alpha=0.75, connectionstyle="arc3,rad=0.08",
        )

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
