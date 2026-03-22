APP_TITLE = "PRIVATEERCAD"
APP_GEOMETRY = "1180x820"

CAD_EXTENSIONS = (".sldprt", ".sldasm")

# ── Umbrales de similitud ──────────────────────────────────────────────────
FEATURE_SIM_THRESHOLD   = 0.92
GRAPH_EDGE_THRESHOLD    = 40   # umbral para dibujar arista en grafo
HIGH_RISK_THRESHOLD     = 50   # "ALTO RIESGO" — hash idéntico (40) + tamaño (5) ya supera
SUSPECT_THRESHOLD       = 30   # "SOSPECHOSO" — misma fecha creación sola (20pt)

# ── Pesos del score de plagio (sobre 100) ─────────────────────────────────
SCORE_HASH_IDENTICO         = 55   # hash SHA idéntico → copia exacta → ALTO RIESGO solo
SCORE_AUTOR_DISTINTO        = 30   # SW-Author ≠ SW-Last Saved By
SCORE_FEATURE_TREE_ALTO     = 35   # similarity ≥ 0.98
SCORE_FEATURE_TREE_MEDIO    = 22   # similarity ≥ 0.90
SCORE_FEATURE_TREE_BAJO     = 12   # similarity ≥ 0.82
SCORE_TIMESTAMP_COLISION    = 25   # misma hora guardado ± 60 s
SCORE_MISMA_MAQUINA         = 18   # mismo hostname
SCORE_MISMO_FEATURE_COUNT   = 7
SCORE_MISMO_TAMANO          = 5
SCORE_METADATA_COMPATIBLE   = 8    # autor_A == last_saved_B

# ── Ventana de colisión de timestamps (segundos) ──────────────────────────
TIMESTAMP_COLLISION_WINDOW_SEC = 60

# ── Features ignorados en árbol de operaciones ────────────────────────────
IGNORED_FEATURE_TYPES = {
    # Carpetas de organización (no son operaciones reales de diseño)
    "HistoryFolder", "OriginProfileFeature", "MaterialFolder",
    "FeatureFolder", "SensorsFolder", "CutListFolder", "DisplayState",
    "AnnotationFolder", "FTRFolder", "BaseBodyFolder",
    "RefSurface", "RefPlane", "RefAxis", "MarkupFolder", "SelectionSetFolder",
    # Carpetas adicionales que SW inserta automáticamente
    "CommentsFolder", "FavoriteFolder", "SwiftAnnotationFolder",
    "DetailCabinet", "BlockFolder", "SketchBlockDefinition",
    "DesignBinder", "DriveWorksFolder", "TableFolder",
}

# ── Propiedades personalizadas a sondear (SW- son las más importantes) ────
PROBE_PROPERTY_NAMES = [
    "SW-Author", "SW-Last Saved By", "SW-Created On", "SW-Last Saved Date",
    "SW-File Name", "SW-Folder Name", "SW-Template Size",
    "Author", "Last Saved By", "LastSavedBy", "Creator", "Created By",
    "Last Edited By", "Modified By", "Revision", "Description",
    "Title", "Subject", "Company", "DrawnBy", "CheckedBy",
    "Autor", "Creado Por", "Modificado Por", "Ultimo Guardado Por",
]

# ── Índices CORRECTOS de SummaryInfo en API SolidWorks ────────────────────

# ── Índices REALES de SummaryInfo según diagnóstico en archivos de alumnos ──
# Verificado con diagnostico_raw.py:
#   [5] = username del autor original (ej: 'lupe', 'gladi')
#   [6] = fecha de creación (string, ej: '10/03/2026 03:07:23 p. m.')
#   [7] = fecha último guardado (string)
#   [8] = fecha de creación formato largo
#   [9] = fecha último guardado formato largo
# Los índices 1-4 estaban vacíos en todos los archivos probados.
SW_SUMMARY_INFO_FIELDS = {
    5: "author",            # ← username del creador (el más importante)
    6: "created_date_sw",   # ← fecha creación
    7: "saved_date_sw",     # ← fecha último guardado
    8: "created_date_sw_long",
    9: "saved_date_sw_long",
}

# También leer índices 1-4 por si acaso en otras versiones de SW
SW_SUMMARY_INFO_EXTRA = {
    1: "title",
    2: "subject",
    3: "author_alt",
    4: "keywords",
}

MAX_SHELL_COLUMNS = 300

# ── Usernames genéricos que NO identifican a una persona real ─────────────
# En PCs de universidades/laboratorios el usuario de Windows suele ser
# "Alumno", "Student", "Usuario", etc. — no sirven para identificar al autor.
# Si el autor SW está en esta lista, buscamos más evidencia antes de concluir.
GENERIC_USERNAMES = {
    "alumno", "alumnos", "student", "students",
    "usuario", "user", "users",
    "admin", "administrator", "administrador",
    "default", "guest", "invitado",
    "pc", "desktop", "lab", "laboratorio",
    "solidworks", "sw", "cad",
    "public", "publico",
}

# ── Ventana para colisión de fecha SW interna (segundos) ─────────────────
# La fecha "Último guardado SW" (SummaryInfo[7]) queda CONGELADA dentro
# del archivo y NO cambia cuando alguien lo copia en otra PC.
# Si dos archivos tienen la misma fecha SW ± esta ventana → mismo archivo.
SW_DATE_COLLISION_WINDOW_SEC = 5   # muy estricto: 5 segundos

