# PrivateerCAD

**Inspector forense de archivos SolidWorks para revisión de similitudes académicas.**

Herramienta diseñada para profesores de ingeniería que necesitan revisar la autoría de trabajos entregados en SolidWorks. Analiza archivos `.sldprt` y `.sldasm`, compara múltiples familias de evidencia y ordena las coincidencias que requieren inspección.

---

## Características

- **Análisis individual** — muestra autor, fechas internas SW y árbol de operaciones
- **Análisis de grupo** — compara todos los archivos de una carpeta con contexto de la práctica
- **Contenido binario** — SHA-256 completo, fragmentos definidos por contenido y flujos OLE
- **Estructura CAD** — secuencia, jerarquía, dimensiones, croquis y componentes de ensamble
- **Geometría** — volumen, área, caja envolvente, cuerpos, caras y aristas
- **Red de similitudes** — diferencia direcciones sustentadas de relaciones ambiguas
- **Exportar CSV** — reporte completo para guardar evidencia

### Criterios de detección

| Indicador | Descripción |
|---|---|
| SHA-256 completo | Confirma un duplicado byte a byte |
| Fragmentos y flujos OLE | Detecta contenido conservado aunque cambien zonas del archivo |
| Árbol estructural | Compara tipos, orden, jerarquía, parámetros y croquis |
| Geometría | Contrasta propiedades físicas y topología del modelo |
| Ensamble | Compara componentes, configuraciones y supresión |
| Fechas y autor | Señales auxiliares; no producen un veredicto por sí solas |
| Frecuencia en el grupo | Reduce el peso de árboles comunes a una plantilla o consigna |

El puntaje indica similitud técnica, no intención académica. Un resultado alto reúne evidencia para revisión; no sustituye el criterio del profesor.

---

## Requisitos

- Windows 10/11
- Python 3.8+ (recomendado 3.12)
- SolidWorks instalado (para lectura completa de metadatos y árbol de operaciones)

> Sin SolidWorks la app compara SHA-256, fragmentos binarios, flujos OLE y metadatos disponibles. El árbol, los parámetros y la geometría requieren la API de SolidWorks.

---

## Instalación para modificación

```bash
git clone https://github.com/SilverPsychoo/PrivateerCAD.git
cd PrivateerCAD

python -m venv .venv
.venv\Scripts\activate

pip install -r requirements.txt
```

---

## Uso

```bash
.venv\Scripts\python.exe src\main.py
```

Al iniciar, la app pregunta si usar la API de SolidWorks. Selecciona **Sí** para obtener el análisis completo.

Los documentos se abren en modo de solo lectura.

## Pruebas

```bash
python -m unittest discover -s tests -v
```

Las pruebas cubren duplicados exactos, operaciones renombradas, cambios binarios localizados, fechas coincidentes sin similitud real, árboles genéricos, patrones comunes de una práctica y dirección de origen ambigua.

---

El ejecutable se genera en `dist/PrivateerCAD/`. Para distribuir, si solo quieres la app para usarla, descarga solo el Release.

---

## Estructura del proyecto

```
PrivateerCAD/
├── src/
│   ├── main.py                  # Interfaz gráfica
│   ├── analizador.py            # Motor de detección de plagio
│   ├── detection_engine.py      # Similitud multiseñal y control de falsos positivos
│   ├── forensics.py             # SHA-256, fragmentos y flujos OLE
│   ├── extractor.py             # Coordinador de extracción
│   ├── extractor_solidworks.py  # Extractor vía API de SolidWorks
│   ├── extractor_fallback.py    # Extractor vía Windows/OLE
│   ├── config.py                # Umbrales y configuración
│   ├── utils.py                 # Utilidades compartidas
│   ├── logo.png                 # Logo de la aplicación
│   └── icon.ico                 # Ícono del ejecutable
├── requirements.txt
└── README.md
```

---

## Licencia

CC0 1.0 Universal. Consulta `LICENSE`.
