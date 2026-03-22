# PrivateerCAD

**Inspector forense de archivos SolidWorks para detección de plagio académico.**

Herramienta diseñada para profesores de ingeniería que necesitan verificar la autoría de trabajos entregados en SolidWorks. Analiza los metadatos internos de los archivos `.sldprt` y `.sldasm` para determinar quién creó cada pieza, cuándo, y si hay copias entre los trabajos del grupo.

---

## Características

- **Análisis individual** — muestra autor, fechas internas SW, árbol de operaciones completo y máquina donde fue creado
- **Análisis de grupo** — compara todos los archivos de una carpeta y detecta copias automáticamente
- **Detección de paciente cero** — identifica al alumno que distribuyó el archivo original
- **Red de distribución** — grafo visual de quién le pasó el archivo a quién
- **Exportar CSV** — reporte completo para guardar evidencia

### Criterios de detección

| Indicador | Descripción |
|---|---|
| Fecha de creación SW idéntica | Dos archivos creados en la misma sesión → mismo origen |
| Fecha de guardado SW idéntica | Copia exacta sin modificar |
| Hash SHA-256 idéntico | Archivos byte a byte iguales |
| Árbol de operaciones idéntico | Misma secuencia de features |
| Mismo número de operaciones | Refuerzo del árbol |
| Mismo usuario SW en varias piezas | Correlación de autoría |

---

## Requisitos

- Windows 10/11
- Python 3.8+ (recomendado 3.12)
- SolidWorks instalado (para lectura completa de metadatos y árbol de operaciones)

> Sin SolidWorks la app funciona en modo básico — lee metadatos del sistema pero no el árbol de operaciones.

---

## Instalación

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

O doble clic en `run.bat` si está en la raíz del proyecto.

Al iniciar, la app pregunta si usar la API de SolidWorks. Selecciona **Sí** para obtener el análisis completo.

---

## Empaquetar como .exe

```bash
build.bat
```

El ejecutable se genera en `dist/PrivateerCAD/`. Para distribuir, comparte toda esa carpeta.

---

## Estructura del proyecto

```
PrivateerCAD/
├── src/
│   ├── main.py                  # Interfaz gráfica
│   ├── analizador.py            # Motor de detección de plagio
│   ├── extractor.py             # Coordinador de extracción
│   ├── extractor_solidworks.py  # Extractor vía API de SolidWorks
│   ├── extractor_fallback.py    # Extractor vía Windows/OLE
│   ├── config.py                # Umbrales y configuración
│   ├── utils.py                 # Utilidades compartidas
│   ├── logo.png                 # Logo de la aplicación
│   └── icon.ico                 # Ícono del ejecutable
├── requirements.txt
├── privateercad.spec            # Configuración de PyInstaller
├── build.bat                    # Script de empaquetado
├── run.bat                      # Script de ejecución
└── README.md
```

---

## Licencia

MIT — libre para uso académico y personal.
