CAD Inspector Pro

Qué hace:
- Intenta abrir archivos SolidWorks con OpenDoc7 + GetOpenDocSpec.
- Si SolidWorks falla, cae automáticamente a Windows/OLE.
- Lee Feature Tree filtrando features genéricos para reducir falsos positivos.
- Compara archivos por árbol, hash y metadatos.
- Grafica una red de posibles copias.

Ejecutar:
  pip install -r requirements.txt
  python main.py

Requisitos:
- Windows
- Python 3.12+
- SolidWorks instalado para el modo API
