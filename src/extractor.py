from __future__ import annotations

import os
from typing import Any, Dict, Optional

import extractor_fallback
import extractor_solidworks
from config import CAD_EXTENSIONS


def extraer_archivo(path: str, usar_solidworks: bool = False,
                    sw_session=None) -> Optional[Dict[str, Any]]:
    if not os.path.isfile(path):
        return None

    # Ignorar archivos temporales de Windows (~$archivo.sldprt)
    if os.path.basename(path).startswith("~"):
        return None

    if not path.lower().endswith(CAD_EXTENSIONS):
        return None

    if usar_solidworks and sw_session is not None:
        data = extractor_solidworks.extract_solidworks_document(path, sw_session)
        if data and not data.get("Error"):
            return data

        # SW falló: usar fallback pero rescatar todo lo que SW sí pudo leer
        fallback = extractor_fallback.extract_fallback_document(path)
        if data:
            fallback["Modo"]             = "solidworks_api+fallback"
            fallback["Open_Method"]      = f"{data.get('Open_Method', '')} -> Windows/OLE"
            fallback["SolidWorks_Error"] = data.get("Error", "")
            for key in ("Hash_Corto", "Tamano_Bytes", "Fecha_Modificacion",
                        "SW_Created_Date", "SW_Saved_Date", "SW_Author_Raw",
                        "Fecha_Creacion_SW", "Fecha_Ultimo_Guardado_SW",
                        "Autor_Original", "Feature_Count", "Feature_Types",
                        "Feature_Names", "Feature_Signature", "Summary_Info",
                        "Propietario_Windows", "Nombre_Maquina", "Confidence"):
                val = data.get(key)
                if val and val not in ("Desconocido", "", 0, {}):
                    fallback[key] = val
        return fallback

    return extractor_fallback.extract_fallback_document(path)
