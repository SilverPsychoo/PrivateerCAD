from __future__ import annotations

import os
import sys
import tempfile
import unittest


ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(ROOT, "src"))
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from analizador import analizar_lote
from extractor_fallback import extract_fallback_document
from test_detection_engine import FEATURES, GEOMETRY, record


def complete(item, author, modified):
    item.update({
        "Autor_Original": author,
        "SW_Author_Raw": author,
        "Tamano_Bytes": 1000,
        "Fecha_Modificacion": modified,
        "Confidence": 90,
    })
    return item


class IntegrationTests(unittest.TestCase):
    def test_low_score_never_gets_a_possible_source(self):
        similar_a = complete(record(
            "Ana.sldprt", FEATURES, geometry=GEOMETRY,
            created="2026-09-01 10:00:00", saved="2026-09-01 10:05:00",
            chunks="aa|bb|cc|dd|ee",
        ), "ana", "2026-09-01 10:30:00")
        similar_b = complete(record(
            "Beto.sldprt", FEATURES, geometry=GEOMETRY,
            created="2026-09-01 10:00:00", saved="2026-09-01 11:05:00",
            chunks="aa|bb|cc|dd|ff",
        ), "beto", "2026-09-01 11:30:00")
        different = complete(record(
            "Carla.sldprt", ["ProfileFeature", "Revolve", "Sweep", "Loft", "Draft", "Rib"],
            geometry={"volume": 2, "surface_area": 5, "bbox_dimensions": [1, 2, 4]},
            created="2026-09-01 12:00:00",
        ), "carla", "2026-09-01 12:30:00")

        frame, _, relations = analizar_lote([similar_a, similar_b, different])
        carla = frame.loc[frame["Archivo"] == "Carla.sldprt"].iloc[0]

        self.assertEqual(carla["Posible_Fuente"], "")
        self.assertEqual(len(relations), 1)

    def test_one_pair_does_not_claim_a_distributor(self):
        left = complete(record(
            "Ana.sldprt", FEATURES, geometry=GEOMETRY,
            created="2026-09-01 10:00:00", saved="2026-09-01 10:05:00",
            chunks="aa|bb|cc|dd|ee",
        ), "ana", "2026-09-01 10:30:00")
        right = complete(record(
            "Beto.sldprt", FEATURES, geometry=GEOMETRY,
            created="2026-09-01 10:00:00", saved="2026-09-01 11:05:00",
            chunks="aa|bb|cc|dd|ff",
        ), "beto", "2026-09-01 11:30:00")

        _, report, _ = analizar_lote([left, right])

        self.assertNotIn("POSIBLE ARCHIVO DE ORIGEN", report)

    def test_fallback_always_produces_full_binary_evidence(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            path = os.path.join(temp_dir, "sample.sldprt")
            with open(path, "wb") as handle:
                handle.write(b"not-a-real-solidworks-file" * 10_000)

            data = extract_fallback_document(path)

        self.assertEqual(len(data["SHA256_Completo"]), 64)
        self.assertEqual(data["Hash_Corto"], data["SHA256_Completo"])
        self.assertGreater(data["Binary_Chunk_Count"], 0)
        self.assertEqual(data["Feature_Structure"], "[]")


if __name__ == "__main__":
    unittest.main()
