from __future__ import annotations

import json
import os
import sys
import unittest


ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(ROOT, "src"))

from detection_engine import build_cohort_context, score_pair


FEATURES = [
    "ProfileFeature", "BossExtrude", "ProfileFeature", "CutExtrude",
    "Fillet", "Chamfer", "Pattern", "Mirror", "HoleWzd", "Shell",
]

GEOMETRY = {
    "volume": 0.001,
    "surface_area": 0.12,
    "mass": 7.8,
    "bbox_dimensions": [0.1, 0.2, 0.3],
    "body_count": 1,
    "face_count": 32,
    "edge_count": 64,
}


def record(
    name: str,
    features=None,
    *,
    geometry=None,
    created: str = "",
    saved: str = "",
    sha256: str = "",
    chunks: str = "",
    names=None,
):
    features = list(features or [])
    names = list(names or [f"op_{index}" for index in range(len(features))])
    structure = [
        {
            "type": feature,
            "depth": 0 if index < 2 else 1,
            "parameters": [round((index + 1) * 0.001, 6)],
            "sketch": {"segments": index + 3, "points": index + 2},
        }
        for index, feature in enumerate(features)
    ]
    return {
        "Archivo": name,
        "Ruta_Completa": name,
        "Feature_Count": len(features),
        "Feature_Types": " > ".join(features),
        "Feature_Names": " > ".join(names),
        "Feature_Structure": json.dumps(structure),
        "Geometry_Data": json.dumps(geometry or {}),
        "SW_Created_Date": created,
        "SW_Saved_Date": saved,
        "SHA256_Completo": sha256,
        "Hash_Corto": sha256,
        "Binary_Chunk_Hashes": chunks,
        "OLE_Stream_Hashes": "{}",
    }


class DetectionEngineTests(unittest.TestCase):
    def test_exact_sha256_is_an_exact_duplicate(self):
        left = record("a.sldprt", FEATURES, sha256="a" * 64)
        right = record("b.sldprt", FEATURES[:-2], sha256="a" * 64)

        result = score_pair(left, right)

        self.assertEqual(result["score"], 100)
        self.assertEqual(result["decision"], "DUPLICADO_EXACTO")
        self.assertTrue(result["exact_hash"])

    def test_same_creation_date_alone_never_flags_plagiarism(self):
        left = record(
            "a.sldprt", FEATURES, geometry=GEOMETRY,
            created="2026-09-01 10:00:00",
        )
        right = record(
            "b.sldprt",
            ["ProfileFeature", "Revolve", "Sweep", "Loft", "Draft", "Rib", "Flex"],
            geometry={"volume": 2, "surface_area": 5, "bbox_dimensions": [1, 2, 4]},
            created="2026-09-01 10:00:00",
        )

        result = score_pair(left, right, build_cohort_context([left, right]))

        self.assertLess(result["score"], 45)
        self.assertEqual(result["comparison_confidence"], "BAJA")

    def test_renaming_features_does_not_hide_a_copy(self):
        left = record(
            "a.sldprt", FEATURES, geometry=GEOMETRY,
            created="2026-09-01 10:00:00", saved="2026-09-01 10:05:00",
            chunks="aa|bb|cc|dd|ee",
            names=[f"Operacion{index}" for index in range(len(FEATURES))],
        )
        right = record(
            "b.sldprt", FEATURES, geometry=GEOMETRY,
            created="2026-09-01 10:00:00", saved="2026-09-01 11:05:00",
            chunks="aa|bb|cc|dd|ff",
            names=[f"Oculto{index}" for index in range(len(FEATURES))],
        )

        result = score_pair(left, right, build_cohort_context([left, right]))

        self.assertGreaterEqual(result["score"], 75)
        self.assertGreaterEqual(result["feature_similarity"], 0.98)
        self.assertGreaterEqual(result["geometry_similarity"], 0.98)

    def test_short_generic_tree_is_not_enough(self):
        left = record("a.sldprt", ["ProfileFeature", "BossExtrude"])
        right = record("b.sldprt", ["ProfileFeature", "BossExtrude"])

        result = score_pair(left, right)

        self.assertLess(result["score"], 45)

    def test_common_assignment_pattern_is_discounted(self):
        cohort = [
            record(
                f"student_{index}.sldprt", FEATURES, geometry=GEOMETRY,
                created=f"2026-09-01 10:{index:02d}:00",
            )
            for index in range(10)
        ]

        result = score_pair(cohort[0], cohort[1], build_cohort_context(cohort))

        self.assertLess(result["score"], 45)
        self.assertTrue(any("patrón compartido" in reason for reason in result["reasons"]))

    def test_different_document_types_are_not_compared_as_copies(self):
        left = record("part.sldprt", FEATURES, geometry=GEOMETRY)
        right = record("assembly.sldasm", FEATURES, geometry=GEOMETRY)

        result = score_pair(left, right)

        self.assertLessEqual(result["score"], 20)

    def test_direction_is_only_claimed_when_dates_support_it(self):
        left = record(
            "older.sldprt", FEATURES, geometry=GEOMETRY,
            created="2026-09-01 10:00:00", saved="2026-09-01 10:05:00",
        )
        right = record(
            "newer.sldprt", FEATURES, geometry=GEOMETRY,
            created="2026-09-01 10:00:00", saved="2026-09-01 11:05:00",
        )

        result = score_pair(left, right, build_cohort_context([left, right]))

        self.assertEqual(result["source_file"], "older.sldprt")
        self.assertGreaterEqual(result["direction_confidence"], 0.60)

    def test_direction_remains_ambiguous_without_temporal_order(self):
        left = record("a.sldprt", FEATURES, geometry=GEOMETRY)
        right = record("b.sldprt", FEATURES, geometry=GEOMETRY)

        result = score_pair(left, right)

        self.assertEqual(result["direction_confidence"], 0.0)


if __name__ == "__main__":
    unittest.main()
