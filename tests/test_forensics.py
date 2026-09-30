from __future__ import annotations

from collections import Counter
import os
import random
import sys
import tempfile
import unittest


ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(ROOT, "src"))

from forensics import collect_binary_evidence, hash_and_chunk_file


class ForensicFingerprintTests(unittest.TestCase):
    def test_full_sha256_changes_when_middle_of_file_changes(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            left = os.path.join(temp_dir, "left.sldprt")
            right = os.path.join(temp_dir, "right.sldprt")
            payload = bytearray(b"A" * 128_000)
            changed = bytearray(payload)
            changed[64_000:64_010] = b"0123456789"
            with open(left, "wb") as handle:
                handle.write(payload)
            with open(right, "wb") as handle:
                handle.write(changed)

            left_evidence = collect_binary_evidence(left)
            right_evidence = collect_binary_evidence(right)

            self.assertNotEqual(left_evidence["SHA256_Completo"], right_evidence["SHA256_Completo"])
            self.assertEqual(len(left_evidence["SHA256_Completo"]), 64)

    def test_content_defined_chunks_survive_an_insertion(self):
        rng = random.Random(20260922)
        original = rng.randbytes(2_000_000)
        modified = original[:900_000] + b"inserted-metadata" * 80 + original[900_000:]

        with tempfile.TemporaryDirectory() as temp_dir:
            left = os.path.join(temp_dir, "left.sldprt")
            right = os.path.join(temp_dir, "right.sldprt")
            with open(left, "wb") as handle:
                handle.write(original)
            with open(right, "wb") as handle:
                handle.write(modified)

            sha_left, chunks_left = hash_and_chunk_file(left)
            sha_right, chunks_right = hash_and_chunk_file(right)

        matches = sum((Counter(chunks_left) & Counter(chunks_right)).values())
        containment = matches / min(len(chunks_left), len(chunks_right))
        self.assertNotEqual(sha_left, sha_right)
        self.assertGreater(containment, 0.70)


if __name__ == "__main__":
    unittest.main()
