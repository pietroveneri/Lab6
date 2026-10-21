"""Controlli sulle coppie e riscontro indipendente delle statistiche."""
from datetime import datetime, timedelta
import hashlib
import math
from pathlib import Path
import statistics
import tempfile
import unittest
from unittest.mock import patch

import analisi_lab6 as analysis


class FakeSheet:
    def __init__(self, rows):
        self.rows = rows

    def iter_rows(self, values_only=False):
        return iter(self.rows)


class FakeWorkbook:
    sheetnames = ["Misure"]

    def __init__(self, rows):
        self.rows = rows

    def __getitem__(self, name):
        return FakeSheet(self.rows)

    def close(self):
        pass


class AnalysisTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.records = []
        for source in analysis.SOURCES:
            rows, _ = analysis.read_source(source, analysis.ROOT)
            cls.records.extend(rows)

    def test_real_data_counts_and_separate_artificial_groups(self):
        expected = {"Artificiali 30 bpm": (40, 39, 39), "Artificiali 150 bpm": (33, 28, 28),
                    "Pietro": (40, 38, 23), "Manfred": (40, 39, 21)}
        self.assertEqual(set(analysis.CATEGORIES), set(expected))
        for category, counts in expected.items():
            rows = [r for r in self.records if r["categoria"] == category]
            result = analysis.summarize(rows, category)
            self.assertEqual((result["righe"], result["coppie_valide"], result["entro_5_bpm"]), counts)
        # Le colonne Garmin vuote negli artificiali non sono letture fallite.
        artificial = [r for r in self.records if r["categoria"].startswith("Artificiali ") and r["coppia_valida"]]
        self.assertTrue(all(r["stato_originale"] == "GARMIN_DA_SCARICARE" for r in artificial))
        self.assertEqual(len({r["file"] for r in artificial}), 2)
        self.assertTrue(any(r["ppg_bpm"] > 150 for r in artificial))

    def test_independent_calculation_and_no_outlier_removal(self):
        for category in analysis.CATEGORIES:
            rows = [r for r in self.records if r["categoria"] == category]
            result = analysis.summarize(rows, category)
            differences = [r["ppg_bpm"] - r["riferimento_bpm"] for r in rows if r["coppia_valida"]]
            self.assertAlmostEqual(result["bias_bpm"], statistics.mean(differences), places=12)
            self.assertAlmostEqual(result["sd_bpm"], statistics.stdev(differences), places=12)
            self.assertAlmostEqual(result["mae_bpm"], sum(map(abs, differences)) / len(differences), places=12)
            self.assertAlmostEqual(result["rmse_bpm"], math.sqrt(sum(d*d for d in differences) / len(differences)), places=12)
            self.assertIsNone(result["ic95_bias_inferiore"])
            self.assertIsNone(result["ic95_bias_superiore"])
        self.assertEqual(max(r["differenza_bpm"] for r in self.records if r["categoria"] == "Pietro" and r["coppia_valida"]), 94)
        self.assertEqual(max(r["differenza_bpm"] for r in self.records if r["categoria"] == "Manfred" and r["coppia_valida"]), 62)

    def test_artificial_statistics_do_not_mix_reference_levels(self):
        self.assertNotIn("Artificiali", analysis.CATEGORIES)
        for category, reference, expected_bias in [("Artificiali 30 bpm", 30, 0), ("Artificiali 150 bpm", 150, 15 / 28)]:
            rows = [r for r in self.records if r["categoria"] == category]
            self.assertEqual({r["riferimento_bpm"] for r in rows}, {reference})
            self.assertEqual(len({r["file"] for r in rows}), 1)
            self.assertAlmostEqual(analysis.summarize(rows, category)["bias_bpm"], expected_bias)

    def test_threshold_boundaries_and_missing_not_zero(self):
        rows = [{"coppia_valida": True, "differenza_bpm": d} for d in [-5, 5, -5.001, 5.001]]
        rows.append({"coppia_valida": False, "differenza_bpm": None})
        result = analysis.summarize(rows, "Esempio")
        self.assertEqual(result["coppie_valide"], 4)
        self.assertEqual(result["escluse"], 1)
        self.assertEqual(result["frazione_entro_5"], .5)
        self.assertFalse(result["target_raggiunto"])
        self.assertIsNone(analysis.number(None))
        self.assertIsNone(analysis.number(float("nan")))
        self.assertIsNone(analysis.number("70"))
        self.assertEqual(analysis.number(0), 0)

    def test_time_pairing_validity_and_original_excel_row(self):
        headers = ["N", "Istante previsto UTC", "Microbit ricevuto UTC", "Microbit BPM", "Garmin BPM", "Garmin allineato UTC", "Stato"]
        t = datetime(2026, 10, 18, 10)
        def row(n, ppg=70, ref=70, shift=0, status=analysis.PAIR_STATUS):
            time = t + timedelta(seconds=10 * n)
            return [n, time, time, ppg, ref, time + timedelta(seconds=shift), status]
        rows = [headers, row(1, shift=1), row(2, shift=1.001), [None]*7,
                row(3, ppg=None), row(4, ref=0), row(5, ppg=170, ref=76),
                row(6, status="GARMIN_DA_SCARICARE"), row(7, ppg=201)]
        with tempfile.TemporaryDirectory() as temporary:
            path = Path(temporary) / "fake.xlsx"
            path.write_bytes(b"fixture")
            with patch.object(analysis, "load_workbook", return_value=FakeWorkbook(rows)):
                records, _ = analysis.read_source(analysis.Source(path.name, "Persona"), path.parent)
        self.assertTrue(records[0]["coppia_valida"])
        self.assertFalse(records[1]["coppia_valida"])
        self.assertIn("FUORI_TOLLERANZA", records[1]["motivi_esclusione"])
        self.assertEqual(records[2]["riga_excel"], 5)
        self.assertFalse(records[2]["coppia_valida"])
        self.assertFalse(records[3]["coppia_valida"])
        self.assertTrue(records[4]["coppia_valida"])
        self.assertEqual(records[4]["differenza_bpm"], 94)
        self.assertFalse(records[5]["coppia_valida"])
        self.assertFalse(records[6]["coppia_valida"])

    def test_original_files_unchanged_by_analysis(self):
        before = {s.filename: hashlib.sha256((analysis.ROOT / s.filename).read_bytes()).hexdigest() for s in analysis.SOURCES}
        for source in analysis.SOURCES:
            rows, _ = analysis.read_source(source, analysis.ROOT)
            analysis.summarize(rows, source.category)
        after = {s.filename: hashlib.sha256((analysis.ROOT / s.filename).read_bytes()).hexdigest() for s in analysis.SOURCES}
        self.assertEqual(before, after)


if __name__ == "__main__":
    unittest.main()
