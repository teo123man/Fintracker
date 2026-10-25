import datetime as dt
import unittest

from harcama import confirmed_coverage, coverage_complete


class ConfirmedCoverageTests(unittest.TestCase):
    def test_fills_only_confirmed_account_gap(self):
        days = {dt.date(2026, 9, i) for i in range(1, 31)}
        imported = {
            "Card": set(days),
            "QNB account": {d for d in days if d.day <= 13},
            "Ziraat account": {d for d in days if d.day <= 15},
        }
        original = {k: set(v) for k, v in imported.items()}
        result = confirmed_coverage(imported, [
            {"card": "Card", "kind": "Card"},
            {"card": "QNB account", "kind": "Account"},
            {"card": "Ziraat account", "kind": "Account"},
        ], [{"source": "QNB account", "start": "2026-09-11", "end": "2026-10-05"}])
        self.assertEqual(len(set.intersection(*imported.values())), 13)
        self.assertEqual(len(set.intersection(*result.values())), 15)
        self.assertTrue(days <= result["QNB account"])
        self.assertIn(dt.date(2026, 10, 5), result["QNB account"])
        self.assertNotIn(dt.date(2026, 10, 6), result["QNB account"])
        self.assertEqual(imported, original)

    def test_no_confirmation_keeps_coverage(self):
        imported = {"Account": {dt.date(2026, 9, 1)}}
        self.assertEqual(confirmed_coverage(imported, [], []), imported)

    def test_excluded_account_does_not_block_but_other_gaps_do(self):
        day = dt.date(2026, 9, 30)
        known = {"Card": {day}, "QNB": {day}, "Ziraat": set()}
        first = {name: dt.date(2026, 9, 1) for name in known}
        self.assertFalse(coverage_complete(day, known, first))
        self.assertTrue(coverage_complete(day, known, first, {"Ziraat"}))
        known["QNB"].clear()
        self.assertFalse(coverage_complete(day, known, first, {"Ziraat"}))
        self.assertFalse(coverage_complete(day, known, first, set(known)))

    def test_rejects_invalid_ranges_and_non_account_sources(self):
        statements = [{"card": "Account", "kind": "Account"}]
        for source, start, end in [
            ("Unknown", "2026-09-01", "2026-09-02"),
            ("Card", "2026-09-01", "2026-09-02"),
            ("Account", "2026-09-02", "2026-09-01"),
            ("Account", "", "2026-09-01"),
        ]:
            with self.subTest(source=source, start=start), self.assertRaises(ValueError):
                confirmed_coverage({"Account": set(), "Card": set()}, statements,
                                   [{"source": source, "start": start, "end": end}])


if __name__ == "__main__":
    unittest.main()
