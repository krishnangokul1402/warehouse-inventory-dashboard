"""Run with: python -m unittest test_agent.py"""
import unittest
from datetime import date
from pathlib import Path
from inventory_agent import load_inventory, metrics, active_shipments, stock_summary, answer_question


class AgentTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.inv = load_inventory(Path(__file__).resolve().parent / "Warehouse_Inventory_Sample.xlsx", as_of=date(2026,11,1))

    def test_total_units(self):
        actual = metrics(self.inv)
        self.assertEqual(actual["on_hand"], 1112)
        self.assertEqual(actual["reserved"], 219)
        self.assertEqual(actual["available"], 893)
        self.assertEqual(actual["plant"], 460)
        self.assertEqual(actual["transit"], 455)
        self.assertEqual(actual["incoming_7d"], 345)
        self.assertEqual(actual["projected_7d"], 1238)
        self.assertEqual(actual["delayed_count"], 2)

    def test_delayed_identification(self):
        late = active_shipments(self.inv)
        self.assertEqual(set(late.loc[late["Delayed"], "Shipment_ID"]), {"SHP-002", "SHP-005"})

    def test_chennai_filter(self):
        m = metrics(self.inv, warehouse_id="WH-CHN")
        self.assertEqual(m["available"], 365)
        self.assertEqual(m["transit"], 130)
        self.assertEqual(m["delayed_count"], 1)

    def test_questions(self):
        ans = answer_question("Which shipments are delayed in Chennai?", self.inv)
        self.assertIn("SHP-002", ans)
        self.assertNotIn("SHP-005", ans)
        ans = answer_question("ETA for SKU-101", self.inv)
        self.assertIn("SHP-001", ans)
        self.assertNotIn("SHP-002", ans)


if __name__ == "__main__":
    unittest.main()
