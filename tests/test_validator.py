"""Smoke/regression tests: run with `python -m unittest discover -s tests`."""
import os
import sys
import unittest

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)

from bpmn_qa_validator import BpmnQAValidator  # noqa: E402

EXAMPLE = os.path.join(ROOT, "examples", "proceso_con_errores.bpmn")


class ValidatorTests(unittest.TestCase):
    def setUp(self):
        self.findings = BpmnQAValidator(EXAMPLE).run()
        self.rules = {h.regla_id for h in self.findings}

    def test_detects_empty_lane(self):
        self.assertIn("R01", self.rules)

    def test_detects_untyped_task(self):
        self.assertIn("R02", self.rules)

    def test_detects_disconnected_node(self):
        self.assertIn("R07", self.rules)

    def test_detects_forbidden_complex_gateway(self):
        self.assertIn("R22", self.rules)
        self.assertIn("R23", self.rules)

    def test_detects_missing_diagram_attributes(self):
        for rule in ("R11", "R12", "R13", "R15"):
            self.assertIn(rule, self.rules)

    def test_digitalisation_index_is_informative(self):
        info = [h for h in self.findings if h.regla_id == "R29"]
        self.assertEqual(len(info), 1)
        self.assertEqual(info[0].severidad, "INFO")


if __name__ == "__main__":
    unittest.main()
