import unittest
import xlsform_to_registro
from xlsform_to_registro import XLSFormToRegistro
from xlsform_to_registro.core.parser import XLSFormParser
from xlsform_to_registro.exporters.excel_styler import ExcelStyler


class TestPackage(unittest.TestCase):
    def test_version(self):
        self.assertEqual(xlsform_to_registro.__version__, "1.0.0")

    def test_exports(self):
        self.assertTrue(hasattr(xlsform_to_registro, "XLSFormToRegistro"))
        self.assertIsNotNone(XLSFormToRegistro)
        self.assertIsNotNone(XLSFormParser)
        self.assertIsNotNone(ExcelStyler)


if __name__ == "__main__":
    unittest.main()
