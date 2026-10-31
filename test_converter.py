import tempfile
import unittest
from pathlib import Path

from openpyxl import load_workbook

from csv_to_excel import detect_delimiter, read_csv_as_text, write_workbook


class ConverterTests(unittest.TestCase):
    def test_delimiters(self):
        self.assertEqual(detect_delimiter("a,b\n1,2"), ",")
        self.assertEqual(detect_delimiter("a;b\n1;2"), ";")
        self.assertEqual(detect_delimiter("a\tb\n1\t2"), "\t")

    def test_values_are_written_as_text(self):
        with tempfile.TemporaryDirectory() as tmp:
            source = Path(tmp) / "sample.csv"
            output = Path(tmp) / "sample.xlsx"
            source.write_text('00123,=1+1,café\n"two, parts",,東京\n', encoding="utf-8")
            write_workbook([source], output)
            sheet = load_workbook(output, data_only=False).active
            self.assertEqual(sheet["A1"].value, "00123")
            self.assertEqual(sheet["A1"].data_type, "s")
            self.assertEqual(sheet["B1"].value, "=1+1")
            self.assertEqual(sheet["B1"].data_type, "s")
            self.assertEqual(sheet["C1"].value, "café")
            self.assertEqual(sheet["A2"].value, "two, parts")
            self.assertEqual(sheet["C2"].value, "東京")


if __name__ == "__main__":
    unittest.main()
