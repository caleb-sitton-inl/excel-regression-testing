"""Tests simple subprocess macros in an excel workbook"""
import os

import xlwings as xw

tests_dir_path = os.path.dirname(__file__)
test_xl_path = os.path.join(os.path.dirname(tests_dir_path), 'workbooks', 'with_sub.xlsm')

def test_sub():
  with xw.Book(test_xl_path) as workbook:
    sample_macro = workbook.macro("sample_macro")
    sample_macro()

    sheet = workbook.sheets[0]
    output_cell = sheet["B2"]
    assert output_cell.value == 2.0

    undo_sample = workbook.macro("undo_sample")
    undo_sample()

    assert output_cell.value == 0.0
    assert output_cell.formula != ''


if __name__ == "__main__":
  test_sub()
