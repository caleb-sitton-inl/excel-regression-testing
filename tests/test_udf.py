"""Tests user-defined functions in an excel workbook"""
import os

import xlwings as xw

tests_dir_path = os.path.dirname(__file__)
test_xl_path = os.path.join(os.path.dirname(tests_dir_path), 'workbooks', 'with_udf.xlsm')

def test_udf():
  with xw.Book(test_xl_path) as workbook:
    linear_1_func = workbook.macro("LINEAR_1")

    # Test that it works with non-negative inputs
    assert linear_1_func(0) == 1
    assert linear_1_func(2) == 5

    # Test that it fails with a negative input
    test_sheet = workbook.sheets.add("test_sheet")
    test_cell = test_sheet["A1"]
    test_cell.formula = "=LINEAR_1(-1)"
    assert test_cell.value is None

if __name__ == "__main__":
  test_udf()
