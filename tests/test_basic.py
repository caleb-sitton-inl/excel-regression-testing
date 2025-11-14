"""Tests sample excel workbook"""
import os

import xlwings as xw

tests_dir_path = os.path.dirname(__file__)
test_xl_path = os.path.join(os.path.dirname(tests_dir_path), 'workbooks', 'basic_workbook.xlsx')

def test_basic():
  with xw.Book(test_xl_path) as workbook:
    sheet = workbook.sheets[0]

    # Set input
    input_cell = sheet["A2"]
    input_cell.value = 2.0

    # Check output
    output_cell = sheet["B2"]
    assert output_cell.value == 4.0


if __name__ == "__main__":
  test_basic()
