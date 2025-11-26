"""Tests sample excel workbook"""
import os

import xlwings as xw

tests_dir_path = os.path.dirname(__file__)
test_xl_path = os.path.join(os.path.dirname(tests_dir_path), 'workbooks', 'with_data_validation.xlsx')

def test_data_validation():
  with xw.Book(test_xl_path) as workbook:
    sheet = workbook.sheets[0]
    input_cell = sheet["A1"]
    input_cell.value = 120  # No timeout or error even though validation restricts input to < 100
    assert True


if __name__ == "__main__":
  test_data_validation()
