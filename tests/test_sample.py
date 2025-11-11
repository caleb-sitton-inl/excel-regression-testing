"""Tests sample excel workbook"""
import os

import xlwings as xw

tests_dir_path = os.path.dirname(__file__)
test_xl_path = os.path.join(os.path.dirname(tests_dir_path), 'workbooks', 'sample_workbook.xlsx')

def test_sample():
  with xw.Book(test_xl_path) as workbook:
    sheet = workbook.sheets[0]

    # Set input
    input_cell = sheet["A2"]
    input_cell.value = 1.0

    # Check output
    output_cell = sheet["B2"]
    assert output_cell.value == 5.0
