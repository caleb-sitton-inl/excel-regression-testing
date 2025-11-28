"""Tests excel workbook with a named range"""
import os

import xlwings as xw

tests_dir_path = os.path.dirname(__file__)
test_xl_path = os.path.join(os.path.dirname(tests_dir_path), 'workbooks', 'with_named_range.xlsx')

def test_named_range():
  with xw.Book(test_xl_path) as workbook:
    sheet = workbook.sheets[0]

    # Set input
    input_cell = sheet["A1"]
    input_cell.value = 2.0

    # Check output
    named_range = workbook.names["results"].refers_to_range
    for i, cell in enumerate(named_range):
      assert cell.value == 2.0 * (i + 1)

if __name__ == "__main__":
  test_named_range()
