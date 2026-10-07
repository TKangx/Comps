"""Companion to CompsModel.xlsx.

xlwings' "Run main" ribbon button runs main() in the Python file that has the same
name as the workbook, so keep this file next to CompsModel.xlsx.
"""
import xlwings as xw

from comps.xl import refresh


def main():
    refresh(xw.Book.caller())


if __name__ == "__main__":
    # Debug from an editor: open CompsModel.xlsx in Excel, then run this file.
    xw.Book("CompsModel.xlsx").set_mock_caller()
    main()
