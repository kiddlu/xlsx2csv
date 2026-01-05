#!/usr/bin/env python3
"""
Generate comprehensive test data for xlsx2csv
This creates Excel files with various edge cases to ensure compatibility
between C and Python versions.
"""

import openpyxl
from openpyxl import Workbook
from openpyxl.styles import numbers
from datetime import datetime, date, time
import os
import math

def create_date_time_test():
    """Test various date and time formats"""
    wb = Workbook()
    ws = wb.active
    ws.title = "Dates and Times"
    
    # Headers
    ws['A1'] = 'Description'
    ws['B1'] = 'Value'
    
    # Date values (Excel stores dates as numbers)
    ws['A2'] = 'Date 2020-01-01'
    ws['B2'] = date(2020, 1, 1)
    
    ws['A3'] = 'Date 1900-01-01'
    ws['B3'] = date(1900, 1, 1)
    
    ws['A4'] = 'Date 2024-12-31'
    ws['B4'] = date(2024, 12, 31)
    
    ws['A5'] = 'DateTime 2020-06-15 14:30:00'
    ws['B5'] = datetime(2020, 6, 15, 14, 30, 0)
    
    ws['A6'] = 'Time 12:30:45'
    ws['B6'] = time(12, 30, 45)
    
    ws['A7'] = 'Time 00:00:00'
    ws['B7'] = time(0, 0, 0)
    
    ws['A8'] = 'Time 23:59:59'
    ws['B8'] = time(23, 59, 59)
    
    # Apply date/time number formats
    ws['B2'].number_format = 'YYYY-MM-DD'
    ws['B3'].number_format = 'YYYY-MM-DD'
    ws['B4'].number_format = 'YYYY-MM-DD'
    ws['B5'].number_format = 'YYYY-MM-DD HH:MM:SS'
    ws['B6'].number_format = 'HH:MM:SS'
    ws['B7'].number_format = 'HH:MM:SS'
    ws['B8'].number_format = 'HH:MM:SS'
    
    wb.save('test_data/date_time.xlsx')
    print("Created: date_time.xlsx")

def create_boolean_test():
    """Test boolean values"""
    wb = Workbook()
    ws = wb.active
    ws.title = "Booleans"
    
    ws['A1'] = 'Description'
    ws['B1'] = 'Value'
    
    ws['A2'] = 'TRUE'
    ws['B2'] = True
    
    ws['A3'] = 'FALSE'
    ws['B3'] = False
    
    ws['A4'] = 'Text TRUE'
    ws['B4'] = 'TRUE'
    
    ws['A5'] = 'Text FALSE'
    ws['B5'] = 'FALSE'
    
    ws['A6'] = 'Number 1'
    ws['B6'] = 1
    
    ws['A7'] = 'Number 0'
    ws['B7'] = 0
    
    wb.save('test_data/boolean.xlsx')
    print("Created: boolean.xlsx")

def create_percentage_test():
    """Test percentage formatting"""
    wb = Workbook()
    ws = wb.active
    ws.title = "Percentages"
    
    ws['A1'] = 'Description'
    ws['B1'] = 'Value'
    
    ws['A2'] = '50%'
    ws['B2'] = 0.5
    ws['B2'].number_format = '0%'
    
    ws['A3'] = '100%'
    ws['B3'] = 1.0
    ws['B3'].number_format = '0%'
    
    ws['A4'] = '0%'
    ws['B4'] = 0.0
    ws['B4'].number_format = '0%'
    
    ws['A5'] = '12.5%'
    ws['B5'] = 0.125
    ws['B5'].number_format = '0.0%'
    
    ws['A6'] = '123.45%'
    ws['B6'] = 1.2345
    ws['B6'].number_format = '0.00%'
    
    ws['A7'] = '-25%'
    ws['B7'] = -0.25
    ws['B7'].number_format = '0%'
    
    wb.save('test_data/percentage.xlsx')
    print("Created: percentage.xlsx")

def create_formula_test():
    """Test formulas (should show calculated values)"""
    wb = Workbook()
    ws = wb.active
    ws.title = "Formulas"
    
    ws['A1'] = 'Description'
    ws['B1'] = 'Formula'
    ws['C1'] = 'Result'
    
    ws['A2'] = 'Addition'
    ws['B2'] = '=2+3'
    ws['C2'] = 5
    
    ws['A3'] = 'Multiplication'
    ws['B3'] = '=4*5'
    ws['C3'] = 20
    
    ws['A4'] = 'Division'
    ws['B4'] = '=10/2'
    ws['C4'] = 5
    
    ws['A5'] = 'Sum'
    ws['B5'] = '=SUM(1,2,3,4,5)'
    ws['C5'] = 15
    
    ws['A6'] = 'Concatenation'
    ws['B6'] = '="Hello"&" "&"World"'
    ws['C6'] = 'Hello World'
    
    wb.save('test_data/formulas.xlsx')
    print("Created: formulas.xlsx")

def create_extreme_numbers_test():
    """Test extreme numeric values"""
    wb = Workbook()
    ws = wb.active
    ws.title = "Extreme Numbers"
    
    ws['A1'] = 'Description'
    ws['B1'] = 'Value'
    
    ws['A2'] = 'Very large positive'
    ws['B2'] = 1.7976931348623157e+308  # Near max double
    
    ws['A3'] = 'Very large negative'
    ws['B3'] = -1.7976931348623157e+308
    
    ws['A4'] = 'Very small positive'
    ws['B4'] = 2.2250738585072014e-308
    
    ws['A5'] = 'Very small negative'
    ws['B5'] = -2.2250738585072014e-308
    
    ws['A6'] = 'Tiny positive'
    ws['B6'] = 1e-100
    
    ws['A7'] = 'Huge integer'
    ws['B7'] = 999999999999999
    
    ws['A8'] = 'Precise decimal'
    ws['B8'] = 0.123456789012345
    
    ws['A9'] = 'Pi'
    ws['B9'] = 3.141592653589793
    
    ws['A10'] = 'E (Euler)'
    ws['B10'] = 2.718281828459045
    
    wb.save('test_data/extreme_numbers.xlsx')
    print("Created: extreme_numbers.xlsx")

def create_mixed_empty_test():
    """Test mixed content with empty cells"""
    wb = Workbook()
    ws = wb.active
    ws.title = "Mixed Empty"
    
    ws['A1'] = 'Col1'
    ws['B1'] = 'Col2'
    ws['C1'] = 'Col3'
    ws['D1'] = 'Col4'
    
    ws['A2'] = 'A'
    # B2 empty
    ws['C2'] = 'C'
    ws['D2'] = 'D'
    
    ws['A3'] = 1
    ws['B3'] = 2
    # C3 empty
    ws['D3'] = 4
    
    # Row 4 all empty
    
    ws['A5'] = ''  # Empty string
    ws['B5'] = 0
    ws['C5'] = False
    ws['D5'] = 'Text'
    
    ws['A6'] = 'Last'
    # B6, C6, D6 all empty
    
    wb.save('test_data/mixed_empty.xlsx')
    print("Created: mixed_empty.xlsx")

def create_unicode_test():
    """Test extended Unicode characters"""
    wb = Workbook()
    ws = wb.active
    ws.title = "Unicode Extended"
    
    ws['A1'] = 'Description'
    ws['B1'] = 'Value'
    
    ws['A2'] = 'Emoji smile'
    ws['B2'] = '😀😃😄😁'
    
    ws['A3'] = 'Emoji various'
    ws['B3'] = '🎉🎊🎈🎁'
    
    ws['A4'] = 'Chinese'
    ws['B4'] = '中文测试数据'
    
    ws['A5'] = 'Japanese'
    ws['B5'] = 'テストデータ'
    
    ws['A6'] = 'Korean'
    ws['B6'] = '테스트 데이터'
    
    ws['A7'] = 'Arabic'
    ws['B7'] = 'بيانات الاختبار'
    
    ws['A8'] = 'Hebrew'
    ws['B8'] = 'נתוני בדיקה'
    
    ws['A9'] = 'Greek'
    ws['B9'] = 'δοκιμαστικά δεδομένα'
    
    ws['A10'] = 'Cyrillic'
    ws['B10'] = 'тестовые данные'
    
    ws['A11'] = 'Math symbols'
    ws['B11'] = '∑∫∂√∞≠≈≤≥'
    
    ws['A12'] = 'Currency symbols'
    ws['B12'] = '€¥£₹₽₩¢'
    
    wb.save('test_data/unicode_extended.xlsx')
    print("Created: unicode_extended.xlsx")

def create_long_strings_test():
    """Test very long strings"""
    wb = Workbook()
    ws = wb.active
    ws.title = "Long Strings"
    
    ws['A1'] = 'Description'
    ws['B1'] = 'Value'
    
    ws['A2'] = 'Short'
    ws['B2'] = 'ABC'
    
    ws['A3'] = '100 chars'
    ws['B3'] = 'A' * 100
    
    ws['A4'] = '1000 chars'
    ws['B4'] = 'B' * 1000
    
    ws['A5'] = '5000 chars'
    ws['B5'] = 'C' * 5000
    
    ws['A6'] = 'With newlines'
    ws['B6'] = 'Line1\nLine2\nLine3\nLine4\nLine5'
    
    ws['A7'] = 'With tabs'
    ws['B7'] = 'Col1\tCol2\tCol3\tCol4'
    
    ws['A8'] = 'Mixed special'
    ws['B8'] = 'Quote"Comma,Tab\tNewline\n'
    
    wb.save('test_data/long_strings.xlsx')
    print("Created: long_strings.xlsx")

def create_escaping_test():
    """Test CSV escaping scenarios"""
    wb = Workbook()
    ws = wb.active
    ws.title = "CSV Escaping"
    
    ws['A1'] = 'Description'
    ws['B1'] = 'Value'
    
    ws['A2'] = 'Simple comma'
    ws['B2'] = 'Hello, World'
    
    ws['A3'] = 'Multiple commas'
    ws['B3'] = 'One, Two, Three, Four'
    
    ws['A4'] = 'Double quotes'
    ws['B4'] = 'He said "Hello"'
    
    ws['A5'] = 'Quote at start'
    ws['B5'] = '"Important"'
    
    ws['A6'] = 'Quote at end'
    ws['B6'] = 'Important"'
    
    ws['A7'] = 'Multiple quotes'
    ws['B7'] = '"Hello" "World" "Test"'
    
    ws['A8'] = 'Comma and quote'
    ws['B8'] = 'Say "Hello, Friend"'
    
    ws['A9'] = 'Leading spaces'
    ws['B9'] = '   Leading'
    
    ws['A10'] = 'Trailing spaces'
    ws['B10'] = 'Trailing   '
    
    ws['A11'] = 'Both spaces'
    ws['B11'] = '  Both  '
    
    ws['A12'] = 'Just spaces'
    ws['B12'] = '     '
    
    wb.save('test_data/escaping.xlsx')
    print("Created: escaping.xlsx")

def create_multisheet_complex_test():
    """Test complex multi-sheet scenarios"""
    wb = Workbook()
    
    # Sheet 1: Data
    ws1 = wb.active
    ws1.title = "Data"
    ws1['A1'] = 'Name'
    ws1['B1'] = 'Value'
    ws1['A2'] = 'Alpha'
    ws1['B2'] = 100
    ws1['A3'] = 'Beta'
    ws1['B3'] = 200
    
    # Sheet 2: Empty sheet
    ws2 = wb.create_sheet("Empty")
    
    # Sheet 3: Single cell
    ws3 = wb.create_sheet("Single")
    ws3['A1'] = 'OnlyOne'
    
    # Sheet 4: Special name with spaces
    ws4 = wb.create_sheet("Sheet With Spaces")
    ws4['A1'] = 'Test'
    ws4['B1'] = 'Data'
    
    # Sheet 5: Numbers in name
    ws5 = wb.create_sheet("Sheet123")
    ws5['A1'] = '123'
    
    wb.save('test_data/multisheet_complex.xlsx')
    print("Created: multisheet_complex.xlsx")

def create_number_formats_test():
    """Test various number format codes"""
    wb = Workbook()
    ws = wb.active
    ws.title = "Number Formats"
    
    ws['A1'] = 'Format'
    ws['B1'] = 'Value'
    
    # General format
    ws['A2'] = 'General'
    ws['B2'] = 12345.6789
    ws['B2'].number_format = 'General'
    
    # Integer format
    ws['A3'] = '0'
    ws['B3'] = 12345.6789
    ws['B3'].number_format = '0'
    
    # Two decimal places
    ws['A4'] = '0.00'
    ws['B4'] = 12345.6789
    ws['B4'].number_format = '0.00'
    
    # Thousands separator
    ws['A5'] = '#,##0'
    ws['B5'] = 12345.6789
    ws['B5'].number_format = '#,##0'
    
    # Thousands with decimals
    ws['A6'] = '#,##0.00'
    ws['B6'] = 12345.6789
    ws['B6'].number_format = '#,##0.00'
    
    # Scientific notation
    ws['A7'] = '0.00E+00'
    ws['B7'] = 12345.6789
    ws['B7'].number_format = '0.00E+00'
    
    # Accounting format
    ws['A8'] = '#,##0.00;(#,##0.00)'
    ws['B8'] = -1234.56
    ws['B8'].number_format = '#,##0.00;(#,##0.00)'
    
    wb.save('test_data/number_formats.xlsx')
    print("Created: number_formats.xlsx")

def main():
    """Generate all test files"""
    os.makedirs('test_data', exist_ok=True)
    
    print("Generating comprehensive test data...")
    print()
    
    create_date_time_test()
    create_boolean_test()
    create_percentage_test()
    create_formula_test()
    create_extreme_numbers_test()
    create_mixed_empty_test()
    create_unicode_test()
    create_long_strings_test()
    create_escaping_test()
    create_multisheet_complex_test()
    create_number_formats_test()
    
    print()
    print("All test files generated successfully!")
    print("Run './test_runner.sh' to execute tests")

if __name__ == '__main__':
    main()
