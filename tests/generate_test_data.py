#!/usr/bin/env python3
"""
Generate test XLSX files for xlsx2csv testing
"""
import sys
import openpyxl
from datetime import datetime, date

def create_basic_test():
    """Create basic test file with various data types"""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "BasicTest"
    
    # Headers
    ws['A1'] = "String"
    ws['B1'] = "Number"
    ws['C1'] = "Float"
    ws['D1'] = "Boolean"
    ws['E1'] = "Date"
    
    # Data rows
    ws['A2'] = "Hello"
    ws['B2'] = 123
    ws['C2'] = 45.67
    ws['D2'] = True
    ws['E2'] = date(2024, 1, 15)
    
    ws['A3'] = "World"
    ws['B3'] = 456
    ws['C3'] = 89.01
    ws['D3'] = False
    ws['E3'] = date(2024, 2, 20)
    
    wb.save('test_data/basic.xlsx')
    print("Created basic.xlsx")

def create_special_chars_test():
    """Create test file with special characters"""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "SpecialChars"
    
    ws['A1'] = "Field with, comma"
    ws['A2'] = 'Field with "quotes"'
    ws['A3'] = "Field with\nnewline"
    ws['A4'] = "Field with\ttab"
    ws['A5'] = "Normal field"
    
    wb.save('test_data/special_chars.xlsx')
    print("Created special_chars.xlsx")

def create_empty_test():
    """Create test file with empty cells and rows"""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "EmptyTest"
    
    ws['A1'] = "Data"
    ws['A2'] = ""
    ws['A3'] = "More"
    ws['A5'] = "Skip row 4"
    
    wb.save('test_data/empty.xlsx')
    print("Created empty.xlsx")

def create_multisheet_test():
    """Create test file with multiple sheets"""
    wb = openpyxl.Workbook()
    
    # Sheet 1
    ws1 = wb.active
    ws1.title = "Sheet1"
    ws1['A1'] = "First Sheet"
    ws1['A2'] = "Data 1"
    
    # Sheet 2
    ws2 = wb.create_sheet("Sheet2")
    ws2['A1'] = "Second Sheet"
    ws2['A2'] = "Data 2"
    
    # Sheet 3
    ws3 = wb.create_sheet("Sheet3")
    ws3['A1'] = "Third Sheet"
    ws3['A2'] = "Data 3"
    
    wb.save('test_data/multisheet.xlsx')
    print("Created multisheet.xlsx")

def create_number_formats_test():
    """Create test file with various number formats"""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Numbers"
    
    ws['A1'] = "Integer"
    ws['A2'] = 100
    ws['A3'] = -50
    
    ws['B1'] = "Float"
    ws['B2'] = 3.14159
    ws['B3'] = 2.71828
    
    ws['C1'] = "Scientific"
    ws['C2'] = 1.23e10
    ws['C3'] = 4.56e-5
    
    wb.save('test_data/numbers.xlsx')
    print("Created numbers.xlsx")

def create_float_format_test():
    """Create comprehensive test file for --floatformat and --sci-float options"""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "FloatFormat"
    
    # Headers
    ws['A1'] = "Description"
    ws['B1'] = "Value"
    
    # Various float scenarios
    # Row 2: Small positive decimal
    ws['A2'] = "Small positive"
    ws['B2'] = 0.001234
    
    # Row 3: Large positive number
    ws['A3'] = "Large positive"
    ws['B3'] = 1234567.89
    
    # Row 4: Small negative decimal
    ws['A4'] = "Small negative"
    ws['B4'] = -0.000456
    
    # Row 5: Large negative number
    ws['A5'] = "Large negative"
    ws['B5'] = -9876543.21
    
    # Row 6: Very small number (scientific notation range)
    ws['A6'] = "Very small"
    ws['B6'] = 1.23e-10
    
    # Row 7: Very large number (scientific notation range)
    ws['A7'] = "Very large"
    ws['B7'] = 9.87e15
    
    # Row 8: Exact decimal (no trailing zeros)
    ws['A8'] = "Exact decimal"
    ws['B8'] = 89.01
    
    # Row 9: Many decimal places
    ws['A9'] = "Many decimals"
    ws['B9'] = 3.141592653589793
    
    # Row 10: Integer as float
    ws['A10'] = "Integer float"
    ws['B10'] = 100.0
    
    # Row 11: Zero
    ws['A11'] = "Zero"
    ws['B11'] = 0.0
    
    # Row 12: Negative zero
    ws['A12'] = "Negative zero"
    ws['B12'] = -0.0
    
    # Row 13: Medium precision
    ws['A13'] = "Medium precision"
    ws['B13'] = 1234.567
    
    # Row 14: Edge case - just below 1
    ws['A14'] = "Just below 1"
    ws['B14'] = 0.999999
    
    # Row 15: Edge case - just above 1
    ws['A15'] = "Just above 1"
    ws['B15'] = 1.000001
    
    wb.save('test_data/float_format.xlsx')
    print("Created float_format.xlsx")

if __name__ == '__main__':
    create_basic_test()
    create_special_chars_test()
    create_empty_test()
    create_multisheet_test()
    create_number_formats_test()
    create_float_format_test()
    print("All test files created successfully")

