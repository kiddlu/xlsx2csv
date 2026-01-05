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

if __name__ == '__main__':
    create_basic_test()
    create_special_chars_test()
    create_empty_test()
    create_multisheet_test()
    create_number_formats_test()
    print("All test files created successfully")

