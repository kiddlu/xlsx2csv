#!/bin/bash
# Test runner for xlsx2csv C version
# Compares output with Python version
# 
# IMPORTANT: This script ALWAYS runs the Python version first to generate
# the expected output, ensuring we test against the actual Python behavior

set -e

# Colors
RED='\033[0;31m'
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
NC='\033[0m' # No Color

PYTHON_XLSX2CSV="../xlsx2csv_python.py"
C_XLSX2CSV="../xlsx2csv"

# Also check build directory
if [ ! -f "$C_XLSX2CSV" ] && [ -f "../build/xlsx2csv" ]; then
    C_XLSX2CSV="../build/xlsx2csv"
fi

if [ ! -f "$C_XLSX2CSV" ]; then
    echo -e "${RED}Error: C xlsx2csv not found at $C_XLSX2CSV${NC}"
    echo "Please run 'make' first"
    exit 1
fi

if [ ! -f "$PYTHON_XLSX2CSV" ]; then
    echo -e "${RED}Error: Python xlsx2csv not found at $PYTHON_XLSX2CSV${NC}"
    echo "Please ensure xlsx2csv_python.py is in the project root directory"
    exit 1
fi

# Create output directories (do NOT create expected - it's generated each time)
mkdir -p actual

echo "====================================="
echo "xlsx2csv Test Suite"
echo "Testing against Python xlsx2csv ($PYTHON_XLSX2CSV)"
echo "====================================="
echo

TESTS_PASSED=0
TESTS_FAILED=0

# Function to run a test
run_test() {
    local test_name="$1"
    local xlsx_file="$2"
    local options="$3"
    
    echo -n "Testing $test_name... "
    
    # ALWAYS run Python version first to generate expected output
    $PYTHON_XLSX2CSV $options "$xlsx_file" "/tmp/expected_${test_name}.csv" 2>/dev/null || {
        echo -e "${YELLOW}SKIP${NC} (Python version failed)"
        return
    }
    
    # Run C version
    $C_XLSX2CSV $options "$xlsx_file" "actual/${test_name}.csv" 2>/dev/null || {
        echo -e "${RED}FAIL${NC} (C version crashed)"
        TESTS_FAILED=$((TESTS_FAILED + 1))
        rm -f "/tmp/expected_${test_name}.csv"
        return
    }
    
    # Compare outputs
    if diff -q "/tmp/expected_${test_name}.csv" "actual/${test_name}.csv" >/dev/null 2>&1; then
        echo -e "${GREEN}PASS${NC}"
        TESTS_PASSED=$((TESTS_PASSED + 1))
        rm -f "/tmp/expected_${test_name}.csv"
    else
        echo -e "${RED}FAIL${NC}"
        echo "  Output differs from Python version"
        echo "  Run: diff /tmp/expected_${test_name}.csv actual/${test_name}.csv"
        TESTS_FAILED=$((TESTS_FAILED + 1))
        # Keep the expected file for debugging
    fi
}

# Function to run a test that expects error (exit code 1)
run_error_test() {
    local test_name="$1"
    local xlsx_file="$2"
    local options="$3"
    
    echo -n "Testing $test_name... "
    
    # Temporarily disable exit on error for these commands
    set +e
    
    # Run Python version (expecting it to fail with exit code 1)
    # Note: options are passed as separate arguments, not as a single string
    $PYTHON_XLSX2CSV $options "$xlsx_file" > "/tmp/expected_${test_name}_stdout.txt" 2> "/tmp/expected_${test_name}_stderr.txt"
    local python_exit=$?
    
    # Run C version (expecting it to fail with exit code 1)
    $C_XLSX2CSV $options "$xlsx_file" > "actual/${test_name}_stdout.txt" 2> "actual/${test_name}_stderr.txt"
    local c_exit=$?
    
    # Re-enable exit on error
    set -e
    
    # Check if both have the same exit code
    if [ $python_exit -ne $c_exit ]; then
        echo -e "${RED}FAIL${NC}"
        echo "  Exit codes differ: Python=$python_exit, C=$c_exit"
        TESTS_FAILED=$((TESTS_FAILED + 1))
        return
    fi
    
    # Compare stdout
    if ! diff -q "/tmp/expected_${test_name}_stdout.txt" "actual/${test_name}_stdout.txt" >/dev/null 2>&1; then
        echo -e "${RED}FAIL${NC}"
        echo "  stdout differs from Python version"
        echo "  Run: diff /tmp/expected_${test_name}_stdout.txt actual/${test_name}_stdout.txt"
        TESTS_FAILED=$((TESTS_FAILED + 1))
        return
    fi
    
    # Compare stderr
    if ! diff -q "/tmp/expected_${test_name}_stderr.txt" "actual/${test_name}_stderr.txt" >/dev/null 2>&1; then
        echo -e "${RED}FAIL${NC}"
        echo "  stderr differs from Python version"
        echo "  Run: diff /tmp/expected_${test_name}_stderr.txt actual/${test_name}_stderr.txt"
        TESTS_FAILED=$((TESTS_FAILED + 1))
        return
    fi
    
    echo -e "${GREEN}PASS${NC}"
    TESTS_PASSED=$((TESTS_PASSED + 1))
    rm -f "/tmp/expected_${test_name}_stdout.txt" "/tmp/expected_${test_name}_stderr.txt"
}

# Basic tests
echo "=== Basic Functionality Tests ==="
run_test "basic" "test_data/basic.xlsx" ""
run_test "special_chars" "test_data/special_chars.xlsx" ""
run_test "empty" "test_data/empty.xlsx" ""
run_test "numbers" "test_data/numbers.xlsx" ""

# Delimiter tests
echo -e "\n=== Delimiter Tests ==="
run_test "basic_tab_delim" "test_data/basic.xlsx" "-d tab"
run_test "basic_semicolon" "test_data/basic.xlsx" "-d ';'"

# Quoting mode tests
echo -e "\n=== Quoting Mode Tests ==="
run_test "basic_quote_all" "test_data/basic.xlsx" "-q all"
run_test "basic_quote_none" "test_data/basic.xlsx" "-q none"
run_test "basic_quote_nonnumeric" "test_data/basic.xlsx" "-q nonnumeric"
run_test "special_quote_all" "test_data/special_chars.xlsx" "-q all"

# Empty line handling
echo -e "\n=== Empty Line Tests ==="
run_test "empty_skip" "test_data/empty.xlsx" "-i"
run_test "empty_keep" "test_data/empty.xlsx" ""

# Multi-sheet tests
echo -e "\n=== Multi-Sheet Tests ==="
run_test "multisheet_sheet1" "test_data/multisheet.xlsx" "-s 1"
run_test "multisheet_sheet2" "test_data/multisheet.xlsx" "-s 2"
run_test "multisheet_sheet3" "test_data/multisheet.xlsx" "-s 3"

# Line terminator tests
echo -e "\n=== Line Terminator Tests ==="
run_test "basic_lf" "test_data/basic.xlsx" "-l '\\n'"
run_test "basic_crlf" "test_data/basic.xlsx" "-l '\\r\\n'"

# UTF-8 test
if [ -f "test_data/utf8_test.xlsx" ]; then
    echo -e "\n=== UTF-8 Tests ==="
    run_test "utf8_test" "test_data/utf8_test.xlsx" ""
fi

# Float format tests (CRITICAL: Must match Python version exactly)
echo -e "\n=== Float Format Tests (CRITICAL) ==="
if [ -f "test_data/float_format.xlsx" ]; then
    # Test --floatformat %.02f (most common use case)
    run_test "float_format_02f" "test_data/float_format.xlsx" "--floatformat %.02f"
    
    # Test --floatformat with different precisions
    run_test "float_format_04f" "test_data/float_format.xlsx" "--floatformat %.04f"
    run_test "float_format_06f" "test_data/float_format.xlsx" "--floatformat %.06f"
    run_test "float_format_08f" "test_data/float_format.xlsx" "--floatformat %.08f"
    
    # Test --sci-float (scientific notation)
    run_test "float_sci_float" "test_data/float_format.xlsx" "--sci-float"
    
    # Test --sci-float with numbers.xlsx (has scientific notation values)
    run_test "numbers_sci_float" "test_data/numbers.xlsx" "--sci-float"
    
    # Test combination: --floatformat with --sci-float (should prioritize floatformat)
    run_test "float_format_02f_sci" "test_data/float_format.xlsx" "--floatformat %.02f --sci-float"
fi

# Test --floatformat with basic numbers file
echo -e "\n=== Float Format with Basic Numbers ==="
run_test "numbers_float_02f" "test_data/numbers.xlsx" "--floatformat %.02f"
run_test "numbers_float_04f" "test_data/numbers.xlsx" "--floatformat %.04f"

# Date and Time tests
if [ -f "test_data/date_time.xlsx" ]; then
    echo -e "\n=== Date and Time Tests ==="
    run_test "date_time_default" "test_data/date_time.xlsx" ""
    run_test "date_time_dateformat" "test_data/date_time.xlsx" "--dateformat %Y-%m-%d"
fi

# Boolean tests
if [ -f "test_data/boolean.xlsx" ]; then
    echo -e "\n=== Boolean Tests ==="
    run_test "boolean_default" "test_data/boolean.xlsx" ""
fi

# Percentage tests
if [ -f "test_data/percentage.xlsx" ]; then
    echo -e "\n=== Percentage Tests ==="
    run_test "percentage_default" "test_data/percentage.xlsx" ""
    run_test "percentage_float02f" "test_data/percentage.xlsx" "--floatformat %.02f"
fi

# Formula tests
if [ -f "test_data/formulas.xlsx" ]; then
    echo -e "\n=== Formula Tests ==="
    run_test "formulas_default" "test_data/formulas.xlsx" ""
fi

# Extreme numbers tests
if [ -f "test_data/extreme_numbers.xlsx" ]; then
    echo -e "\n=== Extreme Numbers Tests ==="
    run_test "extreme_numbers_default" "test_data/extreme_numbers.xlsx" ""
    run_test "extreme_numbers_float02f" "test_data/extreme_numbers.xlsx" "--floatformat %.02f"
    run_test "extreme_numbers_scifloat" "test_data/extreme_numbers.xlsx" "--sci-float"
fi

# Mixed empty cells tests
if [ -f "test_data/mixed_empty.xlsx" ]; then
    echo -e "\n=== Mixed Empty Cells Tests ==="
    run_test "mixed_empty_default" "test_data/mixed_empty.xlsx" ""
    run_test "mixed_empty_skip" "test_data/mixed_empty.xlsx" "-i"
fi

# Unicode extended tests
if [ -f "test_data/unicode_extended.xlsx" ]; then
    echo -e "\n=== Unicode Extended Tests ==="
    run_test "unicode_extended_default" "test_data/unicode_extended.xlsx" ""
    run_test "unicode_extended_quote_all" "test_data/unicode_extended.xlsx" "-q all"
fi

# Long strings tests
if [ -f "test_data/long_strings.xlsx" ]; then
    echo -e "\n=== Long Strings Tests ==="
    run_test "long_strings_default" "test_data/long_strings.xlsx" ""
    run_test "long_strings_quote_all" "test_data/long_strings.xlsx" "-q all"
fi

# CSV escaping tests
if [ -f "test_data/escaping.xlsx" ]; then
    echo -e "\n=== CSV Escaping Tests ==="
    run_test "escaping_minimal" "test_data/escaping.xlsx" "-q minimal"
    run_test "escaping_all" "test_data/escaping.xlsx" "-q all"
    run_test "escaping_nonnumeric" "test_data/escaping.xlsx" "-q nonnumeric"
    run_test "escaping_none" "test_data/escaping.xlsx" "-q none"
fi

# Multi-sheet complex tests
if [ -f "test_data/multisheet_complex.xlsx" ]; then
    echo -e "\n=== Multi-Sheet Complex Tests ==="
    run_test "multisheet_complex_all" "test_data/multisheet_complex.xlsx" ""
    run_test "multisheet_complex_sheet1" "test_data/multisheet_complex.xlsx" "-s 1"
    run_test "multisheet_complex_sheet2" "test_data/multisheet_complex.xlsx" "-s 2"
    run_test "multisheet_complex_sheet3" "test_data/multisheet_complex.xlsx" "-s 3"
    run_test "multisheet_complex_sheet4" "test_data/multisheet_complex.xlsx" "-s 4"
fi

# Number formats tests
if [ -f "test_data/number_formats.xlsx" ]; then
    echo -e "\n=== Number Formats Tests ==="
    run_test "number_formats_default" "test_data/number_formats.xlsx" ""
    run_test "number_formats_float04f" "test_data/number_formats.xlsx" "--floatformat %.04f"
fi

# Excel errors tests (tests error value handling like #VALUE!, #DIV/0!)
if [ -f "test_data/excel_errors.xlsx" ]; then
    echo -e "\n=== Excel Errors Tests ==="
    run_error_test "excel_errors_sheet1" "test_data/excel_errors.xlsx" "-s 1"
    run_error_test "excel_errors_sheet2" "test_data/excel_errors.xlsx" "-s 2"
fi

# Combination tests (stress testing)
echo -e "\n=== Combination Tests ==="
run_test "combo_tab_quote_all" "test_data/basic.xlsx" "-d tab -q all"
run_test "combo_semicolon_skip_empty" "test_data/empty.xlsx" "-d ';' -i"
if [ -f "test_data/float_format.xlsx" ]; then
    run_test "combo_float_tab_quote" "test_data/float_format.xlsx" "--floatformat %.03f -d tab -q nonnumeric"
fi

# Real-world scenario tests
echo -e "\n=== Real-World Scenarios ==="
if [ -f "test_data/stock_data_1107.xlsx" ]; then
    run_test "stock1107_s3" "test_data/stock_data_1107.xlsx" "-s 3 --floatformat %.02f --sci-float"
    run_test "stock1107_s1" "test_data/stock_data_1107.xlsx" "-s 1 --floatformat %.02f --sci-float"
    run_test "stock1107_s2" "test_data/stock_data_1107.xlsx" "-s 2 --floatformat %.02f --sci-float"
    run_test "stock1107_s4" "test_data/stock_data_1107.xlsx" "-s 4 --floatformat %.02f --sci-float"
fi

if [ -f "test_data/stock_data_1121.xlsx" ]; then
    run_test "stock1121" "test_data/stock_data_1121.xlsx" "--floatformat %.02f --sci-float"
fi

if [ -f "test_data/etfx_template.xlsx" ]; then
    run_test "etfx_default" "test_data/etfx_template.xlsx" ""
    run_test "etfx_float" "test_data/etfx_template.xlsx" "--floatformat %.02f"
fi

if [ -f "test_data/sector.xlsx" ]; then
    run_test "sector_default" "test_data/sector.xlsx" ""
    run_test "sector_float04f" "test_data/sector.xlsx" "--floatformat %.04f"
fi

if [ -f "test_data/financial_report.xlsx" ]; then
    run_test "financial" "test_data/financial_report.xlsx" "--floatformat %.02f"
fi

if [ -f "test_data/portfolio_tracking.xlsx" ]; then
    run_test "portfolio_s1" "test_data/portfolio_tracking.xlsx" "-s 1 --floatformat %.02f"
    run_test "portfolio_s2" "test_data/portfolio_tracking.xlsx" "-s 2 --floatformat %.02f"
    run_test "portfolio_holdings" "test_data/portfolio_tracking.xlsx" "-n Holdings --floatformat %.02f"
fi

echo
echo "====================================="
echo "Test Results"
echo "====================================="
echo -e "Tests passed: ${GREEN}$TESTS_PASSED${NC}"
echo -e "Tests failed: ${RED}$TESTS_FAILED${NC}"
echo

if [ $TESTS_FAILED -eq 0 ]; then
    echo -e "${GREEN}All tests passed!${NC}"
    exit 0
else
    echo -e "${RED}Some tests failed${NC}"
    exit 1
fi

