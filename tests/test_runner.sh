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

PYTHON_XLSX2CSV="/usr/bin/xlsx2csv"
C_XLSX2CSV="../xlsx2csv"

if [ ! -f "$C_XLSX2CSV" ]; then
    echo -e "${RED}Error: C xlsx2csv not found at $C_XLSX2CSV${NC}"
    echo "Please run 'make' first"
    exit 1
fi

if [ ! -f "$PYTHON_XLSX2CSV" ]; then
    echo -e "${RED}Error: Python xlsx2csv not found at $PYTHON_XLSX2CSV${NC}"
    echo "Please install Python xlsx2csv"
    exit 1
fi

# Create output directories (do NOT create expected - it's generated each time)
mkdir -p actual

echo "====================================="
echo "xlsx2csv Test Suite"
echo "Testing against Python xlsx2csv $(xlsx2csv --version 2>&1 | head -1)"
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

