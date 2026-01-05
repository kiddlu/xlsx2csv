#!/bin/bash
# Quick test statistics script

echo "========================================="
echo "xlsx2csv Test Statistics"
echo "========================================="
echo ""

cd tests 2>/dev/null || exit 1

# Run tests and capture output
OUTPUT=$(./test_runner.sh 2>&1)

# Extract statistics
PASSED=$(echo "$OUTPUT" | grep "Tests passed:" | grep -o '[0-9]*' | head -1)
FAILED=$(echo "$OUTPUT" | grep "Tests failed:" | grep -o '[0-9]*' | head -1)
TOTAL=$((PASSED + FAILED))

if [ $TOTAL -eq 0 ]; then
    echo "No tests found or tests not run"
    exit 1
fi

PASS_RATE=$(awk "BEGIN {printf \"%.1f\", ($PASSED/$TOTAL)*100}")

echo "Total Tests:  $TOTAL"
echo "Passed:       $PASSED ✅"
echo "Failed:       $FAILED ❌"
echo "Pass Rate:    $PASS_RATE%"
echo ""

if [ $FAILED -eq 0 ]; then
    echo "🎉 All tests passed!"
    exit 0
else
    echo "Failed tests:"
    echo "$OUTPUT" | grep "FAIL" | sed 's/\[0;31m//g' | sed 's/\[0m//g' | sed 's/Testing /  - /'
    echo ""
    echo "Run 'make test' for details"
    exit 1
fi

