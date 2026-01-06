cmake_minimum_required(VERSION 3.10)
project(xlsx2csv C)

# Set C standard
set(CMAKE_C_STANDARD 11)
set(CMAKE_C_STANDARD_REQUIRED ON)

# Compiler flags
set(CMAKE_C_FLAGS "${CMAKE_C_FLAGS} -O2")

# Find required packages
find_package(PkgConfig REQUIRED)

# Find expat (XML parsing)
pkg_check_modules(EXPAT REQUIRED expat)

# Find libzip
pkg_check_modules(LIBZIP REQUIRED libzip)

# Source files
set(SOURCES
    src/main.c
    src/xlsx2csv.c
    src/zip_reader.c
    src/xml_parser.c
    src/csv_writer.c
    src/format_handler.c
    src/utils.c
)

# Header files
set(HEADERS
    src/xlsx2csv.h
)

# Include directories
include_directories(
    ${EXPAT_INCLUDE_DIRS}
    ${LIBZIP_INCLUDE_DIRS}
    src
)

# Create executable
add_executable(xlsx2csv ${SOURCES} ${HEADERS})

# Link libraries
target_link_libraries(xlsx2csv
    ${EXPAT_LIBRARIES}
    ${LIBZIP_LIBRARIES}
    m
)

# Compiler warning options (strict mode, ordered by importance)
# === Critical: Security and correctness ===
# -Wall: Enable common warnings
# -Wextra: Enable extra warnings
# -Werror: Treat all warnings as errors
# -Wformat=2: Enable strict format string checking (level 2) - prevents buffer overflow vulnerabilities
# -Warray-bounds=2: Warn if an array is accessed out of bounds (level 2) - prevents buffer overflow
# -Wstringop-overflow=4: Warn about buffer overflow in string operations (level 4) - critical security
# -Wstringop-truncation: Warn about truncation in string operations - prevents data loss
# -Wnull-dereference: Warn if dereferencing a NULL pointer may lead to undefined behavior - prevents crashes
# -Wsign-conversion: Warn for implicit conversions between signed and unsigned - prevents subtle bugs
# -Wundef: Warn if an undefined identifier is evaluated in an #if directive - prevents unexpected behavior
# === Important: Function safety ===
# -Wstrict-prototypes: Function declarations must be prototypes (func(void) not func())
# -Wmissing-prototypes: Warn about functions without prototypes
# -Wcast-qual: Warn about casts that discard qualifiers (const/volatile) - prevents undefined behavior
# === Important: Control flow safety ===
# -Wswitch-default: Warn about switch statements missing a default case - prevents unhandled cases
# -Wimplicit-fallthrough=5: Warn when a switch case falls through (level 5) - prevents accidental fallthrough
# === Moderate: Code quality ===
# -Wshadow: Warn when a variable shadows another - prevents confusion
# -Wduplicated-cond: Warn about duplicated conditions in if-else-if chains - code quality
# -Wduplicated-branches: Warn about duplicated branches in if-else statements - code quality
# -Wlogical-op: Warn about suspicious uses of logical operators - code quality
# === Style: Strict compliance ===
# -Wpedantic: Issue warnings needed for strict ISO C compliance - style enforcement
# Note: -Wfloat-equal is excluded as it conflicts with negative zero handling (signbit checks)
target_compile_options(xlsx2csv PRIVATE
    ${EXPAT_CFLAGS_OTHER}
    ${LIBZIP_CFLAGS_OTHER}
    -Wall
    -Wextra
    -Werror
    -Wformat=2
    -Warray-bounds=2
    -Wstringop-overflow=4
    -Wstringop-truncation
    -Wnull-dereference
    -Wsign-conversion
    -Wundef
    -Wstrict-prototypes
    -Wmissing-prototypes
    -Wcast-qual
    -Wswitch-default
    -Wimplicit-fallthrough=5
    -Wshadow
    -Wduplicated-cond
    -Wduplicated-branches
    -Wlogical-op
    -Wpedantic
)

