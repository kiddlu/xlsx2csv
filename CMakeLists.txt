cmake_minimum_required(VERSION 3.10)

project(xlsx2csv C)

set(CMAKE_C_STANDARD 11)
set(CMAKE_C_STANDARD_REQUIRED ON)

add_definitions (
	-O3
)

#add_compile_options(-g -O0)
#add_link_options(-g)
#set(CMAKE_C_FLAGS "${CMAKE_C_FLAGS} -fno-omit-frame-pointer -fsanitize=address")
#set(CMAKE_EXE_LINKER_FLAGS "${CMAKE_EXE_LINKER_FLAGS} -fno-omit-frame-pointer -fsanitize=address")

set(SOURCES
${PROJECT_SOURCE_DIR}/src/main.c
${PROJECT_SOURCE_DIR}/src/xlsx2csv.c
${PROJECT_SOURCE_DIR}/src/zip_reader.c
${PROJECT_SOURCE_DIR}/src/xml_parser.c
${PROJECT_SOURCE_DIR}/src/csv_writer.c
${PROJECT_SOURCE_DIR}/src/format_handler.c
${PROJECT_SOURCE_DIR}/src/utils.c
)

#static link
#set(CMAKE_EXE_LINKER_FLAGS "${CMAKE_EXE_LINKER_FLAGS} -static")

include_directories (${PROJECT_SOURCE_DIR}/src/)

add_executable(xlsx2csv ${SOURCES})
# Compiler warning options
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
# Note: -Wfloat-equal is excluded as it conflicts with negative zero handling (signbit checks)
# -Wduplicated-cond: Warn about duplicated conditions in if-else-if chains - code quality
# -Wduplicated-branches: Warn about duplicated branches in if-else statements - code quality
# -Wlogical-op: Warn about suspicious uses of logical operators - code quality
# === Style: Strict compliance ===
# -Wpedantic: Issue warnings needed for strict ISO C compliance - style enforcement
target_compile_options(xlsx2csv PRIVATE
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

target_link_libraries(xlsx2csv -lexpat -lzip -lm)
