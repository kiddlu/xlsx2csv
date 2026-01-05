cmake_minimum_required(VERSION 3.10)
project(xlsx2csv C)

# Set C standard
set(CMAKE_C_STANDARD 11)
set(CMAKE_C_STANDARD_REQUIRED ON)

# Compiler flags
set(CMAKE_C_FLAGS "${CMAKE_C_FLAGS} -Wall -Wextra -O2 -g")

# Find required packages
find_package(PkgConfig REQUIRED)

# Find libxml2
pkg_check_modules(LIBXML2 REQUIRED libxml-2.0)

# Find libzip
pkg_check_modules(LIBZIP REQUIRED libzip)

# Find libcsv (may not have pkg-config, use direct link)
find_library(CSV_LIB csv)

if(NOT CSV_LIB)
    message(FATAL_ERROR "libcsv not found")
endif()

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
    ${LIBXML2_INCLUDE_DIRS}
    ${LIBZIP_INCLUDE_DIRS}
    src
)

# Create executable
add_executable(xlsx2csv ${SOURCES} ${HEADERS})

# Link libraries
target_link_libraries(xlsx2csv
    ${LIBXML2_LIBRARIES}
    ${LIBZIP_LIBRARIES}
    ${CSV_LIB}
    m
)

# Set compile flags
target_compile_options(xlsx2csv PRIVATE
    ${LIBXML2_CFLAGS_OTHER}
    ${LIBZIP_CFLAGS_OTHER}
)

# Install target
install(TARGETS xlsx2csv
    RUNTIME DESTINATION bin
)

