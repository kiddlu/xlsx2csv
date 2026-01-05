# xlsx2csv C version Makefile
# Compatible with Python xlsx2csv v0.8.3

CC = gcc
CFLAGS = -std=c11 -Wall -Wextra -O2 -g
LDFLAGS = -lzip -lxml2 -lcsv -lm

# pkg-config for libxml2 and libzip
CFLAGS += $(shell pkg-config --cflags libxml-2.0 libzip 2>/dev/null || pkg-config --cflags libxml-2.0)
LDFLAGS += $(shell pkg-config --libs libxml-2.0 libzip 2>/dev/null || echo "-lzip -lxml2")

TARGET = xlsx2csv
SRC_DIR = src
OBJ_DIR = obj
TEST_DIR = tests

SOURCES = $(SRC_DIR)/main.c \
          $(SRC_DIR)/xlsx2csv.c \
          $(SRC_DIR)/zip_reader.c \
          $(SRC_DIR)/xml_parser.c \
          $(SRC_DIR)/csv_writer.c \
          $(SRC_DIR)/format_handler.c \
          $(SRC_DIR)/utils.c

OBJECTS = $(patsubst $(SRC_DIR)/%.c,$(OBJ_DIR)/%.o,$(SOURCES))
HEADERS = $(SRC_DIR)/xlsx2csv.h

.PHONY: all clean test install

all: $(TARGET)

$(TARGET): $(OBJECTS)
	$(CC) $(OBJECTS) $(LDFLAGS) -o $@

$(OBJ_DIR)/%.o: $(SRC_DIR)/%.c $(HEADERS) | $(OBJ_DIR)
	$(CC) $(CFLAGS) -c $< -o $@

$(OBJ_DIR):
	mkdir -p $(OBJ_DIR)

clean:
	rm -rf $(OBJ_DIR) $(TARGET)
	rm -f $(TEST_DIR)/actual/*.csv

test: $(TARGET)
	@echo "Running tests..."
	@cd $(TEST_DIR) && bash test_runner.sh

install: $(TARGET)
	install -m 755 $(TARGET) /usr/local/bin/

uninstall:
	rm -f /usr/local/bin/$(TARGET)

