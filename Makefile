# xlsx2csv C version Makefile
# Wraps CMake build system with code formatting
# Compatible with Python xlsx2csv v0.8.3

# Detect CPU count for parallel builds
CPUS := $(shell cat /proc/cpuinfo 2>/dev/null | grep "processor" | wc -l)
ifeq ($(CPUS),)
  CPUS := $(shell sysctl -n hw.ncpu 2>/dev/null || echo 4)
endif

PWD := $(shell pwd)
BUILD_DIR := $(PWD)/build
TARGET := xlsx2csv
TEST_DIR := tests

# Make options (can be overridden)
MAKE_OPT ?=

# 1. Pre-build step: Generate build system with cmake
define shcmd-makepre
	@echo "[shcmd-makepre] Generating build system..."
	@mkdir -p $(BUILD_DIR)
	@cd $(BUILD_DIR) && cmake ..
endef

# 2. Build step: Compile with make
define shcmd-make
	@echo "[shcmd-make] Building..."
	@cd $(BUILD_DIR) && make -j$(CPUS) $(MAKE_OPT) 2>&1 | grep -v "^make\[[0-9]\]:"
endef

# 3. Clean step: Clean build artifacts
define shcmd-makeclean
	@echo "[shcmd-makeclean] Cleaning build artifacts..."
	@if [ -d $(BUILD_DIR) ]; then \
		(cd $(BUILD_DIR) && make clean && echo "## Clean build ##"); \
	fi
endef

# 4. Remove build directory completely
define shcmd-makerm
	@echo "[shcmd-makerm] Removing build directory..."
	@if [ -d $(BUILD_DIR) ]; then \
		rm -rf $(BUILD_DIR); \
	fi
endef

# 5. Pre-build custom step: Format code before building
define shcmd-pre-make-custom
	@echo "[shcmd-pre-make-custom] Formatting code..."
	@if command -v clang-format >/dev/null 2>&1; then \
		find src -type f \( -name "*.c" -o -name "*.h" \) -print0 | \
			xargs -0 -P $(CPUS) clang-format -i 2>/dev/null || true; \
		echo "Code formatting completed."; \
	else \
		echo "Warning: clang-format not found, skipping code formatting."; \
	fi
endef

# 6. Post-build custom step (optional)
define shcmd-post-make-custom
	@echo "[shcmd-post-make-custom] Build completed."
	@if [ -f $(BUILD_DIR)/$(TARGET) ]; then \
		cp $(BUILD_DIR)/$(TARGET) $(PWD)/$(TARGET); \
		echo "Executable copied to $(PWD)/$(TARGET)"; \
	fi
endef

# Main targets
.PHONY: all clean rm pre test format

# Default target: format, configure, and build
all: pre
	$(call shcmd-pre-make-custom)
	$(call shcmd-make)
	$(call shcmd-post-make-custom)

# Only configure cmake (don't build)
pre:
	$(call shcmd-makepre)

# Clean build artifacts
clean:
	$(call shcmd-makeclean)
	@rm -f $(TARGET)
	@rm -f $(TEST_DIR)/actual/*.csv

# Remove build directory completely
rm: clean
	$(call shcmd-makerm)

# Format code only (without building)
format:
	$(call shcmd-pre-make-custom)

# Run tests
test: $(TARGET)
	@echo "Running tests..."
	@cd $(TEST_DIR) && bash test_runner.sh

# Ensure target exists before test
$(TARGET):
	@if [ ! -f $(TARGET) ]; then \
		echo "Error: $(TARGET) not found. Run 'make' first."; \
		exit 1; \
	fi
