CPUS=$(shell cat /proc/cpuinfo | grep "processor" | wc -l)
PWD=$(shell pwd)
BUILD_DIR=$(PWD)/build
TEST_DIR=$(PWD)/tests
MAKE_OPT=

define shcmd-makepre
	@echo "[shcmd-makepre]"
	mkdir -p $(BUILD_DIR)
	cd $(BUILD_DIR) && cmake ..
endef

define shcmd-make
	@echo "[shcmd-make]"
	@cd $(BUILD_DIR) && make -j$(CPUS) $(MAKE_OPT) | grep -v "^make\[[0-9]\]:"
endef

define shcmd-makeclean
	@echo "[shcmd-makeclean]"
	@if [ -d $(BUILD_DIR) ]; then (cd $(BUILD_DIR) && make clean && echo "##Clean build##"); fi
endef

define shcmd-makerm
	@echo "[shcmd-makerm]"
	@if [ -d $(BUILD_DIR) ]; then (rm -rf $(BUILD_DIR)); fi
endef

define shcmd-pre-make-custom
	@echo "[shcmd-pre-make-custom]"
	@echo Codelines Sum:
	@find src -type f \( -name "*.c" -o -name "*.h" \) -exec wc -l {} + | tail -1
	@find src -type f \( -name "*.c" -o -name "*.h" \) -print0 | xargs -0 -P $(CPUS) clang-format -i
endef

define shcmd-post-make-custom
	@echo "[shcmd-post-make-custom]"
endef

define shcmd-test
	@echo "[shcmd-test]"
	@echo "Running tests..."
	@cd $(TEST_DIR) && bash test_runner.sh
endef

.PHONY: all clean rm pre test
all: pre
	$(call shcmd-pre-make-custom)
	$(call shcmd-make)
	$(call shcmd-post-make-custom)

clean:
	$(call shcmd-makeclean)

rm:
	$(call shcmd-makerm)

pre:
	$(call shcmd-makepre)

test:
	$(call shcmd-test)
