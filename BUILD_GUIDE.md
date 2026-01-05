# xlsx2csv 构建指南

## 📋 构建系统架构

本项目使用 **Make 封装 CMake** 的构建方式：

- **CMakeLists.txt**: 负责实际的构建配置（编译选项、源文件、链接库等）
- **Makefile**: 作为包装层，提供简化的构建接口，封装 CMake 的调用流程

## 🚀 快速开始

### 基本构建

```bash
# 完整构建（格式化代码 + CMake 配置 + 编译）
make

# 仅配置 CMake（不编译）
make pre

# 仅格式化代码（不构建）
make format

# 清理构建产物
make clean

# 完全删除构建目录
make rm
```

### 构建流程

执行 `make` 时会自动执行以下步骤：

1. **代码格式化** (`shcmd-pre-make-custom`)
   - 使用 `clang-format` 格式化所有 `.c` 和 `.h` 文件
   - 基于项目根目录的 `.clang-format` 配置文件
   - 并行处理，利用多核 CPU 加速

2. **CMake 配置** (`shcmd-makepre`)
   - 在 `build/` 目录生成构建系统
   - 检测依赖库（libzip, libxml2, libcsv）
   - 配置编译选项

3. **编译** (`shcmd-make`)
   - 使用 `make -j$(CPUS)` 并行编译
   - 自动检测 CPU 核心数

4. **后处理** (`shcmd-post-make-custom`)
   - 将可执行文件从 `build/` 复制到项目根目录

## 📦 依赖要求

### 构建工具

- **CMake** >= 3.10
- **GCC** 或 **Clang** (支持 C11)
- **clang-format** (用于代码格式化)
- **pkg-config** (用于检测依赖库)

### 运行时依赖

- **libzip** - ZIP 文件处理
- **libxml2** - XML 解析
- **libcsv** - CSV 字段转义

### 安装依赖 (Debian/Ubuntu)

```bash
sudo apt-get install -y \
    cmake \
    build-essential \
    clang-format \
    pkg-config \
    libzip-dev \
    libxml2-dev \
    libcsv-dev
```

## 🔧 构建选项

### 并行编译

Makefile 会自动检测 CPU 核心数并使用并行编译：

```bash
# 使用所有 CPU 核心（默认）
make

# 手动指定并行数
make MAKE_OPT="-j4"
```

### 调试构建

```bash
# 使用 CMake 的调试模式
cd build && cmake -DCMAKE_BUILD_TYPE=Debug .. && make
```

### 发布构建

```bash
# 使用 CMake 的发布模式
cd build && cmake -DCMAKE_BUILD_TYPE=Release .. && make
```

## 📁 目录结构

```
xlsx2csv/
├── CMakeLists.txt      # CMake 构建配置
├── Makefile            # Make 包装层
├── .clang-format       # 代码格式化配置
├── build/              # 构建目录（自动生成，已加入 .gitignore）
│   ├── CMakeCache.txt
│   ├── CMakeFiles/
│   └── xlsx2csv        # 编译后的可执行文件
├── src/                # 源代码目录
└── tests/              # 测试目录
```

## 🧪 测试

```bash
# 构建并运行测试
make test

# 或手动运行
cd tests && bash test_runner.sh
```

## 📝 代码格式化

### 自动格式化

构建时会自动格式化代码：

```bash
make  # 自动格式化 + 构建
```

### 手动格式化

```bash
# 仅格式化，不构建
make format

# 或直接使用 clang-format
find src -type f \( -name "*.c" -o -name "*.h" \) -print0 | \
    xargs -0 clang-format -i
```

### 格式化配置

格式化规则由 `.clang-format` 文件定义，基于 Chromium 风格。

## 🔍 故障排除

### CMake 找不到依赖库

```bash
# 检查 pkg-config 是否能找到库
pkg-config --modversion libxml-2.0
pkg-config --modversion libzip

# 如果找不到，安装开发包
sudo apt-get install -y libzip-dev libxml2-dev libcsv-dev
```

### clang-format 未找到

```bash
# 安装 clang-format
sudo apt-get install -y clang-format

# 或跳过格式化（不推荐）
# 编辑 Makefile，注释掉格式化步骤
```

### 构建失败

```bash
# 清理并重新构建
make rm
make

# 查看详细错误信息
cd build && make VERBOSE=1
```

## 📚 参考

- [CMake 官方文档](https://cmake.org/documentation/)
- [clang-format 文档](https://clang.llvm.org/docs/ClangFormat.html)
- 项目 styleguide: `../styleguide/C.md`
