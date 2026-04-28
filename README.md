# 文档页数统计工具

一个简单好用的桌面应用程序，用于批量统计文件夹中 Word 和 PDF 文档的页数，并支持导出 Excel 报表。

## 功能特性

- 选择任意文件夹，自动扫描子目录中的文档
- 支持 Word (.docx) 和 PDF (.pdf) 文件
- 可自由选择只扫描 Word、只扫描 PDF，或两者都扫描
- 实时显示扫描进度和结果列表
- 一键导出 Excel 报表（包含文件名、类型、页数、状态、完整路径）
- 无需安装额外软件，双击即可运行

## 环境准备

本项目需要 Python 3.10 或更高版本。如果你尚未安装，请前往 [Python 官网](https://www.python.org/downloads/) 下载。

**创建虚拟环境（推荐）**

虚拟环境用来隔离项目依赖，避免和系统 Python 冲突。

```bash
# 进入项目根目录（page_counter/）
cd page_counter

# 创建虚拟环境
python3 -m venv venv

# 激活虚拟环境
source venv/bin/activate        # macOS / Linux
# 或 venv\Scripts\activate      # Windows

# 安装依赖
pip install -r src/requirements.txt
```

## 使用说明

### macOS 用户

#### 方法一：直接下载 app（最简单）

1. 从 Releases 或 Actions 产物中下载 `PageCounter_macOS.zip`
2. 解压后得到 `PageCounter.app`
3. 双击即可运行
4. 如遇"无法打开"提示，请前往 **系统设置 > 隐私与安全性** 中允许运行

#### 方法二：本地源码运行

完成上面的【环境准备】后，执行：

```bash
source venv/bin/activate
python src/page_counter.py
```

#### 方法三：本地重新打包

如果你想修改代码后重新构建可执行文件：

```bash
source venv/bin/activate
./build_mac.sh
```

构建完成后，产物位于 `page_counter/dist/PageCounter_macOS.zip`。

### Windows 用户

#### 方法一：GitHub Actions 自动构建（推荐）

将本仓库推送到 GitHub，Actions 会自动构建 Windows 可执行文件。构建完成后在 Actions 页面下载 `PageCounter-Windows` 产物即可。

#### 方法二：本地一键打包

1. 确保已安装 Python 3.10 或更高版本
2. 完成【环境准备】（Windows 下激活命令为 `venv\Scripts\activate`）
3. 在项目根目录双击运行 `src/build.bat`
4. 生成的 `PageCounter.exe` 位于 `dist` 文件夹中

#### 方法三：直接运行源码

```bash
venv\Scripts\activate
python src/page_counter.py
```

## 技术说明

- **Word 页数读取原理**：从 .docx 文件的 `docProps/app.xml` 中读取 Microsoft Word 保存的页数统计信息。如果文档不是由 Word 保存的（例如某些在线工具生成），可能无法读取到页数。
- **PDF 页数读取原理**：直接读取 PDF 文件结构中的页数信息，准确率接近 100%。

## 项目结构

```
page_counter/
├── src/
│   ├── page_counter.py        # 主程序源码
│   ├── requirements.txt       # Python 依赖列表
│   ├── build.bat              # Windows 本地打包脚本
│   └── README.md              # 项目说明
├── build_mac.sh               # macOS 本地打包脚本
├── .github/workflows/build.yml # GitHub Actions 自动构建配置
└── .gitignore
```

## 开源协议

MIT License
