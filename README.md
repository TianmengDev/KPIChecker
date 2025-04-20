# KPI检查器 (KPI Checker)

[![Python Version](https://img.shields.io/badge/python-3.7%2B-blue.svg)](https://www.python.org/downloads/)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

KPI检查器是一个用于检查Word文档中KPI自评语句的工具。它能扫描指定目录下的所有.docx文件，检查是否包含类似"本季度KPI自评XX分"的语句，并生成报告。该工具特别适用于人力资源部门检查员工季度工作总结中是否包含规范的KPI自评。

<p align="center">
  <img src="https://raw.githubusercontent.com/TianmengDev/KPIChecker/refs/heads/main/screenshot.png" alt="KPI检查器截图" width="600">
</p>

## 📋 目录

- [功能特点](#功能特点)
- [安装](#安装)
- [使用方法](#使用方法)
- [输出结果](#输出结果)
- [配置文件](#配置文件)
- [开发指南](#开发指南)
- [常见问题](#常见问题)
- [许可证](#许可证)

## ✨ 功能特点

- 📄 批量扫描目录中的所有.docx文件
- 🔍 支持递归扫描子目录
- 🔎 通过正则表达式识别多种格式的KPI自评语句，包括：
  - `2024年第四季度KPI考核自评85分`
  - `季度KPI考核自评96分`
  - `KPI 自评得分 94 分`
  - `自评得分 94 分`
  - `KPI考核自评82分`
- 📊 生成详细的检查报告（控制台、TXT或Excel格式）
- 🔧 提供自动修复功能，为缺少KPI自评的文档添加标准模板语句
- ⚙️ 支持通过JSON配置文件自定义检查规则和报告格式

## 🚀 安装

### 环境要求

- Python 3.6+

### 依赖安装

**重要提示：** 请确保安装所有依赖包，特别是`openpyxl`，它是生成Excel报告的必要库。

```bash
# 安装所有依赖
pip install -r requirements.txt

# 或者单独安装必要的包
pip install python-docx pandas openpyxl
```

### 从源码安装

```bash
# 克隆仓库
git clone https://github.com/yourusername/kpichecker.git

# 进入项目目录
cd kpichecker

# 安装依赖
pip install -r requirements.txt
```

## 📖 使用方法

### 基本用法

将KPI检查器放在包含要检查的`.docx`文件的目录中，然后运行：

```bash
python main.py
```

默认情况下，程序会扫描当前目录下的所有.docx文件，检查是否包含KPI自评语句，并生成Excel格式的报告。

### 命令行选项

```
usage: main.py [-h] [--dir DIR] [--recursive] [--output OUTPUT]
               [--format {console,txt,excel}] [--config CONFIG]
               [--paragraphs PARAGRAPHS] [--fix] [--fix-list FIX_LIST]
               [--no-backup] [--fix-preview] [--version] [--check-deps]

KPI自评检查工具
```

#### 常用选项

| 选项 | 简写 | 描述 |
|------|------|------|
| `--dir DIR` | `-d DIR` | 要扫描的目录路径，默认为当前目录 |
| `--recursive` | `-r` | 是否递归扫描子目录 |
| `--output OUTPUT` | `-o OUTPUT` | 报告输出目录，默认为当前目录 |
| `--format {console,txt,excel}` | `-f {console,txt,excel}` | 报告格式 |
| `--config CONFIG` | `-c CONFIG` | 配置文件路径 |
| `--fix` | | 修复缺少KPI自评的文件 |
| `--fix-preview` | | 预览修复效果，不实际修改文件 |
| `--check-deps` | | 检查依赖库是否已安装 |
| `--version` | `-v` | 显示版本信息 |

### 使用示例

1. 扫描指定目录并生成Excel报告：

```bash
python main.py --dir /path/to/documents --format excel --output /path/to/output
```

2. 递归扫描目录及其子目录：

```bash
python main.py --dir /path/to/documents --recursive
```

3. 预览修复效果：

```bash
python main.py --dir /path/to/documents --fix-preview
```

4. 自动修复缺少KPI自评的文件：

```bash
python main.py --dir /path/to/documents --fix
```

## 📊 输出结果

KPI检查器可以生成三种格式的报告：控制台输出、TXT文本文件和Excel电子表格。

### 控制台输出

控制台输出包含检查的总结信息，如：
- 总文件数
- 包含KPI自评的文件数
- 缺少KPI自评的文件数
- 平均KPI分数

以及详细的文件列表和检测结果。

### TXT文本报告

TXT报告与控制台输出内容相同，但保存为文本文件，默认文件名为`KPI检查报告.txt`。

### Excel报告（默认）

Excel报告更加结构化，包含多个工作表：

1. **摘要表**：
   - 总文件数
   - 包含KPI自评的文件数
   - 缺少KPI自评的文件数
   - 平均KPI分数

2. **详细数据表**：
   - 文件名
   - 文件路径
   - 是否包含KPI自评
   - KPI自评分数
   - 匹配文本

3. **缺少KPI自评的文件表**：
   - 文件名
   - 文件路径

示例Excel报告结构：

```
KPI检查报告.xlsx
├── 摘要
│   ├── 总文件数
│   ├── 包含KPI自评的文件数
│   ├── 缺少KPI自评的文件数
│   └── 平均分
│
├── 详细数据
│   ├── 文件名
│   ├── 文件路径
│   ├── 包含KPI自评
│   ├── KPI自评分数
│   └── 匹配文本
│
└── 缺少KPI自评的文件
    ├── 文件名
    └── 文件路径
```

## ⚙️ 配置文件

您可以通过JSON配置文件自定义程序的行为：

```json
{
  "check_last_paragraphs": 10,
  "kpi_patterns": [
    "(\\d{4})年?第?[一二三四]季度KPI考核自评(\\d{1,3})分",
    "季度KPI考核自评(\\d{1,3})分",
    "KPI\\s*自评得分\\s*(\\d{1,3})\\s*分",
    "自评得分\\s*(\\d{1,3})\\s*分",
    "KPI考核自评(\\d{1,3})分"
  ],
  "report": {
    "default_format": "excel",
    "excel_filename": "KPI检查报告.xlsx",
    "txt_filename": "KPI检查报告.txt"
  },
  "fixer": {
    "enabled": false,
    "template": "{year}年第{quarter}季度KPI考核自评__分。"
  }
}
```

### 配置项说明

| 配置项 | 描述 | 默认值 |
|--------|------|--------|
| `check_last_paragraphs` | 检查文档末尾的段落数量 | `10` |
| `kpi_patterns` | KPI自评语句的正则表达式模式 | 多种格式 |
| `report.default_format` | 默认报告格式 | `excel` |
| `report.excel_filename` | Excel报告的文件名 | `KPI检查报告.xlsx` |
| `report.txt_filename` | TXT报告的文件名 | `KPI检查报告.txt` |
| `fixer.enabled` | 是否默认启用修复功能 | `false` |
| `fixer.template` | 修复模板 | `{year}年第{quarter}季度KPI考核自评__分。` |

使用配置文件：

```bash
python main.py --config /path/to/config.json
```

## 🛠️ 开发指南

### 项目结构

```
kpichecker/
├── .gitignore                 # Git忽略配置
├── LICENSE                    # MIT许可证
├── README.md                  # 项目文档
├── config.json                # 默认配置文件
├── example-docs/              # 示例文档目录
├── kpichecker/                # 主要代码模块
│   ├── __init__.py            # 包初始化文件
│   ├── config.py              # 配置处理模块
│   ├── docx_parser.py         # 文档解析模块
│   ├── docx_scanner.py        # 文档扫描模块
│   ├── file_fixer.py          # 文档修复模块
│   ├── kpi_checker.py         # KPI检查模块
│   └── report_generator.py    # 报告生成模块
├── main.py                    # 主程序入口
├── requirements.txt           # 依赖项列表
└── setup.py                   # 安装脚本
```

### 扩展开发

如需添加新的功能或修改现有代码，请参考以下步骤：

1. 添加新的KPI识别模式：修改`config.py`中的`kpi_patterns`列表
2. 支持新的报告格式：扩展`report_generator.py`模块
3. 增强文档扫描能力：修改`docx_scanner.py`模块

## ❓ 常见问题

### 找不到openpyxl模块

**问题**：运行程序时报错：`ModuleNotFoundError: No module named 'openpyxl'`

**解决方案**：
```bash
pip install openpyxl
```

### 如何在Anaconda环境中运行？

**解决方案**：
```bash
conda create -n kpichecker python=3.8
conda activate kpichecker
pip install -r requirements.txt
python main.py
```


## 🗒️ 贡献指南

我们欢迎任何形式的贡献！如果您想为项目做出贡献，请遵循以下步骤：

1. Fork 项目
2. 创建您的特性分支 (`git checkout -b feature/AmazingFeature`)
3. 提交您的更改 (`git commit -m 'Add some AmazingFeature'`)
4. 推送到分支 (`git push origin feature/AmazingFeature`)
5. 打开一个 Pull Request

## ✅ 许可证

本项目采用 MIT 许可证 - 详情请参阅 [LICENSE](LICENSE) 文件

## 🚩 免责声明

本项目仅供学习和研究使用，请勿用于任何商业用途。使用本工具时请遵守相关平台的使用条款和规定。因使用本工具而产生的任何问题，开发者不承担任何责任。

## ☎️ 联系方式

如有任何问题或建议，请通过以下方式联系我们：
- 提交 Issue


