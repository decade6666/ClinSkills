# 标准项目目录结构

本文件是临床试验数据审核报告项目的标准目录结构。
`build-metadata` skill 的 Step 1 以此为基准校验项目目录，如有偏离则自动修正。

## 目录树

```
<project_root>/
├── 01 rawdata/                  # 原始数据（不入 Git）
├── 02 metadata/                 # EDC 元数据 Excel + build-metadata 生成的 JSON
├── 03 output/                   # 生成的报表（不入 Git）
├── 04 scripts/                  # 分析代码（平放，不分子文件夹）
├── .claude/                     # Claude Code 配置（不随项目变动）
│   ├── agents/                    # 自定义 Agent 定义（metadata-explorer / python-reviewer）
│   ├── hooks/                     # Claude Code PreToolUse/PostToolUse 钩子
│   │   ├── raw_read_guard.py    # PreToolUse: 禁止直接读 rawdata（Claude Code only）
│   │   └── syntax_check.py      # PostToolUse: 脚本语法检查（Claude Code only）
│   ├── rules/
│   │   └── constraints.md       # 强制约束
│   ├── settings.json            # 权限配置（permission deny 保护 rawdata；无 hooks）
│   └── skills/
│       ├── build-metadata/      # 元数据解析技能
│       └── write-script/        # 脚本编写技能
├── utils/                       # 公共工具函数
│   ├── __init__.py
│   ├── loaders.py               # 数据读取层（load_sheet / load_rand）
│   └── output_format.py         # 报表输出函数（三线表、xlsx）
├── .gitattributes               # Git 行尾与二进制规则
├── .gitignore                   # Git 忽略规则
├── CLAUDE.md                    # 项目说明与 Claude Code 约定
├── config.py                    # 路径配置加载器（从 config.yaml 读取）
├── config.yaml                  # 数据路径配置
└── requirements.txt             # Python 依赖
```

## 目录命名规则

数据目录使用 `序号 名称` 格式（序号为两位数字，后跟空格），确保文件管理器中排序一致：

| 序号 | 目录名 | 用途 | 入 Git |
|------|--------|------|--------|
| 01 | rawdata | EDC 导出的原始 Excel 数据 | 否 |
| 02 | metadata | 元数据 Excel + 解析后的 JSON | 是 |
| 03 | output | 生成的报表文件 | 否 |
| 04 | scripts | 数据核查 Python 脚本 | 否 |

## 骨架文件

以下文件是项目的基础配置，初始化时如不存在则生成默认模板：

### `.gitattributes`

```text
# Auto detect text files and normalize line endings to LF in repository
* text=auto eol=lf

# Explicitly declare text types
*.py text eol=lf
*.md text eol=lf
*.yaml text eol=lf
*.yml text eol=lf
*.json text eol=lf
*.txt text eol=lf

# Binary files
*.xlsx binary
*.docx binary
*.png binary
*.jpg binary
```

### `.gitignore`

```text
# Python cache
__pycache__/
*.pyc

# Generated output
03 output/

# Raw data (large / sensitive)
01 rawdata/

# OS
.DS_Store
Thumbs.db

.venv/
```

### `CLAUDE.md`

模板内容维护在独立文件，初始化时直接读取：

```
.claude/skills/build-metadata/reference/CLAUDE.md.template
```

模板包含 Overview、Directory Structure、Permissions、Conventions（含 EDC 类型表头占位注释）、Agent 清单。
`build-metadata` Step 2b 读取此文件后，根据 Step 1 确定的 EDC 类型替换 `<!-- EDC_TYPE_HEADER_START/END -->` 区块内的表头约定行，再写入项目根目录的 `CLAUDE.md`。

### `config.py`

```python
"""
config.py — 项目路径配置

从 config.yaml 加载数据路径，供各脚本直接 import。
"""

from pathlib import Path
import yaml

_PROJECT_ROOT = Path(__file__).resolve().parent
_CONFIG = _PROJECT_ROOT / "config.yaml"

with open(_CONFIG, "r", encoding="utf-8") as _f:
    _cfg = yaml.safe_load(_f)

raw_path    = str(_PROJECT_ROOT / _cfg["path"]["raw_path"])
pd_path     = str(_PROJECT_ROOT / _cfg["path"]["pd_path"])
code_path   = str(_PROJECT_ROOT / _cfg["path"]["code_path"])
remark_path = str(_PROJECT_ROOT / _cfg["path"]["remark_path"])
timewin_path = str(_PROJECT_ROOT / _cfg["path"]["timewin_path"])
output_path = str(_PROJECT_ROOT / _cfg["path"]["output_path"])
```

### `config.yaml`

```yaml
path:
  # 各原始数据 Excel 路径（相对项目根），按新项目实际填写
  raw_path: "01 rawdata/<日期>/<受试者数据导出>.xlsx"
  pd_path: "01 rawdata/<日期>/<方案偏离>.xlsx"
  code_path: "01 rawdata/<日期>/<医学编码报告>.xlsx"
  remark_path: "01 rawdata/<日期>/<备注明细>.xlsx"
  timewin_path: "01 rawdata/time window.xlsx"
  output_path: "03 output/<日期>"
```

### `requirements.txt`

```text
pandas>=2.0
numpy>=1.24
python-docx>=1.0
openpyxl>=3.1
scipy>=1.11
PyYAML>=6.0
XlsxWriter>=3.1
```

## 目录重命名时的路径同步清单

当数据目录（01-04）发生重命名时，以下文件中的路径引用必须同步更新：

| 文件 | 引用内容 |
|------|---------|
| `config.yaml` | `raw_path` / `pd_path` / `code_path` / `remark_path` / `timewin_path` / `output_path` |
| `CLAUDE.md` | 目录树、Permissions 节 |
| `.claude/settings.json` | permission deny / allow 规则 |
| `.claude/rules/constraints.md` | 强制约束（rawdata 保护、验证规则） |
| `.claude/hooks/syntax_check.py` | `parts[0]` 检查 + docstring（Claude Code 兼容） |
| `.claude/hooks/raw_read_guard.py` | `_under_dir()` 参数 + docstring（Claude Code 兼容） |
| `.claude/skills/write-script/SKILL.md` | 验证命令路径 |
| `.claude/skills/build-metadata/SKILL.md` | `metadata/` 路径描述 |
| `.claude/skills/write-script/scripts/query_metadata.py` | `_resolve_metadata_dir()` 默认路径 |
| `make_template.py` | `EMPTY_DIRS` + 骨架模板 |
| `.gitignore` | `rawdata/` / `output/` 规则 |
| `.claude/skills/write-script/reference/coding-guide.md` | `04 scripts/` 引用 |
| `.claude/skills/build-metadata/reference/project-structure.md` | 目录树 + sync checklist 自身 |
