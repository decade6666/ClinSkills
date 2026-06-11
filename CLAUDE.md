# Project: 脊痛宁数据审核报告

## Overview
临床试验数据审核报告项目。通过 Python + pandas 处理 EDC 导出的 Excel 数据，生成 .docx/.xlsx 报表。

## Directory Structure
```
├── env.py              # 项目环境入口（加载依赖和 config.yaml）
├── config.yaml         # 数据路径配置
├── 注意事项.md          # 业务规则说明
├── scripts/            # 分析代码（.py, jupytext percent 格式）← Claude Code 编辑这里
├── notebooks/          # Jupyter notebooks（.ipynb，从 scripts 自动生成）← 不入 Git
├── utils/              # 公共工具函数
├── raw/                # 原始数据（不入 Git）
└── output/             # 生成的报表（不入 Git）
```

## Jupytext Workflow
`scripts/*.py` 是 Git 源文件，`notebooks/*.ipynb` 是运行时文件（不入 Git）。

- **编辑**: Claude Code 编辑 `scripts/*.py`（纯代码，无输出污染）
- **运行**: JupyterHub 中打开 `notebooks/*.ipynb`（jupytext 自动同步 `.py` 变更）
- **同步**: `jupytext --sync notebooks/*.ipynb`（从 scripts 重建 notebooks）

## Permissions
- `.ipynb` 文件：Claude Code 禁止编辑（只读）
- `scripts/`、`utils/`、`env.py`、`config.yaml`：Claude Code 可编辑
- `raw/`、`output/`：不在 Git 中

## Conventions
- Notebook 文件名格式: `[字母].中文名`
- scripts 中通过 `%run ../env.py` 加载环境
- 报表函数来自 `utils/output_format.py`
- 生成文件路径由 `config.yaml` 的 `output_path` 控制
