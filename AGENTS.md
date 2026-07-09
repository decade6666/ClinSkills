# 脊痛宁数据审核报告 — opencode 项目规则

> 本文件与 `.claude/CLAUDE.md` + `.claude/rules/constraints.md` 保持同步。修改任一方时，必须同步修改另一方。

## Overview

临床试验数据审核报告项目。通过 Python + pandas 处理 EDC 导出的 Excel 数据，生成 .docx/.xlsx 报表。

## Directory Structure

```
├── config.py           # 项目路径配置（从 config.yaml 加载路径变量）
├── config.yaml         # 数据路径配置
├── requirements.txt    # Python 依赖
├── 注意事项.md          # 业务规则说明
├── 04 scripts/            # 分析代码（按章节拆分）
│   ├── a.试验整体情况小结/
│   ├── b.入排标准/
│   ├── c.方案偏离/
│   ├── d.超窗/
│   ├── e.缺失/
│   ├── f.用药后异常有临床意义/
│   ├── g.不良事件相关汇总/
│   ├── h.常规附件清单/
│   ├── i.方案禁止的合并用药清单/
│   └── j.依从性计算/
├── utils/              # 公共工具函数
│   ├── loaders.py      # 数据读取层（load_sheet / load_rand 等）
│   └── output_format.py # 报表输出函数（三线表、xlsx 等）
├── 01 rawdata/                # 原始数据（不入 Git）
└── 03 output/             # 生成的报表（不入 Git）
```

## 权限边界（不可逾越）

- `04 scripts/`、`utils/`、`config.py`、`config.yaml`：可编辑
- `01 rawdata/`、`03 output/`：不在 Git 中；**严禁直接读取 01 rawdata/ 下的表格文件**
- `.claude/**`：可编辑（Claude Code 配置目录）

## Conventions

编码规范（变量前缀、列名集中管理、八步操作模型、脚本模板等）详见 `.claude/skills/write-script/reference/coding-guide.md`。以下为跨生态的通用约定：

- 表头结构：`header=0, skiprows=[1]`（第 1 行中文列名，第 2 行英文列名被跳过）
- 报表函数来自 `utils/output_format.py`
- 数据读取函数来自 `utils/loaders.py`（`load_sheet` / `load_rand` 等）
- 生成文件路径由 `config.yaml` 的 `output_path` 控制（`config.py` 自动解析为绝对路径）
- 虚拟环境位于 `.venv/`，安装依赖：`pip install -r requirements.txt`

## 强制约束（三条铁律）

1. **改脚本前先加载 write-script skill**：任何对 `04 scripts/` 下 Python 脚本的编写或修改，必须先执行 `/write-script` 并按其流程执行，不得跳过直接动手。例外：仅限单行明确拼写 / 注释修正可跳过 skill。如 skill 不可用（调用失败或被拒绝），暂停并告知用户，不得自行推进。

2. **查数据形状用 query_metadata.py**：确认字段名、编码表、列名结构时，先查元数据（`.claude/skills/write-script/scripts/query_metadata.py`），query_metadata.py 无结果或报错时回退读 Excel 表头；字段名对不上时以 Excel 实际列名为准。**严禁直接读取 `01 rawdata/` 下的数据文件**。如需查看实际数据值，通过运行 `04 scripts/` 下的脚本间接获取。

3. **改脚本后必须验证**：`04 scripts/` 下脚本编写或修改后，必须实跑（数据不可用时做语法检查）确认通过，才提醒用户复核；同一错误修复尝试达 2 次仍失败时，停止并将完整报错交给用户定夺。语法检查用 `python -c "import ast; ast.parse(open(r'<路径>', encoding='utf-8').read()); print('OK')"`。

> opencode 环境说明：opencode 不支持 Claude Code 的 `PreToolUse`/`PostToolUse` hooks 机制。`.claude/hooks/` 下的脚本仅作为 Claude Code 兼容保留，在 opencode 中不生效。rawdata 保护依赖 `.opencode/plugins/` 下的 JS 插件 + `opencode.json` 权限规则 + 本 AGENTS.md 的行为指引。

## 超窗指标分类

### 疗效评价指标超窗
| 页面名称 | sheet名称 | 日期列 |
|---|---|---|
| 中医证候积分量表 | QS_TCM | 评估日期 |
| 脊柱疼痛量表 | QS_SPI | 评估日期 |
| BASDAI | QS_DAI | 评估日期 |
| BASFI | QS_SFI | 评估日期 |
| 患者总体评价（PGA） | RS | 评估日期 |

### 安全性评价指标超窗
| 页面名称 | sheet名称 | 日期列 |
|---|---|---|
| 生命体征 | VS | 检查日期 |
| 体格检查 | PE | 检查日期 |
| 血妊娠 | LB_HCG1 | 采样日期 |
| 尿妊娠 | LB_HCG2 | 采样日期 |
| C反应蛋白/超敏C反应蛋白 | LB_CRP | 采样日期 |
| 红细胞沉降率 | LB_ESR | 采样日期 |
| 血常规 | LB_HEM | 采样日期 |
| 尿常规 | LB_URI | 采样日期 |
| 尿沉渣镜检 | LB_MIC | 采样日期 |
| 随机尿微量白蛋白 | LB_UACR | 采样日期 |
| 血生化 | LB_CHEM | 采样日期 |
| 12导联心电图 | EG | 检查日期 |

### 其他指标超窗
| 页面名称 | sheet名称 | 日期列 |
|---|---|---|
| 访视日期 | SV | 访视日期 |
| 身高体重 | VS_HW | 检查日期 |
| 试验药物发放记录（发药日期） | DA_DD1 | 发药日期 |
| 试验药物发放记录（首次用药时间） | DA_DD1 | 受试者首次用药时间 |
| 试验药物回收记录 | DA_DR1 | 回收日期 |

时间窗值统一与随机时间叠加计算（正=往后，负=往前）。
