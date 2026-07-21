# ClinSkills

临床试验数据审核（DMR）报告的 Claude Code 技能集：EDC 元数据解析、数据核查脚本编写、改动审查、harness 审计、skill 编写。

## 一句话安装（全局，跨项目可用）

**Windows（PowerShell 7+）：**

```powershell
irm https://raw.githubusercontent.com/Doraemon-code/ClinSkills/master/install.ps1 | iex
```

**macOS / Linux：**

```bash
curl -fsSL https://raw.githubusercontent.com/Doraemon-code/ClinSkills/master/install.sh | bash
```

装完 `~/.claude/` 下会有 skills、agents、hooks，并注册语法检查 + raw 数据保护 hook（全局安全版）。**同名 skill 会被覆盖更新**，重跑即更新。依赖：`git`、`python`（在 PATH 上）。

## 包含内容

| 类型 | 名称 |
|---|---|
| Skills | `build-metadata`、`write-script`、`review-changes`、`audit-harness`、`build-skill` |
| Agents | `metadata-explorer`、`python-reviewer` |
| Hooks | `syntax_check`、`raw_read_guard`（均全局注册；`raw_read_guard` 为**全局安全版**，仅当命令确切指向 raw 时才拦） |

## 用法

- **新临床项目**：进入项目目录，触发 `build-metadata`——校验/脚手架目录结构、解析 EDC 元数据为 JSON。
- **写核查脚本**：`write-script`（口述需求或给输出示例）。
- **提交前审查**：`review-changes`。
- **审 harness / 写新 skill**：`audit-harness` / `build-skill`。

## 卸载

删除 `~/.claude/skills/` 下 `build-metadata`、`write-script`、`review-changes`、`audit-harness`、`build-skill`，`~/.claude/agents/` 下 `metadata-explorer`、`python-reviewer`，`~/.claude/hooks/` 下 `syntax_check.py`、`raw_read_guard.py`，并从 `~/.claude/settings.json` 移除对应的 `hooks` 条目与 `deny Read(01 rawdata/**)`。

## 设计说明

- **项目无需自带 `.claude/`**：skills / agents、语法检查 hook、raw 数据保护（`deny Read(01 rawdata/**)` + `raw_read_guard`）全部注册进全局 `~/.claude/`，跨项目生效。
- **全局安全**：`raw_read_guard` 只在命令**确切指向 raw**（出现 `01 rawdata/…xlsx` 字面路径，或读调用 + `raw_path` 变量）时才拦，非临床项目里普通的 `read_excel(` 不受影响；`syntax_check` 只作用于 `04 scripts/` 与 `utils/`。两者优先用 `CLAUDE_PROJECT_DIR` 定位当前项目。
- `utils/`（`loaders` 读取 / `output_docx`·`output_xlsx` 输出层）是项目运行时被 import 的代码，由 `build-metadata` 脚手架进目标项目（源仓库 `utils/` 单一源，安装时置入 build-metadata skeleton）。
