# Skills 项目说明

本目录包含两个 Claude Code skills：`read_documents`（读取文档）和 `template-write`（按模板填写文档）。

## 测试规则

**所有测试一律在 `test/` 文件夹中进行。**

- 测试输入文件（模板 docx、截图、操作文档等）放在 `test/`
- 测试产生的输出文件、临时文件（temp/、sections.md 等）也放在 `test/` 下
- **禁止在 skill 源码目录（`read_documents/`、`template-write/`）中创建测试文件**——这些目录只放 skill 本身的代码和文档
- 测试完成后清理 `test/temp/` 等临时产物，保留有价值的测试用例

`test/` 已在 `.gitignore` 中排除，测试文件不会被提交到版本库。

## 测试用例（本地，不入版本库）

`test/` 已 gitignore，**不在仓库中**。克隆或换机后该目录可能不存在或为空——这正常，不是缺件。

本机若备有以下样例，可走端到端验证（文件名仅供参考，以实际存在为准）：

- `实验六_WIFI密码破解操作文档.md` — 实验操作文档（内容参考）
- `实验六图片/` — 10 张实验截图（待嵌入）
- `网络与系统安全报告.docx` — 报告模板（含实验一至八，实验六部分为空）

典型场景：依操作文档与截图，为报告中实验六填充"实验过程及分析"和"实验结果总结"两部分。

**`test/` 为空时**：不要自行创建或臆造测试文件——向用户确认输入文件，或用用户指明的文件。

## 分支与提交规则

仓库维护两个长期分支，职责不同：

| 分支          | 内容                                                                                                          |
| ------------- | ------------------------------------------------------------------------------------------------------------- |
| `opencode`    | 全部改动：代码 + 文档（`.md`、`README.md`、`AGENTS.md`、`docs/`）+ `.gitignore`                               |
| `claude-code` | 仅`.py` 代码同步 + `test/` 占位文件；`CLAUDE.md` / `README.md` / `.gitignore` 由 `scripts/sync_docs.py` 从 opencode 对应文件生成 |

### 提交流程

1. **在 `opencode` 分支提交**：代码 + 文档一起提交，`git add` + `commit` + `push origin opencode`。
2. **同步到 `claude-code` 分支**：同步 `.py` 文件，再生成 `.md`。

   ```bash
   git checkout claude-code
   git checkout refs/heads/opencode -- <path1>.py <path2>.py ...   # 逐个或批量同步 .py
   python scripts/sync_docs.py                                     # 从 opencode 生成 CLAUDE.md + README.md + .gitignore
   git add CLAUDE.md README.md .gitignore
   git commit -m "sync code from opencode: <概述>"
   git push origin claude-code
   git checkout opencode                                           # 切回默认分支
   ```

3. **验证**：同步后做 `import` 冒烟测试（如 `python -c "from fill_template.docx import ..."`），确认代码可用。

### 要点

- **不要**手动编辑 `claude-code` 分支的 `CLAUDE.md` / `README.md`——它们由 `scripts/sync_docs.py` 生成，改 opencode 的 `AGENTS.md` / `README.md` 后重新生成即可。
- **不要**在 `claude-code` 分支做独立开发——所有改动先落在 `opencode`，再单向同步过去。
- 同步前先在 `opencode` 完成提交并 push，确保 `refs/heads/opencode` 是最新状态。
- 提交信息用简洁英文/中文混排，概述改动主题；同步提交以 `sync code from opencode:` 开头。

### 文档生成替换表

`scripts/sync_docs.py` 从 `refs/heads/opencode` 读取 `AGENTS.md` / `README.md`，做以下机械替换后写入 claude-code 分支的 `CLAUDE.md` / `README.md`：

| opencode（源）              | claude-code（生成）    | 说明              |
| --------------------------- | ---------------------- | ----------------- |
| `Claude Code skills`（散文）   | `Claude Code skills`   | 工具名            |
| `# OpenCode Skills`         | `# Claude Skills`      | README 标题       |
| `OpenCode 技能集合`         | `Claude Code 技能集合` | README 副标题     |
| `opencode 技能目录`         | `Claude Code 技能目录` | 安装注释          |
| `~/.config/opencode/skills` | `~/.claude/skills`     | 安装路径 + 结构树 |
| `├── AGENTS.md`（结构树）   | `├── CLAUDE.md`        | AI 指引文件名     |
| `*.opencode`                | `*.claude`             | 项目 AI 配置文件  |

> **注意**：分支名 `opencode` / `claude-code`（git 命令、表格中的分支引用）**不做替换**——两个分支共用同一套分支名。仅在散文中指代工具名时替换。新增涉及工具名的文本时，需同步更新此表。
