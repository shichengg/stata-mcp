# Stata GUI MCP

通过 Windows Stata Automation COM，让 Claude Code、Codex、OpenCode、Claude Desktop、Cursor 和 VS Code 控制**可见的 Stata GUI**，并把每次执行的真实 text log 自动返回给 AI。

A Windows MCP server that controls a **visible Stata GUI** through Stata Automation COM and automatically returns the real Stata text log for every execution.

**[中文](#中文) | [English](#english)**

![Stata GUI MCP Demo](https://github.com/shichengg/stata-mcp/releases/download/v1.0.0/demo.gif)

---

## 中文

### 项目特点

- **真实 GUI**：命令在可见的 Stata GUI 中运行，而不是隐藏的 Python 模拟环境。
- **人工接管**：AI 执行后，用户可以直接在同一个 Stata 窗口中检查、修改和继续分析。
- **持久 Session**：同一个 `session_id` 复用同一个 Stata GUI；数据、宏、矩阵、估计结果和工作目录可继续使用。
- **自动返回结果**：运行若干行代码或完整 do 文件后，MCP 自动读取本次 text log 并返回给 AI。
- **完整代码块**：多行选择通过临时 do 文件整体执行，支持循环、程序定义和续行语法。
- **多任务窗口**：不同 `session_id` 可以维护不同的 Stata GUI 和项目日志。
- **项目本地记录**：session、do 文件和最新日志关系保存在项目的 `.stata-mcp/` 中。
- **不改原始脚本**：运行完整 do 文件时不会重写用户源文件。
- **pip 安装**：无需下载仓库或在 MCP 配置中填写源码路径。

### 适用范围

| 要求 | 说明 |
|---|---|
| 操作系统 | Windows 10/11 |
| Stata | 已安装并获得授权的 Stata 17/18/19，MP、SE 或 BE 均可 |
| Automation | Stata COM 接口已注册 |
| Python | 3.10–3.13 |
| MCP 传输 | 本地 stdio |

本项目当前**不支持 macOS/Linux**，也不使用 PyStata、Stata CLI 或 batch 后端。pip 包只包含 MCP Python 代码，不包含 Stata 软件、许可证或第三方 ado 包。

## 安装

### 1. 安装 Python 包

在 PowerShell 中运行：

```powershell
pip install stata-gui-mcp
```

pip 会同时安装 MCP Python SDK 和 Windows 所需的 `pywin32`。用户不需要克隆 GitHub 仓库。

### 2. 注册 Stata Automation

以**管理员身份**打开 PowerShell，按实际安装路径运行：

```powershell
Start-Process -FilePath "C:\Program Files\Stata18\StataMP-64.exe" -ArgumentList "/Register" -Wait
```

常见可执行文件名包括：

```text
StataMP-64.exe
StataSE-64.exe
StataBE-64.exe
```

请根据 Stata 版本、edition 和安装位置调整路径。不要直接在 Git Bash 中运行 `/Register`，因为 Git Bash 可能把它改写成文件路径。

正常启动时，MCP 直接调用已注册的 COM ProgID `stata.StataOLEApp`，不要求手工填写 Stata 安装路径。

## MCP 客户端配置

所有客户端都启动同一个本地 stdio 命令：

```text
stata-gui-mcp
```

| 客户端 | 推荐方式 | 配置键/文件 |
|---|---|---|
| **Claude Code** | `claude mcp add` | user scope 或 `.mcp.json` |
| Codex CLI | `codex mcp add` | `%USERPROFILE%\.codex\config.toml` |
| OpenCode | 配置文件或交互式添加 | `%USERPROFILE%\.config\opencode\opencode.json` |
| Claude Desktop | Desktop 配置 | `%APPDATA%\Claude\claude_desktop_config.json` |
| Cursor | `mcp.json` | `%USERPROFILE%\.cursor\mcp.json` 或 `.cursor\mcp.json` |
| VS Code | `mcp.json` | 用户 MCP 配置或 `.vscode\mcp.json` |

### Claude Code

推荐注册到 user scope，使当前用户的所有项目都能使用：

```powershell
claude mcp add --transport stdio --scope user stata -- stata-gui-mcp
```

命令含义：

```text
--transport stdio   使用本地标准输入输出传输
--scope user        当前用户的所有项目可用
stata               Claude Code 中显示的 server 名称
--                  分隔 Claude 参数和服务器命令
stata-gui-mcp       pip 安装的服务器命令
```

如果只希望项目内使用，在项目根目录创建 `.mcp.json`：

```json
{
  "mcpServers": {
    "stata": {
      "type": "stdio",
      "command": "stata-gui-mcp",
      "args": [],
      "env": {}
    }
  }
}
```

重新启动 Claude Code 或重新加载 MCP 后，在工具列表中应看到以 `mcp__stata__` 开头的工具。

### Codex CLI

通过命令添加本地 stdio server：

```powershell
codex mcp add stata -- stata-gui-mcp
```

也可以编辑用户配置 `%USERPROFILE%\.codex\config.toml`，或项目级 `.codex\config.toml`：

```toml
[mcp_servers.stata]
command = "stata-gui-mcp"
args = []
```

保存配置后重新启动 Codex。`codex mcp list` 可列出已注册的 server。

### OpenCode

OpenCode 可以交互式添加：

```powershell
opencode mcp add stata
```

选择 local server，并将命令设为 `stata-gui-mcp`。也可以编辑用户配置 `%USERPROFILE%\.config\opencode\opencode.json`，或项目根目录的 `opencode.json`：

```json
{
  "$schema": "https://opencode.ai/config.json",
  "mcp": {
    "stata": {
      "type": "local",
      "command": ["stata-gui-mcp"],
      "enabled": true
    }
  }
}
```

OpenCode 使用 `mcp` 键，并把本地命令写成数组；它与 Claude Desktop 的 `mcpServers` 格式不同。保存后重新启动 OpenCode，或运行 `opencode mcp list` 检查状态。

### Claude Desktop

打开以下 Windows 配置文件：

```text
%APPDATA%\Claude\claude_desktop_config.json
```

加入：

```json
{
  "mcpServers": {
    "stata": {
      "command": "stata-gui-mcp",
      "args": []
    }
  }
}
```

如果文件中已有其他 server，只添加 `stata` 项，不要覆盖其他配置。完全退出 Claude Desktop 后重新打开，再在 Developer/MCP 页面检查工具。

### Cursor

全局配置文件：

```text
%USERPROFILE%\.cursor\mcp.json
```

项目配置文件：

```text
<project>\.cursor\mcp.json
```

内容：

```json
{
  "mcpServers": {
    "stata": {
      "command": "stata-gui-mcp",
      "args": []
    }
  }
}
```

也可以从 Cursor 的 **Settings → Tools & MCP** 添加。保存后重新加载 Cursor 窗口并启用 `stata` server。

### VS Code

项目级配置文件：

```text
<project>\.vscode\mcp.json
```

内容：

```json
{
  "servers": {
    "stata": {
      "type": "stdio",
      "command": "stata-gui-mcp",
      "args": []
    }
  }
}
```

用户级配置可通过命令面板运行 **MCP: Open User Configuration** 打开。VS Code 使用 `servers` 键，不是 `mcpServers`。保存后在 MCP server 列表中启动 `stata`。

### 命令不在 PATH 中

如果客户端报告找不到 `stata-gui-mcp`，可以改用同一 Python 环境的模块入口：

```json
{
  "command": "python",
  "args": ["-m", "stata_mcp"]
}
```

Claude Code 对应命令：

```powershell
claude mcp add --transport stdio --scope user stata -- python -m stata_mcp
```

使用虚拟环境时，把 `python` 换成该环境解释器的绝对路径，例如：

```json
{
  "command": "D:\\Python\\python.exe",
  "args": ["-m", "stata_mcp"]
}
```

## 快速开始

### 运行一个完整 do 文件

```text
stata_run_dofile(
  path="D:/research/project/analysis.do",
  session_id="main",
  role="entry"
)
```

MCP 会：

1. 通过 COM 启动或连接一个可见的 Stata GUI；
2. 将工作目录切换到 do 文件目录；
3. 覆盖该 session 的最新运行日志；
4. 执行原始 do 文件；
5. 自动读取并返回本次 Stata 输出。

### 在同一个 session 中继续

```text
stata_run(commands="ereturn list\npredict double yhat\nsummarize yhat")
```

多行字符串作为一个临时 do 文件整体执行，执行完成后临时文件会删除。Stata 内存状态继续保留在同一个 GUI 中。

### 查看数据结构快照

```text
stata_get_data_schema(
  sample_rows=20,
  include_codebook=true,
  include_sample=true,
  include_missing=true
)
```

快照包括：

```stata
describe
codebook, compact
misstable summarize
list in 1/20, abbreviate(20)
notes
label dir
```

## Session 与日志

项目运行后创建：

```text
<project>/.stata-mcp/
├── cache/
│   ├── task_registry.json
│   ├── task_registry.json.lock
│   ├── run_<session>_<hash>.log
│   └── schema_<session>_<hash>.log
└── dofiles/
```

| 文件 | 内容 | 更新方式 |
|---|---|---|
| `task_registry.json` | session、entry/source/current do 和日志路径 | session 操作时更新 |
| `task_registry.json.lock` | 防止多个 MCP 进程同时更新 registry 时丢失数据 | registry 事务期间自动加锁 |
| `run_<session>_<hash>.log` | 最近一次命令代码块或完整 do 文件输出 | 每次 `stata_run`/`stata_run_dofile` 覆盖 |
| `schema_<session>_<hash>.log` | 最近一次数据结构快照 | 每次 `stata_get_data_schema` 覆盖 |
| `dofiles/` | MCP 通过相对文件名生成的 do 文件 | 写文件工具调用时更新 |

日志读取后不会自动删除。运行日志和 schema 日志相互独立。建议把 `.stata-mcp/` 加入分析项目自己的 `.gitignore`。

每个 session 固定使用一个项目缓存目录。不同 `session_id` 的文件名包含不同哈希，不会互相覆盖。

### 输出长度

运行工具默认最多把日志末尾 200,000 个字符返回给 AI。超过时，响应会标记 `output_truncated: true` 并保留完整 `log_path`。磁盘日志不会被截断。

可选环境变量：

```json
{
  "env": {
    "STATA_MCP_MAX_OUTPUT_CHARS": "300000"
  }
}
```

## 工具参考

| 工具 | 用途 |
|---|---|
| `stata_run` | 在最近 session 中将若干行代码作为一个代码块执行，覆盖并返回最新运行日志 |
| `stata_run_dofile` | 在指定 session 中运行完整 do 文件，覆盖并返回最新运行日志 |
| `stata_session` | `list`、`get`、`destroy` 或 `set_recent` session |
| `stata_write_dofile` | 写入 do 文件；相对名称写到项目 `.stata-mcp/dofiles/` |
| `stata_read_dofile` | 读取 do 文件 |
| `stata_append_dofile` | 向已有 do 文件追加代码 |
| `stata_read_log` | 读取最近运行日志或显式日志路径，支持 `full`、`core`、`dict` |
| `stata_install_package` | 在当前 Stata GUI 中执行 `ssc install` 或 `net install` |
| `stata_get_results` | 执行 `return list` 或 `ereturn list` 并返回结果 |
| `stata_get_data_info` | 执行 `describe` 并返回结果 |
| `stata_get_data_schema` | 通过独立 schema 日志返回结构、缺失摘要和样本 |
| `stata_status` | 显示 COM、内存 session 和项目 registry 状态 |

`stata_run_dofile` 仍接受旧版 `log_mode` 参数以保持客户端兼容，但 1.1 版始终使用 `replace`，保证每个 session 只保留最近一次运行输出。

## 与其他 Stata MCP 的比较

不同项目针对不同工作流，没有一个后端适合所有场景。

| 维度 | `stata-gui-mcp`（本项目） | [SepineTam/mcp-for-stata](https://github.com/SepineTam/mcp-for-stata) | [hanlulong/stata-mcp](https://github.com/hanlulong/stata-mcp) |
|---|---|---|---|
| 主要后端 | Windows COM Automation | Stata CLI/批处理导向 | PyStata worker |
| GUI | 可见 Stata GUI | 通常无 GUI | 通常无 Stata GUI |
| 人工接管 | 可直接接管同一窗口 | 不是主要目标 | 不是主要目标 |
| 状态持续 | 同一 COM GUI session 持续 | 取决于其执行方式 | 持久 PyStata worker |
| AI 获取输出 | 每次自动读取 session text log | CLI/log 输出 | PyStata输出和临时日志 |
| 选中多行代码 | 临时 do 文件整体执行 | 支持命令/文件执行 | 支持 |
| 多 session | 多个 COM GUI | 以对应项目当前实现为准 | 多 worker 进程 |
| 主要平台 | Windows | 跨平台 CLI 场景，具体支持见上游 | PyStata 支持的平台，具体支持见上游 |
| 更适合 | 希望看见、核查并手工继续 Stata | 自动化、批处理和服务器工作流 | 无 GUI 持久会话与深度 IDE 集成 |

本项目的主要优势是**可见 GUI + 人工接管 + AI 自动读取真实日志**。如果目标是 Linux 服务器、纯批处理或不显示 GUI，应选择 CLI/PyStata 类型方案。

## 环境变量

普通用户通常无需设置。

| 变量 | 默认值 | 用途 |
|---|---|---|
| `STATA_COM_PROG_ID` | `stata.StataOLEApp` | 非标准 COM ProgID |
| `STATA_EXE` | 未设置 | 特殊安装位置需要显式预启动 Stata 时的可执行文件路径；MCP 只保留一个存活的预启动进程，退出后会按需重启，Automation 会话仍由 COM ProgID 建立 |
| `STATA_MCP_MAX_OUTPUT_CHARS` | `200000` | 自动返回给 AI 的日志字符上限 |
| `STATA_MCP_DIR` | 安装包根目录 | 仅用于状态诊断；项目缓存仍跟随 do 文件 |

默认情况下，COM 注册信息负责定位 Stata，不会把 Python `site-packages` 的上一级误认为 Stata 安装目录。

## 安全说明

- MCP 获得的权限等同于当前 Windows 用户和 Stata GUI。
- 内置危险命令检查只是一层意外操作防护，不是完整安全沙箱。
- `stata_write_dofile` 和 `stata_append_dofile` 会修改明确指定的文件。
- `stata_install_package` 会访问外部源并修改 Stata ado 环境。
- 建议在运行前审阅 AI 生成的分析代码和文件路径。
- MCP 使用具名日志 `__stata_mcp_run` 和 `__stata_mcp_schema`，不会主动执行 `log close _all`。
- 如果用户 do 文件自身执行 `log close _all`，它也会关闭 MCP 日志，后续输出可能无法捕获；建议只关闭用户自己命名的日志。

## 故障排查

### COM 未注册或无法启动 Stata

以管理员 PowerShell 重新执行 `/Register`，然后完全退出旧 Stata 进程并重启 MCP 客户端。

### 找不到 `stata-gui-mcp`

客户端启动环境的 PATH 与安装 pip 包时的终端可能不同。使用上文的 `python -m stata_mcp` 配置，或填入正确 Python 解释器绝对路径。

### `stata_run` 提示没有 session

先运行一个绝对路径 do 文件：

```text
stata_run_dofile(path="D:/research/project/analysis.do", session_id="main")
```

后续无路径命令会发送到最近 session。也可以使用：

```text
stata_session(action="set_recent", session_id="main")
```

### Stata GUI 被关闭

MCP 检测到 COM 断开时会保留旧日志和 registry，重新初始化一个空 GUI，并明确提示内存状态已丢失。中断命令不会自动重放。

### AI 只看到日志末尾

检查响应中的 `output_truncated` 和 `log_path`，再调用 `stata_read_log` 读取完整文件，或调整 `STATA_MCP_MAX_OUTPUT_CHARS`。

## 开发

源码仓库：<https://github.com/shichengg/stata-mcp>

```powershell
git clone https://github.com/shichengg/stata-mcp
cd stata-mcp
python -m venv .venv
.\.venv\Scripts\python -m pip install -e ".[dev]"
.\.venv\Scripts\python -m pytest
```

欢迎提交 issue 和 pull request。

## License

[MIT](LICENSE)

---

## English

### Overview

`stata-gui-mcp` is a Windows MCP server that controls a visible Stata GUI through Stata Automation COM. It keeps Stata state alive per session and automatically returns the actual text log from every selected-code or do-file execution.

### Highlights

- Visible Stata GUI with direct manual takeover.
- Persistent Stata data, macros, matrices, estimates, and working directory within a session.
- Multi-line code is executed as one temporary do-file, not line by line.
- One replace-only latest run log per session, returned automatically to the AI.
- A separate replace-only schema snapshot log.
- Multiple `session_id` values can maintain independent Stata GUI windows.
- Original user do-files are never rewritten.
- Installable from pip without cloning the repository.

### Requirements

- Windows 10/11.
- A licensed Stata 17/18/19 installation (MP, SE, or BE).
- Stata Automation COM registered with `/Register`.
- Python 3.10–3.13.

The package does not include Stata or a Stata license. It does not currently provide macOS/Linux, PyStata, CLI, or batch backends.

## Installation

Install the package in PowerShell:

```powershell
pip install stata-gui-mcp
```

Register the executable that matches your installation from an elevated PowerShell:

```powershell
Start-Process -FilePath "C:\Program Files\Stata18\StataMP-64.exe" -ArgumentList "/Register" -Wait
```

Adjust the Stata version, edition, and path as needed. Normal MCP startup locates Stata through the registered COM ProgID; no source path or default `STATA_EXE` is required.

## MCP client setup

All clients start the same local stdio command: `stata-gui-mcp`.

### Claude Code

Claude Code is listed first because it is the primary supported setup:

```powershell
claude mcp add --transport stdio --scope user stata -- stata-gui-mcp
```

Project-scoped `.mcp.json`:

```json
{
  "mcpServers": {
    "stata": {
      "type": "stdio",
      "command": "stata-gui-mcp",
      "args": [],
      "env": {}
    }
  }
}
```

Restart Claude Code or reload MCP servers and look for tools prefixed with `mcp__stata__`.

### Codex CLI

```powershell
codex mcp add stata -- stata-gui-mcp
```

User configuration `%USERPROFILE%\.codex\config.toml` or project `.codex\config.toml`:

```toml
[mcp_servers.stata]
command = "stata-gui-mcp"
args = []
```

Restart Codex after saving. `codex mcp list` shows configured servers.

### OpenCode

Interactive setup:

```powershell
opencode mcp add stata
```

Select a local server and enter `stata-gui-mcp`, or edit `%USERPROFILE%\.config\opencode\opencode.json` / project `opencode.json`:

```json
{
  "$schema": "https://opencode.ai/config.json",
  "mcp": {
    "stata": {
      "type": "local",
      "command": ["stata-gui-mcp"],
      "enabled": true
    }
  }
}
```

Restart OpenCode or use `opencode mcp list`.

### Claude Desktop

Edit `%APPDATA%\Claude\claude_desktop_config.json`:

```json
{
  "mcpServers": {
    "stata": {
      "command": "stata-gui-mcp",
      "args": []
    }
  }
}
```

Merge the `stata` entry with existing servers, fully quit Claude Desktop, and reopen it.

### Cursor

Use global `%USERPROFILE%\.cursor\mcp.json` or project `.cursor\mcp.json`:

```json
{
  "mcpServers": {
    "stata": {
      "command": "stata-gui-mcp",
      "args": []
    }
  }
}
```

Alternatively use **Settings → Tools & MCP**, then reload the Cursor window.

### VS Code

Create project `.vscode\mcp.json`:

```json
{
  "servers": {
    "stata": {
      "type": "stdio",
      "command": "stata-gui-mcp",
      "args": []
    }
  }
}
```

For a user-level configuration, run **MCP: Open User Configuration** from the Command Palette. VS Code uses `servers`, not `mcpServers`.

### PATH fallback

If a client cannot find the console command, use the Python interpreter that installed the package:

```json
{
  "command": "python",
  "args": ["-m", "stata_mcp"]
}
```

Claude Code equivalent:

```powershell
claude mcp add --transport stdio --scope user stata -- python -m stata_mcp
```

Replace `python` with the virtual-environment interpreter's absolute path when applicable.

## Quick start

Start a persistent GUI session by running a do-file:

```text
stata_run_dofile(
  path="D:/research/project/analysis.do",
  session_id="main",
  role="entry"
)
```

Continue in the same in-memory Stata state:

```text
stata_run(commands="ereturn list\npredict double yhat\nsummarize yhat")
```

Capture the current dataset schema:

```text
stata_get_data_schema(sample_rows=20)
```

## Sessions and logs

```text
<project>/.stata-mcp/
├── cache/
│   ├── task_registry.json
│   ├── task_registry.json.lock
│   ├── run_<session>_<hash>.log
│   └── schema_<session>_<hash>.log
└── dofiles/
```

- `stata_run` and `stata_run_dofile` replace `run_<session>_<hash>.log` and return it automatically.
- `stata_get_data_schema` separately replaces `schema_<session>_<hash>.log`.
- `task_registry.json.lock` serializes registry updates across separate MCP processes.
- Reading either log does not delete it.
- Different session IDs use distinct hashed names.
- Add `.stata-mcp/` to the analysis project's `.gitignore`.

The automatic response includes at most the last 200,000 characters by default. If `output_truncated: true` appears, the complete file remains at `log_path`. Override the limit with `STATA_MCP_MAX_OUTPUT_CHARS`.

## Tool reference

| Tool | Purpose |
|---|---|
| `stata_run` | Execute a complete code block in the recent session and return its latest run log |
| `stata_run_dofile` | Run an original do-file in a session and return its latest run log |
| `stata_session` | List, inspect, destroy, or select the recent session |
| `stata_write_dofile` | Write a do-file |
| `stata_read_dofile` | Read a do-file |
| `stata_append_dofile` | Append Stata code to a do-file |
| `stata_read_log` | Read a log as full text, core text, or parsed command/result JSON |
| `stata_install_package` | Run `ssc install` or `net install` in the GUI |
| `stata_get_results` | Return `return list` or `ereturn list` output |
| `stata_get_data_info` | Return `describe` output |
| `stata_get_data_schema` | Return schema, missing summary, and sample through a separate log |
| `stata_status` | Show COM, in-memory session, and registry status |

`log_mode` remains accepted by `stata_run_dofile` for old clients, but version 1.1 always replaces the latest session run log.

## Comparison

| Dimension | `stata-gui-mcp` | [SepineTam/mcp-for-stata](https://github.com/SepineTam/mcp-for-stata) | [hanlulong/stata-mcp](https://github.com/hanlulong/stata-mcp) |
|---|---|---|---|
| Main backend | Windows COM Automation | Stata CLI/batch-oriented | PyStata workers |
| Visible Stata GUI | Yes | Usually no | Usually no |
| Manual takeover | Same GUI window | Not the primary goal | Not the primary goal |
| Persistent state | Persistent COM GUI session | Depends on execution mode | Persistent PyStata worker |
| AI output | Automatically returned session text log | CLI/log output | PyStata output and temporary logs |
| Multi-session | Multiple COM GUI sessions | See upstream implementation | Multiple worker processes |
| Primary fit | Visible, reviewable, human-in-the-loop Stata | Automation and batch workflows | Headless persistent IDE integration |

The differentiator is **visible GUI + manual takeover + automatic real-log return**. Choose a CLI/PyStata project when a server, headless, or cross-platform workflow matters more.

## Environment variables

| Variable | Default | Purpose |
|---|---|---|
| `STATA_COM_PROG_ID` | `stata.StataOLEApp` | Override the registered COM ProgID |
| `STATA_EXE` | unset | Explicitly prelaunch a non-standard Stata executable; MCP tracks one live prelaunched process and relaunches it after exit, while Automation sessions are still created through the COM ProgID |
| `STATA_MCP_MAX_OUTPUT_CHARS` | `200000` | Maximum log characters returned automatically |
| `STATA_MCP_DIR` | package root | Status diagnostics only |

## Security and limitations

- The server runs with the current Windows user's permissions.
- The dangerous-command check is a guardrail, not a sandbox.
- File-writing tools modify explicitly selected files.
- Package installation changes the Stata ado environment and may use the network.
- Review generated analysis code and paths before execution.
- MCP uses named logs and never intentionally issues `log close _all`.
- A user do-file containing `log close _all` can still close the MCP log and truncate capture; close only user-named logs instead.

## Troubleshooting

- **COM startup fails:** rerun `/Register` from elevated PowerShell and restart Stata/MCP clients.
- **Command not found:** use the `python -m stata_mcp` fallback with the correct interpreter.
- **No current session:** call `stata_run_dofile` with an absolute path first.
- **GUI was closed:** MCP preserves old logs/registry and reports that the newly initialized GUI has empty memory; it never silently replays the interrupted command.
- **Output was truncated:** read `log_path` with `stata_read_log` or raise `STATA_MCP_MAX_OUTPUT_CHARS`.

## Development

```powershell
git clone https://github.com/shichengg/stata-mcp
cd stata-mcp
python -m venv .venv
.\.venv\Scripts\python -m pip install -e ".[dev]"
.\.venv\Scripts\python -m pytest
```

Repository: <https://github.com/shichengg/stata-mcp>

## License

[MIT](LICENSE)
