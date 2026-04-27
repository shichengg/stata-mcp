# Stata MCP Server

通过 Windows Stata Automation COM 让 Claude Code 控制 Stata GUI 的 MCP 服务器。  
A Claude Code MCP server that controls the Stata GUI through Windows Stata Automation COM.

- [中文](#中文)
- [English](#english)

---

## 中文

### 功能

- 通过 Stata Automation COM 在真实 Stata GUI 窗口中运行命令。
- 按 do 文件绝对路径维护 Stata 会话：同一个 do 文件复用同一个 Stata 窗口。
- 不同 do 文件打开不同 Stata 窗口，方便同时浏览多个分析任务。
- 不带路径的命令会发送到最近一次运行的 do 文件会话。
- 支持 do 文件写入、读取、追加和运行。
- 支持读取 Stata text log，供 Claude 分析输出结果。
- 支持读取当前数据结构、缺失情况和样本预览，辅助 Claude 修改 do 文件。

### 演示效果

<video src="https://github.com/shichengg/stata-mcp/releases/download/v1.0.0/demo.mp4" controls="controls" style="max-width: 100%; height: auto;">
</video>

### 目录结构

```text
D:/Stata18/mcp/
├── README.md
├── pyproject.toml
├── .gitignore
├── stata_mcp.py              # 兼容启动器
├── src/
│   └── stata_mcp/
│       ├── __init__.py
│       └── server.py         # MCP 服务器主文件
├── runtime/
│   ├── dofiles/              # Claude/MCP 默认生成 do 文件
│   └── logs/                 # Claude/MCP 默认读取或生成 log
└── examples/                 # 示例 do 文件
```

### 环境要求

- Windows
- Stata 18 MP，已注册 Automation COM
- Python 3.10+
- Python 包：`mcp`, `pywin32`
- Claude Code

### 推荐安装位置

推荐把本项目放在 Stata 安装目录下的 `mcp` 文件夹，例如：

```text
D:/Stata18/mcp
```

如果 Stata 安装在其他位置，也建议放在对应目录下，例如：

```text
C:/Program Files/Stata18/mcp
```

高级用户也可以放在任意稳定目录。无论放在哪里，Claude Code MCP 配置中的脚本路径都必须指向实际的 `stata_mcp.py`。

### 安装步骤

1. 下载或 clone 本仓库到 Stata 安装目录下的 `mcp` 文件夹。

2. 安装 Python 依赖：

```powershell
pip install mcp pywin32
```

也可以在虚拟环境中安装：

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install mcp pywin32
```

3. 以管理员身份打开 PowerShell，注册 Stata Automation。按你的 Stata 安装路径修改命令：

```powershell
Start-Process -FilePath "D:\Stata18\StataMP-64.exe" -ArgumentList "/Register" -Wait
```

4. 验证 COM 是否可用：

```powershell
python -c "import win32com.client; s=win32com.client.Dispatch('stata.StataOLEApp'); s.DoCommand('display 12345')"
```

预期结果：Stata GUI 打开或被连接，并显示 `12345`。

### 添加到 Claude Code（推荐）

推荐使用 `claude mcp add`：

```powershell
claude mcp add stata -- python D:\Stata18\mcp\stata_mcp.py
```

如果安装在其他路径，请替换为实际路径。例如：

```powershell
claude mcp add stata -- python "C:\Program Files\Stata18\mcp\stata_mcp.py"
```

添加后重启 Claude Code，确认 MCP 工具列表中出现 `stata` 工具。

### 手动配置 Claude Code MCP

如果不使用 `claude mcp add`，也可以手动配置 MCP JSON：

```json
"stata": {
  "command": "python",
  "args": ["D:\\Stata18\\mcp\\stata_mcp.py"]
}
```

也可以直接指向主文件：

```json
"stata": {
  "command": "python",
  "args": ["D:\\Stata18\\mcp\\src\\stata_mcp\\server.py"]
}
```

通常建议使用兼容启动器 `stata_mcp.py`，这样后续项目内部结构调整时配置更稳定。

### 可配置环境变量

一般用户不需要设置。如果你的安装路径或 COM 名称特殊，可以设置：

| 变量 | 默认值 | 作用 |
|---|---|---|
| `STATA_MCP_DIR` | 自动识别项目根目录 | MCP 项目目录 |
| `STATA_DIR` | MCP 上一级目录 | Stata 安装目录 |
| `STATA_EXE` | `%STATA_DIR%/StataMP-64.exe` | README 注册提示用 |
| `STATA_COM_PROG_ID` | `stata.StataOLEApp` | Stata Automation ProgID |

### 工具

| 工具 | 功能 |
|---|---|
| `stata_run` | 在最近 do 文件会话中执行命令 |
| `stata_run_dofile` | 按 do 文件绝对路径运行/复用 Stata 会话 |
| `stata_write_dofile` | 写入 do 文件；普通文件名写到 `runtime/dofiles`，完整路径写到指定位置 |
| `stata_read_dofile` | 读取 do 文件内容 |
| `stata_append_dofile` | 追加 do 文件内容 |
| `stata_read_log` | 读取 Stata text log |
| `stata_install_package` | 在 Stata GUI 中安装外部包 |
| `stata_get_results` | 在 Stata GUI 中显示 `return list` 或 `ereturn list` |
| `stata_get_data_info` | 在 Stata GUI 中显示 `describe` |
| `stata_get_data_schema` | 读取当前数据结构、缺失摘要和样本预览并返回给 Claude |
| `stata_status` | 查看 MCP 和 Stata 会话状态 |

### do 文件和会话规则

- `stata_run_dofile(path)` 使用 do 文件绝对路径作为会话标识。
- 同一个 do 文件路径会复用同一个 Stata 窗口。
- 不同 do 文件路径会创建不同 Stata 窗口。
- `stata_run`、`stata_get_results`、`stata_get_data_info`、`stata_get_data_schema` 默认发送到最近一次运行的 do 文件会话。
- 如果还没有运行过 do 文件，无路径命令会提示先运行一个 do 文件。

### 读取当前数据结构和样本

`stata_get_data_schema` 会在最近 do 文件会话中生成一份 text log 快照，并把内容返回给 Claude。默认包含：

- `describe`
- `codebook, compact`
- `misstable summarize`
- `list in 1/20, abbreviate(20)`
- `notes`
- `label dir`

可选参数：

| 参数 | 默认值 | 说明 |
|---|---:|---|
| `sample_rows` | `20` | 样本预览行数，最大 1000 |
| `include_codebook` | `true` | 是否包含 compact codebook |
| `include_sample` | `true` | 是否包含样本预览 |
| `include_missing` | `true` | 是否包含缺失值摘要 |

这让 Claude 不仅能看到变量名、类型和标签，还能看到少量真实数据，从而更准确地辅助修改 do 文件。

### Log 策略

- `stata_get_data_schema` 会覆盖生成自己的 schema 快照 log，因为它表示当前数据结构。
- 其他分析结果需要 Claude 读取时，应让 Stata 写 text log，然后使用 `stata_read_log` 读取。
- 推荐规则：同一 Claude/MCP 会话内追加；下次新会话首次运行同一 do 文件时覆盖，避免 log 无限增长。

### 故障排查

#### `pywin32` 不可用

运行：

```powershell
pip install pywin32
```

#### COM 无法创建 Stata 实例

重新注册 Stata Automation：

```powershell
Start-Process -FilePath "D:\Stata18\StataMP-64.exe" -ArgumentList "/Register" -Wait
```

不要在 Git Bash 中直接运行带 `/Register` 的命令，因为 Git Bash 可能把 `/Register` 改写成路径。

#### `stata_run` 提示还没有 do 文件会话

先运行一个 do 文件：

```text
stata_run_dofile(path="D:/Stata18/mcp/examples/browse_test.do")
```

之后无路径命令会发送到这个最近会话。

#### Claude 不能直接看到 Stata GUI 结果

Claude 不能读取屏幕内容。需要让 Stata 写 text log，然后通过 `stata_read_log` 读取；或者使用 `stata_get_data_schema` 读取当前数据结构快照。

---

## English

### Features

- Runs Stata commands in the real Stata GUI through Stata Automation COM.
- Keeps one Stata session per absolute do-file path: the same do file reuses the same Stata window.
- Different do files open different Stata windows, so multiple analysis tasks can stay visible.
- Pathless commands are sent to the most recently used do-file session.
- Supports writing, reading, appending, and running do files.
- Supports reading Stata text logs so Claude can analyze command output.
- Supports reading the current data structure, missing-value summary, and sample rows to help Claude edit do files.

### Project layout

```text
D:/Stata18/mcp/
├── README.md
├── pyproject.toml
├── .gitignore
├── stata_mcp.py              # compatibility launcher
├── src/
│   └── stata_mcp/
│       ├── __init__.py
│       └── server.py         # main MCP server
├── runtime/
│   ├── dofiles/              # default do files generated by Claude/MCP
│   └── logs/                 # default logs generated or read by Claude/MCP
└── examples/                 # example do files
```

### Requirements

- Windows
- Stata 18 MP with Automation COM registered
- Python 3.10+
- Python packages: `mcp`, `pywin32`
- Claude Code

### Recommended install location

It is recommended to place this project inside your Stata installation directory, for example:

```text
D:/Stata18/mcp
```

If Stata is installed somewhere else, place the project under that Stata directory, for example:

```text
C:/Program Files/Stata18/mcp
```

Advanced users may place it in any stable directory. Wherever you install it, the Claude Code MCP configuration must point to the actual `stata_mcp.py` path.

### Installation

1. Download or clone this repository into the `mcp` folder under your Stata installation directory.

2. Install Python dependencies:

```powershell
pip install mcp pywin32
```

You may also use a virtual environment:

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install mcp pywin32
```

3. Open PowerShell as Administrator, then register Stata Automation. Adjust the path to your Stata installation:

```powershell
Start-Process -FilePath "D:\Stata18\StataMP-64.exe" -ArgumentList "/Register" -Wait
```

4. Verify COM:

```powershell
python -c "import win32com.client; s=win32com.client.Dispatch('stata.StataOLEApp'); s.DoCommand('display 12345')"
```

Expected result: the Stata GUI opens or is connected, and displays `12345`.

### Add to Claude Code (recommended)

Use `claude mcp add`:

```powershell
claude mcp add stata -- python D:\Stata18\mcp\stata_mcp.py
```

If you installed the project somewhere else, replace the path with your actual path. Example:

```powershell
claude mcp add stata -- python "C:\Program Files\Stata18\mcp\stata_mcp.py"
```

Restart Claude Code after adding the MCP server, then confirm that the `stata` tools are available.

### Manual Claude Code MCP configuration

If you do not use `claude mcp add`, you can configure MCP manually with JSON:

```json
"stata": {
  "command": "python",
  "args": ["D:\\Stata18\\mcp\\stata_mcp.py"]
}
```

You can also point directly to the main server file:

```json
"stata": {
  "command": "python",
  "args": ["D:\\Stata18\\mcp\\src\\stata_mcp\\server.py"]
}
```

The compatibility launcher `stata_mcp.py` is usually recommended because it keeps your Claude Code configuration stable if the internal project layout changes later.

### Environment variables

Most users do not need these. If your installation path or COM ProgID is unusual, you can configure:

| Variable | Default | Purpose |
|---|---|---|
| `STATA_MCP_DIR` | auto-detected project root | MCP project directory |
| `STATA_DIR` | parent of MCP directory | Stata installation directory |
| `STATA_EXE` | `%STATA_DIR%/StataMP-64.exe` | used in README registration examples |
| `STATA_COM_PROG_ID` | `stata.StataOLEApp` | Stata Automation ProgID |

### Tools

| Tool | Purpose |
|---|---|
| `stata_run` | Run commands in the most recent do-file session |
| `stata_run_dofile` | Run/reuse a Stata session by absolute do-file path |
| `stata_write_dofile` | Write a do file; simple names go to `runtime/dofiles`, absolute paths write to that location |
| `stata_read_dofile` | Read a do file |
| `stata_append_dofile` | Append content to a do file |
| `stata_read_log` | Read a Stata text log |
| `stata_install_package` | Install external packages in the Stata GUI |
| `stata_get_results` | Display `return list` or `ereturn list` in the Stata GUI |
| `stata_get_data_info` | Display `describe` in the Stata GUI |
| `stata_get_data_schema` | Return current data structure, missing summary, and sample rows to Claude |
| `stata_status` | Show MCP and Stata session status |

### Do-file and session rules

- `stata_run_dofile(path)` uses the absolute do-file path as the session key.
- The same do-file path reuses the same Stata window.
- Different do-file paths create different Stata windows.
- `stata_run`, `stata_get_results`, `stata_get_data_info`, and `stata_get_data_schema` are sent to the most recent do-file session by default.
- If no do file has been run yet, pathless commands will ask you to run a do file first.

### Reading current data structure and sample rows

`stata_get_data_schema` creates a text log snapshot in the most recent do-file session and returns it to Claude. By default, it includes:

- `describe`
- `codebook, compact`
- `misstable summarize`
- `list in 1/20, abbreviate(20)`
- `notes`
- `label dir`

Options:

| Option | Default | Description |
|---|---:|---|
| `sample_rows` | `20` | Number of sample rows, max 1000 |
| `include_codebook` | `true` | Include compact codebook |
| `include_sample` | `true` | Include sample preview |
| `include_missing` | `true` | Include missing-value summary |

This lets Claude see not only variable names, types, and labels, but also a small real-data sample, which helps it write better Stata code.

### Log strategy

- `stata_get_data_schema` overwrites its own schema snapshot log because it represents the current data state.
- For other analysis output, ask Stata to write a text log, then use `stata_read_log` to read it.
- Recommended rule: append within the same Claude/MCP session; overwrite the first time the same do file is run in a new session, so logs do not grow forever.

### Troubleshooting

#### `pywin32` is unavailable

Run:

```powershell
pip install pywin32
```

#### COM cannot create a Stata instance

Register Stata Automation again:

```powershell
Start-Process -FilePath "D:\Stata18\StataMP-64.exe" -ArgumentList "/Register" -Wait
```

Do not run the `/Register` command directly in Git Bash, because Git Bash may rewrite `/Register` as a path.

#### `stata_run` says there is no do-file session

Run a do file first:

```text
stata_run_dofile(path="D:/Stata18/mcp/examples/browse_test.do")
```

After that, pathless commands are sent to the most recent session.

#### Claude cannot directly see Stata GUI output

Claude cannot read screen contents directly. To let Claude analyze Stata output, write a Stata text log and read it with `stata_read_log`; or use `stata_get_data_schema` for a current data snapshot.
