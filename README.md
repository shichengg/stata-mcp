# Stata MCP Server

通过 Windows Stata Automation COM 让 Claude Code 控制 Stata GUI 的 MCP 服务器。  
A Claude Code MCP server that controls the Stata GUI through Windows Stata Automation COM.

**[中文](#中文) | [English](#english)**

---

## 中文

### 🎯 核心功能

- **真实 GUI 控制**：通过 Stata Automation COM 在已打开的 Stata GUI 窗口中直接运行命令
- **Session 会话管理**：按 do 文件路径或自定义 session_id 维护独立的 Stata 窗口，同一任务复用同一窗口
- **多窗口支持**：不同 do 文件/session_id 自动创建不同 Stata 窗口，方便同时监控多个分析任务
- **灵活命令执行**：支持无路径命令自动发送到最近会话，也支持在指定 session 中执行
- **完整 Do 文件管理**：写入、读取、追加、运行 do 文件，支持项目本地存储
- **结构化数据获取**：读取数据描述、缺失摘要、样本预览、变量标签等，辅助 Claude 精准修改分析代码
- **Log 读取与解析**：支持读取 Stata text log，可选输出为完整文本、简洁格式或 JSON 命令-结果对
- **安全防护**：内置危险命令黑名单，防止意外删除文件或执行 shell 命令
- **任务追踪**：本地记录 session、do 文件和 log 的关系，支持任务级别的复现

### 📺 演示效果

![Stata MCP Demo](https://github.com/shichengg/stata-mcp/releases/download/v1.0.0/demo.gif)

### 📦 项目结构

```text
D:/Stata18/mcp/
├── README.md                 # 项目文档
├── pyproject.toml           # Python 项目配置
├── .gitignore
├── stata_mcp.py             # 兼容启动器（推荐配置路径）
├── src/
│   └── stata_mcp/
│       ├── __init__.py
│       └── server.py        # MCP 服务器主程序
└── assets/                  # 项目资源

# 运行时自动创建的项目本地缓存（每个项目目录）
project-dir/.stata-mcp/
├── cache/
│   └── task_registry.json   # Session/do/log 关系记录
└── dofiles/                 # 生成的 do 文件
```

### 🔧 环境要求

| 要求 | 版本 | 备注 |
|-----|------|------|
| **操作系统** | Windows | 必须。依赖 Windows COM 和 pywin32 |
| **Stata** | 18 MP | 须已注册 Automation COM 接口 |
| **Python** | 3.10+ | 推荐 3.11+ |
| **包** | mcp, pywin32 | 通过 pip 安装 |
| **Claude 编辑器** | Latest | 已安装 Claude Code 扩展 |

### 📥 安装步骤

#### 1. 克隆或下载项目

推荐放在 Stata 安装目录下的 `mcp` 文件夹：

```powershell
# 方式一：克隆到 Stata 目录
cd D:\Stata18
git clone https://github.com/shichengg/stata-mcp mcp

# 方式二：手动下载并解压
# 下载 ZIP 后解压到 D:\Stata18\mcp
```

高级用户也可放在任意稳定目录，但后续 Claude Code 配置需指向实际 `stata_mcp.py` 路径。

#### 2. 安装 Python 依赖

**推荐方案（使用虚拟环境）：**

```powershell
cd D:\Stata18\mcp
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install mcp pywin32
```

**快速方案（全局安装）：**

```powershell
pip install mcp pywin32
```

#### 3. 注册 Stata Automation COM

必须以 **管理员身份** 运行 PowerShell，执行：

```powershell
Start-Process -FilePath "D:\Stata18\StataMP-64.exe" -ArgumentList "/Register" -Wait
```

> ⚠️ **注意**：
> - 替换 `D:\Stata18\StataMP-64.exe` 为你的 Stata 安装路径
> - 不要在 Git Bash 中运行此命令，因为 Git Bash 可能改写 `/Register`
> - Windows 7/8 用户可能需要禁用 UAC 或使用 `runas` 命令

#### 4. 验证 COM 接口可用

运行以下命令验证 Stata Automation 是否已正确注册：

```powershell
python -c "import win32com.client; s = win32com.client.Dispatch('stata.StataOLEApp'); s.DoCommand('display 12345')"
```

预期结果：Stata GUI 打开或被连接，并在结果窗口显示 `12345`。

#### 5. 配置 Claude Code

**推荐方案（使用 `claude mcp add` 命令）：**

在 PowerShell 中运行以下命令：

```powershell
claude mcp add stata --scope user -- python D:\Stata18\mcp\stata_mcp.py
```

> 💡 **提示**：
> - `--scope user` 将配置保存到用户级别（可选，默认也是用户级别）
> - 如果 Stata 安装在其他路径，请相应调整：
> ```powershell
> claude mcp add stata --scope user -- python "C:\Program Files\Stata18\mcp\stata_mcp.py"
> ```
> - 也可在 Claude Code 中直接运行此命令

配置后重启 Claude Code，确认工具列表中出现 `stata` MCP。

**替代方案（手动编辑 Claude Code 配置）：**

如果 `claude mcp add` 不可用，可手动编辑 Claude Code 的 MCP 配置文件，添加：

```json
"stata": {
  "command": "python",
  "args": ["D:\\Stata18\\mcp\\stata_mcp.py"]
}
```

> 💡 **配置文件位置**：通常在 `%APPDATA%\Code\User\settings.json` 或通过 Claude Code 的设置界面配置

### 🌍 环境变量（可选）

通常用户无需设置。如果有特殊需求（如非标准安装路径），可在启动 MCP 前设置以下环境变量：

| 环境变量 | 默认值 | 说明 |
|---------|--------|------|
| `STATA_MCP_DIR` | 自动检测项目根目录 | MCP 项目目录 |
| `STATA_DIR` | MCP 上一级目录 | Stata 安装目录 |
| `STATA_EXE` | `%STATA_DIR%/StataMP-64.exe` | Stata 可执行文件路径 |
| `STATA_COM_PROG_ID` | `stata.StataOLEApp` | Stata Automation COM ProgID |

示例：

```powershell
$env:STATA_DIR = "C:\Program Files\Stata18"
$env:STATA_COM_PROG_ID = "stata.StataOLEApp"
python D:\Stata18\mcp\stata_mcp.py
```


---

## 📚 工具详解

Stata MCP 服务器提供以下工具供 Claude Code 调用：

### 核心工具

#### 1. `stata_run` - 在最近 Session 中执行命令

在最近使用的 Stata 会话中执行一条或多条命令。如果还没有运行过 do 文件，会提示先运行一个。

**参数：**
- `commands` (string, required): 要执行的 Stata 命令，多行用换行符分隔

#### 2. `stata_run_dofile` - 运行 Do 文件（Session 管理）

运行一个 do 文件，并按 do 路径或 session_id 维护独立的 Stata 窗口。

**参数：**
- `path` (string, required): do 文件的绝对路径
- `session_id` (string, optional): 自定义 session ID
- `role` (string, optional): do 在 session 中的角色 (entry/source/auxiliary)，默认 entry
- `log_mode` (string, optional): 日志模式 (replace/append)，默认 replace

#### 3. `stata_session` - 管理 Stata 会话

列出、查询、销毁或切换 Stata 会话。

**参数：**
- `action` (string, required): list/get/destroy/set_recent
- `session_id` (string, optional): session ID

#### 4. `stata_write_dofile` - 写入 Do 文件

生成并保存 Stata do 文件到磁盘。

**参数：**
- `content` (string, required): do 文件内容
- `filename` (string, optional): 文件名或完整路径

#### 5. `stata_read_dofile` - 读取 Do 文件

读取 do 文件内容。

**参数：**
- `path` (string, required): do 文件的完整路径

#### 6. `stata_append_dofile` - 追加 Do 文件内容

向已有 do 文件追加代码。

**参数：**
- `path` (string, required): do 文件的完整路径
- `content` (string, required): 要追加的代码

#### 7. `stata_read_log` - 读取 Stata Log

读取 Stata text log 文件。

**参数：**
- `path` (string, optional): log 文件路径，留空则读取最近 session 的日志
- `tail_lines` (integer, optional): 只读取最后 N 行
- `output_format` (string, optional): 输出格式 (full/core/dict)，默认 full

#### 8. `stata_install_package` - 安装包

在 Stata GUI 中安装外部包（如 estout、ivreg2 等）。

**参数：**
- `package` (string, required): 包名称
- `source` (string, optional): 安装来源 (ssc/net)，默认 ssc

#### 9. `stata_get_results` - 获取存储结果

显示 r() 或 e() 存储结果。

**参数：**
- `type` (string, required): 结果类型 (r/e)

#### 10. `stata_get_data_info` - 查看数据信息

显示当前数据集的基本信息。

#### 11. `stata_get_data_schema` - 获取数据结构和样本

读取当前数据集结构、缺失摘要和样本预览。

**参数：**
- `sample_rows` (integer, optional): 样本预览行数，默认 20
- `include_codebook` (boolean, optional): 是否包含 codebook，默认 true
- `include_sample` (boolean, optional): 是否包含样本预览，默认 true
- `include_missing` (boolean, optional): 是否包含缺失值摘要，默认 true

#### 12. `stata_status` - 检查服务器状态

查看 MCP 服务器状态和活跃会话。

---

## 💡 最佳实践

### 快速数据探索

```
# 1. 加载数据
stata_run_dofile(path="D:/project/load_data.do", session_id="explore")

# 2. 查看数据结构
stata_get_data_schema(sample_rows=50)

# 3. 运行探索性分析
stata_run(commands="summarize\ncorrelation")
```

### 多步骤回归分析

```
# 1. 数据清洗
stata_run_dofile(path="D:/analysis/01_clean.do", session_id="proj_v1", role="source")

# 2. 主分析
stata_run_dofile(path="D:/analysis/02_model.do", session_id="proj_v1", role="entry")

# 3. 获取系数
stata_get_results(type="e")
```

### 调试和迭代

```
# 首次运行
stata_run_dofile(path="D:/project/analysis.do", session_id="debug_v1")

# 检查结果
stata_read_log(output_format="dict")

# 修改代码
stata_append_dofile(path="D:/project/analysis.do", content="* Fixed: ...")

# 重新运行（同一 session，append 模式）
stata_run_dofile(path="D:/project/analysis.do", session_id="debug_v1", log_mode="append")
```

---

## 🔍 故障排查

### ❌ `pywin32` 不可用

```powershell
pip install --upgrade pywin32
```

### ❌ COM 无法创建 Stata 实例

以 **管理员身份** 重新注册 Stata Automation：

```powershell
Start-Process -FilePath "D:\Stata18\StataMP-64.exe" -ArgumentList "/Register" -Wait
```

### ❌ `stata_run` 提示没有 do 文件会话

先运行一个 do 文件：

```
stata_run_dofile(path="D:/project/analysis.do")
```

之后无路径命令会发送到这个最近会话。

### ❌ Claude 不能看到 Stata GUI 结果

Claude 不能读取屏幕内容。使用以下方案之一：

1. 让 Stata 写 text log，然后通过 `stata_read_log` 读取
2. 使用 `stata_get_data_schema()` 读取当前数据结构快照
3. 使用 `stata_get_results()` 获取存储结果

---

## English

### 🎯 Key Features

- **Real GUI Control**: Execute commands directly in the active Stata GUI via Stata Automation COM
- **Session Management**: Maintain independent Stata windows per do-file path or custom session_id
- **Multi-Window Support**: Different do files/session_ids create separate Stata windows
- **Flexible Command Execution**: Send pathless commands to recent session or specific session
- **Complete Do-File Management**: Write, read, append, and run do files with project-local storage
- **Structured Data Retrieval**: Get data descriptions, missing-value summaries, sample previews
- **Log Reading and Parsing**: Read Stata text logs in multiple formats (full/clean/JSON)
- **Safety Protections**: Built-in blacklist for dangerous commands
- **Task Tracking**: Local records of session/do/log relationships for reproducibility

### 📺 Demo

![Stata MCP Demo](https://github.com/shichengg/stata-mcp/releases/download/v1.0.0/demo.gif)

### 📦 Project Structure

```text
D:/Stata18/mcp/
├── README.md                 # Documentation
├── pyproject.toml           # Python configuration
├── .gitignore
├── stata_mcp.py             # Compatibility launcher
├── src/
│   └── stata_mcp/
│       ├── __init__.py
│       └── server.py        # Main MCP server
└── assets/                  # Resources

project-dir/.stata-mcp/
├── cache/
│   └── task_registry.json
└── dofiles/
```

### 🔧 Requirements

| Requirement | Version | Notes |
|-------------|---------|-------|
| **OS** | Windows | Required for Windows COM |
| **Stata** | 18 MP | Must have Automation COM registered |
| **Python** | 3.10+ | Recommended 3.11+ |
| **Packages** | mcp, pywin32 | Install via pip |
| **Claude Editor** | Latest | With Claude Code extension |

### 📥 Installation

#### 1. Clone or Download

```powershell
cd D:\Stata18
git clone https://github.com/shichengg/stata-mcp mcp
```

#### 2. Install Dependencies

```powershell
cd D:\Stata18\mcp
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install mcp pywin32
```

#### 3. Register Stata Automation COM

Run PowerShell as **Administrator**:

```powershell
Start-Process -FilePath "D:\Stata18\StataMP-64.exe" -ArgumentList "/Register" -Wait
```

#### 4. Verify COM Interface

```powershell
python -c "import win32com.client; s = win32com.client.Dispatch('stata.StataOLEApp'); s.DoCommand('display 12345')"
```

#### 5. Configure Claude Code

Run the following command in PowerShell:

```powershell
claude mcp add stata --scope user -- python D:\Stata18\mcp\stata_mcp.py
```

> 💡 **Tips**:
> - `--scope user` saves the configuration to user level (optional, default is user level)
> - If Stata is installed elsewhere, adjust the path accordingly:
> ```powershell
> claude mcp add stata --scope user -- python "C:\Program Files\Stata18\mcp\stata_mcp.py"
> ```
> - You can also run this command directly in Claude Code

Restart Claude Code and verify `stata` appears in the MCP tools list.

### 🌍 Environment Variables (Optional)

| Variable | Default | Purpose |
|----------|---------|---------|
| `STATA_MCP_DIR` | Auto-detect | MCP project directory |
| `STATA_DIR` | Parent of MCP | Stata installation directory |
| `STATA_EXE` | Default path | Stata executable path |
| `STATA_COM_PROG_ID` | stata.StataOLEApp | COM ProgID |

---

## 📚 Tools Reference

Stata MCP provides 12 main tools:

1. **stata_run** - Execute commands in recent session
2. **stata_run_dofile** - Run do file with session management
3. **stata_session** - Manage sessions (list/get/destroy/set_recent)
4. **stata_write_dofile** - Write do file to disk
5. **stata_read_dofile** - Read do file content
6. **stata_append_dofile** - Append to do file
7. **stata_read_log** - Read Stata log file
8. **stata_install_package** - Install packages (estout, ivreg2, etc.)
9. **stata_get_results** - Get r() or e() stored results
10. **stata_get_data_info** - View data info
11. **stata_get_data_schema** - Get data structure and sample
12. **stata_status** - Check server status

---

## 💡 Best Practices

### Quick Data Exploration

```
stata_run_dofile(path="D:/project/load_data.do", session_id="explore")
stata_get_data_schema(sample_rows=50)
stata_run(commands="summarize\ncorrelation")
```

### Multi-Step Regression

```
stata_run_dofile(path="D:/analysis/01_clean.do", session_id="proj_v1", role="source")
stata_run_dofile(path="D:/analysis/02_model.do", session_id="proj_v1", role="entry")
stata_get_results(type="e")
```

### Debug and Iterate

```
stata_run_dofile(path="D:/project/analysis.do", session_id="debug_v1")
stata_read_log(output_format="dict")
stata_append_dofile(path="D:/project/analysis.do", content="* Fixed: ...")
stata_run_dofile(path="D:/project/analysis.do", session_id="debug_v1", log_mode="append")
```

---

## 🔍 Troubleshooting

### ❌ `pywin32` not available

```powershell
pip install --upgrade pywin32
```

### ❌ Stata COM fails to initialize

Re-register Stata Automation as **Administrator**:

```powershell
Start-Process -FilePath "D:\Stata18\StataMP-64.exe" -ArgumentList "/Register" -Wait
```

### ❌ "No do file session" error

Run a do file first:

```
stata_run_dofile(path="D:/project/analysis.do")
```

### ❌ Claude can't see Stata results

Use `stata_read_log()` or `stata_get_data_schema()`:

```
stata_read_log(output_format="dict")
stata_get_data_schema()
```

---

## 📄 License

Open source. See LICENSE file for details.

## 🤝 Contributing

Contributions welcome! Please open issues and pull requests on GitHub.

---

**Repository**: [github.com/shichengg/stata-mcp](https://github.com/shichengg/stata-mcp)

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
