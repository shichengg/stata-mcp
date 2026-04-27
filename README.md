# Stata MCP Server

通过 Windows Stata Automation COM 让 Claude Code 控制 Stata GUI 的 MCP 服务器。

## 特性

- 按 do 文件绝对路径维护 Stata 会话。
- 同一个 do 文件复用同一个 Stata 窗口。
- 不同 do 文件打开不同 Stata 窗口。
- 无路径命令发送到最近一次运行的 do 文件会话。
- 支持 do 文件写入、读取、追加和运行。
- 支持读取 text log，供 Claude 分析结果。

## 目录结构

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

## 环境要求

- Windows
- Stata 18 MP，已注册 Automation COM
- Python 3.10+
- Python 包：`mcp`, `pywin32`

## 推荐安装位置

推荐把本项目放在 Stata 安装目录下：

```text
D:/Stata18/mcp
```

如果 Stata 安装在其他位置，也建议放在对应目录下，例如：

```text
C:/Program Files/Stata18/mcp
```

高级用户也可以放在任意稳定目录，只要 Claude Code MCP 配置指向正确的 `stata_mcp.py`。

## 安装步骤

1. 下载或 clone 本仓库到 Stata 安装目录下的 `mcp` 文件夹。
2. 安装 Python 依赖：

```powershell
pip install mcp pywin32
```

3. 注册 Stata Automation。按你的 Stata 安装路径修改命令：

```powershell
Start-Process -FilePath "D:\Stata18\StataMP-64.exe" -ArgumentList "/Register" -Wait
```

4. 验证 COM：

```powershell
python -c "import win32com.client; s=win32com.client.Dispatch('stata.StataOLEApp'); s.DoCommand('display \"COM OK\"')"
```

## 可配置环境变量

一般用户不需要设置。如果你的安装路径或 COM 名称特殊，可以设置：

| 变量 | 默认值 | 作用 |
|---|---|---|
| `STATA_MCP_DIR` | 自动识别项目根目录 | MCP 项目目录 |
| `STATA_DIR` | MCP 上一级目录 | Stata 安装目录 |
| `STATA_EXE` | `%STATA_DIR%/StataMP-64.exe` | README 注册提示用 |
| `STATA_COM_PROG_ID` | `stata.StataOLEApp` | Stata Automation ProgID |

## Claude Code MCP 配置

现有配置可以继续使用兼容启动器：

```json
"stata": {
  "command": "python",
  "args": ["D:\\Stata18\\mcp\\stata_mcp.py"]
}
```

也可以直接指向新主文件：

```json
"stata": {
  "command": "python",
  "args": ["D:\\Stata18\\mcp\\src\\stata_mcp\\server.py"]
}
```

## 工具

| 工具 | 功能 |
|---|---|
| `stata_run` | 在最近 do 文件会话中执行命令 |
| `stata_run_dofile` | 按 do 文件绝对路径运行/复用 Stata 会话 |
| `stata_write_dofile` | 写入 do 文件；普通文件名写到 `runtime/dofiles`，完整路径写到指定位置 |
| `stata_read_dofile` | 读取 do 文件内容 |
| `stata_append_dofile` | 追加 do 文件内容 |
| `stata_read_log` | 读取 text log |
| `stata_get_results` | 在最近会话显示 `return list` 或 `ereturn list` |
| `stata_get_data_info` | 在最近会话显示 `describe` |
| `stata_get_data_schema` | 读取当前数据集结构并返回给 Claude 分析 |
| `stata_status` | 查看 MCP 和 Stata 会话状态 |

## 读取当前数据结构和样本

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

## Log 策略

`stata_get_data_schema` 会覆盖生成自己的 schema 快照 log，因为它表示当前数据结构。

其他分析结果需要 Claude 读取时，应让 Stata 写 text log，然后使用 `stata_read_log` 读取。后续自动分析 log 推荐规则：同一 Claude/MCP 会话内追加；下次新会话首次运行同一 do 文件时覆盖。
