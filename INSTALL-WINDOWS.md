# Windows 安装包使用说明

本包为 Outlook Local 0.3.0，包含经典 Outlook 按需启动及 Windows COM 正文搜索优化。正文搜索默认首次直接扫描，后续按需建索引。直接解压安装即可，无需预装 Git、Node.js 或 npm。

这是联网安装包，不是完全离线包：本次分发的 ZIP 已附带 Windows x64 的 Node.js 24.22.0 和固定版本 OST/PST 解析器，无需到 GitHub 下载代码或解析器；安装时仍需联网从 npm registry 获取锁定依赖。安装器保留 Node 和解析器的 SHA-256 校验。Windows ARM64 也可安装，但需额外从 nodejs.org 下载对应 Node，且本地解析器需要 x64 仿真支持。自行运行打包脚本时，需要提供 `-NodeArchive` 和 `-ReaderArchive` 参数才会附带这些文件。

## 一、安装

1. 将 ZIP **全部解压**到普通本地文件夹；不要在压缩包预览窗口内直接运行。
2. 打开**经典 Outlook**，完成邮箱登录和同步。新版 Windows Outlook 不支持 COM；本包的文件后端也不支持其专有缓存。
3. 双击解压目录 `OutlookLocal-Setup` 中的 **install.cmd**，等待安装完成。无需以管理员身份运行。
4. 默认安装到 `%LOCALAPPDATA%\OutlookLocal`。安装器自动准备专用 Node、依赖、解析器、启动入口和本机 MCP 配置，不修改全局 Node 或系统 PATH。
5. 在安装目录运行 **Test-MCP.cmd**，应显示 `mcp: "ready"`，并列出 10 个工具及邮箱后端。打开 **Setup-MCP.cmd** 查看本机接入说明和配置。

若已有本项目源码服务占用默认端口 17379，可以复用原服务，或为新安装指定不同端口；不要删除原目录或强行终止不明进程。已有受管理的安装目录可以再次运行安装器升级，安装器会先停止该目录原服务。无关非空目录会被拒绝。

Agent 无需操作双击界面，可以在解压目录通过 PowerShell 安装：

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\install.ps1
```

需要独立目录和端口时：

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\install.ps1 -InstallDir "$env:LOCALAPPDATA\OutlookLocal-Test" -Port 17389
```

只需经典 Outlook COM 时，可追加 `-SkipLocalReader`。安装结束的邮箱检查默认可以启动经典 Outlook。首次启动若出现登录或配置提示，完成配置和同步后重新运行 `outlook.cmd health`。

## 二、让其他 Agent 使用

安装完成后，在安装目录找到：

- `outlook.cmd`：CLI 入口，使用包安装的专用 Node，无需全局 Node。
- `Start-Outlook.cmd` / `Stop-Outlook.cmd`：启动、停止本机后台服务。
- `mcp-config.json`：该电脑实际路径的 MCP 配置，不含令牌。
- `MCP-SETUP.txt`：带实际路径的中文接入说明。
- `Setup-MCP.cmd` / `Test-MCP.cmd`：查看配置、验证连接。

支持 JSON 配置的 Agent：将 `mcp-config.json` 中的 `mcpServers.outlook-local` 条目合并到客户端配置，保留已有服务。使用表单的客户端选择 **stdio**，照配置填写 `command`、`args` 和 `env`。不要把 `Start-Outlook.cmd` 当作 stdio 命令，不要复制另一台电脑的绝对路径或服务令牌。

重新连接 MCP 后，调用 `outlook_health` 检查 `connected: true`，再调用 `outlook_messages`，参数 `{"folder":"inbox","limit":1}`。同一电脑、同一用户的多个 Agent 可以共享服务。另一台电脑必须安装本包并使用它自己的 Outlook 邮箱；云端 Agent 不能直接访问这台电脑的 localhost。

不能配置 MCP 的 Agent 可以直接执行安装目录内的 CLI：

```powershell
.\outlook.cmd health
.\outlook.cmd today --folder inbox --limit 25
.\outlook.cmd search --folder inbox --query '项目' --search-body true --limit 25
```

正文搜索保留默认 `auto` 策略即可。结果的 `next_offset` 非空时继续翻页；`retry_required` 为 true 时重试当前页；`body_search_incomplete` 为 true 时说明结果不完整。原始邮件只读，不发送、删除、移动邮件或修改已读状态。

本流程默认不设置登录自启，重启或注销后运行 `Start-Outlook.cmd`；明确需要登录自启时安装参数可加 `-AutoStart`。更多查询与分页规则见包内 `docs/agent-install-guide.md` 第 5 节；**本包用户跳过其中从 GitHub 安装源码的步骤。**

## 三、可直接转发给 Agent

> 请使用我提供的 Outlook Local ZIP 安装包，不从 GitHub 克隆源码。先完整解压并阅读 INSTALL-WINDOWS.md，检查已有服务和端口，优先复用已有安装。需要新安装时运行 install.ps1，按需指定独立安装目录和端口；安装依赖需联网。然后验证 outlook.cmd health，读取本机生成的 MCP-SETUP.txt 和 mcp-config.json，保留其他配置后接入 MCP，并验证 9 个工具和邮箱连接。不能修改 MCP 配置时先通过 outlook.cmd 使用。报告安装路径、后端和验证结果，保留默认 auto 正文搜索策略，不上传邮箱数据或令牌。

不要转发安装后的目录：其中可能包含服务令牌、索引和下载附件。向其他人分发原始安装 ZIP 即可。
