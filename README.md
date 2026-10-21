# Outlook CLI

让本机 Agent 通过统一的 **CLI 和 MCP** 读取 Windows、macOS 上的 Outlook 邮件。支持邮件检索、中文正文读取、附件下载，以及按日期查询。

**原始邮件只读**：不发送、删除、移动邮件，也不修改已读状态。邮箱登录和同步由 Outlook 完成；本项目不内置 AI 模型，不需要模型 API Key 或 Microsoft Graph OAuth。

> 跨平台源码版本为 **v0.3.0**，源码安装使用 `main` 分支。Windows 已发布的 v0.2.1 安装包不包含 Mac 读取器和 `outlook_today` 工具。

## 项目背景

希望让本机 Codex 和其他 Agent 帮忙查找邮件、总结当天工作、阅读正文及附件，同时复用 Outlook 已登录、已同步的本地数据。

项目最初在 Windows 上通过经典 Outlook COM 和本地 OST/PST 文件读取邮件，后来增加 macOS Outlook 缓存读取器，并统一 CLI 与 MCP 接口。Agent 使用同一套工具，服务根据操作系统选择读取方式。

Mac 读取器重点解决中文主题解码、邮件与正文和附件的准确关联，以及收件箱、已发送与子目录的范围识别。无法可靠关联的内容会明确报告缺失或未关联，不按相似主题猜测正文。

```text
Codex / 其他 Agent / 命令行
          │
     CLI 或 MCP
          │
     本机后台服务
          │
   ┌──────┴──────┐
 Windows       macOS
 COM / OST/PST  HxStore + Osa 缓存证据
```

## 支持范围

| 项目 | Windows | macOS |
|---|---|---|
| 读取来源 | 已运行的经典 Outlook COM；不可用时读取支持的本地 OST/PST | Outlook 已同步到本机的 HxStore 缓存 |
| 文件夹 | 原有 COM / OST/PST 后端可见目录 | 已确认的收件箱、已发送及其全部已知子目录 |
| 已删除邮件 | 原有后端支持；需要时显式选择目录 | 排除已删除及其子目录、归属不明邮件 |
| 日期筛选 | 接收时间 `received_at` | 接收时间 `received_at`（验证后的 `LastDeliveryTime`） |
| 正文 | 分页文本读取，视后端支持 HTML | 中文文本提取与分页，缺失状态明确返回 |
| 搜索 | 主题、发件人，可选正文 | 主题、发件人 |
| 附件 | 普通文件附件，单个最大 10 MiB | 可可靠关联的缓存附件，单个最大 10 MiB |

共享工具每次查询一个文件夹。读取所有子目录需要逐级遍历，不能把一次 `today` 或 `search` 的结果当成全部邮箱。Linux 可运行协议和合成缓存测试，暂不提供实际邮箱后端。

## 安装

**按默认用户目录安装和使用，无需管理员权限。** Windows 安装包使用当前用户目录和专用 Node.js；macOS 源码应放在自己的用户目录，无需 `sudo`。Mac 首次安装 Node.js、Python 是否需要管理员密码取决于所选安装方式。Windows COM 服务应与 Outlook 使用同一用户和权限级别，建议均以普通用户运行；企业电脑的访问、下载和脚本策略仍需遵循。

### macOS：从源码安装

准备：Git、Node.js **22+**、Python **3.11+**，以及已经登录并同步邮件的 Outlook。Python 读取器使用标准库，无需额外安装 pip 依赖。

```sh
git clone --branch main https://github.com/guanglinhuang99/outlook-cli.git
cd outlook-cli
npm ci --ignore-scripts

node bin/outlook.mjs start
node bin/outlook.mjs health
node bin/outlook.mjs today --folder inbox --limit 25
```

`start` 启动后台服务，`health` 检查邮件后端。Mac 首次建立邮件头、目录和接收时间索引仍需扫描本地缓存，但不会预先复制全部历史正文；正文在 `read` 时按需校验并保存，后续分页复用已保存内容。后续命令在项目目录执行；也可使用 `sh outlook.sh` 代替 `node bin/outlook.mjs`。

默认自动发现唯一的 Outlook profile。存在多个 profile 时，先显式设置再启动：

```sh
export OUTLOOK_MAC_PROFILE='/absolute/path/to/Outlook profile'
node bin/outlook.mjs start
```

若服务已运行，应使用 `restart` 应用环境变量变更。自定义 Python、缓存路径及原生 Python MCP 用法见 [macOS 说明](docs/macos.md)。

### Windows：安装包

1. 在 [Releases](https://github.com/guanglinhuang99/outlook-cli/releases) 下载 `outlook-local-0.2.1-setup.zip`，解压全部文件。
2. 双击 `install.cmd`，默认安装到 `%LOCALAPPDATA%\OutlookLocal`，无需管理员身份。
3. 安装器会准备专用 Node.js、依赖和 OST/PST 解析器，并生成启动入口和 MCP 配置。
4. 在安装目录运行 `outlook.cmd health` 检查后端，双击 `Setup-MCP.cmd` 查看接入说明，`Test-MCP.cmd` 验证连接。

已发布安装包提供 8 个 MCP 工具。需要本 README 中新增的 `today` 命令和 9 个工具时，请使用下面的 v0.3.0 源码安装方式。

### Windows：从源码安装

准备：Git、Node.js **22+**、Windows PowerShell **5.1**，以及已配置的经典 Outlook 或支持的本地 OST/PST 文件。

```powershell
git clone --branch main https://github.com/guanglinhuang99/outlook-cli.git
cd outlook-cli
npm ci --ignore-scripts
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\setup-reader.ps1

.\outlook.cmd start
.\outlook.cmd health
.\outlook.cmd today --folder inbox --limit 25
```

默认优先连接已运行的经典 Outlook，仅在 COM 不可用时转到本地文件读取。服务不会主动启动 Outlook。新版 Windows Outlook 不支持 COM，其自身缓存也不属于本项目 OST/PST 后端的支持范围。

COM 服务应与 Outlook 使用相同的 Windows 用户、交互式会话和权限级别。手动指定 OST/PST、安装参数和本地模式限制见 [详细参考](docs/usage-reference.md)。

## 日常使用

下列 `node bin/outlook.mjs` 命令在两个平台通用。Windows 也可换成 `.\outlook.cmd`。成功结果为 JSON；错误写入 stderr，退出码非零。

### 今天和最近的邮件

```sh
# 今天：收件箱与已发送分别查询
node bin/outlook.mjs today --folder inbox --limit 25
node bin/outlook.mjs today --folder sent --limit 25

# 最近的邮件元数据
node bin/outlook.mjs messages --folder inbox --limit 20

# 指定日期
node bin/outlook.mjs today --date 2026-10-20 --folder inbox
```

“今天”按 `Asia/Singapore` 的零点至次日零点计算。Windows 和 Mac 均按接收时间 `received_at` 筛选。Mac 使用本地 Outlook 的 `LastDeliveryTime`（最后投递时间），并与原始时间戳交叉核对；发送时间仅作为独立字段返回。缺少或存在冲突的接收时间不会用发送时间替代：日期筛选排除这些记录，并通过 `received_time_unknown` 和 `date_filter_incomplete` 提示覆盖不完整。翻页时传回第一页的 `date` 和 `next_offset`，避免跨午夜切换日期。

### 查找邮件并阅读正文

```sh
node bin/outlook.mjs search --folder inbox --query '项目' --limit 25
node bin/outlook.mjs read --id '<message_id>' --max-chars 20000

# 正文未读完时，使用返回的 body.next_offset 继续
node bin/outlook.mjs read --id '<message_id>' --body-offset 20000 --max-chars 20000
```

将示例占位符替换成响应中的真实 ID。检索结果的 `next_offset` 非空时，保持原筛选条件并用 `--offset` 继续；空白页不代表检索结束。Mac 暂不支持未读筛选、正文搜索或 HTML 原文请求。

### 读取收件箱、已发送下的全部子目录

```sh
node bin/outlook.mjs stores
node bin/outlook.mjs folders --folder '<root_id>'
node bin/outlook.mjs folders --folder '<folder_id>'
node bin/outlook.mjs messages --folder '<folder_id>' --limit 25
```

从 `stores` 返回的 `root_id` 开始，逐级调用 `folders`，对每个目标目录调用 `messages`、`today` 或 `search`。目录列表和邮件列表都需要跟随 `next_offset`，直到为 null。多账户无法唯一确定 `inbox` / `sent` 时，使用明确的文件夹 ID。

Mac 只公开已确认的收件箱和已发送目录树。目录名可能被 Outlook 脱敏，服务不会猜测原名称。邮件 ID 属于具体后端和缓存快照，失效时重新检索获取。

### 查看与下载附件

```sh
node bin/outlook.mjs attachments --id '<message_id>'
node bin/outlook.mjs download --id '<message_id>' --index 1 --out './report.xlsx'
```

附件序号从 1 开始。下载自动分块并校验 SHA-256，拒绝覆盖已有文件。只支持可读取的普通文件附件；Mac 附件列表可能不完整，列表为空不证明原邮件没有附件。

### 管理后台服务

```sh
node bin/outlook.mjs status
node bin/outlook.mjs health
node bin/outlook.mjs restart
node bin/outlook.mjs stop
```

`status` 检查服务是否在线，`health` 检查邮箱后端是否可用。机器重启或注销后需重新启动；Windows 安装器可选登录自启。

## 接入 Codex 和其他 Agent

### 通用 MCP 配置

先启动后台服务，然后生成本机配置：

```sh
node bin/outlook.mjs start
node bin/outlook.mjs mcp-config
```

将输出的 `mcpServers.outlook-local` 条目合并到支持此配置格式的客户端中，保留已有 MCP 设置。配置会包含本机 Node、入口脚本和状态目录的绝对路径，不含访问令牌。配置表单型客户端则填写对应的 command、args 和 env。

```json
{
  "mcpServers": {
    "outlook-local": {
      "command": "/absolute/path/to/node",
      "args": ["/absolute/path/to/outlook-cli/bin/outlook.mjs", "mcp"],
      "env": {
        "OUTLOOK_LOCAL_STATE_DIR": "/absolute/path/to/service-state"
      }
    }
  }
}
```

上面仅为格式示例，请使用 `mcp-config` 的实际输出。每个 Agent 启动自己的 stdio 适配器，共享本机后台服务。

### Codex CLI

用 `mcp-config` 输出中的实际路径注册，状态目录必须与后台服务一致：

```sh
codex mcp add outlook-local --env OUTLOOK_LOCAL_STATE_DIR=/absolute/path/to/service-state -- /absolute/path/to/node /absolute/path/to/outlook-cli/bin/outlook.mjs mcp
codex mcp list
```

路径含空格时加引号。注册后重新连接客户端或开启新的 Codex 会话，以加载工具。Windows 同样使用生成配置中的 Windows 绝对路径。

### 提供给 Agent 的工具

| 工具 | 用途 |
|---|---|
| `outlook_health` | 检查后端及缓存状态 |
| `outlook_stores` | 列出邮箱存储和根目录 ID |
| `outlook_folders` | 列出直接子目录 |
| `outlook_messages` | 分页列出邮件元数据 |
| `outlook_today` | 列出指定文件夹今天或指定日期的邮件 |
| `outlook_search` | 检索指定文件夹 |
| `outlook_read` | 读取邮件与分页正文 |
| `outlook_attachments` | 列出附件 |
| `outlook_read_attachment` | 分块读取附件内容 |

可向 Agent 提出这样的请求：

> 请读取今天收件箱、已发送及其全部已知子目录的邮件，逐目录翻页，阅读可用正文后总结主要事项和待办。说明日期依据、缓存缺失和未覆盖范围，不修改邮件。

本项目负责提供数据，归纳和总结由调用它的 Agent 完成。也支持带令牌认证的本机 Streamable HTTP MCP，见 [详细参考](docs/usage-reference.md)。

## 数据与使用边界

- **本地缓存不等于完整邮箱。** 未同步或无法可靠关联的正文、附件不能保证读取；快照时间也不等于服务器同步时间。
- Mac 会在独立私有目录写入索引、正文快照及附件导出，原始 Outlook 缓存只读。默认位于 `~/Library/Caches/outlook-cli/`；不要上传这些数据或服务令牌。
- Mac 列表先按接收时间和文件夹筛选，不按发送时间提前裁剪，因此早先发送、后来投递的邮件仍能进入对应接收日期。列表中的正文待读取状态不代表正文缺失；首次读取才确认可用性。首次读取前源正文变化时会要求刷新，避免将变化后的内容错误关联到旧快照。
- Mac 缓存格式并非微软公开稳定 API，其他 Outlook 版本可能不兼容。服务保留一份当前压缩邮箱容器，刷新构建期间还需一份新容器的可用空间（容器部分通常为一份、峰值两份），另需索引、按需正文和附件空间。当前没有正文缓存总字节配额。源容器更新后旧快照 ID / 游标失效，应重新检索。
- 服务仅监听 `127.0.0.1` 并使用令牌认证，适用于本机 Agent，不直接提供局域网远程访问。
- Windows、macOS、Linux 自动化检查覆盖协议和合成测试；Mac 已完成真实缓存验证，统一版本的 Windows 真实邮箱验证仍待完成。

升级到接收时间筛选版本后，请运行 `node bin/outlook.mjs restart` 重新加载读取器并刷新索引。原生 Python MCP 用户需更新源码并重启对应客户端。

## 常见问题

| 现象 | 处理方式 |
|---|---|
| `status` 正常但无法读取 | 执行 `health`，检查邮箱后端；首次 Mac 索引需要等待 |
| Mac profile 不唯一 | 设置 `OUTLOOK_MAC_PROFILE` 后重新启动服务 |
| `FOLDER_AMBIGUOUS` | 通过 `stores` → `folders` 获取明确目录 ID |
| 正文或附件缺失 | 检查返回的状态和诊断，确认 Outlook 已同步；不要把缺失当成空内容 |
| Mac 参数不支持 | 移除 `unread`、`search_body: true` 或 `format: html` |
| Windows COM 不可用 | 检查经典 Outlook 和服务的登录会话，或使用已配置的本地 OST/PST |
| Agent 连接失败 | 确认服务已启动，配置中的绝对路径和状态目录一致 |

## 开发与详细文档

```sh
npm ci --ignore-scripts
npm test
# macOS / Linux：合成缓存测试，不访问真实邮箱
python3 -m unittest discover -s macos -p 'test*.py'
```

- [macOS 配置、目录范围与原生 Python MCP](docs/macos.md)
- [Windows 安装参数、后端、分页协议与 HTTP API](docs/usage-reference.md)
- [ost-mcp](https://github.com/3rg0n/ost-mcp)：Windows OST/PST 后端使用的解析器，固定版本和限制见详细参考。
