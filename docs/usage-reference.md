# Outlook CLI 详细使用参考

首次使用请先看 [项目首页](../README.md)。本页保留安装参数、后端限制和协议细节。

Windows 和 macOS Outlook 的本机只读服务，共用 **CLI、MCP stdio、MCP Streamable HTTP**。不内置 AI，不需要模型 API Key，不使用 Graph OAuth。邮箱登录及同步仍由 Outlook 完成。

源码 v0.3.0 按平台自动选择读取器，共享 9 个 `outlook_*` 工具。已有 Windows COM / OST/PST 用法保持兼容；Mac 的 HxStore 格式由独立 Python 读取器处理，邮件 ID 不跨后端通用。

| 平台 | 默认后端 | 日期筛选依据 | 目录范围 |
|---|---|---|---|
| Windows | 已运行的经典 Outlook COM；不可用时读本地 OST/PST | `received_at` | 原有 Outlook / 文件后端范围 |
| macOS | `mac-cache`，只读本地 HxStore 和 Osa 目录证据 | `received_at`（已验证的 `LastDeliveryTime`） | 已确认的收件箱、已发送及其全部已知子目录；排除已删除及归属不明邮件 |

### Mac 从源码使用

需要 Node.js 22+、Python 3.11+，以及已在本机同步邮件的 Outlook。无需安装 Windows 解析器。

```sh
git clone --branch codex/macos-windows-unified https://github.com/guanglinhuang99/outlook-cli.git
cd outlook-cli
npm ci --ignore-scripts
node bin/outlook.mjs start
node bin/outlook.mjs health
node bin/outlook.mjs today --folder inbox --limit 25
node bin/outlook.mjs today --folder sent --limit 25
node bin/outlook.mjs stores
node bin/outlook.mjs mcp-config
```

Mac 使用细节、通用 MCP 配置和原生 Python MCP 兼容入口见 [macOS 说明](macos.md)。共享命令每次查询一个文件夹；用 `stores` → `folders` 逐级遍历子目录。Mac 缓存无法验证的已读状态和 HTML 正文会明确报错；正文搜索需显式启用 `search_body`，状态通过 `fulltext` 和 `body_search_incomplete` 返回。首次建索引可能需要几十秒，索引和附件副本仅写入服务自己的缓存目录，原始邮件只读。

## Windows：Outlook 未打开时读取本地数据

默认 `auto` 模式：先尝试连接已运行的经典 Outlook；仅当 COM 返回 `OUTLOOK_UNAVAILABLE` 时转到本地文件解析。不会为了读取而启动 Outlook，也不会在 Outlook 忙、权限错误或某封邮件读取失败时悄悄换数据源。每次响应包含 `backend`（`com` / `local`）及 `data_scope`。

```powershell
.\outlook.cmd health
.\outlook.cmd messages --limit 10
# 显式指定后端，方便排查或在 Outlook 打开时验证文件读取
.\outlook.cmd health --backend local
.\outlook.cmd messages --backend local --folder inbox --limit 10
.\outlook.cmd search --backend local --query '项目' --search-body true
.\outlook.cmd messages --backend com --limit 10
```

所有 MCP 邮箱工具也接受 `backend: "auto" | "com" | "local" | "mac-cache"`。可用服务环境变量 `OUTLOOK_LOCAL_BACKEND` 设置默认值。Windows 的 `local` 读取 OST/PST，Mac 的 `local` 是 `mac-cache` 别名。使用本地文件 ID 的请求在 Windows auto 模式下仍路由到本地后端，因此 Outlook 重新打开后可以继续处理同一 ID。**各后端的 ID 不互通**；Outlook 关闭后若继续传 COM ID，会返回 `BACKEND_ID_MISMATCH`，应重新列出邮件获取本地 ID。

文件由 `ost-mcp` 从 Outlook 配置注册表和默认缓存目录发现，第一个文件作为内置文件夹别名的默认存储。需要手动指定文件或默认顺序时，在安装目录创建 **`local-stores.json`**：

```json
[
  "C:\\Users\\YOUR_USER\\AppData\\Local\\Microsoft\\Outlook\\your-mailbox.ost",
  "D:\\MailArchive\\archive.pst"
]
```

也可以将这个 JSON 数组设置到 `OUTLOOK_LOCAL_FILES`。显式配置会替代自动发现；不接受网络共享路径。文件夹别名支持常见中文和英文名称；无法唯一识别时会要求通过 `stores` → `root_id` → `folders` 选择明确文件夹，不猜测目标。多个文件中不支持的文件会在 `stores` 中单独报告错误。

本地模式的边界：

- **仅覆盖已缓存在磁盘上的内容**；没有下载的历史邮件、正文或附件无法凭空恢复。`local_file_modified_at` 是文件修改时间，不是最近同步时间。
- 使用固定版本 `ost-mcp 0.1.0`（MIT），安装时校验 ZIP 和 EXE 的 SHA-256。支持其能够解析的 OST/PST；加密格式及一些 PST 变体不支持，遇到错误会明确报告。已有本地文件时，本地后端无需运行 Outlook；新版 Outlook 自己的缓存不属于此后端的支持范围。
- 默认文本读取会将 HTML 缓存转换为纯文本，`body.derived_from` 标明来源。不会运行脚本或加载远程图片。解析器优先返回一种正文版本，若无法满足 `format: html` 会报错；仅有 RTF 的正文暂不转换。
- 每封最多解析前 400 万个原始正文字符、返回其中前 200 万个转换后的字符。超过上限时 `body.source_truncated` 为 true，`total_chars` 为 null；`next_offset` 只用于当前可读部分的分页，为 null 也可能仍有未能解析的内容。
- 本地搜索每次最多扫描 50 封候选邮件，可跟随 `next_offset` 继续。正文搜索使用独立 SQLite 索引，最多读取每封 1,600 万个原始字符；只有确认完整的正文才入索引，缺失、截断或不支持的正文会使 `body_search_incomplete` 为 true。普通 `read` 的 200 万字符上限不变。内部发件人及其他元数据无法解析时返回 null，不能视作空值或零。
- 附件沿用 10 MiB 上限和分块协议。部分嵌入对象不能直接读取。
- 每次查询重新打开文件，并比较操作前后的大小和修改时间；检测到变化返回 `OFFLINE_STORE_CHANGED`。这不是完整的并发一致性保证。默认在 Outlook 关闭后使用文件模式；Outlook 正在同步时优先用 COM。

现有源码部署可先运行 `powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\setup-reader.ps1` 安装解析器，然后安装新依赖并 `restart`。一键安装器已经包含这一步。

## Windows：在另一台电脑上一键安装

1. 从 [GitHub Releases](https://github.com/guanglinhuang99/outlook-cli/releases/latest) 下载 `outlook-local-0.2.1-setup.zip`，复制到目标 Windows 电脑，**先解压全部文件**。
2. 确认已配置经典 Outlook 邮箱，或有可读取的本地 OST/PST。双击解压目录中的 **`install.cmd`**，无需管理员身份。
3. 完成后，双击安装目录中的 **`Setup-MCP.cmd`**，查看中文接入说明和这台电脑的 `mcp-config.json`。
4. 在 Agent 的 MCP 设置中导入配置，或只合并 `mcpServers.outlook-local` 条目，保留已有设置。使用配置表单的客户端，按说明填写 stdio 的 command、args、env。
5. 保存并重新连接 Agent，然后双击 **`Test-MCP.cmd`** 验证 MCP 连接和邮箱后端。

**MCP 服务端随一键安装器一起安装，不需要额外执行 npm install。** 已发布的 v0.2.1 安装包提供 8 个工具；本分支源码增加 `outlook_today`，重新打包后检查 9 个工具。中文 `MCP-SETUP.txt` 包含实际安装路径、手动填写字段、验证方法和常见错误处理。`Setup-MCP.cmd` 打开说明与配置供你接入客户端，不会自动修改客户端配置。`Test-MCP.cmd` 检查实际配置、协议和邮箱健康状态，不读取邮件正文；使用 `-NoStart` 安装时，请先运行 `Start-Outlook.cmd` 再测试。

默认安装到当前用户的 `%LOCALAPPDATA%\OutlookLocal`，自动安装专用便携 Node.js 24.22.0、按 `package-lock.json` 安装生产依赖和固定版本 OST/PST 解析器、生成启动/停止入口及本机 MCP 配置，并启动服务验证后端。Node/COM 支持 Windows x64/ARM64；文件解析器提供 x64 程序，ARM64 需 Windows x64 仿真支持。可用 `-SkipLocalReader` 只安装 COM 功能。不会修改全局 Node.js、系统 PATH 或系统 PowerShell 执行策略。

首次安装需要连接 `nodejs.org`、`registry.npmjs.org` 和 GitHub 发布文件下载域名。Node ZIP 使用脚本内固定的官方 SHA-256 校验，npm 依赖使用锁文件完整性校验；不会关闭 HTTPS 证书验证。安装包不包含邮件数据、服务令牌、原电脑的本地配置或 node_modules。

安装后可双击 `Start-Outlook.cmd`、`Stop-Outlook.cmd`，也可在命令行调用安装目录中的 `outlook.cmd`。默认不修改登录自启。Outlook 未运行时服务可以完成安装，打开经典 Outlook 后再运行 `outlook.cmd health` 检查。

可选参数（在解压后的安装包目录运行）：

```powershell
# 自定义安装目录和端口；安装位置必须与解压源目录分开
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\install.ps1 -InstallDir 'D:\Apps\OutlookLocal' -Port 17380

# 安装并添加当前用户的登录自启（不需要管理员权限）
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\install.ps1 -AutoStart

# 安装但不启动服务
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\install.ps1 -NoStart

# 提前下载 Node 官方 ZIP 后使用本地文件，仍会校验 SHA-256；npm 依赖仍需联网
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\install.ps1 -NodeArchive 'D:\Downloads\node-v24.22.0-win-x64.zip'
```

重复安装同一目录会先停止该安装的服务，再更新程序和依赖，重新生成 MCP 配置。安装失败可修复问题后重试。安装器拒绝覆盖未标记的非空目录；端口被其他实例占用时会提示换端口。`-AutoStart` 在当前用户的启动文件夹生成 `Outlook Local Service.lnk`，删除该快捷方式即可取消自启。

安装时终端提示使用英文，以兼容 Windows PowerShell 5.1 的脚本编码。企业电脑若限制脚本执行、下载或应用运行，需要遵循其现有管理策略。

开发者重新打包：

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\build-installer.ps1
```

该命令使用文件白名单生成 ZIP 和旁边的 `.sha256` 文件，排除本机运行数据。

```text
其他 Agent ── CLI / MCP stdio / MCP HTTP
                         │
                127.0.0.1:17379
                         │
                 按平台选择工作进程
                  ┌──────┴──────┐
               Windows        macOS
          COM / 本地 OST、PST  Python HxStore 缓存
```

## Windows 源码快速使用

推荐使用上述一键安装包。开发者也可克隆源码（需要 Node.js 22.13 或更新版本，推荐 24 LTS），在普通桌面终端运行：

```powershell
git clone --branch codex/macos-windows-unified https://github.com/guanglinhuang99/outlook-cli.git
cd outlook-cli
npm ci --ignore-scripts
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\setup-reader.ps1
.\outlook.cmd start
.\outlook.cmd status
.\outlook.cmd health
.\outlook.cmd messages --limit 5
.\outlook.cmd search --query "项目" --limit 10
.\outlook.cmd restart
.\outlook.cmd stop
```

`start` 启动隐藏的后台进程，重复执行会复用已经运行的服务。`status` 检查服务，`health` 检查 Outlook 或本地文件后端。需要已配置并运行经典 Outlook，或有支持的本地 OST/PST 文件。CLI 成功输出 JSON；失败在 stderr 输出 JSON，退出码非零。`serve` 可在前台运行，便于排查问题。

COM 模式下，服务应在与 Outlook **相同的 Windows 用户、交互式登录会话及权限级别**下启动；本地模式无需 Outlook 运行，只需读取文件的权限。通常不需要管理员身份。本项目采用用户会话后台进程，因为 Outlook COM 不适合 Session 0 的 Windows 系统服务。服务不会主动启动或关闭 Outlook。注销、重启后需要重新执行 `start`；可在安装时选择 `-AutoStart`。

若在隔离的 Agent 环境中启动出现 `OUTLOOK_UNAVAILABLE`，请在普通桌面终端启动服务；Agent 随后可连接该本机服务。

## MCP 接入

### stdio：推荐

运行以下命令得到适合本机的完整配置，复制到其他 Agent 的 MCP 配置中：

```powershell
.\outlook.cmd mcp-config
```

一键安装后的可复制版本在安装目录的 `mcp-config.json`，中文接入说明在 `MCP-SETUP.txt`；源码部署可用上述命令生成配置。配置包含 Node 可执行文件、入口脚本和状态目录的绝对路径，不包含访问令牌。通用形式如下：

```json
{
  "mcpServers": {
    "outlook-local": {
      "command": "C:\\path\\to\\node.exe",
      "args": ["D:\\work\\outlook-cli\\bin\\outlook.mjs", "mcp"]
    }
  }
}
```

**先执行 `start`。** 每个 Agent 启动自己的轻量 stdio 适配器，共享后台服务和按平台选择的读取进程。Mac 将示例中的命令换成 `node bin/outlook.mjs mcp-config`，它会输出本机的绝对路径。适配器不写非协议文本到 stdout。

### Streamable HTTP

- URL：`http://127.0.0.1:17379/mcp`
- 请求头：`Authorization: Bearer <token>`
- `<token>` 从服务状态目录的 `service.json` 读取；Windows 默认为 `.runtime/service`，Mac 为 `~/Library/Caches/outlook-cli/service`。不要将这个文件提交到代码仓库或复制进公共配置。
- 使用无会话状态的 Streamable HTTP，POST 请求；不提供旧版独立 SSE 端点。
- 令牌在服务重启时更换。stdio 适配器会自动读取新令牌，使用固定 HTTP 配置的客户端需更新请求头。

只监听 IPv4 回环地址，不接受局域网连接。HTTP 检查 Host、拒绝带 Origin 的浏览器请求，并要求令牌。本机令牌目录在 Windows 使用 ACL 仅授予当前用户和 SYSTEM 访问权；Mac 创建目录使用 0700、令牌文件使用 0600。此鉴权不隔离同一用户下已能读取配置的进程。

## 工具和参数

CLI 命令与 MCP 工具一一对应，MCP 工具名增加 `outlook_` 前缀。

| CLI / 操作 | MCP 工具 | 功能 |
|---|---|---|
| `health` | `outlook_health` | Outlook 连接、版本、离线状态 |
| `stores` | `outlook_stores` | 邮件存储及各自的 `root_id` |
| `folders` | `outlook_folders` | 指定文件夹的直接子文件夹及数量 |
| `messages` | `outlook_messages` | 新到旧排列的邮件元数据 |
| `today` | `outlook_today` | 按新加坡时区列出今天在指定文件夹的邮件 |
| `search` | `outlook_search` | 一个文件夹内按主题、发件人匹配，可选搜索正文 |
| `read` | `outlook_read` | 单封邮件及分页正文 |
| `attachments` | `outlook_attachments` | 附件名称、类型、大小、序号 |
| `read_attachment` / `download` | `outlook_read_attachment` | 分块读取附件；CLI download 合并并保存文件 |

没有发送、删除、移动、修改已读状态等接口。服务端只读；CLI `download` 会写入用户指定的本机文件，拒绝覆盖已有文件。Outlook 自身仍会正常同步邮箱。

### 文件夹和邮件 ID

默认文件夹是 `inbox`。`sent` 表示已发送。Windows 还支持 `drafts`、`deleted`、`junk`、`outbox`；Mac 拒绝这些目录。别名针对默认存储，Mac 多个账户无法唯一确定默认值时会要求显式选择。其他账户和文件夹先调用 `stores`，用返回的 `root_id` 调用 `folders`，逐级取得文件夹 `id`。

ID 编码了后端和对象身份，Mac ID 还绑定缓存快照；调用方应原样传回。邮件移动、目录权限范围或存储变化可能导致 ID 失效，需要重新列出获取。它们不是全局永久 ID。

```powershell
.\outlook.cmd stores
.\outlook.cmd folders --folder '<root_id>'
.\outlook.cmd messages --folder '<folder_id>' --limit 20
.\outlook.cmd read --id '<message_id>' --max-chars 20000
.\outlook.cmd read --id '<message_id>' --body-offset 20000 --max-chars 20000
.\outlook.cmd attachments --id '<message_id>'
.\outlook.cmd download --id '<message_id>' --index 1 --out 'D:\work\report.xlsx'
```

### 筛选、搜索和分页

```powershell
.\outlook.cmd messages --unread true --after '2026-10-01T00:00:00+08:00' --before '2026-11-01T00:00:00+08:00'
.\outlook.cmd search --query '预算' --folder inbox --search-body true --scan-limit 500 --limit 25
```

- 日期范围是 `[after, before)`，须带时区。Windows 和 Mac 均按接收时间筛选，`date_basis` 为 `received_at`。Mac 使用经时间戳核对、且身份关联可靠的 Osa `LastDeliveryTime`；缺失或冲突时返回 null，不回退到发送时间。
- Mac 返回 `received_time_status`、`received_time_method` 标明接收时间证据。无日期条件时仍可检索、阅读缺失接收时间的邮件；有日期条件时排除这些邮件，并返回 `received_time_unknown` 及 `date_filter_incomplete`，不会伪装成完整日期结果。按接收时间倒序排列，缺失时间排在末尾。
- `today` 使用 `Asia/Singapore` 的当天零点至次日零点。 显式 `date` 支持 `1970-01-01` 至 `9999-12-30`，按历史时区规则计算零点。翻页时传回第一页返回的 `date` 和 `next_offset`，防止跨午夜后切换日期。例如 `node bin/outlook.mjs today --date 2026-10-20 --offset 25`。默认只查收件箱根目录；已发送用 `--folder sent`，子目录传对应 ID。
- 每页默认 25，最多 200。每次最多检查 `scan_limit` 个项目，默认 500，上限 2,000。
- **搜索只覆盖指定文件夹。** 查询是大小写不敏感的字面子串，不是 SQL 或正则表达式。`search_body` 默认 false。
- Mac 的 `search_body: true` 会先应用文件夹、接收时间和人员条件，再搜索完整纯文本；元数据命中或正文命中均可返回，人员条件仍按 AND 组合，正文命中附带 `body_snippet`。正文只在显式启用时读取，附件内容不纳入索引。
- Mac 共享搜索的 `fulltext` 统计包含 `candidates`、`indexed`、`missing`、`errors`、`pending`、`updated`、`reused`、`complete`、`revision` 和引擎信息。冷查询受 30 秒、128 MiB 的单次构建预算限制；`pending` 大于 0 或有错误时，重复相同第一页参数，直到 pending 和 errors 都为 0 后再用 `next_offset` 翻页；持续读取错误需要先刷新快照并检查正文诊断。`missing` 表示当前快照下正文未缓存或未能可靠关联；刷新后出现的新正文会重新参与索引，不是 pending；缺失正文、错误或未完成索引会通过 `body_search_incomplete` 明确返回。
- Mac 正文索引以私有 SQLite 文件持久保存完整纯文本，跨进程复用，数据库上限为 512 MiB；源缓存刷新后会重新考虑新增或变化的正文。三字符及以上 Unicode 查询使用 trigram，一、二字符查询对候选正文做字面扫描。Python 的 SQLite 构建需支持 FTS5/trigram，不支持或达到上限时返回明确索引错误。Windows 使用 Node 内置 SQLite 的独立实现，见下一条；两端均不提供实时服务器状态保证。
- Windows 的 `search_body: true` 同样支持中文短词、字面匹配、增量持久化和 `body_snippet`。仅传人员、不传关键词时仍走元数据搜索，不读取或索引正文。源码部署需 Node 22.13+，推荐 Node 24 LTS；安装器内置版本已满足要求。索引位于服务状态目录的 `fulltext` 子目录，继承当前用户/SYSTEM 的 Windows ACL，数据库上限 512 MiB。只写服务索引，不调用邮件保存、移动或标记已读操作。
- Windows 每次最多扫描 50 个当前候选项目；先应用收到时间、文件夹和人员条件，再索引完整正文。COM 用邮件 change key 与下载状态识别变化；没有可靠 change key 时本次读取可检索，但不持久复用。离线解析器未暴露逐邮件 change key，因此 OST/PST 的文件大小/修改时间/身份变化后，当前候选正文会重新校验，即使邮件 `modified` 没变。每次都重新枚举候选邮件，不用发送时间或修改时间高水位排除新邮件。
- Windows 正文索引仅允许已验证的收件箱、已发送及子目录，排除已删除。COM 使用 Outlook 默认目录 ID，离线模式使用解析器的标准目录树和中英文目录名；无法确认范围时返回 `FULLTEXT_FOLDER_SCOPE`。共享接口仍一次查询一个文件夹，不自动递归。
- Windows `fulltext` 返回本次候选的 `indexed`、`reused`、`updated`、`missing`、`errors`、`pending` 等统计。`retry_required=true` 时用本页原参数重试，`next_offset` 保持本页起点；全部可处理后再翻页。缺失、截断、格式不支持的正文会报告不完整，并在以后查询重试。`complete` 仅表示分页结束，`body_search_incomplete` 表示正文覆盖不足。单次目标预算 30 秒/128 MiB，其中预留 5 秒复核来源，单封最多 1,600 万 UTF-16 字符；超限正文不保存为完整结果。附件不参与索引。
- 热查询返回前也会复核来源：COM 重新核对候选邮件的变更标识和当前目录，离线模式重新检查文件版本。检测到同步、编辑或移动会明确报错，重试当前页即可；分页仍不是邮箱快照。COM 的 `Body` 属性会先读取整段正文，然后才检查长度上限，因此 1,600 万字符是索引上限，不是 Outlook 读取时的内存上限。
- `search --person '姓名或邮箱' --person-role any|sender|to|cc` 支持单独按人员搜索；可与 `--query` 同用，按 AND 组合。`any` 在三个角色间按 OR 匹配，不包含密送人。命中角色见 `matched_person_roles`；缺少角色证据时 `person_search_incomplete=true`，不能把未命中解释为不存在。人员只来自结构化邮件头/收件人字段，不从正文猜测。
- 人员搜索也先应用文件夹和接收时间范围。继续读取所有 `next_offset` 页面；`complete` 只表示本轮候选扫描结束，不能覆盖本地未缓存的邮件或参与人信息。
- 返回 `next_offset` 非空时，用相同筛选参数、该 offset 再请求。**本页没有结果不代表搜索完成**，应检查 `next_offset` 或 `complete`。
- offset 是经过日期/未读粗筛后检查的项目位置，不是匹配结果数量。`filtered_item_count` 包含非邮件项目及精确日期复核前的候选项目，不是最终命中数。
- 邮箱是实时变化的，分页不是快照。新邮件、移动、删除可能使不同页之间重复或遗漏；抓取时按 `id` 去重，重要批量导出应选择稳定时间窗口。
- `read` 默认返回最多 20,000 个正文字符，上限 200,000。跟随 `body.next_offset` 读取后续内容，单位为 UTF-16 字符。Windows 可使用 `--format html`；Mac 返回提取的纯文本，并保留缺失正文的状态。服务不渲染 HTML、不加载邮件中的远程图片。
- Exchange 发件人尽量解析为 `smtp_address`；无法解析时保留原始 `address` 和 `address_type`，SMTP 地址返回 null。`to`、`cc` 是 Outlook 的显示文本。

### 附件分块读取

`outlook_read_attachment` 参数：`id`、1 起始的 `index`、可选 `byte_offset`、`chunk_bytes`（默认 64 KiB，上限 1 MiB）、`max_bytes`（整个附件大小上限，最高 10 MiB）。返回 base64 `data`、`offset`、`chunk_size`、`next_offset`、整个文件的 `size_bytes` 和 `sha256`。

跟随 `next_offset` 合并各块，最后校验整个文件的 SHA-256。CLI `download` 自动完成这个过程。Windows 每块通过后端重新读取附件；Mac 从经过指纹检查的缓存附件复制到服务导出目录再分块。若附件在读取期间变化，CLI 会报错。仅支持普通文件附件；内嵌 Outlook 邮件、OLE 对象和超过 10 MiB 的附件会明确报错。Mac 缓存可能缺少部分附件，返回 `attachments_complete: false`。

## 本地 JSON API

同样使用令牌，`POST /rpc` 接受 `{ "method": "messages", "params": { "limit": 5 } }`，返回 `{ "result": ... }`。参数名称与 MCP 一致。以下 PowerShell 示例不会打印令牌：

```powershell
$state = Get-Content .\.runtime\service\service.json -Raw | ConvertFrom-Json
$headers = @{ Authorization = "Bearer $($state.token)" }
$body = @{ method = 'health'; params = @{} } | ConvertTo-Json
Invoke-RestMethod "http://127.0.0.1:$($state.port)/rpc" -Method Post -Headers $headers -ContentType 'application/json' -Body $body
```

也可以用 CLI 传完整 JSON 参数：`outlook.cmd rpc messages --json '{"limit":5}'`，注意所用 shell 的引号规则。请求体上限 64 KiB。最多排队 32 个调用，Windows 工作进程执行超时 45 秒，Mac 因首次索引需要更久，超时为 180 秒。超时会重建工作进程，等待中的请求返回错误，由调用方决定重试。

常见错误：`OUTLOOK_UNAVAILABLE`（经典 Outlook 未运行或会话不匹配）、`OUTLOOK_TIMEOUT`（Outlook 忙或弹窗）、`INVALID_ARGUMENT`、`INVALID_ID`、`ATTACHMENT_TOO_LARGE`、`UNSUPPORTED_ATTACHMENT`、`BUSY`。

## 安装、开发和验证

共享入口需要 Node.js 22+，Windows 全文索引需 22.13+（推荐 24 LTS）。Windows 需要系统自带 Windows PowerShell 5.1，以及已配置的经典 Outlook 或受支持的本地 OST/PST；新版 Outlook 不支持 COM。Mac 需要 Python 3.11+ 和可解析的 Outlook HxStore 本地缓存，支持范围见 [macOS 说明](macos.md)。Windows PowerShell COM 工作进程使用 STA，仅在该进程中允许执行项目脚本，不修改系统执行策略。

标准 Node 环境使用：

```powershell
npm ci --ignore-scripts
npm test
# macOS / Linux 上运行合成缓存测试，不访问真实邮箱
python3 -m unittest discover -s macos -p 'test*.py'
```

仓库提供 `pnpm-lock.yaml`；可用 `pnpm install --frozen-lockfile` 复现锁定版本。本机使用 Codex 附带的 Node/pnpm 安装；`outlook.cmd` 会优先使用 PATH 中的 Node，否则尝试当前用户的 Codex 运行时路径。

- `test/*.test.mjs`：参数边界、拒绝写操作、HTTP 鉴权、官方 SDK MCP 交互、并发串行化、超时恢复。
- `macos/test*.py`：合成 HxStore 解析、中文主题与正文关联、目录范围和已删除排除、附件指纹、纯 Python MCP、共享协议兼容性；无需邮件账户。
- `.github/workflows/ci.yml`：Windows、macOS、Linux 运行 Node 测试，macOS 和 Linux 另跑合成缓存测试。Linux 仅用于协议和解析测试，实际邮箱后端只支持 Windows/macOS。
- Windows 索引的合成性能验证：`node scripts/benchmark-windows-fulltext.mjs`。检查 300 封邮件在冷查询、热查询、服务重启和单封变更后的结果与完整扫描一致，并验证正文读取次数为 300、0、0、1；耗时不包含真实 Outlook COM 或 OST/PST 解析延迟。
- `scripts/smoke.mjs`：Windows 真实邮箱检查；`scripts/smoke-mac.mjs`：Mac 真实缓存检查。只输出计数/布尔值。两者检查 CLI/MCP、正文和附件分页/哈希及拒绝覆盖；Mac 还检查接收日期口径、目录限制和正文搜索 pending 重试。下载测试使用服务状态目录内的临时文件，测试后删除。不显示主题、地址、正文或令牌。
- `Test-OutlookRead.ps1`：最初的独立可行性探测脚本，服务不依赖它。
- `scripts/smoke-offline.mjs`：真实 OST 文件的正文/附件/筛选/MCP 验证，以及模拟 COM 不可用时自动调用真实文件后端的测试；不会关闭 Outlook。
- 服务状态目录的 `service.log` 只记录启动/服务错误。Windows 正常使用不建立邮件副本或全文索引；Mac 在独立服务缓存中保存本地索引、可恢复正文、可选全文索引和按需导出的附件，不写回原始缓存。

服务端口可通过 `OUTLOOK_LOCAL_PORT` 设置；状态目录可通过 `OUTLOOK_LOCAL_STATE_DIR` 设置。变更后应在服务端和客户端使用一致的状态目录。MCP SDK 的确切依赖版本记录在 package.json 和锁文件中。

缓存范围仍由 Outlook 决定。本项目已验证当前在线缓存会话的读取，不保证未缓存历史邮件在断网时可读，也未覆盖所有企业加密/权限保护邮件。

## 参考

- [Microsoft：Outlook 接口选择](https://learn.microsoft.com/en-us/office/client-developer/outlook/selecting-an-api-or-technology-for-developing-solutions-for-outlook)
- [Microsoft：日期时间筛选](https://learn.microsoft.com/en-us/office/vba/outlook/how-to/search-and-filter/filtering-items-using-a-date-time-comparison)
- [Microsoft：新版和经典 Outlook 的功能对比](https://support.microsoft.com/en-us/outlook/getstarted/feature-comparison-between-new-outlook-and-classic-outlook)
- [官方 MCP TypeScript SDK](https://github.com/modelcontextprotocol/typescript-sdk)
- [ost-mcp 固定版本及已知限制](https://github.com/3rg0n/ost-mcp/tree/v0.1.0)；许可证随安装包放在 `third-party/ost-mcp-LICENSE.txt`。
