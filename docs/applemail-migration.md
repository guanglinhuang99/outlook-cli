# Apple Mail CLI 迁移到 Outlook CLI

在项目目录运行 `node bin/outlook.mjs`。本地服务和 Microsoft Graph 使用不同的邮件 ID。下表列出可迁移的工作流和使用的后端。

| 工作流 | Outlook CLI | 来源 |
|---|---|---|
| 跨目录候选检索、当日报告、批量取信 | `context-bundle`、`today-bundle`、`messages-bundle` | 本地已验证的 Outlook 缓存 |
| 更多目录和未读筛选 | `stores`、`folders`、`messages --unread true` | 本地目录和 Osa 未读观察，账户范围可显式开启 |
| 追溯会话、持续监控、分拣任务 | `conversation`、`delta-bundle`、`task-bundle` | 本地固定快照及可重复使用的检查点 |
| 运行附件摘要或一次导出整封邮件的附件 | `attachment-text`、`download-all` | 本地 Outlook 缓存 |
| 按语义查找正文 | `semantic sync`、`semantic search`、`semantic status` | 本机模型与可验证的缓存正文 |
| 一次读取 50 封邮件与指定附件，查询和导出证据 | `knowledge bundle`、`knowledge sync/search/read/source/export` | 本地缓存派生文本、私有索引与 SHA-256 |
| 查看日历和日程 | `calendar status/calendars/agenda` | 已授权的 macOS EventKit |
| 查服务器目录、未读邮件、正文和附件 | `graph folders`、`graph list`、`graph read`、`graph attachments`、`graph attachment-read` | 已授权的 Microsoft 邮箱 |
| 获取服务器提供的邮件 MIME | `graph mime` | 已授权的 Microsoft 邮箱 |
| 设置已读、旗标、移动及批量更新 | `graph mark-read`、`graph flag`、`graph move`、`graph batch-update` | 已授权的 Microsoft 邮箱 |
| 草稿、回复、转发、发送 | `graph drafts`、`graph compose`、`graph draft-update`、`graph draft-attach`、`graph draft-discard`、`graph reply`、`graph reply-all`、`graph forward`、`graph draft-send` | 已授权的 Microsoft 邮箱 |

跨目录查询按 `received_at` 和新加坡时区筛选。邮件缓存范围与分页规则见 [macOS 批量工具](macos.md#跨目录查询与批量读取)。现有 12 个 macOS MCP 工具可继续用于本地只读查询和正文。下面的新增命令通过 CLI 调用。

## 账户目录和未读

启用账户范围时，`OUTLOOK_MAC_FOLDER_SCOPE=account` 在启动服务前设置。给该实例单独分配 `OUTLOOK_LOCAL_STATE_DIR`、`OUTLOOK_MAC_STATE_DIR` 和 `OUTLOOK_LOCAL_PORT`。通过 `stores` 和 `folders` 找到已验证的邮箱和联机存档目录。账户范围包括有本地证据的草稿、垃圾邮件和已删除项。

```sh
node bin/outlook.mjs messages --folder inbox --unread true --limit 25
node bin/outlook.mjs context-bundle --folder all --unread true --limit 100
```

未知未读状态会从筛选结果排除，并计入 `unread_unknown_count`。每封邮件的 `unread_evidence.observed_at` 是最新通过校验的本地观察时间。

## 会话、增量与任务

```sh
node bin/outlook.mjs conversation --json '{"id":"<local_id>","folder":"all","subject_fallback":false}'
node bin/outlook.mjs delta-bundle --json '{"folder":"inbox"}'
node bin/outlook.mjs delta-bundle --json '{"folder":"inbox","since":"<checkpoint>"}'
node bin/outlook.mjs task-bundle --json '{"folder":"inbox","preset":"unread","limit":25}'
```

三个工作流都支持 `folder`、`include_subfolders`、`from_date`、`to_date` 和 `scan_limit`。默认最多扫描 5,000 封。`conversation` 用可验证的 Message-ID 及从 HxStore 解析的关联 ID 串起会话，主题匹配会标为回退。可设 `include_body:true` 读取首批最多 25 封的缓存正文。

`delta-bundle` 返回新增邮件、元数据变更及从同一范围的本地快照消失的邮件。首次读取用 `added` 返回基线。扫描完整时才返回新 `checkpoint`；后续传 `since` 时保留同样的目录、子目录和日期条件。本地消失可能源于移动、缓存淘汰或同步变化，服务器删除状态需要独立证据。任务预设为 `actionable`、`risk`、`unread` 和 `folder-summary`；前两种按主题或已有预览中的明确措辞列出命中理由。

## 本地附件

```sh
node bin/outlook.mjs attachments --id '<local_id>'
node bin/outlook.mjs download-all --id '<local_id>' --out-dir './downloads'
node bin/outlook.mjs attachment-text --id '<local_id>' --index 1 --max-chars 8000
```

输出目录需已存在。批量下载逐个检查大小和 SHA-256，保留原有文件，并逐个返回 `downloaded` 和 `errors`。`complete` 表示本批已列出文件的导出结果，`attachments_complete` 表示后端能否证明附件清单完整。每封最多处理 100 个附件，单文件最多 10 MiB，批次最多 100 MiB。

文本提取支持纯文本、PDF、DOCX 和 XLSX。PDF 需本机有 `pdftotext`，扫描件需另行 OCR。Office 格式需要本机 Python 3 和标准库。输出中的 `truncated` 表示文本达到限额。

## 本地语义检索

```sh
node bin/outlook.mjs semantic sync --json '{"folder":"all","from_date":"2026-11-01","scan_limit":250}'
node bin/outlook.mjs semantic search --json '{"query":"下月的预算审批需要谁确认","limit":10}'
node bin/outlook.mjs semantic status
```

macOS 默认用 Apple NaturalLanguage 的离线句向量。需可用的 Swift 工具链。`OUTLOOK_APPLE_EMBEDDING_LANGUAGE` 默认 `zh-Hans`，也可设为 `en`。如选择 `OUTLOOK_EMBEDDING_PROVIDER=ollama`，设置 `OUTLOOK_EMBEDDING_MODEL` 为已安装的本地模型。Ollama 必须运行在 HTTP 回环地址，并关闭云功能。索引保存在所属服务状态目录的 `semantic` 子目录，记录快照、模型身份、正文哈希和向量。正文原文不写入语义索引。[Ollama 本地向量接口](https://docs.ollama.com/api/embed)

每次 `sync` 可选收件箱、已发送、全部已验证目录或指定 ID，最多扫描 500 封。一个状态目录保留一份当前范围的索引。结果明确区分完整正文、截断、缺失和无法验证。搜索返回前会复核候选正文。快照、范围或模型变化后，使用相应范围重新同步。

## 邮件与附件知识库

```sh
node bin/outlook.mjs knowledge bundle --json '{"ids":["<local_id>"],"attachment_requests":[{"id":"<local_id>","index":1}]}'
node bin/outlook.mjs knowledge sync --json '{"ids":["<local_id>"],"semantic":true}'
node bin/outlook.mjs knowledge search --json '{"query":"预算审批","kind":"all","mode":"hybrid"}'
node bin/outlook.mjs knowledge source --json '{"archive_id":"<archive_id>"}'
node bin/outlook.mjs knowledge export --json '{"ids":["<archive_id>"],"out_dir":"./exports"}'
```

`bundle` 和 `sync` 每次最多接受同一快照的 50 封邮件及 25 个明确指定的附件。正文和附件共享 `text_budget_chars`，默认 60,000 UTF-16 字符。`sync` 将已选证据写入服务状态目录的私有 `knowledge` 索引。设置 `semantic:true` 为可用文本建立 Apple 本地向量。搜索支持 `literal`、`semantic`、`hybrid`，可限定 `kind` 为 `mail` 或 `attachment`。

`status` 显示索引大小、快照和覆盖范围，`read` 读取索引记录。导出为带 SHA-256 的派生证据 JSON，保留来源、快照、缺失和截断状态；本地 Outlook 缓存无法提供字节级原始 RFC 822 邮件。输出目录需已存在。

## 只读日历

```sh
node bin/outlook.mjs calendar status
node bin/outlook.mjs calendar authorize
node bin/outlook.mjs calendar calendars
node bin/outlook.mjs calendar agenda --json '{"date":"2026-11-25","time_zone":"Asia/Singapore"}'
```

连接的是当前用户的 macOS 日历数据。Outlook 账户的日程需已同步到 macOS Calendar。首次授权需在有图形会话的 Mini 上运行 `calendar authorize` 并授予完整日历访问；普通查询遇到权限缺失时会返回 `CALENDAR_AUTH_REQUIRED`。日程支持精确的 `calendar_ids`、日期范围、重复和全天事件、分页，以及显式开启的 `include_details`。

## Microsoft Graph 授权与写操作

使用组织批准的 Microsoft Entra 公共客户端注册，并允许设备代码登录。在本机环境设置 `OUTLOOK_GRAPH_CLIENT_ID` 和 `OUTLOOK_GRAPH_TENANT_ID`。登录默认申请 `Mail.ReadWrite`，只读可传 `{"read_only":true}`。确需发送时传 `{"send":true}` 单独请求 `Mail.Send`。也可从组织批准的身份工具提供委托访问令牌 `OUTLOOK_GRAPH_ACCESS_TOKEN`。设备登录令牌和刷新凭证仅保存在当前服务状态目录的私有 `graph-auth` 子目录。[Microsoft Entra 设备代码授权](https://learn.microsoft.com/en-us/entra/identity-platform/v2-oauth2-device-code) · [Outlook 邮件 API](https://learn.microsoft.com/en-us/graph/outlook-create-send-messages)

```sh
node bin/outlook.mjs graph login --json '{"send":true}'
node bin/outlook.mjs graph status --json '{"verify":true}'
node bin/outlook.mjs graph folders
node bin/outlook.mjs graph list --json '{"folder":"inbox","unread":true,"limit":25}'
node bin/outlook.mjs graph resolve --json '{"internet_message_id":"<message@example.invalid>"}'
```

如果从本地查询开始，用返回的 `internet_message_id` 做 `graph resolve`，并核对唯一命中的 Graph 邮件。`graph list` 的 `after` 和 `before` 使用带时区的 ISO 接收时间，分别含起点和不含终点。写操作都须提供每次操作唯一的 `request_id`。移动、丢弃、上传到草稿和发送还要求明确的 `confirm:true`。如果请求回执不确定，先在 Outlook 核对结果并保留原请求 ID。发送前可用 `graph read` 核对已有草稿。正文和附件内容优先通过 `--json-file <path>` 或 `--json-file -` 传入，避免放在命令行参数里。[Microsoft Graph 草稿发送](https://learn.microsoft.com/en-us/graph/api/message-send?view=graph-rest-1.0)

| 动作 | JSON 字段 |
|---|---|
| `mark-read` | `id`, `is_read`, `request_id` |
| `flag` | `id`, `flag`, `request_id`，旗标值为 `notFlagged`、`flagged` 或 `complete` |
| `move` | `id`, `folder`, `request_id`, `confirm` |
| `batch-update` | `ids`, `request_id`，任选 `is_read`、`flag`、`move_to`；移动时提供 `confirm` |
| `compose` | `request_id`，可选 `subject`, `body`, `to`, `cc`, `bcc`，收件人为地址数组 |
| `draft-update` | `id`, `request_id`，以及需要改动的草稿字段 |
| `reply` / `reply-all` / `forward` | `id`, `request_id`，可选 `comment`；转发还需 `to` |
| `draft-attach` | `id`, `request_id`, `confirm`, `name`, `content_base64`，可选 `content_type`，文件小于 3 MB |
| `draft-discard` / `draft-send` | `id`, `request_id`, `confirm` |

`compose`、`reply`、`reply-all` 和 `forward` 创建草稿。只有 `draft-send` 发送。`graph attachments` 列服务器附件，`graph attachment-read` 读取一个普通附件的 Base64 和 SHA-256，最多 10 MiB。`graph logout` 清理此 CLI 缓存的设备授权，调用方通过环境提供的访问令牌仍由调用方管理。

`graph mime --json '{"id":"<graph_id>"}'` 返回 Graph 提供的 MIME 字节，附 Base64、字节数和 SHA-256，最多 10 MiB。使用已获准的 Graph 邮件读取权限，来源标记为 `graph_mime_representation`。[Microsoft Graph MIME 接口](https://learn.microsoft.com/en-us/graph/outlook-get-mime-message)

在 Mini 切换脚本时，先用同账号、同日期、同目录的两套只读命令核对邮件、正文和附件覆盖，再启用需要 Graph 的写操作。
