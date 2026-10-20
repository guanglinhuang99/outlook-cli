# macOS Outlook 只读读取器

共享 Node 入口根据平台选择后端。Mac 使用 Python 标准库解析 Outlook 的 HxStore 缓存和 Osa 日志中的目录关系；不会启动 Outlook，不会修改邮件、已读状态或附件。邮件同步仍由 Outlook 负责。该格式并非微软公开稳定 API，其他 Outlook 版本的缓存布局可能不兼容。

## 安装与接入

需要 Node.js 22+ 和 Python 3.11+。在源码目录执行：

```sh
npm ci --ignore-scripts
node bin/outlook.mjs start
node bin/outlook.mjs health
node bin/outlook.mjs stores
node bin/outlook.mjs today --folder inbox --limit 25
node bin/outlook.mjs today --folder sent --limit 25
node bin/outlook.mjs mcp-config
```

`sh outlook.sh` 等价于 `node bin/outlook.mjs`。首次调用会创建索引，可能需要几十秒；`status` 只检查 Node 服务是否在线，`health` 才检查缓存。停止测试服务使用 `node bin/outlook.mjs stop`。

大于 64 MiB 的 HxStore 会分块复制到私有临时目录，校验后通过只读映射解析；没有 2 GiB 源文件上限。首次索引需要足够磁盘空间容纳临时副本及恢复的正文，空间不足会报错。进程超时退出留下的构建目录会在下一次刷新时清理，仍在运行的构建会保留。

`mcp-config` 输出本机绝对路径，通用配置如下。先启动后台服务，再把此条目合并进客户端配置。现有客户端中的其他 MCP 配置应保留。

```json
{
  "mcpServers": {
    "outlook-local": {
      "command": "/absolute/path/to/node",
      "args": ["/absolute/path/to/outlook-cli/bin/outlook.mjs", "mcp"],
      "env": { "OUTLOOK_LOCAL_STATE_DIR": "/absolute/path/to/service-state" }
    }
  }
}
```

这套配置暴露与 Windows 相同的 9 个 `outlook_*` 工具，包含 `outlook_today`。端口、令牌鉴权和 Streamable HTTP 与 Windows 共用。`OUTLOOK_LOCAL_STATE_DIR` 必须与启动服务时一致；通常直接复制 `mcp-config` 的输出即可。 自定义服务状态路径的父目录须由当前用户或 root 控制，拒绝可被其他用户替换的共享目录。

服务状态会保留启动时的自定义 profile 保护路径，后续下载命令无需重复设置该环境变量。升级旧版本后先执行 `node bin/outlook.mjs restart`；旧状态缺少保护路径时，下载会要求重启。MCP 工具可能创建、刷新或清理私有缓存，因此安全注解会标明本地文件写入；这不改变原始 Outlook 邮件只读的约束。

`OUTLOOK_LOCAL_STATE_DIR` 必须是专用目录的绝对路径，首次应指定尚不存在的目录。服务只认领新目录或经验证的旧服务目录，拒绝更改无关现有目录的权限。

| 环境变量 | 默认值 / 用途 |
|---|---|
| `OUTLOOK_LOCAL_BACKEND` | `auto`；Mac 选择 `mac-cache`，`local` 也是它的别名 |
| `OUTLOOK_MAC_PYTHON` | `python3`；可设为 Python 可执行文件绝对路径 |
| `OUTLOOK_MAC_PROFILE` | 自动发现唯一的 HxStore profile；多个 profile 时要求显式指定绝对目录或 `~/` 路径 |
| `OUTLOOK_MAC_STATE_DIR` | `~/Library/Caches/outlook-cli/mac-cache`；共享服务的 Python 索引、正文快照、附件导出 |
| `OUTLOOK_LOCAL_STATE_DIR` | `~/Library/Caches/outlook-cli/service`；Node 服务令牌、日志 |
| `OUTLOOK_LOCAL_PORT` | `17379`；同机多实例应各用不同状态目录和端口 |

默认 profile 从 `~/Library/Group Containers/UBF8T346G9.Office/Outlook/Outlook 15 Profiles` 下发现。状态目录和下载目标不得位于原始 Outlook 缓存、`~/Library/Mail` 或显式 profile 内，祖先符号链接也会检查。服务可以写自己的索引和导出目录；原始邮件只读。 状态目录及其受管理子目录必须归当前用户所有且权限为 `0700`；不安全的现有目录会被拒绝，不会自动修改权限，可改用新的私有状态目录。请勿将服务状态、索引、正文或附件副本上传到 GitHub。

## 目录与时间口径

- 只公开目录证据确认的收件箱、已发送及其全部已知子目录；已删除及其子目录、归属不明和存在冲突的邮件均排除。旧快照 ID 还会核对最新本地目录范围，范围变更后须重新列出。
- `stores` 返回缓存账户分组；用 `root_id` 调用 `folders` 可获得收件箱和已发送根目录，按 `is_inbox` / `is_sent` 识别，再逐级获取子目录。部分原始目录名被 Outlook 脱敏为 `pii:`；返回名称状态，不猜测中文名称或账户名称。
- `inbox` / `sent` 仅在能唯一确定缓存账户时使用。判断包含根目录和全部已知子目录；多个账户有缓存时返回 `FOLDER_AMBIGUOUS`，应选择显式文件夹 ID。
- `messages`、`search`、`today` 每次查一个文件夹，不递归。遍历全部邮件需逐级调用 `folders`，逐目录跟随 `next_offset`；分页为 offset 模式，跨页邮箱变化可能带来重复或遗漏，按 ID 去重。
- Mac 已验证的时间字段是发送时间：`date_basis: "sent_at"`，`received_at: null`。`today` 以新加坡时区当天 `[00:00, 次日00:00)` 筛选发送时间，不能据此声称是当天新接收的全部邮件。翻页传回结果的 `date` 和 `next_offset`。 共享工具的显式 `date` 范围为 `1970-01-01` 至 `9999-12-30`；使用历史日期时按 `Asia/Singapore` 的对应时区规则筛选。
- 字段 271 的时间按毫秒对齐的 .NET DateTime ticks 解码，保留毫秒用于排序和时间过滤；支持从公元 1 年起、至 UTC `9999-12-31 16:00:00` 之前的格式有效值，上界预留新加坡时区转换空间，没有按当前年份或 2030 年截断。
- 只能读取已缓存、且能可靠关联的内容。`captured_at` 是快照生成时间，不是服务器同步时间；`coverage` 说明本地证据范围，不保证服务器邮箱完整。
- 每次成功刷新后保留当前目录快照和最近两个历史快照，清理不再引用的正文快照和附件导出。旧 ID/游标超过保留范围会失效，或在并发刷新清理时返回读取错误，应重新检索；原始 Outlook 缓存不参与清理。任何相关 Osa 日志读取失败时刷新报错，旧目录范围也不再视为有效。

## 中文正文和附件

正文通过 Message-ID、EntryID 等精确证据与邮件关联，不用主题相似度猜测。无可靠正文时保留 `body.available: false`、`text: null`、`status` 和 `diagnosis`，区分可确认的缓存缺失与解析尚未关联；空字符串不替代缺失正文。

共享 `read` 返回从缓存 HTML 提取的纯文本，处理中文编码；不渲染 HTML 或访问远程资源。正文分页单位与 Windows 一致，为 UTF-16 字符，跟随 `body.next_offset` 拼接。Mac 尚不能验证的功能会报错：

| 参数 | 错误 |
|---|---|
| `unread`（包含 false） | `UNSUPPORTED_FILTER` |
| `search_body: true` | `UNSUPPORTED_FILTER` |
| `format: "html"` | `UNSUPPORTED_FORMAT` |
| 已删除、草稿、垃圾邮件、发件箱 | `SCOPE_DENIED` |

未验证的已读状态、收件人等字段为 null。附件列表只包含识别到的记录，`attachments_complete: false`；不能由列表为空推断原邮件无附件。读取前核对附件指纹，按需复制到服务自己的导出目录，分块返回 base64 与整文件 SHA-256。私有附件导出最多保留 64 个文件、总计 512 MiB，淘汰后可重新从仍有效的原始缓存读取。CLI `download` 校验大小和哈希，使用独占创建拒绝覆盖，单附件上限 10 MiB。

## 保留的原生 Python MCP

已安装原生 Mac MCP 的客户端可继续使用 `python3 /absolute/path/to/outlook-cli/macos/mcp_server.py`。它直接提供原有 9 个 `outlook_local_*` 工具，不需要 Node 后台服务，可用 `--state`、`--profile`、`--from-date`、`--refresh-seconds` 参数配置，详见 `macos/mcp_server.py --help`。 使用原生安装器且有多个 profile 时，可运行 `python3 macos/install_mcp.py --profile /absolute/path/to/profile`，生成的通用配置会保留该选择。

原生工具提供跨收件箱/已发送目录树搜索、快照游标、批量正文和 `outlook_local_today`；共享 Node 工具沿用 Windows 的单目录、offset 契约。两者工具名、参数和 ID 不相同，请按工具声明调用。现有原生 MCP 配置无需因本次整合改动。

共享服务默认使用独立缓存目录，避免和已安装的旧版原生 MCP 相互覆盖索引。只有相同版本的读取器才应显式共用状态目录；升级原生安装后需重启对应客户端。

## 验证

```sh
npm test
python3 -m unittest discover -s macos -p 'test*.py'
# 真实缓存检查：先 start，输出仅含计数和布尔值
node scripts/smoke-mac.mjs
node bin/outlook.mjs mcp-config > /tmp/outlook-mcp-config.json
node scripts/check-mcp.mjs /tmp/outlook-mcp-config.json --health
```

Python 测试使用合成缓存，不需要邮件账户。真实缓存检查只读邮件，临时附件下载位于服务状态目录并在结束后删除；不会遍历任意本机文件或修改 Outlook。
