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

大于 64 MiB 的 HxStore 会分块复制到私有目录，校验后通过只读映射解析；没有 2 GiB 源文件上限。按需正文版本会保留这份压缩容器快照，小文件也保存一份，避免 Outlook 同步重排数据块后刚检索的邮件无法读取。服务仅保留当前源快照的容器副本；正常保留一份，刷新期间旧副本与新副本同时存在，容器部分峰值为两份。刷新前需额外留出新容器大小的可用空间，还要另留索引、按需正文和附件空间；空间不足会报错并保留此前可用视图。升级时旧版本遗留的历史快照在首次成功刷新后清理。进程超时退出留下的构建目录会在下一次刷新时清理，仍在运行的构建会保留。

共享服务和原生 MCP 都先建立邮件身份、文件夹和接收时间索引，再执行日期及目录筛选。索引阶段不复制和解码全部历史正文；`read` 或原生批量正文工具才按需恢复目标邮件的正文，保存在私有缓存中供后续分页复用。为了识别延迟投递，仍需扫描历史邮件头，不会以发送日期提前排除。首次全量扫描的耗时仍取决于 HxStore 大小和日志量；本次优化不包含块级增量解析，也没有增加正文缓存总字节配额。

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

服务状态会保留启动时的自定义 profile 保护路径，后续下载命令无需重复设置该环境变量。升级旧版本后先执行 `node bin/outlook.mjs restart`；旧状态缺少保护路径时，下载会要求重启。接收时间筛选版本使用新的目录快照版本，重启后会重新构建索引。MCP 工具可能创建、刷新或清理私有缓存，因此安全注解会标明本地文件写入；这不改变原始 Outlook 邮件只读的约束。

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
- Mac 与 Windows 一样按接收时间筛选：`date_basis: "received_at"`。读取 Osa `MessageHeader.LastDeliveryTime`，与同一记录的 `ReceivedOrRenewTime` 原始 .NET ticks 核对，并沿用 EntryID / Message-ID 与发送时间的身份验证；不使用日志的响应时间或文件修改时间。`today` 以新加坡时区当天 `[00:00, 次日00:00)` 筛选接收时间。这里的接收时间是 Outlook 最后投递时间，不保证是邮件最初的 SMTP 接收时间。翻页传回结果的 `date` 和 `next_offset`。 共享工具的显式 `date` 范围为 `1970-01-01` 至 `9999-12-30`；使用历史日期时按 `Asia/Singapore` 的对应时区规则筛选。
- 接收时间缺失、无效或证据冲突时，`received_at` 为 null，`received_time_status` 说明原因；不回退到发送时间。无日期条件仍可检索和读取这些邮件，有日期条件时排除并通过 `received_time_unknown` / `date_filter_incomplete` 提示范围不完整。结果按已验证接收时间倒序，缺失时间置后。接收时间只覆盖本地日志可可靠关联的记录，不保证全部服务器邮件。
- 字段 271 仍单独保留为 `sent_at`，用于显示与身份核对，不参与邮件日期筛选。该字段按毫秒对齐的 .NET DateTime ticks 解码，保留毫秒用于排序和时间过滤；支持从公元 1 年起、至 UTC `9999-12-31 16:00:00` 之前的格式有效值，上界预留新加坡时区转换空间，没有按当前年份或 2030 年截断。
- 只能读取已缓存、且能可靠关联的内容。`captured_at` 是快照生成时间，不是服务器同步时间；`coverage` 说明本地证据范围，不保证服务器邮箱完整。
- 每次成功刷新后最多保留当前目录快照和共用同一源容器的最近两个历史视图；源快照变化时清理旧源及其视图、正文和附件导出。失败的目录刷新回退源指针并清理未发布视图的容器副本。旧 ID/游标超过保留范围会失效，或在并发刷新清理时返回读取错误，应重新检索；原始 Outlook 缓存不参与清理。任何相关 Osa 日志读取失败时刷新报错，旧目录范围也不再视为有效。

## 搜索与重复读取性能

原生 MCP 的 `outlook_local_search` 支持 `query`、`field`、`person`、`person_role`，人员角色为 `any` / `sender` / `to` / `cc`；关键词和人员条件按 AND 组合。Osa 收件人被脱敏或缺失时，服务读取本地 HxStore 的 Person 对象，按 UTF-16LE 解码真实姓名、邮箱，并通过邮件头及 To/Cc 父集合引用关联邮件；不解码 `pii:` 占位符，不从正文猜测，也不依赖网络 API。只接受已验证的对象布局，排除阅读窗对象、密送和歧义关联。

`participants_status` 按角色标记 `verified`、`partial`、`missing`、`redacted` 或 `conflict` 等状态；`participants_evidence` 说明 HxStore 来源、解析人数和日志人数。HxStore 人员只有在没有拒绝记录、且人数与最新明确完整的 Osa 观察一致时才标为 `verified`；没有完整日志或部分对象解析失败时，可搜索已确认人员，但标为 `partial`。对象版本冲突不返回该角色名单。`verified` 是本地证据一致，不保证服务端或整个邮箱完整。`person_search_incomplete` 提醒本地证据可能漏项；原始 Exchange 地址不等于已验证的 SMTP 别名。共享 CLI 用 `search --query '预算' --person '张三' --person-role any`，按文件夹分别检索。

常驻进程最多缓存 3 个已验证索引，文件内容/身份/权限变化会使缓存失效；每次请求仍检查 profile、目录权限及原始邮件隔离边界。宽范围搜索只为返回页生成完整元数据，正文继续按需解析。首次加载、Outlook 更新后的索引刷新、进程重启后仍有开销，默认 300 秒的刷新周期没有放宽。升级后的目录快照格式为 5，首次使用需重建目录/人员索引；旧邮件不会按发送日期提前排除。升级已安装 MCP 后须重新连接客户端，让常驻进程加载新代码。

查询在内部复用已验证的索引，只复制实际返回的数据，避免每次搜索复制整个邮箱索引。Osa 日志的解析元数据通过 Python 标准库 SQLite 保存到私有状态目录的 `osa-metadata.sqlite3`，新进程也可复用。缓存最多保留 4096 个文件、64 MiB 序列化元数据，数据库文件另限 68 MiB；不保存正文和附件内容，不需要安装数据库服务。首次使用需要填充缓存。

每次刷新仍完整枚举日志，并核对文件身份、大小、修改/变更时间和权限，以及扫描前后的清单。新增、修改和删除的日志会重新参与目录、人员和接收时间关联，延迟收到的旧邮件不会被时间水位过滤掉。SQLite 缓存包含解析器版本、源文件指纹及内容校验和；损坏、占用或不安全的缓存路径会回退到源日志解析。缓存更新在成功扫描后以事务写入，不会修改原始 Outlook 文件。

HxStore 的 LZ4 解压改用批量复制，保留重叠复制语义、CRC 和边界检查，可减少首次构建及容器变化后的重建耗时。源容器变化后仍需全量解析并重新核对跨块关联；这里尚未实现按块增量索引，也没有把邮件查询接口迁移到 SQL。现有索引仍需在新进程中首次加载和校验。

## 中文正文和附件

已支持本机验证的 type-202 内嵌正文变体 5，以及变体 4/5 的额外字符串池布局。新增布局必须同时通过对象归属、描述符边界、编码和非空摘要一致性校验；摘要缺失或不一致时仍返回不可用，不猜测正文。解析索引版本升级到 20，升级后首次调用会重建索引，避免复用旧版“尚未关联”的正文结果。正文继续按需读取。

正文通过 Message-ID、EntryID 等精确证据与邮件关联，不用主题相似度猜测。无可靠正文时保留 `body.available: false`、`text: null`、`status` 和 `diagnosis`，区分可确认的缓存缺失与解析尚未关联；空字符串不替代缺失正文。

列表中的 `body_status: "deferred"` 表示尚未验证正文内容，不表示缺失；状态摘要中的 `deferred_bodies` 也是索引生成时的待读取数量。首次读取从私有容器按位置校验目标数据块、帧哈希和邮件归属，再生成正文和纯文本快照；外部正文文件则校验索引时记录的指纹，若外部文件已变化，重新检索/刷新后再读取。已成功保存的正文继续使用快照校验，后续分页无需重新解析整封邮件。

共享 `read` 返回从缓存 HTML 提取的纯文本，处理中文编码；不渲染 HTML 或访问远程资源。正文分页单位与 Windows 一致，为 UTF-16 字符，跟随 `body.next_offset` 拼接。Mac 尚不能验证的功能会报错：

| 参数 | 错误 |
|---|---|
| `unread`（包含 false） | `UNSUPPORTED_FILTER` |
| `search_body: true` | `UNSUPPORTED_FILTER` |
| `format: "html"` | `UNSUPPORTED_FORMAT` |
| 已删除、草稿、垃圾邮件、发件箱 | `SCOPE_DENIED` |

未验证的已读状态、收件人等字段为 null。附件列表只包含识别到的记录，`attachments_complete: false`；不能由列表为空推断原邮件无附件。读取前核对附件指纹，按需复制到服务自己的导出目录，分块返回 base64 与整文件 SHA-256。私有附件导出最多保留 64 个文件、总计 512 MiB，淘汰后可重新从仍有效的原始缓存读取。CLI `download` 校验大小和哈希，使用独占创建拒绝覆盖，单附件上限 10 MiB。

原生 MCP 批量正文读取的 64 MiB 预算计入实际读取的外部正文文件、解码后的 HTML 字节和实际解压的数据块；gzip 解压受剩余预算限制。已恢复的正文缓存命中不消耗该源处理预算，同次读取命中块缓存不重复计费，淘汰后再次解压会再次计费。另设 64 MiB 的批量缓存校验预算，限制重启后首次读取的正文哈希和文本块校验；进程内已校验且指纹未变的内容不重复计费。超出预算的邮件返回明确错误，可改为单封读取。

## 保留的原生 Python MCP

已安装原生 Mac MCP 的客户端可继续使用 `python3 /absolute/path/to/outlook-cli/macos/mcp_server.py`。它直接提供 10 个 `outlook_local_*` 工具，不需要 Node 后台服务，可用 `--state`、`--profile`、`--from-date`、`--refresh-seconds` 参数配置，详见 `macos/mcp_server.py --help`。 使用原生安装器且有多个 profile 时，可运行 `python3 macos/install_mcp.py --profile /absolute/path/to/profile`，生成的通用配置会保留该选择。

原生工具提供跨收件箱/已发送目录树搜索、快照游标、批量正文和 `outlook_local_today`；共享 Node 工具沿用 Windows 的单目录、offset 契约。两者工具名、参数和 ID 不相同，请按工具声明调用。现有原生 MCP 配置无需因本次整合改动。

共享服务默认使用独立缓存目录，避免和已安装的旧版原生 MCP 相互覆盖索引。只有相同版本的读取器才应显式共用状态目录；升级原生安装后需重启对应客户端。

### 周报、日报的精简读取

Mac 原生 MCP 新增 `outlook_local_digest`，合并搜索和正文读取，供 agent 生成摘要。示例参数：

```json
{"from_date":"2026-10-15","to_date":"2026-10-21","limit":25,"body_max_chars":12000}
```

日期按新加坡时区的接收时间，两端日期包含整天；默认范围为收件箱、已发送及其已知子目录，排除已删除。也支持 search 的关键词、人员和文件夹筛选。需要立刻检查新邮件时先调用 `outlook_local_refresh`，否则沿用默认 300 秒刷新周期。

返回 `items` 为精简邮件信息，保留中文主题、人员及其验证状态、接收时间、目录、附件名称与正文状态。正文置于 `bodies` 字典，邮件用 `body.content_id` 引用；只有完整纯文本哈希一致且返回内容一致时才合并。每封邮件的 ID、日期、人员仍独立保留，不按主题合并，也不自动删掉引用或转发正文。`body.truncated` 表示截断，`body.status` / `diagnosis` 区分缺失和未关联；有缓存的空正文与缺失正文不同。附件只列出，内容仍按需另读。详细解析证据可用原有 `read` 查询。

Agent 按下面方式减少重复传输和重复摘要：

1. 跟随 `next_cursor` 读完所有页，保留每页 `page_ids`。同一次分页保持相同筛选条件。
2. 保存已收到的 `bodies`，后续页可传其键列表为 `known_bodies`，避免跨页重复正文；只有仍保存正文时才能传入。
3. 为每封成功处理的邮件保存 `version` 和对应摘要。下次从第一页重新查询同一日期/目录范围，传 `known_versions`；仅变化的邮件进入 `items`，其他邮件进入 `unchanged`。新增缓存正文、正文截断范围变化、人员或目录变化会使版本失效。服务本身不调用模型，也不保存 AI 摘要。
4. 只有所有分页成功且每页 `errors` 为空，才能把当前 `page_ids` 并集与上次相同范围的 ID 集合比较，识别移出范围的邮件；不能从单页缺席推断删除。失败时保留旧摘要并重试，不能把失败邮件标为已处理。游标失效则从第一页重新开始。

每页最多 25 封，每封最多 20,000 字符，返回正文最多 500,000 字符；源文件读取和缓存验证各保留 64 MiB 预算。预算不足会逐封报错，可降低分页大小重试。两个 known 列表各最多 5,000 项；调用前需自行限制为不超过 5,000 项的子集，超长请求会被拒绝（遗漏项会重新返回，正确性不受影响）。增量方式每次仍检查整个查询范围，不用“上次发送时间”水位跳过延迟收到的邮件。

该工具只存在于 Mac 原生 MCP；共享 Node/Windows 的 9 个工具接口不变。安装后重新连接 MCP 才能发现新工具。

本机实测（2026-10-21，按接收时间读取 10 月 15—21 日共 731 封邮件，正文上限同为 12,000 字符）：原搜索＋批量正文的 MCP 返回数据为 16.82 MB，digest 为 8.58 MB，减少 49%；传入全部已处理版本后为 0.27 MB，减少 98.4%。这些是响应字节，不包括客户端发送 known 列表的请求字节。热缓存读取分别为 0.46 秒、0.40 秒和 0.38 秒；digest 在新进程中的第一遍读取为 2.16 秒，另需索引刷新 15.63 秒。精简保留字段、附件名称/状态和正文与原接口逐封一致；319 封正文未缓存、28 封正文截断的标记保持一致。412 条可用正文记录合并为 360 个内容条目，但正文字符仅减少约 0.5%，因为重复正文多为空或很短。这次主要减少元数据冗余及重复摘要输入，尚未重新计时完整模型摘要过程，不能把响应缩小比例当作端到端加速比例。

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
