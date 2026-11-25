# Outlook 复制邮件 → MSG 文件剪贴板（Windows）

在经典 Outlook 的邮件列表里选中邮件并按 **Ctrl+C**，后台程序将复制的邮件保存为 `.msg`，然后将文件放到 Windows 文件剪贴板。之后可以去 WorkBuddy 测试 **Ctrl+V**，也可以在资源管理器里粘贴检查结果。

## 启动与停止

无需 Codex 或 Node.js 的独立入口：双击项目目录中的 `Start-MailCopy.cmd` 启动，`Status-MailCopy.cmd` 查看状态，`Stop-MailCopy.cmd` 停止。启动成功后窗口关闭，后台继续运行。这些入口与下列 Node CLI 共用状态目录，可以互相管理。

需要发给其他电脑时，运行 `powershell.exe -NoProfile -ExecutionPolicy Bypass -File scripts/build-clipboard-package.ps1`，生成 `dist/Outlook-MailCopy-standalone.zip`。ZIP 仅包含 Windows PowerShell/.NET 源码、三个双击入口和说明，不含 Node.js、邮件或账号数据。完整解压后即可使用；操作步骤及可选登录自启见 [独立版说明](mailcopy-standalone.txt)。

在项目目录执行：

```powershell
.\outlook.cmd clipboard start
.\outlook.cmd clipboard status
.\outlook.cmd clipboard stop
```

启动命令返回后即可复制邮件。程序在后台运行，无需启动邮件查询服务；重复启动不会增加监听进程。重启电脑或注销后需要再次启动，不会注册开机启动。

程序与 Outlook 应使用同一 Windows 用户、交互式会话和权限级别，通常都用普通权限。仅支持 Windows 经典 Outlook（`OUTLOOK.EXE`）。新版 Outlook、网页版和 macOS 不在此功能范围。

## 验证方法

1. 启动后，在 Outlook **邮件列表**中选择一封邮件，按 Ctrl+C。不要只选择正文里的文字。
2. 稍等导出完成。小邮件通常很快，大附件需要更多时间；`clipboard status` 的 `exported_count` 增长表示文件已准备好，`last_files` 给出路径。
3. 去 WorkBuddy 输入框按 Ctrl+V，检查是否出现 MSG 附件。此程序只负责导出和文件剪贴板；WorkBuddy 的文件粘贴与 MSG 解析能力需要实际验证。
4. 若 WorkBuddy 没有反应，可以先在资源管理器的空文件夹里 Ctrl+V，检查 MSG 是否可正常粘贴和打开，或直接通过 WorkBuddy 的上传按钮选择导出的文件。

## 数据与行为

- 默认文件位于 `.runtime/service/clipboard/exports/<时间-随机标识>/`，状态中的 `export_dir` 是实际绝对路径。设置 `OUTLOOK_LOCAL_STATE_DIR` 时，位置随服务状态目录改变。
- 直接从本次复制的 OLE 数据对象保存 MSG，保留源文件里的正文、格式和内嵌附件，不从当前选中邮件重新读取，也不转换为 TXT。不强制改变 Outlook 原本提供的 MSG 编码。
- 不发送、删除、移动邮件或修改已读状态。只处理 Outlook 提供的整批 `.msg` 虚拟文件；普通文字、图片、其他应用的复制和现有文件剪贴板不会触发导出。仅把路径作为文字复制也不会触发。
- 支持一次复制 1–20 封邮件，整批导出大小最多 256 MiB。结构化存储在保存后检查大小。任一封失败，不发布不完整批次，错误显示在 `last_error`。
- 每次复制使用独立目录；同主题邮件用序号区分，不覆盖已有文件。检查 MSG 的复合文件头，但不重写邮件内容。
- 如果导出期间复制了其他内容，保留新剪贴板；已导出的文件可能留在导出目录。剪贴板写入被占用时会短暂重试。
- 成功后剪贴板变为普通文件列表，原 Outlook 专用剪贴板格式被替换。在 Outlook 内部粘贴的行为因此可能与原生复制不同；需要原生复制时先停止本程序。
- 导出文件不会自动删除，避免 WorkBuddy 尚未读取时文件消失。用完后可手动清理上述 `exports` 目录；清理后旧文件剪贴板将不再有效。停止监听不会删除文件。
- `responsive: false` 表示超过 15 秒没有心跳，可能正等待 Outlook 提供大型邮件；停止请求会在当前操作返回后生效。不会强制结束 Outlook。

诊断日志位于状态目录下的 `clipboard/clipboard.log`。程序不记录邮件正文。启动前已有的剪贴板内容不会处理，请启动后重新复制邮件。

## 实现与测试

`powershell/ClipboardBridge.cs` 在 STA 线程中读取 `FileGroupDescriptorW`（或 ANSI 描述符）和 `FileContents`，支持 `IStream`、`IStorage`、`HGLOBAL`，输出真实 MSG 文件，然后原子检查剪贴板序列号并发布 `CF_HDROP`。PowerShell 负责心跳与停止请求，Node CLI 负责后台进程启动和状态显示。

```powershell
node --test test/clipboard.test.mjs
```

测试使用合成数据验证中文文件名、重复文件名、路径清理、无效描述符、文件剪贴板结构，以及三种存储介质导出；验证结构化存储中的嵌套附件数据保持不变。测试不会写入系统剪贴板或读取真实邮箱。
