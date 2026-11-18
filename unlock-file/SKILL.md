---
name: unlock-file
description: 写入/替换文件失败并报 PermissionError、WinError 32、"另一个程序正在使用此文件"、EBUSY，或需要修改正被 Excel/Word/WPS 打开的 xlsx/docx/pptx 时使用。自动定位占用进程，优雅关闭文档窗口（不会静默丢失未保存修改），再重试写入。
license: MIT
compatibility: Windows 10/11; Python 3.8+; pywin32 optional (COM route is more precise)
---

# 解除文件占用

## 何时使用

出现以下任一情况时，先用本技能解除占用，再执行原本的写入/覆盖操作：

- 写文件、编辑文件或脚本保存时（例如 Python 的 `openpyxl.save()`、`shutil.move()`）报 `PermissionError`、`[WinError 32]`、"另一个程序正在使用此文件"、`EBUSY`
- 目标文件旁存在 `~$文件名` 临时锁文件
- 已知用户正开着 Excel / Word / WPS 且需要改写该文件

## 安全红线（必须遵守）

1. **默认只做优雅关闭**：发送窗口关闭消息，等价于点右上角 ×。有未保存修改时，应用会自己弹「是否保存」，不会静默丢数据。
2. **绝不擅自 `--force-kill`**。只有用户明确说"强杀/结束进程/放弃修改"，才可加这个参数；执行前再确认一次。
3. **绝不删除 `~$` 锁文件**。应用还开着时删锁文件可能损坏文档，脚本只报告、不处理。
4. **`needs_decision` / `needs_save_confirm` 必须交还给用户**：使用当前客户端提供的提问/选择界面，让用户选「保存并关闭 / 放弃修改并关闭 / 取消」。如果客户端没有这类工具，就在回复中列出三个选项并等待用户选择；不要自己替用户决定。
5. 被关闭的是**用户正在编辑的文档**，操作前先说明一句"检测到该文件正被 xxx 打开，我先帮你关掉"，让用户有预期。

## 标准流程

本技能的所有命令都以 `SKILL.md` 所在目录为技能根目录。先解析出该目录，再执行：

```bash
python scripts/unlock_file.py "<文件绝对路径>" --json
```

不要假定技能安装在某个固定路径（例如 `~/.claude/skills`）；不同客户端的技能目录不同。如果你的运行环境不会自动把工作目录切到技能根目录，请先 `cd` 到本 `SKILL.md` 所在目录，或使用客户端提供的技能目录变量。

1. 执行上面的检测命令（默认就会尝试优雅关闭）。
2. 按返回的 `status` 处理：

| status | 含义 | 下一步 |
|---|---|---|
| `not_locked` | 没被占用 | 直接重试原写入操作 |
| `closed` | 已关闭，锁已释放 | 重试原写入操作，并告知用户文档已关闭 |
| `needs_decision` | COM 能读到文档，但**有未保存修改** | 询问用户：保存并关闭 / 放弃修改关闭 / 取消 |
| `needs_save_confirm` | 已发关闭请求，但仍锁着——多半弹了保存对话框 | 告诉用户在 WPS/Office 里点一下「保存」或「不保存」，然后重试 |
| `locked_unknown` | 找不到可关的窗口（后台进程/其他软件占用） | 报告占用进程；询问用户是否允许 `--force-kill` |
| `locked` | `--no-close` 模式下的检测结果 | 仅报告，不做处理 |
| `not_found` | 路径不对 | 核对路径 |

3. 写入成功后，简要汇报：关掉了哪个文档、写入是否成功。

## 常用命令

```bash
# 默认：检测 + 优雅关闭（推荐）
python scripts/unlock_file.py "D:/x/报告.xlsx" --json

# 只检测，不做任何关闭动作
python scripts/unlock_file.py "D:/x/报告.xlsx" --no-close --json

# 用户已同意“保存后关闭”（仅对 COM 能读到的文档有效）
python scripts/unlock_file.py "D:/x/报告.xlsx" --save --json

# 用户明确同意强杀（最后手段，会丢未保存修改）
python scripts/unlock_file.py "D:/x/报告.xlsx" --force-kill --json
```

退出码：`0` 已解除/未占用，`2` 需要用户决定，`3` 无法安全解除，`1` 脚本错误。

## 实现说明（排障用）

- **占用检测**：以写模式打开文件，捕获共享冲突。
- **占用进程识别**：Windows Restart Manager API（`rstrtmgr.dll`），能识别 WPS 的 `et.exe` / `wps.exe` 等子进程。
- **优雅关闭**：
  - 优先走 COM（`Excel.Application` / `KET.Application` / `Word.Application` / `KWps.Application` / `KWPP.Application`），能精确判断 `Saved` 状态并只关目标文档；
  - COM 不可用时按**窗口标题里的文件名**找窗口发 `WM_CLOSE`（WPS 是多进程架构，窗口进程 ≠ 持锁进程，按 PID 找会失败）。
- **强杀**：`taskkill /F /PID`，仅在用户明确授权时使用。

## 典型场景

> 用户：帮我把台账加上这两笔，但是 Excel 开着改不了。
>
> 助手：检测到文件正被 WPS 打开，我先帮你关掉再写。
> → 执行脚本 → `closed` → 用 openpyxl 写入 → 回报"已关闭 WPS 中的台账并写入成功，重新打开即可看到"。

> 返回 `needs_decision`（`modified: true`）
>
> 助手：台账里有未保存的修改，你要「保存并关闭」还是「放弃修改并关闭」？