# DOCM → DOCX 批量无损转换器

把电脑上的 `.docm`(启用宏的 Word 文档)批量转成 `.docx`,**内容无损**。
带图形界面、实时进度、可取消、可导出日志,并内置一套「生成测试文档 → 转换 → 机器校验 → 人工复核」的自检功能。

---

## 〇、一键入口

| 方式 | 说明 |
| --- | --- |
| **桌面快捷方式** | 双击桌面上的 **DOCM转DOCX转换器** 图标即可(已在本机创建好) |
| **`run.bat`** | 程序目录里双击,效果相同 |
| **命令行** | `python docm2docx.py` 与上面两者等价 |

一键启动的内部流程:

1. 先跑一次**环境体检** —— Python / pywin32 / tkinter / Microsoft Word 是否就绪;
2. **优先使用内置的便携运行时** `runtime\python\pythonw.exe`(没有才退回系统 Python),全程不弹黑色控制台窗口;
3. 万一启动崩溃,会**弹窗报错**,并把完整堆栈写到程序目录下的 `启动错误.log`。

想手动体检:

```bat
python docm2docx.py --check
```

换了机器或移动了目录后,双击 `创建桌面快捷方式.bat` 可重新生成桌面图标(图标名称取自 `shortcut_name.txt`)。不需要桌面图标的话,直接删掉那个 `.lnk` 即可,不影响程序。

---

## 〇之二、便携运行时(免安装)

程序目录下的 `runtime\` 是**内置的便携运行环境**,整个文件夹拷到任何 Windows 电脑都能直接跑,**不需要装 Python、不需要装 pywin32、不需要装 Everything**:

```
runtime\
├─ python\      精简版 Python 3.14(含 tkinter 图形库 + pywin32 + pip),约 66 MB
└─ everything\  Everything 主程序,约 3 MB
```

- `run.bat` 会**自动优先使用** `runtime\python\pythonw.exe`。
- **Everything 内核**:若本机没有正在运行的 Everything,程序会**自动拉起** `runtime\everything\Everything.exe` 并等它的 IPC 就绪;退出时把由本工具启动的那个关掉(你自己开的实例不受影响)。若已有 Everything 实例但 IPC 不可用,则不会启动第二个去抢窗口,而是自动降级到 Windows 搜索索引。
- `runtime\` 由 `build-runtime.bat` 生成(它会从本机 Python 精简复制一份,并收集 Everything;Everything 没装则从 voidtools.com 下载)。换机器时先跑一次这个脚本即可。
- `runtime\` 约 70MB,**不入 Git**,需要时用脚本重新生成。

> 想让内置的 Everything 拥有完整 NTFS 索引,请用管理员身份运行一次
> `runtime\everything\Everything.exe -install-service`,装成系统服务后索引最全。

---

## 一、特性

| 能力 | 说明 |
| --- | --- |
| 扫描范围自由 | 整机所有分区、单个分区、任意文件夹、或手动挑选单个/多个 `.docm` 文件,可混选 |
| **高速搜索内核** | 内置三档内核:**Everything 索引**(毫秒级)、**Windows 搜索索引**(快)、直接遍历(兜底)。同一时刻结果与遍历完全一致,速度快一个数量级 |
| 内容无损 | 通过 Word 引擎另存,标题层级、样式、列表、表格、图片、形状、超链接、域、页眉页脚、脚注、分页、字符级格式全部保留 |
| 宏安全 | 打开 `.docm` 时强制禁用宏(`AutomationSecurity=ForceDisable`),不会执行任何 VBA |
| 不动原件 | 全程只读打开,`.docm` 原文件保持不变,`.docx` 另存为新文件 |
| 进程干净 | 用完即关:转换/校验结束后后台 Word 实例自动退出,不残留 `WINWORD.EXE`(每个约 170 MB) |
| 实时进度 | 进度条 + 当前文件 + 成功/跳过/失败计数 + 速度(个/分)+ 总耗时 |
| 可中断 | 扫描和转换都能随时取消,已完成的成果保留 |
| 输出策略 | 与源文件同目录 / 指定统一目录(保留原目录结构) |
| 源文件处理 | 转换成功后可选**不处理 / 移到回收站(可恢复) / 永久删除**,默认不处理;跳过与失败的文件一律保留 |
| 冲突策略 | 目标已存在时:跳过 / 覆盖 / 自动重命名 `name (1).docx` |
| 日志导出 | 一键把完整日志存成 `.log` |
| 自检闭环 | 生成含 33 类元素的测试 `.docm` → 转换 → 13 项机器校验 → 在 Word 中并排打开供人工复核 |

---

## 二、环境要求

- Windows + Microsoft Word(2010 及以上,已安装且能正常打开)
- Python 3.8+(已在 3.14 实测)
- 依赖:`pywin32`(仅此一项)

```bat
pip install -r requirements.txt
```

> 必须用 **CPython**,不能用 IronPython/PyPy —— 需要真实调用 Word COM。

---

## 三、快速开始

### 图形界面

双击 `run.bat`(或桌面快捷方式),或:

```bat
python docm2docx.py
```

### 命令行

```bat
:: 扫描某个分区里的 .docm
python docm2docx.py scan D:\

:: 批量转换(输出到统一目录,同名则自动重命名)
python docm2docx.py convert D:\资料 --out E:\输出 --overwrite rename

:: 转换成功后把原 .docm 移到回收站(可恢复);会先要求确认
python docm2docx.py convert D:\资料 --delete-source recycle

:: 转换成功后永久删除原 .docm(不可恢复);-y 跳过确认
python docm2docx.py convert D:\资料 --delete-source delete -y

:: 一键自检:生成测试文档 → 转换 → 机器校验
python docm2docx.py selftest --dir .\selftest_output --keep
```

`selftest` 全部检查通过时退出码为 `0`,有问题时为 `3`,方便接 CI。

输出目录不存在时会自动创建;`--out` 指定统一目录时,会保留原有的子目录结构。

---

## 四、界面说明

### 页签一:批量转换

```
① 扫描目标   分区下拉框(可选"全部分区")+ 添加文件夹 / 添加文件,可混选
② 扫描       递归子目录、跳过系统目录、移除阻止标记、**搜索内核选择**;实时显示已找到数量与耗时
③ 文件清单   文件名 / 所在目录 / 大小 / 修改时间 / 状态 / 说明(状态实时更新、绿黄红三色)
④ 转换选项   输出位置(同目录 / 指定目录)、目标已存在时的策略、**转换成功后如何处理原 .docm**、开始 / 取消 / 导出日志
⑤ 进度       进度条、当前文件、成功/跳过/失败/总计、用时与速度
⑥ 日志       完整过程日志,可导出
```

操作顺序:**加目标 → 开始扫描 → 设选项 → 开始转换**。

### 页签二:测试与校验

| 按钮 | 作用 |
| --- | --- |
| ① 生成测试 DOCM | 造一份含标题/多级标题/项目符号/编号/两张表格/PNG 图片/矩形形状/超链接/日期域/页码域/页眉页脚/分页符/字符格式/脚注/中英文混排/特殊字符的 `.docm` |
| ② 转换为 DOCX | 用与正式批处理**完全相同**的引擎转换 |
| ③ 机器校验 | 13 项自动检查,见下表 |
| ④ 人工复核 | 在 Word 中同时打开 `.docm` 与 `.docx`,肉眼逐项比对 |

#### 机器校验的 13 项检查

| # | 检查项 | 判据 |
| --- | --- | --- |
| 1 | 源文件存在 | — |
| 2 | 产物文件存在 | — |
| 3 | DOCX 包结构完整 | ZIP CRC 校验通过 + 全部 XML 部件可解析 + 必需部件齐全 |
| 4 | 宏工程已按预期移除 | 产物不含 `vbaProject.bin` |
| 5 | 正文文本逐字一致 | 全文逐字比对,相似度需 100% |
| 6 | 逐段文本顺序一致 | 逐段文本 + 段落顺序完全一致 |
| 7 | 内容结构元素计数一致 | 段落/表格/行/单元格/图片/域/书签/节/脚注引用/字符数 |
| 8 | run 数(仅提示) | Word 另存时会重排 run(如把简写域展开成复杂域),只提示不判失败 |
| 9 | 附件部件数量一致 | 页眉/页脚/媒体/脚注/尾注/嵌入对象 |
| 10 | 页眉页脚文本一致 | 页眉页脚文字逐字相同 |
| 11 | Word 可正常打开产物 | Word 真实打开不报错、不触发修复流程 |
| 12 | Word 层面内容统计一致 | 段落/表格/内嵌图/形状/域/超链接/脚注/节/页数/字符数 |
| 13 | Word 读取的正文内容一致 | Word 打开后读到的正文起始内容相同 |

---

## 五、高速搜索内核

扫描 `.docm` 曾经是整个工具最慢的一环:一个 E 盘遍历要好几分钟。现在内置三档内核,可以在界面上直接切换:

| 内核 | 原理 | 实测(E 盘,同一时刻) | 前提 |
| --- | --- | --- | --- |
| **Everything 索引** | 走 Everything 的 IPC 索引(基于 NTFS USN 日志,常驻) | **1.3 秒 / 36 个** | 本机装有并运行着 Everything;没有的话会自动拉起内置的那份 |
| **Windows 搜索索引** | OLEDB `Search.CollatorDSO` 查系统索引 | 0.2 秒 / 更少 | Windows 搜索服务(WSearch)运行中 |
| **直接遍历** | 自己走文件系统 | 14 秒(冷缓存 126 秒)/ **同样 36 个** | 无,永远可用 |

两种内核在**同一时刻实测结果完全一致**(36/36),Everything 只是快了一个数量级。注意 Everything 的索引里也包含回收站中的 `$I*.docm` / `$R*.docm` 元数据文件,程序会自动剔除。

**默认是「自动」**:同时跑 Everything 与 Windows 索引并**取并集**——两者覆盖范围不同(实测 Everything 136 个 vs Windows 索引 47 个),合并才不漏;只有当两个快内核都没给出结果时,才回退到遍历。这样即使目标目录不在任何索引范围内,也不会误判成"没有文件"。

内核实现在 `d2d/search_kernel.py`,思路移植自 Flow Launcher 的 `Flow.Launcher.Plugin.Explorer/Search`(三档内核、Everything SDK 调用方式、Windows 索引的 SQL 构造都是照搬它的做法)。Everything 的 IPC 用它附带的官方 SDK 助手库 `d2d/native/Everything.dll`,通过 ctypes 调用。

内核结果会做三轮过滤:剔除 `$RECYCLE.BIN` / `System Volume Information` 里的条目、剔除 Word 的 `~$` 临时文件、剔除索引里已失效(文件已不存在)的条目。

命令行也能指定:

```bat
python docm2docx.py scan E:\ --kernel everything
python docm2docx.py convert E:\ --kernel auto
```

---

## 六、"无损"到底指什么

- **保留**:文字、排版样式、多级标题、项目符号与编号、表格(行列/边框/合并)、图片与形状、超链接、域(可继续自动更新)、页眉页脚、脚注尾注、分节与分页、字符级格式(字体/字号/颜色/高亮/下划线)。
- **丢弃**:VBA 宏工程(转成 `.docx` 必然丢失,这是格式本身的要求,不是缺陷)。
- **可能变化**:`run` 数量、XML 内部序列化方式 —— 例如 `<w:fldSimple>` 简写域在另存时被展开成 `fldChar`/`instrText` 复杂域。这是 Word 自身的重写行为,渲染结果与文本完全一致,自检里因此只作提示。

实现方式:用 `Word.Application` 独立后台实例,只读打开 `.docm`,再 `SaveAs2(FileFormat=16)` 另存为 `.docx` —— 由 Word 自己完成格式转换,而不是拼 XML,因此兼容性最好。

---

## 七、常见问题

**Q:提示"无法启动 Microsoft Word"?**
A: 未安装 Word,或被安全策略禁用。安装 Office 后重试。

**Q:个别文件失败,提示"文档被锁定"?**
A: 该 `.docm` 正被其他程序打开(常见于正开在 Word 里)。关掉后重新转换,或在清单里重新扫描。

**Q:从网上下载的文档打不开/转换失败?**
A: 勾选「移除阻止标记(Protected View)」。该选项只删除 NTFS 的 `Zone.Identifier` 备用流,不动文档内容。

**Q:整盘扫描很慢?**
A: 正常 —— 磁盘是瓶颈。默认已跳过回收站、Windows、Program Files、`$Recycle.Bin`、版本库等目录。想全盘逐字节找就取消「跳过系统目录」。扫描阶段可随时取消。

**Q:转换速度?**
A: 受 Word 另存速度限制,实测约 20–60 个/分钟(单个 Word 实例串行处理)。大文档会明显更慢。

**Q:转换后想把原 .docm 删掉,怎么操作?**
A:「④ 转换选项 → 转换成功后处理原 .docm」有三个档位:

| 档位 | 行为 |
| --- | --- |
| 不处理(保留原 .docm) | 默认,源文件原封不动 |
| 移到回收站(可恢复,推荐) | 源文件进 Windows 回收站,需要时可还原 |
| 永久删除(不可恢复!) | 直接抹掉,选它会弹二次确认 |

**只有转换成功的文件会被处理。** 因为同名 `.docx` 已存在而"跳过"的、或者转换失败的,源文件一律保留 —— 不会出现"没转成却把源文件删了"的情况。删除失败(文件被占用、只读等)也会记进日志,源文件保留。

**Q:会覆盖我原来的 `.docx` 吗?**
A: 默认**不覆盖**。目标同名时按你选的策略处理:跳过 / 覆盖 / 自动重命名。

**Q:会不会执行文档里的宏?**
A: 不会。打开时强制 `AutomationSecurity=ForceDisable`,宏既不执行也不弹窗。

---

## 八、项目结构

```
docmtodocx/
├─ docm2docx.py          入口(默认启动 GUI,带参数走 CLI)
├─ run.bat               一键启动(环境体检 + 优先用内置便携运行时)
├─ build-runtime.bat     生成便携运行时(runtime\python + runtime\everything)
├─ build-runtime.ps1     上面那个脚本调用的实际逻辑
├─ 创建桌面快捷方式.bat   在桌面创建"DOCM转DOCX转换器"图标
├─ create_shortcut.ps1   上面那个脚本调用的实际逻辑
├─ shortcut_name.txt     桌面快捷方式的中文名称
├─ requirements.txt
├─ README.md
└─ d2d/
   ├─ constants.py       Word/OOXML 常量
   ├─ scanner.py         文件系统遍历(兜底内核)
   ├─ search_kernel.py   高速搜索内核(Everything / Windows 索引 / 遍历)
   ├─ native/            Everything 官方 IPC SDK 助手库
   ├─ engine.py          Word COM 转换引擎
   ├─ ooxml.py           OOXML 包解析(文本提取、结构统计、完整性校验)
   ├─ testgen.py         测试用 .docm 生成器(纯 Python 造 PNG,无额外依赖)
   ├─ verify.py          13 项机器校验
   ├─ gui.py             Tkinter 图形界面
   └─ cli.py             命令行 / 自检
```

---

## 九、注意事项

- 需要本机装有 **Microsoft Word**,本工具不解析或改写 OOXML,所有转换都由 Word 引擎完成。
- 只处理 `.docm`;`.doc` / `.docx` 不在范围内。
- 建议先在少量文件上试跑并用「测试与校验」页确认效果,再执行整盘批量转换。
- 「移除阻止标记」会修改源文件的 NTFS 备用数据流(不影响文档内容与时间戳),介意可取消勾选。
- `d2d/native/Everything.dll` 是 Everything 官方 IPC SDK 的助手库(取自 Flow Launcher 附带的 `Plugins/Flow.Launcher.Plugin.Explorer/EverythingSDK`,voidtools 公开授权),只用于向本机正在运行的 Everything 查询索引;不启动 Everything、不上传任何数据。
