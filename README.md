# 文件元数据编辑器 · macOS 版

本地 macOS 桌面工具，用于查看与编辑新版 Office 文档和 PDF 的常用文件属性。
从 [Abysstutu/file-metadata-editor](https://github.com/Abysstutu/file-metadata-editor)（Windows 版）移植而来，纯标准库实现，无需联网、无需第三方 Python 包。

## 运行

双击 `FileMetadataEditor.app`，或在终端执行：

```sh
python3 app.py      # 直接运行
sh run.sh           # 等价于上面，对应 Windows 版的 run.cmd
sh build_app.sh     # 重新生成 .app
```

需要自带 Tkinter 的 Python 3.8+。[python.org 官方安装包](https://www.python.org/downloads/macos/) 与 Homebrew 的 `python-tk` 均可；macOS 系统自带的 `/usr/bin/python3` 通常没有 Tk，`launcher` 会自动跳过它。

首次双击若提示「无法验证开发者」，执行一次：

```sh
xattr -dr com.apple.quarantine FileMetadataEditor.app
```

## 支持范围

- **Office Open XML**：`.docx`、`.xlsx`、`.pptx` 及其宏/模板变体
  - 核心属性、应用属性、自定义属性
  - 文件系统创建时间与修改时间
- **PDF**：常用 Info 元数据及文件系统时间（不支持加密 PDF）
- 旧式 `.doc`、`.xls`、`.ppt` 为二进制复合文档，只提示不支持，请在 Word、Excel、Pages、Numbers 或 WPS 中另存为新版格式后再编辑。

保存前可选择创建带时间戳的 `.metadata-backup` 备份，重复备份不会互相覆盖。

## 测试

```sh
python3 tests/test_engine.py   # 读写、包结构、时间设置、错误处理
python3 tests/test_ui.py       # 界面渲染、菜单、快捷键、滚轮、保存链路
```

界面测试会短暂弹出窗口并自动关闭。

---

## 移植改动说明

Windows 与 macOS 在「创建时间」这件事上差异极大，这是移植的主要工作量所在。

### 1. 创建时间的读取：`st_ctime` ≠ 创建时间

原代码用 `stat.st_ctime` 表示创建时间。这在 Windows 上成立（NTFS 的 ctime 就是 creation time），但在 macOS/Linux 上 `st_ctime` 是 **inode 变更时间**——`chmod`、改名、写入都会刷新它。实测：对一个文件执行 `chmod` 后，`st_ctime` 变成当前时刻，而真实创建时间不变。

改为优先读取 `st_birthtime`（APFS/HFS+ 提供），取不到时再按平台回退。见 `engine._birthtime()`。

### 2. 创建时间的写入：`setattrlist` + 一个 24 字节的坑

macOS 没有 `SetFileTime`，对应能力是 `setattrlist()` 的 `ATTR_CMN_CRTIME`。这里有个不报错的陷阱：

`struct attrlist` 的 `reserved` 字段是 `u_int16_t`。若按 `c_uint32` 声明，结构体总长变成 28 字节而非内核要求的 24 字节，调用会**一律返回 EINVAL，且不给出任何原因提示**——很容易误判成「macOS 不允许改创建时间」或权限问题。实测在沙箱内外表现一致，排除权限因素后才定位到布局问题。

修正声明后三种场景均验证通过：往前调、往后调、只改创建时间而保留修改时间。属性值还必须按 `attr.h` 中位号从小到大的顺序紧密排列（CRTIME 在 MODTIME 之前）。见 `engine._setattrlist_times()`。

平台分支：macOS 用 `setattrlist`，Windows 保留原 `SetFileTime`，Linux 及不支持的卷（部分网络卷、exFAT）回退到 `os.utime` 两步法。

> 补充：`os.utime` 两步法也能改创建时间，原理是 APFS/HFS+ 强制「创建时间 ≤ 修改时间」，把 mtime 往前调会连带拉低 birthtime。但它**无法把创建时间调到比现值更晚**，因此只作为回退方案。

### 3. 「清除嵌入属性不改动文件系统时间」——原来是个未兑现的承诺

原项目文档写了这条承诺，但代码并没有实现：

- Office 保存走 `os.replace(tmp, path)`，新文件的创建/修改时间会变成「刚刚」；
- PDF 保存走 `path.write_bytes(...)`，会刷新修改时间。

移植时用测试确认了这一点（裸写入后创建时间确实被重置）。现在由 `capture_times()` / `restore_times()` 在写入前后显式存取快照，`ui.clear()` 负责调用，承诺真正成立。快照/还原放在 UI 层而不是引擎层：`write_office` 只负责写内容，是否保留时间由调用方决定。

### 4. 界面适配

| 项目 | Windows | macOS |
| --- | --- | --- |
| 界面字体 | Microsoft YaHei UI | PingFang SC（自动探测，回退 Hiragino Sans GB / Heiti SC） |
| 等宽字体 | Consolas | Menlo（自动探测） |
| 高分屏 | 需 `SetProcessDpiAwareness` + manifest | 系统自动处理，无需任何设置 |
| 剪贴板快捷键 | 默认生效 | **必须有 Edit 菜单**，否则输入框里 ⌘C/⌘V/⌘X 全部无效 |
| 滚轮 delta | ±120 的整数倍 | ±1..±15 的小值 |
| 窗口激活 | 自动前台 | 从 .app 启动时常排在其他应用后面 |
| 打开对话框 | `*.*` | `*` |

其中滚轮那条原代码 `abs(delta) // 120` 在 macOS 上恒为 0，只能靠 `max(1, ...)` 兜底，快速滚动丢失加速；现在按量级分流，小 delta 走 1 步、大 delta 保留整数倍加速（已测试 `1200 → 10 步`）。

另外补了：窗口初始高度扣除菜单栏与标题栏（否则窗口会顶出屏幕）、⌘O/⌘S 快捷键、`-topmost` 抢占前台。

字体探测用 `tkfont.families()` 判断真实存在性，因为 Tk 对不存在的字体名是**静默回退**到 `.AppleSystemUIFont`，不会报错——直接写死 Windows 字体名会「看起来正常」但字重与间距不对。

### 5. 顺带修复的跨平台 bug

`read_pdf` 原用 `(.*?)` 匹配字符串字面量，遇到值内含转义括号会被截断：标题写作「报告 (v2)」（转义后 `报告 \(v2\)`）时只能读回前半段。改为 `((?:\\.|[^)\\])*)`，先吃转义对再吃普通字节。这个问题与平台无关，Windows 版同样存在。

## 未改动的部分

Office/PDF 的元数据解析与写入逻辑、备份策略、界面布局与配色均与原版保持一致，只调整了平台相关部分。`write_office` 中「关闭源 zip 前先读完 entries」的注释也保留了——虽然那是 Windows 的文件占用限制，但提前读取在两个平台上都是更稳妥的写法。
