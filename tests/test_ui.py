"""界面冒烟测试：实例化窗口、加载真实文件、检查渲染结果，但不进入 mainloop。"""
import datetime as dt
import os
import shutil
import sys
import tempfile
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent
APP = ROOT.parent
sys.path.insert(0, str(APP))

from metadata_editor import engine  # noqa: E402
from metadata_editor.ui import Editor  # noqa: E402

FAILURES = []


def check(label, condition, detail=""):
    print(f"  [{'PASS' if condition else 'FAIL'}] {label}")
    if not condition:
        print(f"         {detail}")
        FAILURES.append(label)


def make_docx(path):
    core = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
<dc:title>原始标题</dc:title><dc:creator>原作者</dc:creator>
</cp:coreProperties>"""
    app = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties">
<Application>Microsoft Office Word</Application><Company>原始公司</Company>
</Properties>"""
    custom = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
<property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="2" name="项目编号"><vt:lpwstr>A-100</vt:lpwstr></property>
</Properties>"""
    ct = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="xml" ContentType="application/xml"/>
</Types>"""
    rels = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
</Relationships>"""
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", ct)
        zf.writestr("_rels/.rels", rels)
        zf.writestr("docProps/core.xml", core)
        zf.writestr("docProps/app.xml", app)
        zf.writestr("docProps/custom.xml", custom)


def make_pdf(path):
    body = (
        "%PDF-1.4\n"
        "1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj\n"
        "4 0 obj\n<< /Title (原始 PDF 标题) /Author (原作者) /Producer (测试生成器) "
        "/CreationDate (D:20190101000000Z) >>\nendobj\n"
    ).encode("utf-8")
    path.write_bytes(
        body
        + b"xref\n0 5\n0000000000 65535 f \ntrailer\n<< /Size 5 /Root 1 0 R /Info 4 0 R >>\nstartxref\n"
        + str(len(body)).encode()
        + b"\n%%EOF\n"
    )


work = Path(tempfile.mkdtemp(dir=str(ROOT)))
try:
    docx = work / "ui-sample.docx"
    make_docx(docx)
    pdf = work / "ui-sample.pdf"
    make_pdf(pdf)

    print("\n【1】窗口实例化与几何")
    app = Editor()
    app.update_idletasks()
    geom = app.geometry()
    check("窗口几何已设置", geom.startswith("1"), geom)
    w, rest = geom.split("x", 1)
    h = rest.split("+")[0]
    check("宽度在合理区间", 760 <= int(w) <= 1240, geom)
    check("高度在合理区间", 400 <= int(h) <= 780, geom)
    print(f"         实际几何: {geom}  屏幕: {app.winfo_screenwidth()}x{app.winfo_screenheight()}")

    print("\n【2】字体解析（不能回退到 .AppleSystemUIFont）")
    check("界面字体已选中真实中文字体", app.ui_font != ".AppleSystemUIFont" and bool(app.ui_font), app.ui_font)
    check("等宽字体已选中真实字体", app.mono_font != "Courier" and bool(app.mono_font), app.mono_font)
    print(f"         界面字体={app.ui_font}  等宽字体={app.mono_font}")

    print("\n【3】菜单栏（macOS 剪贴板快捷键依赖 Edit 菜单）")
    menu = app.nametowidget(app.cget("menu"))
    labels = [menu.entrycget(i, "label") for i in range(menu.index("end") + 1)]
    check("存在文件菜单", "文件" in labels, str(labels))
    check("存在编辑菜单", "编辑" in labels, str(labels))
    edit = app.nametowidget(menu.entrycget(labels.index("编辑"), "menu"))
    edit_labels = [edit.entrycget(i, "label") for i in range(edit.index("end") + 1)]
    check("编辑菜单含复制/粘贴", {"复制", "粘贴", "剪切", "全选"} <= set(edit_labels), str(edit_labels))

    print("\n【4】加载 docx 并渲染表单")
    app.open(str(docx))
    app.update_idletasks()
    check("状态栏显示已读取", "已读取" in app.status.cget("text"), app.status.cget("text"))
    check("kind 为 Office", app.meta.kind == "Office", str(app.meta.kind))
    var_keys = set(app.vars)
    check("表单含文件系统创建时间", "filesystem_created" in var_keys, str(sorted(var_keys)))
    check("表单含标题字段", "title" in var_keys, str(sorted(var_keys)))
    check("创建时间值已填入", app.vars["filesystem_created"].get() != "", app.vars["filesystem_created"].get())
    check("标题值已填入", app.vars["title"].get() == "原始标题", app.vars["title"].get())
    check("自定义属性框已渲染", hasattr(app, "custom_text") and "A-100" in app.custom_text.get("1.0", "end"), app.custom_text.get("1.0", "end") if hasattr(app, "custom_text") else "无 custom_text")
    tree_rows = app.tree.get_children()
    check("左侧树有三个分组", len(tree_rows) == 3, str(len(tree_rows)))

    print("\n【5】加载 PDF 并渲染（不应出现自定义属性框）")
    app.open(str(pdf))
    app.update_idletasks()
    check("kind 为 PDF", app.meta.kind == "PDF", app.meta.kind)
    check("PDF 表单含 producer", "producer" in app.vars, str(sorted(app.vars)))
    check("PDF 无自定义属性框", not hasattr(app, "custom_text") or not app.custom_text.winfo_exists(), "custom_text 仍存在")

    print("\n【6】编辑并保存（走完整 UI 保存路径）")
    app.vars["title"].set("来自 UI 的新标题")
    app.vars["author"].set("UI 测试者")
    created = dt.datetime(2003, 4, 5, 6, 7, 8)
    app.vars["filesystem_created"].set(created.strftime(engine.TIME_FORMAT))
    app.backup_var.set(False)
    app.save()
    app.update_idletasks()
    check("保存后状态为成功", "保存成功" in app.status.cget("text"), app.status.cget("text"))
    check("标题已写入", app.vars["title"].get() == "来自 UI 的新标题", app.vars["title"].get())
    stat = os.stat(pdf)
    birth = dt.datetime.fromtimestamp(stat.st_birthtime if hasattr(stat, "st_birthtime") else stat.st_ctime)
    check("UI 保存后创建时间生效", birth == created, f"{birth} != {created}")
    raw = engine.load(str(pdf))
    check("PDF 内嵌标题已更新", raw.values["title"] == "来自 UI 的新标题", raw.values["title"])

    print("\n【7】滚轮适配（直接监听滚动步数，不依赖画布当前位置）")
    app.update()  # 必须真正映射窗口，否则 winfo_rootx/y 恒为 0，坐标判断会失真
    cx = app.form_canvas.winfo_rootx() + app.form_canvas.winfo_width() // 2
    cy = app.form_canvas.winfo_rooty() + app.form_canvas.winfo_height() // 2

    calls = []
    original_scroll = app.form_canvas.yview_scroll

    def spy(*args):
        calls.append(args)
        return original_scroll(*args)

    app.form_canvas.yview_scroll = spy

    def wheel(delta):
        calls.clear()
        app._mousewheel(type("E", (), {"x_root": cx, "y_root": cy, "delta": delta})())
        app.update_idletasks()
        return calls[0][0] if calls else None

    # 符号约定沿用原版（未改动）：正 delta = 滚轮远离用户 = 内容向上滚。
    check("正 delta 向上滚动（steps 为负）", wheel(3) == -1, f"calls={calls}")
    check("负 delta 向下滚动（steps 为正）", wheel(-3) == 1, f"calls={calls}")

    # 量级适配是本次改动点：macOS 触控板 delta 只有 ±1..±15，
    # 原式 abs(delta)//120 恒为 0，只能靠 max(1,...) 兜底；现在小 delta 走 1 步，
    # 大 delta 仍按 120 的整数倍加速。
    check("macOS 量级 delta=3 → 1 步", wheel(3) == -1, f"calls={calls}")
    check("macOS 量级 delta=-1 → 1 步", wheel(-1) == 1, f"calls={calls}")
    check("Windows 量级 delta=120 → 1 步", wheel(120) == -1, f"calls={calls}")
    check("Windows 量级 delta=-360 → 3 步（保留加速）", wheel(-360) == 3, f"calls={calls}")
    check("Windows 量级 delta=1200 → 10 步", wheel(1200) == -10, f"calls={calls}")

    app.form_canvas.yview_scroll = original_scroll

    print("\n【7b】树视图滚轮")
    tx = app.tree.winfo_rootx() + app.tree.winfo_width() // 2
    ty = app.tree.winfo_rooty() + app.tree.winfo_height() // 2
    tree_calls = []
    original_tree_scroll = app.tree.yview_scroll
    app.tree.yview_scroll = lambda *a: (tree_calls.append(a), original_tree_scroll(*a))[1]
    result = app._mousewheel(type("E", (), {"x_root": tx, "y_root": ty, "delta": -3})())
    check("指针在树上时返回 break（消费事件）", result == "break", repr(result))
    check("滚动作用在树上而非画布上", len(tree_calls) == 1, f"tree={tree_calls}")
    app.tree.yview_scroll = original_tree_scroll

    print("\n【7c】指针在空白区域时不滚动")
    blank_calls = []
    app.form_canvas.yview_scroll = lambda *a: blank_calls.append(a)
    app._mousewheel(type("E", (), {"x_root": 2, "y_root": 2, "delta": -3})())
    check("指针不在滚动区内时不触发滚动", not blank_calls, f"calls={blank_calls}")
    app.form_canvas.yview_scroll = original_scroll

    print("\n【8】快捷键绑定")
    binds = app.bind_all()
    check("已绑定 ⌘S 或 Ctrl+S", any("s>" in b.lower() or "Key-s" in b for b in binds), str(binds))
    check("已绑定 ⌘O 或 Ctrl+O", any("o>" in b.lower() or "Key-o" in b for b in binds), str(binds))

    app.destroy()
finally:
    shutil.rmtree(work, ignore_errors=True)

print("\n" + "=" * 60)
if FAILURES:
    print(f"共 {len(FAILURES)} 项失败：")
    for item in FAILURES:
        print(f"  - {item}")
    sys.exit(1)
print("界面测试全部通过。")
