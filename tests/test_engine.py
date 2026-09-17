"""端到端测试：真实构造 docx 与 PDF，走一遍 engine 的读取/写入/时间设置流程。"""
import datetime as dt
import os
import shutil
import sys
import tempfile
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

ROOT = Path(__file__).resolve().parent.parent  # 项目根（tests/ 的上一级）
sys.path.insert(0, str(ROOT))

from metadata_editor import engine  # noqa: E402

FAILURES = []


def check(label, got, expected):
    ok = got == expected
    print(f"  [{'PASS' if ok else 'FAIL'}] {label}")
    if not ok:
        print(f"         实际: {got!r}")
        print(f"         期望: {expected!r}")
        FAILURES.append(label)


def check_true(label, condition, detail=""):
    print(f"  [{'PASS' if condition else 'FAIL'}] {label}")
    if not condition:
        print(f"         {detail}")
        FAILURES.append(label)


def make_docx(path):
    """构造一个最小可用的 docx（含 core/app/custom 三个属性包）。"""
    core = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
<dc:title>原始标题</dc:title><dc:creator>原作者</dc:creator>
<dcterms:created xsi:type="dcterms:W3CDTF">2019-01-01T00:00:00Z</dcterms:created>
<dcterms:modified xsi:type="dcterms:W3CDTF">2019-01-02T00:00:00Z</dcterms:modified>
</cp:coreProperties>"""
    app = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties">
<Application>Microsoft Office Word</Application><Company>原始公司</Company><TotalTime>42</TotalTime>
</Properties>"""
    custom = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
<property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="2" name="项目编号"><vt:lpwstr>A-100</vt:lpwstr></property>
</Properties>"""
    ct = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
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


def make_pdf(path, title="原始 PDF 标题"):
    """构造最小可用的未压缩 PDF。title 里若含转义括号，可用来验证解析正则。"""
    body = (
        "%PDF-1.4\n"
        "1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj\n"
        "2 0 obj\n<< /Type /Pages /Kids [3 0 R] /Count 1 >>\nendobj\n"
        "3 0 obj\n<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] >>\nendobj\n"
        f"4 0 obj\n<< /Title ({title}) /Author (原作者) /Producer (测试生成器) "
        "/CreationDate (D:20190101000000Z) /ModDate (D:20190102000000Z) >>\nendobj\n"
    ).encode("utf-8")
    xref = (
        b"xref\n0 5\n"
        b"0000000000 65535 f \n0000000009 00000 n \n0000000058 00000 n \n"
        b"0000000115 00000 n \n0000000190 00000 n \n"
        b"trailer\n<< /Size 5 /Root 1 0 R /Info 4 0 R >>\n"
        b"startxref\n" + str(len(body)).encode() + b"\n%%EOF\n"
    )
    path.write_bytes(body + xref)


work = Path(tempfile.mkdtemp(dir=str(ROOT)))
try:
    # ------------------------------------------------------------------ Office
    print("\n【1】Office 读取")
    docx = work / "sample.docx"
    make_docx(docx)
    meta = engine.load(str(docx))
    check("kind", meta.kind, "Office")
    check("标题", meta.values["title"], "原始标题")
    check("作者", meta.values["creator"], "原作者")
    check("公司", meta.values["company"], "原始公司")
    check("总编辑时长", meta.values["totalTime"], "42")
    check("创建内容时间", meta.values["created"], "2019-01-01 00:00:00")
    check("自定义属性", meta.custom.get("项目编号"), "A-100")
    check_true("文件系统创建时间非空", bool(meta.values["filesystem_created"]), meta.values)

    print("\n【2】Office 写入 + 文件系统时间")
    before_stat = docx.stat()
    created = dt.datetime(2005, 6, 7, 8, 9, 10)
    modified = dt.datetime(2015, 11, 12, 13, 14, 15)
    values = {k: "" for k, _ in engine.FIELDS}
    values.update({
        "title": "新标题", "creator": "新作者", "company": "新公司",
        "created": "2005-06-07 08:09:10", "modified": "2015-11-12 13:14:15",
        "filesystem_created": created.strftime(engine.TIME_FORMAT),
        "filesystem_modified": modified.strftime(engine.TIME_FORMAT),
    })
    engine.write_office(meta, values, {"项目编号": "B-200", "阶段": "已交付"}, False)
    engine.set_filesystem_times(docx, values["filesystem_created"], values["filesystem_modified"])

    reread = engine.load(str(docx))
    check("回读标题", reread.values["title"], "新标题")
    check("回读作者", reread.values["creator"], "新作者")
    check("回读公司", reread.values["company"], "新公司")
    check("回读文档创建时间", reread.values["created"], "2005-06-07 08:09:10")
    check("回读自定义属性", reread.custom, {"项目编号": "B-200", "阶段": "已交付"})
    check("回读文件系统创建时间", reread.values["filesystem_created"], "2005-06-07 08:09:10")
    check("回读文件系统修改时间", reread.values["filesystem_modified"], "2015-11-12 13:14:15")

    print("\n【3】Office 包结构完整性（写入未破坏 zip）")
    with zipfile.ZipFile(docx) as zf:
        bad = zf.testzip()
        names = zf.namelist()
    check_true("zip 无损坏条目", bad is None, f"损坏: {bad}")
    check_true("保留 [Content_Types].xml", "[Content_Types].xml" in names, str(names))
    check_true("保留 _rels/.rels", "_rels/.rels" in names, str(names))
    check_true("三个属性包都在", {"docProps/core.xml", "docProps/app.xml", "docProps/custom.xml"} <= set(names), str(names))
    with zipfile.ZipFile(docx) as zf:
        root = ET.fromstring(zf.read("docProps/core.xml"))
    check_true("core.xml 无重复节点", len(root.findall("{http://purl.org/dc/elements/1.1/}title")) == 1, "title 出现多次")

    print("\n【4】清除嵌入属性：写入会重置时间，快照+还原才能保持不变")
    preserved_times = engine.capture_times(docx)
    stat_before = os.stat(docx)
    birth_before = stat_before.st_birthtime if hasattr(stat_before, "st_birthtime") else stat_before.st_ctime
    mtime_before = stat_before.st_mtime

    # 4a: 裸写入（不还原）确实会改变文件系统时间 —— 证明还原是必要的
    bare = work / "bare.docx"
    shutil.copy2(docx, bare)
    bare_meta = engine.load(str(bare))
    engine.write_office(bare_meta, {k: "" for k in bare_meta.values}, {}, True)
    bare_stat = os.stat(bare)
    bare_birth = bare_stat.st_birthtime if hasattr(bare_stat, "st_birthtime") else bare_stat.st_ctime
    check_true(
        "裸写入会重置创建时间（因此需要还原）",
        abs(bare_birth - birth_before) > 0.5 or abs(bare_stat.st_mtime - mtime_before) > 0.5,
        f"birth {dt.datetime.fromtimestamp(bare_birth)} vs {dt.datetime.fromtimestamp(birth_before)}",
    )
    cleared_bare = engine.load(str(bare))
    check("裸写入确实清空了标题", cleared_bare.values["title"], "")

    # 4b: 快照 + 写入 + 还原 = 文件系统时间保持不变
    check_true("capture_times 返回二元组", len(preserved_times) == 2, str(preserved_times))
    engine.write_office(reread, {k: "" for k in reread.values}, {}, True)
    check_true("restore_times 返回成功", engine.restore_times(docx, preserved_times) is True)
    after = os.stat(docx)
    after_created = after.st_birthtime if hasattr(after, "st_birthtime") else after.st_ctime
    check("还原后创建时间不变", dt.datetime.fromtimestamp(after_created), dt.datetime.fromtimestamp(birth_before))
    check("还原后修改时间不变", dt.datetime.fromtimestamp(after.st_mtime), dt.datetime.fromtimestamp(mtime_before))
    cleared = engine.load(str(docx))
    check("还原后标题为空", cleared.values["title"], "")
    check("还原后自定义属性为空", cleared.custom, {})
    check_true("还原后文档仍可读且未损坏", zipfile.ZipFile(docx).testzip() is None, "zip 损坏")

    # --------------------------------------------------------------------- PDF
    print("\n【5】PDF 读取")
    pdf = work / "sample.pdf"
    make_pdf(pdf)
    pmeta = engine.load(str(pdf))
    check("kind", pmeta.kind, "PDF")
    check("标题", pmeta.values["title"], "原始 PDF 标题")
    check("作者", pmeta.values["author"], "原作者")
    check("Producer", pmeta.values["producer"], "测试生成器")
    check("创建内容时间", pmeta.values["created"], "20190101000000Z")

    print("\n【6】PDF 增量写入")
    pvalues = {k: "" for k, _ in engine.PDF_FIELDS}
    pvalues.update({"title": "新 PDF 标题 (含括号)", "author": "新作者", "created": "20050607080910"})
    engine.write_pdf(pmeta, pvalues, False)
    again = engine.load(str(pdf))
    check("回读标题", again.values["title"], "新 PDF 标题 (含括号)")
    check("回读作者", again.values["author"], "新作者")
    check("回读创建时间", again.values["created"], "20050607080910")
    check_true("增量写入保留原字节", pdf.read_bytes().startswith(b"%PDF-1.4\n"), "文件头被破坏")
    check_true("追加了新 xref", pdf.read_bytes().count(b"startxref") == 2, "xref 段数异常")

    print("\n【6b】PDF 写入后还原文件系统时间")
    pdf_times = engine.capture_times(pdf)
    engine.write_pdf(again, {k: "" for k in again.values}, True)
    check_true("PDF 还原成功", engine.restore_times(pdf, pdf_times) is True)
    pdf_stat = os.stat(pdf)
    check("PDF 还原后创建时间不变",
          dt.datetime.fromtimestamp(pdf_stat.st_birthtime if hasattr(pdf_stat, "st_birthtime") else pdf_stat.st_ctime),
          pdf_times[0])
    check("PDF 还原后修改时间不变", dt.datetime.fromtimestamp(pdf_stat.st_mtime), pdf_times[1])

    # ------------------------------------------------------------------- 备份
    print("\n【7】备份")
    target = engine.backup(pdf)
    check_true("备份文件已创建", target.exists(), str(target))
    check_true("备份扩展名正确", target.name.endswith(".metadata-backup"), target.name)
    check_true("备份内容与原件一致", target.read_bytes() == pdf.read_bytes(), "内容不同")
    second = engine.backup(pdf)
    check_true("重复备份不覆盖（时间戳不同）", second != target and second.exists(), f"{target} / {second}")

    # --------------------------------------------------------------- 时间边界
    print("\n【8】文件系统时间边界情况")
    probe = work / "times.txt"
    probe.write_text("x")

    engine.set_filesystem_times(probe, "", "")
    check_true("两个参数都为空时不报错", True)

    future = (dt.datetime.now() + dt.timedelta(days=30)).replace(microsecond=0)
    engine.set_filesystem_times(probe, future.strftime(engine.TIME_FORMAT), future.strftime(engine.TIME_FORMAT))
    stat = os.stat(probe)
    birth = stat.st_birthtime if hasattr(stat, "st_birthtime") else stat.st_ctime
    check("创建时间可往后调（未来时间）", dt.datetime.fromtimestamp(birth), future)
    check("修改时间同步为未来时间", dt.datetime.fromtimestamp(stat.st_mtime), future)

    engine.set_filesystem_times(probe, "", "2001-03-04 05:06:07")
    stat = os.stat(probe)
    check("只填修改时间时生效", dt.datetime.fromtimestamp(stat.st_mtime), dt.datetime(2001, 3, 4, 5, 6, 7))

    try:
        engine.set_filesystem_times(probe, "2001/03/04", "")
        check_true("非法时间格式应报错", False, "没有抛出 ValueError")
    except ValueError as exc:
        check_true("非法时间格式应报错", "格式" in str(exc), str(exc))

    print("\n【9】只填创建时间、保留修改时间")
    probe2 = work / "times2.txt"
    probe2.write_text("x")
    keep_mtime = dt.datetime.fromtimestamp(os.stat(probe2).st_mtime)
    engine.set_filesystem_times(probe2, "2001-03-04 05:06:07", "")
    stat = os.stat(probe2)
    birth = stat.st_birthtime if hasattr(stat, "st_birthtime") else stat.st_ctime
    check("创建时间已改", dt.datetime.fromtimestamp(birth), dt.datetime(2001, 3, 4, 5, 6, 7))
    # mtime 完全没被触碰，仍保留原始微秒；按秒比较即可。
    check("修改时间保持原值", dt.datetime.fromtimestamp(stat.st_mtime).replace(microsecond=0), keep_mtime.replace(microsecond=0))

    # --------------------------------------------------------------- 错误处理
    print("\n【10】错误处理")
    legacy = work / "old.doc"
    legacy.write_bytes(b"\xd0\xcf\x11\xe0")
    try:
        engine.load(str(legacy))
        check_true("旧版 .doc 应拒绝", False, "没有报错")
    except ValueError as exc:
        check_true("旧版 .doc 应拒绝", "另存为" in str(exc), str(exc))

    unknown = work / "note.txt"
    unknown.write_text("hi")
    try:
        engine.load(str(unknown))
        check_true("不支持的扩展名应拒绝", False, "没有报错")
    except ValueError as exc:
        check_true("不支持的扩展名应拒绝", "仅支持" in str(exc), str(exc))

    encrypted = work / "secret.pdf"
    encrypted.write_bytes(b"%PDF-1.4\ntrailer\n<< /Encrypt 5 0 R >>\nstartxref\n0\n%%EOF\n")
    try:
        engine.load(str(encrypted))
        check_true("加密 PDF 应拒绝", False, "没有报错")
    except ValueError as exc:
        check_true("加密 PDF 应拒绝", "加密" in str(exc), str(exc))

finally:
    shutil.rmtree(work, ignore_errors=True)

print("\n" + "=" * 60)
if FAILURES:
    print(f"共 {len(FAILURES)} 项失败：")
    for item in FAILURES:
        print(f"  - {item}")
    sys.exit(1)
print("全部测试通过。")
