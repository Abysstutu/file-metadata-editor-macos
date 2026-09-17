from __future__ import annotations

import datetime as dt
import os
import re
import shutil
import sys
import tempfile
import zipfile
from dataclasses import dataclass, field
from pathlib import Path
from xml.etree import ElementTree as ET

CORE = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
DC = "http://purl.org/dc/elements/1.1/"
DCTERMS = "http://purl.org/dc/terms/"
CP = "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties"
VT = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"
APP = "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
NS = {"cp": CORE, "dc": DC, "dcterms": DCTERMS, "app": APP}
for prefix, uri in {"cp": CORE, "dc": DC, "dcterms": DCTERMS, "vt": VT}.items():
    ET.register_namespace(prefix, uri)

OFFICE_SUFFIXES = {".docx", ".xlsx", ".pptx", ".docm", ".xlsm", ".pptm", ".dotx", ".xltx", ".potx", ".dotm", ".xltm", ".potm", ".ppsx", ".ppsm"}
LEGACY_SUFFIXES = {".doc", ".xls", ".ppt"}

FIELDS = [
    ("title", "标题"), ("subject", "主题"), ("creator", "作者"), ("lastModifiedBy", "最后保存者"),
    ("keywords", "标记 / 关键词"), ("category", "类别"), ("description", "备注"), ("contentStatus", "状态"),
    ("revision", "修订号"), ("created", "创建内容时间"), ("modified", "最后一次保存时间"),
    ("lastPrinted", "最后打印时间"), ("totalTime", "总编辑时长（分钟）"), ("application", "程序名称"),
    ("company", "公司"), ("manager", "管理者"), ("hyperlinkBase", "超链接基址"), ("template", "文档模板"),
]
PDF_FIELDS = [("title", "标题"), ("author", "作者"), ("subject", "主题"), ("keywords", "关键词"), ("creator", "Creator"), ("producer", "Producer"), ("created", "创建内容时间"), ("modified", "修改内容时间")]

TIME_FORMAT = "%Y-%m-%d %H:%M:%S"


@dataclass
class Metadata:
    path: Path
    kind: str
    values: dict[str, str] = field(default_factory=dict)
    custom: dict[str, str] = field(default_factory=dict)
    readonly: dict[str, str] = field(default_factory=dict)


def iso_display(value: str) -> str:
    return value.replace("T", " ").replace("Z", "")[:19] if value else ""


def iso_store(value: str) -> str:
    if not value:
        return ""
    value = value.strip().replace(" ", "T")
    if "T" in value and not value.endswith("Z") and "+" not in value[10:]:
        return value + "Z"
    return value


# ---------------------------------------------------------------------------
# 文件系统时间
#
# 三个平台对「创建时间」的表达完全不同，这是移植的主要难点：
#   Windows : os.stat_result.st_ctime 就是创建时间，写入用 SetFileTime。
#   macOS   : st_ctime 是 inode 变更时间（chmod、改名都会刷新它），创建时间
#             在 st_birthtime，写入用 setattrlist 的 ATTR_CMN_CRTIME。
#   Linux   : 标准 os.stat 拿不到创建时间，也写不了，只能处理修改时间。
# ---------------------------------------------------------------------------


def _birthtime(stat: os.stat_result) -> float:
    """取出该平台的「创建时间」时间戳，取不到时退回修改时间。"""
    birth = getattr(stat, "st_birthtime", None)
    if birth is not None:
        return birth
    if os.name == "nt":
        return stat.st_ctime
    return stat.st_mtime


if sys.platform == "darwin":
    import ctypes

    class _AttrList(ctypes.Structure):
        """<sys/attr.h> 的 struct attrlist，总长必须是 24 字节。

        reserved 是 u_int16_t；若按 c_uint32 声明，结构体会变成 28 字节，
        内核校验 bitmapcount 之外的长度就会直接返回 EINVAL，且不给出任何提示。
        """

        _fields_ = [
            ("bitmapcount", ctypes.c_uint16),
            ("reserved", ctypes.c_uint16),
            ("commonattr", ctypes.c_uint32),
            ("volattr", ctypes.c_uint32),
            ("dirattr", ctypes.c_uint32),
            ("fileattr", ctypes.c_uint32),
            ("forkattr", ctypes.c_uint32),
        ]

    class _TimeSpec(ctypes.Structure):
        _fields_ = [("tv_sec", ctypes.c_int64), ("tv_nsec", ctypes.c_int64)]

    _ATTR_BIT_MAP_COUNT = 5
    _ATTR_CMN_CRTIME = 0x00000200
    _ATTR_CMN_MODTIME = 0x00000400

    _libsystem: ctypes.CDLL | None = None

    def _libc() -> ctypes.CDLL:
        global _libsystem
        if _libsystem is None:
            library = ctypes.CDLL("libSystem.dylib", use_errno=True)
            library.setattrlist.argtypes = [ctypes.c_char_p, ctypes.POINTER(_AttrList), ctypes.c_void_p, ctypes.c_size_t, ctypes.c_uint]
            library.setattrlist.restype = ctypes.c_int
            _libsystem = library
        return _libsystem

    def _setattrlist_times(path: Path, created: dt.datetime | None, modified: dt.datetime | None) -> None:
        attrs = _AttrList()
        attrs.bitmapcount = _ATTR_BIT_MAP_COUNT
        # 属性值必须按 attr.h 中位号从小到大的顺序紧密排列：CRTIME 在 MODTIME 之前。
        payload = b""
        for mask, value in ((_ATTR_CMN_CRTIME, created), (_ATTR_CMN_MODTIME, modified)):
            if value is None:
                continue
            attrs.commonattr |= mask
            stamp = _TimeSpec(int(value.timestamp()), value.microsecond * 1000)
            payload += ctypes.string_at(ctypes.byref(stamp), ctypes.sizeof(stamp))
        if not attrs.commonattr:
            return
        buffer = (ctypes.c_char * len(payload)).from_buffer_copy(payload)
        ctypes.set_errno(0)
        if _libc().setattrlist(os.fsencode(path), ctypes.byref(attrs), buffer, ctypes.c_size_t(len(payload)), ctypes.c_uint(0)) != 0:
            error = ctypes.get_errno()
            raise OSError(error, os.strerror(error), str(path))


def _utime_times(path: Path, created: dt.datetime | None, modified: dt.datetime | None) -> None:
    """os.utime 回退实现，也是 Linux 的唯一实现。

    APFS/HFS+ 强制「创建时间 ≤ 修改时间」：把 mtime 往前调时 birth 会被一并拉低。
    于是可以先拿 created 当 mtime 写一次（顺带把 birth 拉到 created），再写真正的
    mtime。代价是无法把创建时间调到比文件现有创建时间更晚，Linux 上则完全没有创建时间。
    """
    stat = path.stat()
    atime = stat.st_atime
    if created is not None:
        os.utime(path, (atime, created.timestamp()))
    os.utime(path, (atime, modified.timestamp() if modified is not None else stat.st_mtime))


def _windows_times(path: Path, created: dt.datetime | None, modified: dt.datetime | None) -> None:
    import ctypes

    # Windows FILETIME counts 100-nanosecond intervals since 1601-01-01 UTC.
    class FILETIME(ctypes.Structure):
        _fields_ = [("dwLowDateTime", ctypes.c_ulong), ("dwHighDateTime", ctypes.c_ulong)]

    def as_filetime(value: dt.datetime | None):
        if value is None:
            return None
        ticks = int((value.astimezone(dt.timezone.utc) - dt.datetime(1601, 1, 1, tzinfo=dt.timezone.utc)).total_seconds() * 10_000_000)
        return FILETIME(ticks & 0xFFFFFFFF, ticks >> 32)

    creation, write = as_filetime(created), as_filetime(modified)
    handle = ctypes.windll.kernel32.CreateFileW(str(path), 0x100, 0, None, 3, 0x80, None)
    if handle == ctypes.c_void_p(-1).value:
        raise OSError("无法打开文件以更新文件系统时间。")
    try:
        if not ctypes.windll.kernel32.SetFileTime(handle, ctypes.byref(creation) if creation else None, None, ctypes.byref(write) if write else None):
            raise OSError("更新文件系统时间失败；请确认文件未被其他程序占用。")
    finally:
        ctypes.windll.kernel32.CloseHandle(handle)


def _apply_times(path: Path, created: dt.datetime | None, modified: dt.datetime | None) -> None:
    if created is None and modified is None:
        return
    if sys.platform == "darwin":
        try:
            _setattrlist_times(path, created, modified)
            return
        except OSError:
            # 非 APFS/HFS+ 卷（部分网络卷、exFAT）不支持 ATTR_CMN_CRTIME。
            _utime_times(path, created, modified)
            return
    if os.name == "nt":
        _windows_times(path, created, modified)
        return
    _utime_times(path, created, modified)


def filesystem_values(path: Path) -> dict[str, str]:
    stat = path.stat()
    return {
        "filesystem_created": dt.datetime.fromtimestamp(_birthtime(stat)).strftime(TIME_FORMAT),
        "filesystem_modified": dt.datetime.fromtimestamp(stat.st_mtime).strftime(TIME_FORMAT),
    }


def set_filesystem_times(path: Path, created: str, modified: str) -> None:
    """Set creation and last-write timestamps without touching document data."""
    try:
        created_dt = dt.datetime.strptime(created, TIME_FORMAT) if created else None
        modified_dt = dt.datetime.strptime(modified, TIME_FORMAT) if modified else None
    except ValueError as exc:
        raise ValueError("文件时间格式必须为 YYYY-MM-DD HH:MM:SS。") from exc
    _apply_times(path, created_dt, modified_dt)


def capture_times(path: Path) -> tuple[dt.datetime, dt.datetime]:
    """返回 (创建时间, 修改时间) 快照，用于在写入后原样恢复。

    Office 保存走 os.replace，新文件的创建时间会变成「刚刚」；PDF 保存走
    write_bytes，会刷新修改时间。所以「清除嵌入属性不动文件系统时间」这个
    承诺必须在写入前后显式地存、还，不能指望文件系统自己保持。
    """
    stat = path.stat()
    return (
        dt.datetime.fromtimestamp(_birthtime(stat)).replace(microsecond=0),
        dt.datetime.fromtimestamp(stat.st_mtime).replace(microsecond=0),
    )


def restore_times(path: Path, times: tuple[dt.datetime, dt.datetime]) -> bool:
    """把 capture_times 拿到的快照写回文件；失败返回 False 而不抛异常。"""
    try:
        _apply_times(path, times[0], times[1])
    except OSError:
        return False
    return True


def _text(root: ET.Element | None, xpath: str) -> str:
    node = root.find(xpath, NS) if root is not None else None
    return node.text or "" if node is not None else ""


def _xml_from_zip(zf: zipfile.ZipFile, name: str) -> ET.Element | None:
    try:
        return ET.fromstring(zf.read(name))
    except KeyError:
        return None


def read_office(path: Path) -> Metadata:
    result = Metadata(path, "Office")
    result.values.update(filesystem_values(path))
    with zipfile.ZipFile(path) as zf:
        core, app, custom = (_xml_from_zip(zf, "docProps/core.xml"), _xml_from_zip(zf, "docProps/app.xml"), _xml_from_zip(zf, "docProps/custom.xml"))
    core_map = {"title": "dc:title", "subject": "dc:subject", "creator": "dc:creator", "lastModifiedBy": "cp:lastModifiedBy", "keywords": "cp:keywords", "category": "cp:category", "description": "dc:description", "contentStatus": "cp:contentStatus", "revision": "cp:revision", "created": "dcterms:created", "modified": "dcterms:modified"}
    app_map = {"application": "app:Application", "company": "app:Company", "manager": "app:Manager", "hyperlinkBase": "app:HyperlinkBase", "template": "app:Template", "lastPrinted": "app:LastPrinted", "totalTime": "app:TotalTime"}
    for key, xp in core_map.items(): result.values[key] = iso_display(_text(core, xp)) if key in {"created", "modified"} else _text(core, xp)
    for key, xp in app_map.items(): result.values[key] = iso_display(_text(app, xp)) if key == "lastPrinted" else _text(app, xp)
    if custom is not None:
        for prop in custom.findall(f"{{{CP}}}property"):
            value = next(iter(prop), None)
            result.custom[prop.get("name", "未命名属性")] = value.text or "" if value is not None else ""
    return result


def _set(root: ET.Element, tag: str, value: str) -> None:
    node = root.find(tag, NS)
    if node is None:
        node = ET.SubElement(root, tag.replace("cp:", f"{{{CORE}}}").replace("dc:", f"{{{DC}}}").replace("dcterms:", f"{{{DCTERMS}}}").replace("app:", f"{{{APP}}}"))
    node.text = value


def write_office(meta: Metadata, new_values: dict[str, str], custom: dict[str, str], clear: bool) -> None:
    with zipfile.ZipFile(meta.path) as original:
        core, app, custom_root = (_xml_from_zip(original, "docProps/core.xml"), _xml_from_zip(original, "docProps/app.xml"), _xml_from_zip(original, "docProps/custom.xml"))
        if core is None: raise ValueError("此文件缺少 Office 核心属性包，无法安全写入。")
        if app is None: app = ET.Element(f"{{{APP}}}Properties")
        if custom_root is None: custom_root = ET.Element(f"{{{CP}}}Properties")
        core_map = {"title": "dc:title", "subject": "dc:subject", "creator": "dc:creator", "lastModifiedBy": "cp:lastModifiedBy", "keywords": "cp:keywords", "category": "cp:category", "description": "dc:description", "contentStatus": "cp:contentStatus", "revision": "cp:revision", "created": "dcterms:created", "modified": "dcterms:modified"}
        app_map = {"application": "app:Application", "company": "app:Company", "manager": "app:Manager", "hyperlinkBase": "app:HyperlinkBase", "template": "app:Template", "lastPrinted": "app:LastPrinted", "totalTime": "app:TotalTime"}
        for key, tag in core_map.items():
            value = "" if clear else new_values.get(key, "")
            _set(core, tag, iso_store(value) if key in {"created", "modified"} else value)
        for key, tag in app_map.items():
            value = "" if clear else new_values.get(key, "")
            _set(app, tag, iso_store(value) if key == "lastPrinted" else value)
        for child in list(custom_root): custom_root.remove(child)
        if not clear:
            for pid, (name, value) in enumerate(custom.items(), start=2):
                prop = ET.SubElement(custom_root, f"{{{CP}}}property", {"fmtid": "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}", "pid": str(pid), "name": name})
                ET.SubElement(prop, f"{{{VT}}}lpwstr").text = value
        # Read entries before closing the source archive. Windows won't replace a
        # path while its ZipFile handle is open.
        entries = [(item, original.read(item.filename)) for item in original.infolist()
                   if item.filename not in {"docProps/core.xml", "docProps/app.xml", "docProps/custom.xml"}]
        had_custom = "docProps/custom.xml" in original.namelist()
    descriptor, tmp_name = tempfile.mkstemp(suffix=meta.path.suffix, dir=meta.path.parent)
    os.close(descriptor)
    tmp = Path(tmp_name)
    try:
        with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as out:
            for item, contents in entries: out.writestr(item, contents)
            out.writestr("docProps/core.xml", ET.tostring(core, encoding="utf-8", xml_declaration=True))
            out.writestr("docProps/app.xml", ET.tostring(app, encoding="utf-8", xml_declaration=True))
            if custom or had_custom: out.writestr("docProps/custom.xml", ET.tostring(custom_root, encoding="utf-8", xml_declaration=True))
        os.replace(tmp, meta.path)
    finally:
        if tmp.exists(): tmp.unlink(missing_ok=True)


def _pdf_unescape(value: bytes) -> str:
    return re.sub(rb"\\([\\()])", rb"\1", value).decode("utf-8", "replace")


def _pdf_escape(value: str) -> str:
    return value.replace("\\", "\\\\").replace("(", "\\(").replace(")", "\\)")


# 字符串字面量体：先吃掉「反斜杠 + 任意一个字节」的转义对，再吃掉普通字节。
# 原写法 (.*?) 遇到值内含转义括号（例如标题写作 "报告 (v2)"，转义后是 "报告 \(v2\)"）
# 会在第一个 ")" 处提前截断，只读回 "报告 (v2" 的前半段。
_PDF_LITERAL = rb"((?:\\.|[^)\\])*)"


def read_pdf(path: Path) -> Metadata:
    data = path.read_bytes()
    if b"/Encrypt" in data: raise ValueError("加密 PDF 不支持读取或写入。")
    result = Metadata(path, "PDF")
    result.values.update(filesystem_values(path))
    mapping = {"title": b"Title", "author": b"Author", "subject": b"Subject", "keywords": b"Keywords", "creator": b"Creator", "producer": b"Producer", "created": b"CreationDate", "modified": b"ModDate"}
    for key, pdf_key in mapping.items():
        matches = list(re.finditer(rb"/" + pdf_key + rb"\s*\(" + _PDF_LITERAL + rb"\)", data, re.S))
        value = _pdf_unescape(matches[-1].group(1)) if matches else ""
        result.values[key] = value.replace("D:", "") if key in {"created", "modified"} else value
    return result


def _pdf_date(value: str) -> str:
    digits = re.sub(r"\D", "", value)
    return "D:" + (digits[:14] if digits else dt.datetime.now().strftime("%Y%m%d%H%M%S"))


def write_pdf(meta: Metadata, values: dict[str, str], clear: bool) -> None:
    data = meta.path.read_bytes()
    if b"/Encrypt" in data: raise ValueError("加密 PDF 不支持写入。")
    starts = list(re.finditer(rb"startxref\s*(\d+)", data))
    trailers = list(re.finditer(rb"trailer\s*<<(.*?)>>", data, re.S))
    if not starts or not trailers: raise ValueError("此 PDF 结构不支持安全的增量属性写入。")
    trailer, previous = trailers[-1].group(1), int(starts[-1].group(1))
    root = re.search(rb"/Root\s+(\d+\s+\d+\s+R)", trailer)
    size = re.search(rb"/Size\s+(\d+)", trailer)
    if not root or not size: raise ValueError("未找到 PDF 根对象，无法写入。")
    obj = int(size.group(1))
    pairs = [("Title", "title"), ("Author", "author"), ("Subject", "subject"), ("Keywords", "keywords"), ("Creator", "creator"), ("Producer", "producer"), ("CreationDate", "created"), ("ModDate", "modified")]
    body = [f"{obj} 0 obj", "<<"]
    for pdf_key, key in pairs:
        value = "" if clear else values.get(key, "")
        if value:
            value = _pdf_date(value) if key in {"created", "modified"} else value
            body.append(f"/{pdf_key} ({_pdf_escape(value)})")
    body += [">>", "endobj"]
    appendix = ("\n".join(body) + "\n").encode("utf-8")
    offset = len(data) + len(appendix)
    xref = f"xref\n{obj} 1\n{len(data):010d} 00000 n \ntrailer\n<< /Size {obj + 1} /Root {root.group(1).decode()} /Info {obj} 0 R /Prev {previous} >>\nstartxref\n{offset}\n%%EOF\n".encode()
    meta.path.write_bytes(data + appendix + xref)


def load(path: str) -> Metadata:
    item = Path(path)
    suffix = item.suffix.lower()
    if suffix in OFFICE_SUFFIXES: return read_office(item)
    if suffix == ".pdf": return read_pdf(item)
    if suffix in LEGACY_SUFFIXES: raise ValueError("旧版 Office 二进制文件暂不支持。请在 Word、Excel、Pages、Numbers 或 WPS 中另存为 DOCX/XLSX/PPTX 后再编辑。")
    raise ValueError("仅支持新版 Office 文档和 PDF。")


def backup(path: Path) -> Path:
    # Never overwrite an earlier backup: network-sync folders and document apps
    # often keep the old .metadata-backup file open or read-only.
    stamp = dt.datetime.now().strftime("%Y%m%d-%H%M%S-%f")
    target = path.with_name(f"{path.name}.{stamp}.metadata-backup")
    try:
        shutil.copy2(path, target)
    except PermissionError as exc:
        raise PermissionError(
            "无法在当前文件夹创建备份。请确认该文件夹可写、文件未被 iCloud 同步或 Office/WPS/Pages 占用，"
            "或者在界面中将“保存前备份”设为“否”。"
        ) from exc
    return target
