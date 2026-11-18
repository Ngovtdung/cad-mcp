"""Kiểm tệp .xls fixture không chứa dữ liệu cá nhân. In CHỈ số đếm, không bao giờ in tên.
  python -I -B tests\\verify-xls-fixtures.py <thư mục fixtures-xls>
Kiểm: SummaryInformation PID 4 (Author) và PID 8 (LastAuthor) rỗng hoặc "test"; tìm byte (ASCII + UTF-16LE) tên Office hiện tại,
tên người dùng Windows (%USERNAME%). Mã thoát = tổng số lần trùng + số PID sai."""
import os
import struct
import sys

sys.dont_write_bytecode = True  # không ghi vendor\xlrd\__pycache__: .pyc sẽ bị nạp thay cho .py đã kiểm hash
sys.path.append(os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "vendor"))
import xlrd  # noqa: E402
from xlrd import compdoc  # noqa: E402


def office_user():
    try:
        import winreg
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"Software\Microsoft\Office\Common\UserInfo") as k:
            return winreg.QueryValueEx(k, "UserName")[0]
    except OSError:
        pass
    try:
        import winreg
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"Software\Microsoft\Office\16.0\Common\UserInfo") as k:
            return winreg.QueryValueEx(k, "UserName")[0]
    except OSError:
        return ""


def props(stream):
    """{pid: str} từ property set (chỉ LPSTR/LPWSTR)."""
    out = {}
    nsec = struct.unpack_from("<I", stream, 24)[0]
    for s in range(nsec):
        off = struct.unpack_from("<I", stream, 28 + s * 20 + 16)[0]
        nprops = struct.unpack_from("<I", stream, off + 4)[0]
        for i in range(nprops):
            pid, po = struct.unpack_from("<II", stream, off + 8 + i * 8)
            t = struct.unpack_from("<I", stream, off + po)[0] & 0xFFFF
            if t == 0x1E:
                n = struct.unpack_from("<I", stream, off + po + 4)[0]
                out[pid] = stream[off + po + 8:off + po + 8 + n].split(b"\0")[0].decode("latin-1")
            elif t == 0x1F:
                n = struct.unpack_from("<I", stream, off + po + 4)[0]
                out[pid] = stream[off + po + 8:off + po + 8 + n * 2].decode("utf-16-le").split("\0")[0]
    return out


names = [n for n in {office_user(), os.environ.get("USERNAME", "")} if n]
bad = 0
for fn in sorted(os.listdir(sys.argv[1])):
    if not fn.lower().endswith(".xls"):
        continue
    data = open(os.path.join(sys.argv[1], fn), "rb").read()
    hits = 0
    for n in names:
        for enc in ("ascii", "utf-16-le"):
            hits += data.lower().count(n.lower().encode(enc, "ignore")) if n.isascii() or enc != "ascii" else 0
    pid_bad = []
    try:
        doc = compdoc.CompDoc(data, logfile=open(os.devnull, "w"))
        ss = doc.get_named_stream("\x05SummaryInformation")
        p = props(ss) if ss else {}
        for pid in (4, 8):
            v = p.get(pid, "")
            if v not in ("", "test"):
                pid_bad.append(pid)
    except Exception as e:
        print(fn, "SummaryInformation: khong doc duoc (%s)" % type(e).__name__)
    bad += hits + len(pid_bad)
    print("%s: byte-trung=%d pid48-sai=%d" % (fn, hits, len(pid_bad)))
print("TONG vi-pham=%d" % bad)
sys.exit(1 if bad else 0)
