"""Đọc .xls (Excel 97-2003) từ stdin, in TSV theo sheet ra stdout (UTF-8) cho palette Claude.
Gọi (Attachments.cs RunXls): <python.exe tuyệt đối> -I -S -B -X utf8 xls2tsv.py   (stdin = byte tệp .xls; không nhận đường dẫn)
Dòng đầu stdout: "XLS2TSV 1 <số sheet> <số sheet ẩn đã gửi> <0|1 đã cắt>". Mã thoát: 0 ok, 2 không đọc được, 3 watchdog.
Bố cục giống Attachments.cs (Xlsx.Sheet). Spec: docs\\spec-attachments.md §2.2"""
import os
import sys
import threading

_wd = threading.Timer(20, os._exit, (3,))  # tự huỷ nếu treo (palette còn kill ở 15 s)
_wd.daemon = True
_wd.start()
# -S: không có site-packages. append (không insert(0)): vendor\ không được che module chuẩn của stdlib
sys.path.append(os.path.join(os.path.dirname(os.path.abspath(__file__)), "vendor"))
import datetime  # noqa: E402
import xlrd  # noqa: E402

MAX_BYTES = 10 * 1024 * 1024
MAX_SHEETS = 50
MAX_COLS = 100
CELL_MAX = 1000
TEXT_MAX = 100000


def col_name(n):
    s = ""
    while n > 0:
        n -= 1
        s = chr(65 + n % 26) + s
        n //= 26
    return s


def clean(s):
    s = s.replace("\r\n", "\\n").replace("\n", "\\n").replace("\r", "\\n").replace("\t", " ")
    return s[:CELL_MAX] + "…" if len(s) > CELL_MAX else s


def num(v):
    return format(v, ".15g")  # 0.30000000000000004 -> 0.3; 42.0 -> 42


def date(v, mode):
    try:
        d = xlrd.xldate.xldate_as_datetime(v, mode)
    except Exception:
        return num(v)
    d = (d + datetime.timedelta(microseconds=500000)).replace(microsecond=0)
    if v < 1:
        return d.strftime("%H:%M:%S")
    return d.strftime("%Y-%m-%d") if d.time() == datetime.time(0) else d.strftime("%Y-%m-%d %H:%M:%S")


def cell(sh, r, c, mode):
    t = sh.cell_type(r, c)
    v = sh.cell_value(r, c)
    if t == xlrd.XL_CELL_TEXT:
        return clean(v)
    if t == xlrd.XL_CELL_NUMBER:
        return num(v)
    if t == xlrd.XL_CELL_DATE:
        return date(v, mode)
    if t == xlrd.XL_CELL_BOOLEAN:
        return "TRUE" if v else "FALSE"
    if t == xlrd.XL_CELL_ERROR:
        return xlrd.error_text_from_code.get(v, "#ERR")
    return ""


def main():
    data = sys.stdin.buffer.read(MAX_BYTES + 4)
    # .NET Framework: Process.StandardInput tự ghi BOM UTF-8 vào đầu luồng khi console có codepage 65001 (BOM không thể là đầu của tệp OLE)
    if data[:3] == b"\xef\xbb\xbf":
        data = data[3:]
    if len(data) > MAX_BYTES:
        sys.stderr.write("too big\n")
        return 2
    try:
        # logfile=devnull: xlrd mặc định in cảnh báo ra sys.stdout, sẽ làm bẩn kết quả
        # ragged_rows=True: mỗi dòng chỉ cấp phát tới ô cuối có dữ liệu (không đệm cả ncols)
        book = xlrd.open_workbook(file_contents=data, logfile=open(os.devnull, "w"), on_demand=True,
                                  formatting_info=False, ragged_rows=True)
    except Exception as e:  # hỏng, có mật khẩu, không phải BIFF
        sys.stderr.write("%s: %s\n" % (type(e).__name__, str(e)[:200]))
        return 2
    names = book.sheet_names()
    out = []   # các khối sheet đã xong
    done = 0   # số ký tự của "\n\n".join(out)
    hidden = 0
    cut = len(names) > MAX_SHEETS
    for i, name in enumerate(names[:MAX_SHEETS]):
        if out:
            done += 2
        try:
            sh = book.sheet_by_index(i)
        except Exception:
            block = "## Sheet: %s (biểu đồ hoặc không đọc được, bỏ qua)" % name
            out.append(block)
            done += len(block)
            continue
        tag = " (ẩn)" if sh.visibility else ""
        hidden += 1 if sh.visibility else 0
        lines, used, max_col, max_row, stop = [], 0, 0, 0, 0
        for r in range(sh.nrows):
            cells = [cell(sh, r, c, book.datemode) for c in range(min(sh.row_len(r), MAX_COLS))]
            while cells and cells[-1] == "":
                cells.pop()
            if not cells:
                continue
            line = "%d\t%s" % (r + 1, "\t".join(cells))
            if done + used + len(line) + 600 > TEXT_MAX:
                stop = r + 1
                break
            lines.append(line)
            used += len(line) + 1
            max_col = max(max_col, len(cells))
            max_row = r + 1
        part = ["## Sheet: " + name + tag + (" (A1:%s%d)" % (col_name(max_col), max_row) if lines else " (trống)")]
        if sh.ncols > MAX_COLS:
            part.append("[chỉ gửi %d cột đầu]" % MAX_COLS)
        if lines:
            part.append("".join("\t" + col_name(k) for k in range(1, max_col + 1)))
        part += lines
        book.unload_sheet(i)
        if stop:
            part.append("[... cắt: dừng ở dòng %d của sheet %s]" % (stop, name))
            rest = names[i + 1:]
            if rest:
                part.append("[... các sheet sau không gửi: " + ", ".join(rest[:20]) + (", …" if len(rest) > 20 else "") + "]")
            out.append("\n".join(part))
            cut = True
            break
        block = "\n".join(part)
        out.append(block)
        done += len(block)
    body = "XLS2TSV 1 %d %d %d\n" % (len(names), hidden, 1 if cut else 0) + "\n\n".join(out)
    sys.stdout.buffer.write(body.encode("utf-8"))  # byte: không đổi \n thành CRLF
    sys.stdout.buffer.flush()
    return 0


if __name__ == "__main__":
    sys.exit(main())
