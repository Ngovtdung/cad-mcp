"""spec-3.2-tools: gói 1 (mốc undo, undo/redo, mirror, offset, array, ellipse, block, save guard).
Không cần AutoCAD: COM giả. Chạy: python -I -B tests\\test_tools32.py   (mã thoát = số ca sai)"""
import importlib.util
import math
import os
import sys
import tempfile
import types
from unittest import mock

sys.stdout.reconfigure(encoding="utf-8", errors="replace")
SERVER = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "server.py")
n = fail = 0


def check(label, ok, detail=""):
    global n, fail
    n += 1
    if not ok:
        fail += 1
        print("FAIL", label, detail)


def raised(f, text):
    try:
        f()
    except Exception as e:
        return text in str(e)
    return False


class ComError(Exception):
    def __init__(self, hresult=-1):
        super().__init__(hresult)
        self.hresult = hresult


def load():
    for m in ("pythoncom", "win32com", "win32com.client", "win32gui", "win32ui", "PIL", "PIL.Image",
              "mcp", "mcp.server", "mcp.server.mcpserver"):
        sys.modules[m] = mock.MagicMock()
    sys.modules["win32com"].client = sys.modules["win32com.client"]
    sys.modules["pythoncom"].com_error = ComError
    sys.modules["win32com.client"].DispatchBaseClass = type("DispatchBaseClass", (), {})
    sys.modules["win32com.client"].VARIANT = lambda t, v: v  # _pt -> tuple (x, y, z): so sánh được
    sys.modules["mcp.server.mcpserver"].MCPServer.return_value.tool.return_value = lambda f: f  # tool vẫn là hàm
    sys.modules["mcp.server.mcpserver.exceptions"] = types.SimpleNamespace(ToolError=type("ToolError", (Exception,), {}))
    for k in ("CAD_CLAUDE_ACAD_PID", "CAD_CLAUDE_PIN_FILE"):
        os.environ.pop(k, None)
    spec = importlib.util.spec_from_file_location("server_under_test", SERVER)
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod


class Ent:
    def __init__(self, doc, kind="Line", args=()):
        doc.seq += 1
        self.doc, self.ObjectName, self.args, self.Layer = doc, kind, args, "0"
        self.Handle, self._o, self.ops = "H%d" % doc.seq, self, []
        doc.ents[self.Handle] = self

    def _op(self, name, *a):
        d = self.doc
        d.log.append(name)
        d.calls[name] = d.calls.get(name, 0) + 1
        if d.fail == (name, d.calls[name]):
            raise ComError(-2145386390)
        self.ops.append((name,) + a)

    def Copy(self):
        self._op("Copy")
        return Ent(self.doc, self.ObjectName)

    def Move(self, a, b):
        self._op("Move", a, b)

    def Rotate(self, c, ang):
        self._op("Rotate", c, ang)

    def Mirror(self, a, b):
        self._op("Mirror", a, b)
        return Ent(self.doc, self.ObjectName)

    def Offset(self, dist):
        self._op("Offset", dist)
        return (Ent(self.doc), Ent(self.doc))

    def Delete(self):
        self._op("Delete")
        self.doc.ents.pop(self.Handle, None)


class MS:
    def __init__(self, d):
        self.d = d

    def __getattr__(self, name):  # AddLine, AddCircle, AddEllipse, AddLightWeightPolyline...
        if not name.startswith("Add"):
            raise AttributeError(name)

        def add(*a):
            self.d.log.append(name)
            return Ent(self.d, name[3:], a)
        return add


class Blk:
    def __init__(self, bl, name, origin=(0.0, 0.0, 0.0), count=0):
        self.bl, self.Name, self.Origin, self.items, self._o = bl, name, origin, [None] * count, self

    @property
    def Count(self):
        return len(self.items)

    def Delete(self):
        self.bl.lst.remove(self)


class Blocks:
    def __init__(self):
        self.lst, self.added = [], 0
        for nm, c in (("*Model_Space", 9), ("*Paper_Space", 0), ("*U1", 2), ("DOOR", 3)):
            self.lst.append(Blk(self, nm, count=c))

    @property
    def Count(self):
        return len(self.lst)

    def Item(self, i):
        if isinstance(i, int):
            return self.lst[i]
        for b in self.lst:
            if b.Name.lower() == i.lower():
                return b
        raise ComError(-2145386476)

    def Add(self, pt, name):
        self.added += 1
        b = Blk(self, name, pt)
        self.lst.append(b)
        return b


class Doc:
    def __init__(self, full=r"C:\w\a.dwg", name="a.dwg"):
        self.FullName, self.Name, self._o = full, name, self
        self.seq, self.ents, self.log, self.sent, self.calls, self.fail = 0, {}, [], [], {}, None
        self.cmd, self.undoctl, self.mark_error, self.copy_error = 0, 5, False, False
        self.ModelSpace, self.Blocks = MS(self), Blocks()
        self.Utility = types.SimpleNamespace(Prompt=lambda t: None)

    def GetVariable(self, name):
        return {"CMDACTIVE": self.cmd, "UNDOCTL": self.undoctl}.get(name, 0)

    def StartUndoMark(self):
        if self.mark_error:
            raise ComError(-2147352567)
        self.log.append("start")

    def EndUndoMark(self):
        self.log.append("end")
        if getattr(self, "end_error", False):
            raise ComError(-2147352567)

    def HandleToObject(self, h):
        if h not in self.ents:
            raise ComError(-2145386484)
        return self.ents[h]

    def SendCommand(self, s):
        self.sent.append(s)

    def CopyObjects(self, objs, owner):
        self.log.append("CopyObjects")
        if self.copy_error:
            raise ComError(-2145386390)
        owner.items += list(objs)

    def Save(self):
        self.log.append("Save")

    def SaveAs(self, p):
        self.log.append("SaveAs")
        self.FullName, self.Name = p, os.path.basename(p)


s = load()


def fresh(**kw):
    d = Doc(**kw)
    s._doc = lambda: s._DocR(d)
    return d


# --- Task 1: mốc undo ---
d = fresh()
s.line(0, 0, 10, 0)
check("T1 line: start, AddLine, end", d.log == ["start", "AddLine", "end"], d.log)
d = fresh()
s.rectangle(0, 0, 10, 5)
check("T1 rectangle->polyline lồng: đúng 1 cặp mốc", d.log == ["start", "AddLightWeightPolyline", "end"], d.log)
d = fresh()
s.run_python("line(0, 0, 1, 1)\nline(0, 0, 2, 2)")
check("T1 Q2 run_python: tool con KHÔNG đặt mốc (không có mốc nào, không nhóm mở dở khi bị Dừng)",
      d.log == ["AddLine", "AddLine"], d.log)
d = fresh()
s.run_python("pass")
check("T1 Q2 run_python tự nó không đặt mốc", d.log == [], d.log)
d = fresh()
s.run_python("ms.AddLine(1, 2)")
check("T1 Q2 run_python gọi COM trực tiếp: không mốc (không để nhóm mở dở khi bị Dừng)", d.log == ["AddLine"], d.log)
d = fresh()
d.cmd = 1
s.line(0, 0, 1, 1)
check("T1 CMDACTIVE≠0: không mốc, vẫn vẽ", d.log == ["AddLine"], d.log)
d = fresh()
d.undoctl = 5 | 8
s.line(0, 0, 1, 1)
check("T1 UNDOCTL có nhóm đang mở: không mốc", d.log == ["AddLine"], d.log)
d = fresh()
d.undoctl = 0
s.line(0, 0, 1, 1)
check("T1 UNDO tắt: không mốc", d.log == ["AddLine"], d.log)
d = fresh()
d.mark_error = True
s.line(0, 0, 1, 1)
check("T1 StartUndoMark lỗi COM: tool vẫn chạy, không End", d.log == ["AddLine"], d.log)
d = fresh()
check("T1 tool lỗi: vẫn EndUndoMark", raised(lambda: s.delete("NOPE"), "ComError") and d.log == ["start", "end"], d.log)
check("T1 sau lỗi: _undo_depth về 0", s._undo_depth == 0, s._undo_depth)
d = fresh()
s.message("x")
check("T1 tool không sửa hình (message): không mốc", d.log == [], d.log)

# --- Task 2: undo / redo ---
s._is_active = lambda d: True
d = fresh()
r = s.undo(3)
check("T2 undo(3): gửi 3 lệnh U", d.sent == ["_.U\n_.U\n_.U\n"] and "3" in r, (d.sent, r))
d = fresh()
s.undo()
check("T2 undo(): mặc định 1", d.sent == ["_.U\n"], d.sent)
for bad in (0, 21, -1):
    d = fresh()
    check(f"T2 undo({bad}): từ chối, không gửi", raised(lambda: s.undo(bad), "steps") and d.sent == [], d.sent)
d = fresh()
s.redo()
check("T2 redo: gửi REDO", d.sent == ["_.REDO\n"], d.sent)
d = fresh()
d.cmd = 1
check("T2 undo khi user dở lệnh: từ chối", raised(lambda: s.undo(), "đang chạy dở") and d.sent == [], d.sent)
s._is_active = lambda d: False
d = fresh()
check("T2 undo trên bản vẽ nền: từ chối", raised(lambda: s.undo(), "không phải tab đang mở") and d.sent == [], d.sent)
check("T2 redo trên bản vẽ nền: từ chối", raised(lambda: s.redo(), "không phải tab đang mở") and d.sent == [], d.sent)
check("T2 undo/redo không đặt mốc", "start" not in d.log, d.log)

# --- Task 3: mirror / offset ---
d = fresh()
h = s.line(0, 0, 10, 0)
d.log.clear()
m = s.mirror(h, 0, 5, 10, 5)
src = d.ents[h]
check("T3 mirror: trả handle mới, giữ bản gốc", m != h and m in d.ents and h in d.ents, (m, list(d.ents)))
check("T3 mirror: trục đúng điểm", src.ops[-1] == ("Mirror", (0.0, 5.0, 0.0), (10.0, 5.0, 0.0)), src.ops)
check("T3 mirror: có mốc undo", d.log[0] == "start" and d.log[-1] == "end", d.log)
d.log.clear()
m2 = s.mirror(h, 0, 5, 10, 5, keep_original=False)
check("T3 mirror keep_original=False: xoá gốc SAU khi tạo bản mới",
      h not in d.ents and m2 in d.ents and d.log.index("Mirror") < d.log.index("Delete"), d.log)
d = fresh()
h = s.line(0, 0, 10, 0)
for args in ((1, 1, 1, 1), (1, 1, 1 + 1e-7, 1), (0, 0, float("nan"), 1), (0, 0, float("inf"), 1)):
    check(f"T3 mirror trục suy biến {args}: từ chối", raised(lambda: s.mirror(h, *args), "suy biến"))
check("T3 mirror suy biến: không gọi Mirror", "Mirror" not in d.log, d.log)
d = fresh()
h = s.circle(0, 0, 5)
out = s.offset(h, 2.5)
check("T3 offset: trả mọi handle mới", len(out) == 2 and all(x in d.ents for x in out) and h in d.ents, out)
check("T3 offset: đúng khoảng cách", d.ents[h].ops[-1] == ("Offset", 2.5), d.ents[h].ops)
for bad in (0, 1e-9, -1e-9, float("nan"), float("inf")):
    check(f"T3 offset({bad}): từ chối", raised(lambda: s.offset(h, bad), "distance"))
check("T3 offset từ chối: không gọi Offset", d.ents[h].ops == [("Offset", 2.5)], d.ents[h].ops)

# --- Task 4: array_rect / array_polar ---
d = fresh()
h = s.line(0, 0, 1, 0)
out = s.array_rect(h, 2, 3, 10, 20)
moves = sorted(d.ents[x].ops[0][2] for x in out)
check("T4 rect 2x3: 5 bản sao đúng vị trí WCS", len(out) == 5 and moves == sorted(
    [(20.0, 0.0, 0.0), (40.0, 0.0, 0.0), (0.0, 10.0, 0.0), (20.0, 10.0, 0.0), (40.0, 10.0, 0.0)]), moves)
check("T4 rect: gốc còn nguyên", h in d.ents and d.ents[h].ops == [("Copy",)] * 5, d.ents[h].ops)
check("T4 rect: 1 cặp mốc", d.log.count("start") == 2 and d.log[-1] == "end", d.log)  # 1 của line + 1 của array
out = s.array_rect(h, 1, 3, 0, -5)
check("T4 rect 1x3 spacing âm, row_spacing=0 được", sorted(d.ents[x].ops[0][2] for x in out) == [(-10.0, 0.0, 0.0), (-5.0, 0.0, 0.0)])
for args, txt in (((1, 1, 1, 1), "rows"), ((0, 5, 1, 1), "rows"), ((2, -1, 1, 1), "rows"), ((1001, 1, 1, 1), "1000"),
                  ((40, 26, 1, 1), "1000"), ((2, 1, 0, 1), "row_spacing"), ((1, 2, 1, 0), "col_spacing"),
                  ((2, 2, float("nan"), 1), "hữu hạn")):
    before = len(d.ents)
    check(f"T4 rect {args}: từ chối, không tạo gì", raised(lambda: s.array_rect(h, *args), txt) and len(d.ents) == before)
d = fresh()
h = s.line(0, 0, 1, 0)
d.fail = ("Move", 3)
check("T4 rect lỗi giữa chừng: báo lỗi", raised(lambda: s.array_rect(h, 1, 5, 0, 10), "ComError"))
check("T4 rect lỗi giữa chừng: xoá sạch bản sao", set(d.ents) == {h}, list(d.ents))
check("T4 rect lỗi: vẫn đóng mốc", d.log[-1] == "end", d.log)
d = fresh()
h = s.line(0, 0, 1, 0)
out = s.array_polar(h, 4, 5, 5)
angs = [d.ents[x].ops[0][2] for x in out]
check("T4 polar 4/360: 3 bản xoay 90/180/270 quanh (5,5)",
      len(out) == 3 and all(math.isclose(a, math.radians(b)) for a, b in zip(angs, (90, 180, 270)))
      and all(d.ents[x].ops[0][1] == (5.0, 5.0, 0.0) for x in out), (angs, out))
out = s.array_polar(h, 3, 0, 0, 180)
check("T4 polar 3/180: bước 90", [round(math.degrees(d.ents[x].ops[0][2]), 9) for x in out] == [90, 180])
out = s.array_polar(h, 4, 0, 0, -360)
check("T4 polar 4/-360: bước -90", [round(math.degrees(d.ents[x].ops[0][2]), 9) for x in out] == [-90, -180, -270])
for args, txt in (((1, 0, 0), "count"), ((1001, 0, 0), "count"), ((4, 0, 0, 0), "angle_deg"), ((4, 0, 0, 361), "angle_deg"),
                  ((4, float("inf"), 0), "hữu hạn")):
    before = len(d.ents)
    check(f"T4 polar {args}: từ chối", raised(lambda: s.array_polar(h, *args), txt) and len(d.ents) == before)
d.fail = ("Rotate", d.calls.get("Rotate", 0) + 2)
before = set(d.ents)
check("T4 polar lỗi giữa chừng: báo lỗi + dọn sạch", raised(lambda: s.array_polar(h, 6, 0, 0), "ComError") and set(d.ents) == before)

# --- Task 5: ellipse ---
d = fresh()
h = s.ellipse(1, 2, 10, 5)
c, ax, ratio = d.ents[h].args
check("T5 ellipse rx>ry: trục lớn theo X, ratio 0.5", c == (1.0, 2.0, 0.0) and ax == (10.0, 0.0, 0.0) and ratio == 0.5, d.ents[h].args)
h = s.ellipse(0, 0, 5, 10)
c, ax, ratio = d.ents[h].args
check("T5 ellipse ry>rx: trục lớn theo Y", math.isclose(ax[0], 0, abs_tol=1e-12) and math.isclose(ax[1], 10) and ratio == 0.5, ax)
h = s.ellipse(0, 0, 10, 5, 30, layer="L1")
c, ax, ratio = d.ents[h].args
check("T5 ellipse xoay 30 độ + layer", math.isclose(ax[0], 10 * math.cos(math.radians(30)))
      and math.isclose(ax[1], 5.0) and d.ents[h].Layer == "L1", (ax, d.ents[h].Layer))
h = s.ellipse(0, 0, 4, 4)
check("T5 ellipse rx=ry: ratio 1", d.ents[h].args[2] == 1.0)
for args in ((0, 0, 0, 5), (0, 0, 5, -1), (0, 0, float("nan"), 5), (float("inf"), 0, 5, 5)):
    check(f"T5 ellipse {args}: từ chối", raised(lambda: s.ellipse(*args), "rx"))

# --- Task 6: create_block / list_blocks ---
d = fresh()
check("T6 list_blocks: bỏ tên '*'", s.list_blocks() == [{"name": "DOOR", "count": 3}], s.list_blocks())
check("T6 list_blocks limit=0: kẹp về 1 (review NOTE)", s.list_blocks(0) == [{"name": "DOOR", "count": 3}])
check("T6 list_blocks: không mốc", "start" not in d.log, d.log)
h1, h2 = s.line(0, 0, 1, 0), s.circle(5, 5, 1)
for bad in ("", "   ", "A<B", "a/b", "x*", "*U9", 'q"', "a" * 256, "bell\x07"):
    check(f"T6 tên xấu {bad[:12]!r}: từ chối", raised(lambda: s.create_block(bad, [h1]), "tên block"))
check("T6 tên trùng (khác hoa thường): từ chối", raised(lambda: s.create_block("door", [h1]), "đã tồn tại"))
check("T6 handles rỗng: từ chối", raised(lambda: s.create_block("WIN", []), "handle"))
check("T6 handle sai: từ chối", raised(lambda: s.create_block("WIN", [h1, "NOPE"]), "ComError"))
check("T6 các ca từ chối: chưa Blocks.Add", d.Blocks.added == 0, d.Blocks.added)
r = s.create_block(" WIN ", [h1, h2], 5, 5)
b = d.Blocks.Item("WIN")
check("T6 tạo block: trả tên đã strip, gốc (5,5)", r == "WIN" and b.Origin == (5.0, 5.0, 0.0), (r, b.Origin))
check("T6 tạo block: chép đúng 2 đối tượng, gốc còn", b.items == [d.ents[h1], d.ents[h2]] and h1 in d.ents and h2 in d.ents)
check("T6 tạo block: Việt có dấu được", s.create_block("Cửa sổ 1", [h1]) == "Cửa sổ 1")
d.copy_error = True
check("T6 CopyObjects lỗi: báo lỗi", raised(lambda: s.create_block("WIN2", [h1]), "ComError"))
check("T6 CopyObjects lỗi: xoá định nghĩa rỗng", raised(lambda: d.Blocks.Item("WIN2"), "") and
      all(x.Name != "WIN2" for x in d.Blocks.lst))

# --- Task 7: save ---
tmp = tempfile.mkdtemp()
other = os.path.join(tmp, "other.dwg")
open(other, "w").close()
d = fresh(full="", name="Drawing1.dwg")
check("T7 chưa từng lưu + không path: từ chối", raised(lambda: s.save(), "chưa từng được lưu") and "Save" not in d.log)
d = fresh()
check("T7 đã có tên + không path: Save tại chỗ", s.save() == r"C:\w\a.dwg" and "Save" in d.log, d.log)
for bad, txt in (("x.dwg", "tuyệt đối"), (r"\x.dwg", "tuyệt đối"), ("C:x.dwg", "tuyệt đối"), (os.path.join(tmp, "x.dxf"), ".dwg"),
                 (r"\\srv\share\x.dwg", "mạng"), ("//srv/share/x.dwg", "mạng"), (r"\\?\C:\x.dwg", "mạng")):
    check(f"T7 path {bad}: từ chối", raised(lambda: s.save(bad), txt) and "SaveAs" not in d.log, d.log)
check("T7 tệp khác đã tồn tại: từ chối", raised(lambda: s.save(other), "đã tồn tại") and "SaveAs" not in d.log)
check("T7 overwrite=True: cho ghi", s.save(other, overwrite=True) == other and "SaveAs" in d.log)
own = os.path.join(tmp, "own.dwg")
open(own, "w").close()
d = fresh(full=own, name="own.dwg")
check("T7 SaveAs đè chính tệp của bản vẽ (khác hoa thường): không cần cờ", s.save(own.upper()) and "SaveAs" in d.log)
d = fresh(full="", name="Drawing1.dwg")
s._PIN = "Drawing1.dwg"
new = os.path.join(tmp, "new.dwg")
s.save(new)
check("T7 sau SaveAs: ghim theo tên mới", s._PIN == new, s._PIN)
check("T7 save: không mốc", "start" not in d.log, d.log)


# --- review C2 / N-a / N-b ---
d = fresh()
s.run_python("for i in range(3):\n    line(0, 0, i + 1, 1)")
check("C2 run_python vòng 3 line: không mốc nào, depth về 0", d.log == ["AddLine"] * 3 and s._undo_depth == 0, (d.log, s._undo_depth))
d = fresh()
s.run_python("raise ValueError('x')")
check("C2 run_python lỗi trong mã: depth vẫn về 0", s._undo_depth == 0, s._undo_depth)
s.line(0, 0, 1, 1)
check("C2 sau run_python lỗi: tool thường lại có mốc", d.log[-3:] == ["start", "AddLine", "end"], d.log)
d = fresh()
d.end_error = True
r = None
try:
    s.line(0, 0, 1, 1)
except Exception as e:
    r = str(e)
check("N-a tool OK mà EndUndoMark lỗi: báo RuntimeError hướng dẫn UNDO -> End", r is not None and "RuntimeError" in r and "UNDO" in r and "End" in r, r)
check("N-a sau đó depth về 0", s._undo_depth == 0, s._undo_depth)
d = fresh()
d.end_error = True
r = None
try:
    s.layer("L1")  # Doc giả không có Layers -> AttributeError (lỗi gốc khác ComError của EndUndoMark)
except Exception as e:
    r = str(e)
check("N-a tool lỗi + EndUndoMark lỗi: giữ lỗi gốc, không che", r is not None and "AttributeError" in r and "UNDO" not in r, r)
d = fresh()
class BadBlk(Blk):
    @property
    def Count(self):
        raise ComError(-2147352567)
d.Blocks.lst.append(BadBlk(d.Blocks, "XREF1"))
r = s.list_blocks()
check("N-b list_blocks: block đọc lỗi vẫn được liệt kê kèm ghi chú, block khác không hỏng",
      r[0] == {"name": "DOOR", "count": 3} and r[1]["name"] == "XREF1" and r[1]["count"] is None and r[1].get("note"), r)

# --- review security: save chuẩn hoá đường dẫn + một luật ổ đĩa; ADS / tên thiết bị; list_blocks kẹp limit ---
SEEN = []
real_exists = os.path.exists


def spy_exists(p):  # đường dẫn chưa qua luật ổ đĩa mà bị exists() thì có thể mở kết nối SMB: ghi lại, không chạm mạng
    SEEN.append(p)
    return False


BAD_SAVE = [
    (r"\/srv\share\x.dwg", "tuyệt đối"), (r"/\srv/s/x.dwg", "tuyệt đối"), (r"\/?\C:\x.dwg", "tuyệt đối"),
    ("//srv/s/x.dwg", "tuyệt đối"), (r"\\?\C:\x.dwg", "tuyệt đối"), (r"\\.\C:\x.dwg", "tuyệt đối"),
    ("C:x.dwg", "tuyệt đối"), ("x.dwg", "tuyệt đối"),
    (r"C:\a\b.txt:x.dwg", "không hợp lệ"), (r"C:\a\COM1.dwg", "không hợp lệ"), (r"C:\a\nul.dwg", "không hợp lệ"),
    (r"C:\a\x.dwg.", "không hợp lệ"), (r"C:\a\x .dwg ", "không hợp lệ"), (r"C:\a\NUL.tar.dwg", "không hợp lệ"),
    (r"C:\CON\x.dwg", "không hợp lệ"), (r"C:\a\x" + chr(7) + ".dwg", "không hợp lệ"), (r"C:\a\LPT9 .dwg", "không hợp lệ"),
    ("C:\\", "không hợp lệ"), (r"C:\a\x" + chr(0) + ".dwg", "không hợp lệ"), (r"\\?\UNC\srv\s\x.dwg", "tuyệt đối"),
    ("C\uff1a\\x.dwg", "tuyệt đối"), (" C:\\x.dwg", "tuyệt đối"),  # U+FF1A fullwidth ':' ; khoảng trắng đầu
    (r"C:\a\COM0.dwg", "không hợp lệ"), (r"C:\a\lpt0.dwg", "không hợp lệ"),
    # A-fix1: luật theo thành phần áp trên chuỗi GỐC, ".." không che được vi phạm; ký tự cấm <>"|?*
    (r"C:\a \..\b\x.dwg", "không hợp lệ"), (r"C:\a\nul\..\b.dwg", "không hợp lệ"), (r"C:\a:b\..\c.dwg", "không hợp lệ"),
    (r"C:\a\x" + chr(7) + r"\..\b.dwg", "không hợp lệ"), (r"C:\a\x<\..\b.dwg", "không hợp lệ"), (r"C:\a\..", "không hợp lệ"),
    (r"C:\a\k<l.dwg", "không hợp lệ"), (r"C:\a\k>l.dwg", "không hợp lệ"), ('C:\\a\\k"l.dwg', "không hợp lệ"),
    (r"C:\a\k|l.dwg", "không hợp lệ"), (r"C:\a\k?l.dwg", "không hợp lệ"), (r"C:\a\k*l.dwg", "không hợp lệ"),
]
d = fresh()
with mock.patch.object(os.path, "exists", spy_exists):
    for bad, txt in BAD_SAVE:
        check(f"SEC save {bad!r}: từ chối, không SaveAs", raised(lambda: s.save(bad), txt) and "SaveAs" not in d.log, d.log)
        check(f"SEC save {bad!r}: không exists()/mạng", SEEN == [], SEEN)
        SEEN.clear()
    r1 = s.save("C:/a/b/x.dwg")
    check("SEC save C:/a/b/x.dwg: chuẩn hoá về dấu \\ rồi nhận", r1 == "C:\\a\\b\\x.dwg" and d.FullName == "C:\\a\\b\\x.dwg", (r1, d.FullName))
    r2 = s.save("C:\\a\\x.DWG")
    check("SEC save C:\\a\\x.DWG: nhận", r2 == "C:\\a\\x.DWG", r2)
    r3 = s.save("C:\\a\\..\\z\\y.dwg")
    check("SEC save: normpath trước mọi kiểm (.. thu gọn)", r3 == "C:\\z\\y.dwg", r3)
    r4 = s.save(r"C:\\\\srv\share\x.dwg")
    check("SEC save C:\\\\\\\\srv\\share\\x.dwg (nhiều dấu \\ liền): thu gọn rồi nhận", r4 == "C:\\srv\\share\\x.dwg" and d.FullName == r4, (r4, d.FullName))
    vn = "C:\\D\u1ef1 \u00e1n (A&B)\\b\u1ea3n v\u1ebd #1 'v2'.dwg"   # F6: tên tiếng Việt có ngoặc, &, #, ' phải qua _save_path nguyên văn
    r5 = s.save(vn)
    check("F6 save tên tiếng Việt có (A&B) #1 'v2': nhận, trả nguyên văn", r5 == vn and d.FullName == vn, (r5, d.FullName))
    check("SEC save hợp lệ: exists chỉ gọi trên đường dẫn đã chuẩn hoá", all(x.startswith("C:\\") and "/" not in x for x in SEEN), SEEN)
d = fresh()
d.Blocks.lst += [Blk(d.Blocks, "B%d" % i) for i in range(2100)]
check("NOTE list_blocks limit=5000: kẹp 2000", len(s.list_blocks(5000)) == 2000, len(s.list_blocks(5000)))
check("NOTE list_blocks limit=-3: kẹp 1", len(s.list_blocks(-3)) == 1)
check("NOTE list_blocks limit=5: giữ 5", len(s.list_blocks(5)) == 5)


print("%d/%d PASS" % (n - fail, n))
sys.exit(fail)
