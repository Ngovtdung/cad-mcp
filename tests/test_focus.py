"""Kiểm lỗi 3.0.1 (docs\bug-3.0.1-focus.md §4): chặn Activate / đổi tab, _retry cả RETRYLATER, không che lỗi tool,
_app() không tạo AutoCAD ẩn. Không cần AutoCAD: module COM được thay bằng giả. Ca 4 chạy tiến trình con với mcp thật.
Chạy: python -I -B tests\test_focus.py   (mã thoát = số ca sai)"""
import asyncio
import importlib.util
import os
import subprocess
import sys
import time
import types
from unittest import mock

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

SERVER = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "server.py")
ENV = "CAD_CLAUDE_PIN_FILE"
n = fail = 0


def check(label, ok, detail=""):
    global n, fail
    n += 1
    if not ok:
        fail += 1
        print("FAIL", label, detail)


class ComError(Exception):
    def __init__(self, hresult):
        super().__init__(hresult)
        self.hresult = hresult


def load(real_mcp=False):
    """Nạp server.py MỚI. real_mcp=True: giữ mcp thật, chỉ giả pywin32 + PIL."""
    mods = ["pythoncom", "win32com", "win32com.client", "win32gui", "win32ui", "PIL", "PIL.Image"]
    if not real_mcp:
        mods += ["mcp", "mcp.server", "mcp.server.mcpserver"]
    for m in mods:
        sys.modules[m] = mock.MagicMock()
    sys.modules["win32com"].client = sys.modules["win32com.client"]  # `import win32com.client` đi qua thuộc tính này
    sys.modules["pythoncom"].com_error = ComError
    sys.modules["win32com.client"].DispatchBaseClass = type("DispatchBaseClass", (), {})
    if not real_mcp:
        sys.modules["mcp.server.mcpserver"].MCPServer.return_value.tool.return_value = lambda f: f  # tool vẫn là hàm thật
        sys.modules["mcp.server.mcpserver.exceptions"] = types.SimpleNamespace(ToolError=type("ToolError", (Exception,), {}))
    os.environ.pop(ENV, None)
    spec = importlib.util.spec_from_file_location("server_under_test", SERVER)
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod


def raised(f, exc, text):
    try:
        f()
    except exc as e:
        return text in str(e)
    except Exception:
        return False
    return False


def real_mcp_case():
    """Ca 4 (tiến trình riêng): tools/list giữ nguyên tên + schema với mcp thật."""
    s = load(real_mcp=True)
    tools = {t.name: t for t in asyncio.run(s.mcp.list_tools())}
    check("ca 4: đủ 40 tool", len(tools) == 40, len(tools))
    for name in ("run_python", "screenshot", "zoom_extents", "undo", "redo", "mirror", "offset", "array_rect",
                 "array_polar", "ellipse", "create_block", "list_blocks", "save_skill"):
        check("ca 4: có tool " + name, name in tools)
    rp = tools.get("run_python")
    sch = rp.input_schema if rp and hasattr(rp, "input_schema") else getattr(rp, "inputSchema", {})
    check("ca 4: run_python có code (string, bắt buộc)",
          sch.get("properties", {}).get("code", {}).get("type") == "string" and sch.get("required") == ["code"], sch)
    sc = tools.get("screenshot")
    sch = sc.input_schema if hasattr(sc, "input_schema") else getattr(sc, "inputSchema", {})
    check("ca 4: screenshot có zoom_extents (boolean), không bắt buộc",
          sch.get("properties", {}).get("zoom_extents", {}).get("type") == "boolean" and not sch.get("required"), sch)
    print("%d/%d PASS" % (n - fail, n))
    sys.exit(fail)


if "--real" in sys.argv:
    real_mcp_case()

s = load()
ToolError = sys.modules["mcp.server.mcpserver.exceptions"].ToolError

# --- ca 1: _retry thử lại cả RPC_E_SERVERCALL_RETRYLATER ---
calls = []


def busy():
    calls.append(1)
    if len(calls) < 3:
        raise ComError(-2147417846)
    return 5


with mock.patch.object(time, "sleep", lambda *_: None):
    try:
        r = s._retry(busy)
    except Exception as e:
        r = e
check("ca 1: _retry thử lại -2147417846 hai lần rồi trả 5", r == 5, repr(r))
with mock.patch.object(time, "sleep", lambda *_: None):
    check("ca 1: hresult khác vẫn ném ngay", raised(lambda: s._retry(lambda: (_ for _ in ()).throw(ComError(-1))), ComError, ""))


# --- ca 2: _R chặn Activate / gán ActiveDocument, Visible, WindowState ---
class Fake:
    Visible = True
    ActiveDocument = "d"
    activated = 0

    def Activate(self):
        Fake.activated += 1


fake = Fake()
check("ca 2: Activate() bị chặn", raised(lambda: s._R(fake).Activate(), RuntimeError, "không được đổi tab"))
check("ca 2: fake.Activate không được gọi", Fake.activated == 0)
for attr in ("ActiveDocument", "WindowState"):
    check("ca 2: gán " + attr + " bị chặn", raised(lambda a=attr: setattr(s._R(fake), a, 1), RuntimeError, "không được đổi tab"))
class IAcadApplication:  # tên lớp makepy của ứng dụng
    Visible = True


ent = Fake()  # entity: Visible hợp lệ
s._R(ent).Visible = False
check("ca 2: gán entity.Visible được phép", ent.Visible is False)
check("ca 2: gán Visible trên AcadApplication bị chặn", raised(lambda: setattr(s._R(IAcadApplication()), "Visible", 0), RuntimeError, "không được đổi tab"))
check("ca 2: đọc .Visible vẫn được", s._R(Fake()).Visible is True)
check("ca 2: đọc .ActiveDocument vẫn được", s._R(Fake()).ActiveDocument == "d")

# --- ca 3: lỗi tool không bị che: run_python ném ToolError có chữ lỗi thật ---
s._doc = lambda: (_ for _ in ()).throw(RuntimeError("bản vẽ X"))
check("ca 3: run_python ném ToolError chứa 'bản vẽ X'", raised(lambda: s.run_python("1"), ToolError, "bản vẽ X"))
check("ca 3: ToolError có tên lỗi gốc", raised(lambda: s.run_python("1"), ToolError, "RuntimeError"))

# --- ca 4: mcp thật, tiến trình riêng ---
p = subprocess.run([sys.executable, "-I", "-B", os.path.abspath(__file__), "--real"], capture_output=True, text=True,
                   encoding="utf-8", errors="replace")
out = p.stdout.strip().splitlines()
for line in out:
    if line.startswith("FAIL"):
        print(line)
m = out[-1] if out else ""
try:
    ok, tot = (int(x) for x in m.split()[0].split("/"))
except (ValueError, IndexError):
    ok = tot = 0
    print("FAIL ca 4: tiến trình con lỗi", p.stderr[-500:])
n += max(tot, 1)
fail += (tot - ok) if tot else 1

# --- ca 5: _app() không bao giờ tạo AutoCAD ---
s = load()
com, ed = sys.modules["win32com.client"], sys.modules["win32com.client"].gencache.EnsureDispatch
com.GetActiveObject.side_effect = ComError(-2147221021)
check("ca 5: _app() báo 'không thấy AutoCAD'", raised(s._app, RuntimeError, "không thấy AutoCAD"))
check("ca 5: EnsureDispatch không được gọi", not ed.called)
com.GetActiveObject.side_effect = None
s._app()
check("ca 5: không bao giờ EnsureDispatch('AutoCAD.Application')",
      all(c.args != ("AutoCAD.Application",) for c in ed.call_args_list), ed.call_args_list)
check("ca 5: có app thì dùng GetActiveObject", com.GetActiveObject.called)

print("%d/%d PASS" % (n - fail, n))
sys.exit(fail)
