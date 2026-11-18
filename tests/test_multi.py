"""server.py 3.1.0: xoá biến xác thực, kiểm pid AutoCAD, không chen lệnh, khớp bản vẽ hai lượt.
Chạy: python -I -B tests\\test_multi.py   (mã thoát = số ca sai)"""
import importlib.util, os, sys, types
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


def raises(f, text):
    try:
        f()
    except RuntimeError as e:
        return text in str(e)
    return False


class FakeApp:
    HWND = 4242


def load(env):
    for m in ("pythoncom", "win32com", "win32com.client", "win32gui", "win32ui", "PIL", "PIL.Image",
              "mcp", "mcp.server", "mcp.server.mcpserver"):
        sys.modules[m] = mock.MagicMock()
    sys.modules["mcp.server.mcpserver.exceptions"] = types.SimpleNamespace(ToolError=Exception)
    for k in ("CAD_CLAUDE_ACAD_PID", "CAD_CLAUDE_PIN_FILE"):
        os.environ.pop(k, None)
    os.environ.update(env)
    spec = importlib.util.spec_from_file_location("server_under_test", SERVER)
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    mod.pythoncom.com_error = type("com_error", (Exception,), {"hresult": 0})
    mod.win32com.client.DispatchBaseClass = type("DispatchBaseClass", (), {})
    mod.win32com.client.GetActiveObject = lambda name: object()
    mod.win32com.client.gencache.EnsureDispatch = lambda o: FakeApp()
    return mod


class FakeDoc:
    def __init__(self, full, name, cmd=0):
        self.FullName, self.Name, self.cmd, self.sent = full, name, cmd, []
        self._o = self

    def GetVariable(self, name):
        return self.cmd if name == "CMDACTIVE" else 0

    def SendCommand(self, s):
        self.sent.append(s)


# 1. xoá biến xác thực (F1, tiền tố, hoa thường)
for k in ("ANTHROPIC_API_KEY", "ANTHROPIC_AUTH_TOKEN", "ANTHROPIC_CUSTOM_HEADERS", "CLAUDE_CODE_OAUTH_TOKEN",
          "CLAUDE_CODE_USE_BEDROCK", "AWS_BEARER_TOKEN_BEDROCK", "CLAUDE_API_KEY", "anthropic_foundry_x"):
    os.environ[k] = "dummy"
os.environ["CLAUDE_CODE_GIT_BASH_PATH"] = "keep"
load({})
left = [k for k in os.environ if k.upper().startswith(("ANTHROPIC_", "CLAUDE_CODE_USE_", "CLAUDE_CODE_OAUTH_"))
        or k.upper() in ("AWS_BEARER_TOKEN_BEDROCK", "CLAUDE_API_KEY")]
check("server.py xoá mọi biến xác thực", left == [], left)
check("server.py giữ biến khác", os.environ.get("CLAUDE_CODE_GIT_BASH_PATH") == "keep")

# 2. pid
pin = os.path.join(os.environ.get("TEMP", "."), "cad_claude_doc_test_multi.txt")
mod = load({"CAD_CLAUDE_ACAD_PID": "999", "CAD_CLAUDE_PIN_FILE": pin})
seen = []
mod._hwnd_pid = lambda h: (seen.append(h), 999)[1]
check("cùng pid: trả app", isinstance(mod._app()._o, FakeApp))
check("so pid theo app.HWND", seen == [4242], seen)
mod._hwnd_pid = lambda h: 1000
check("khác pid: từ chối", raises(mod._app, "không phải AutoCAD có palette"))
for bad in ({}, {"CAD_CLAUDE_ACAD_PID": "abc"}):
    m2 = load(dict(bad, CAD_CLAUDE_PIN_FILE=pin))
    check(f"N4: có PIN_FILE, pid {bad or 'thiếu'}: từ chối", raises(m2._app, "CAD_CLAUDE_ACAD_PID"))
m3 = load({})


def boom(h):
    raise AssertionError("không được gọi _hwnd_pid ở chế độ CLI")


m3._hwnd_pid = boom
check("CLI (không biến): không kiểm pid", isinstance(m3._app()._o, FakeApp))

# 3. CMDACTIVE
mod._is_active = lambda d: True
busy = FakeDoc(r"C:\x\a.dwg", "a.dwg", 1)
check("CMDACTIVE≠0: từ chối", raises(lambda: mod._DocR(busy).SendCommand("LINE\n"), "đang chạy dở") and busy.sent == [], busy.sent)
idle = FakeDoc(r"C:\x\a.dwg", "a.dwg", 0)
mod._DocR(idle).SendCommand("LINE\n")
check("CMDACTIVE=0: gửi", idle.sent == ["LINE\n"], idle.sent)
mod._is_active = lambda d: False
bg = FakeDoc(r"C:\x\a.dwg", "a.dwg")
check("không active: vẫn từ chối", raises(lambda: mod._DocR(bg).SendCommand("LINE\n"), "không phải tab đang mở") and bg.sent == [])

# 4. F5: khớp hai lượt
saved = FakeDoc(r"C:\jobs\Drawing1.dwg", "Drawing1.dwg")
unsaved = FakeDoc("", "Drawing1.dwg")
mod._app = lambda: types.SimpleNamespace(Documents=types.SimpleNamespace(Count=2, Item=lambda i: [saved, unsaved][i]))
mod._pin_name = lambda app: "Drawing1.dwg"
check("F5: tên ngắn khớp tuyệt đối bản chưa lưu", mod._doc()._o is unsaved)
mod._pin_name = lambda app: r"C:\jobs\Drawing1.dwg"
check("F5: đường dẫn đầy đủ khớp bản đã lưu", mod._doc()._o is saved)
a1, a2 = FakeDoc(r"C:\a\x.dwg", "x.dwg"), FakeDoc(r"C:\b\x.dwg", "x.dwg")
mod._app = lambda: types.SimpleNamespace(Documents=types.SimpleNamespace(Count=2, Item=lambda i: [a1, a2][i]))
mod._pin_name = lambda app: "x.dwg"
check("F5: tên ngắn trùng 2 bản: từ chối", raises(mod._doc, "cùng tên"))
mod._app = lambda: types.SimpleNamespace(Documents=types.SimpleNamespace(Count=1, Item=lambda i: a1))
check("F5: tên ngắn đúng 1 bản: nhận", mod._doc()._o is a1)

print(f"{n - fail}/{n} PASS")
sys.exit(fail)
