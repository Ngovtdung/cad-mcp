"""Kiểm cách server.py ghim bản vẽ làm việc (C1). Không cần AutoCAD: các module COM / mcp được thay bằng giả.
Chạy: python -I -B tests\\test_pin.py   (mã thoát = số ca sai)"""
import importlib.util
import os
import sys
import tempfile
import time
import types
from unittest import mock

sys.stdout.reconfigure(encoding="utf-8", errors="replace")  # nhãn tiếng Việt, console cp1252

SERVER = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "server.py")
ENV = "CAD_CLAUDE_PIN_FILE"
n = fail = 0


def check(label, ok, detail=""):
    global n, fail
    n += 1
    if not ok:
        fail += 1
        print("FAIL", label, detail)


def load(env_file):
    """Nạp server.py MỚI (PIN_FILE đọc lúc nạp). env_file=None = không đặt biến."""
    for m in ("pythoncom", "win32com", "win32com.client", "win32gui", "win32ui", "PIL", "PIL.Image",
              "mcp", "mcp.server", "mcp.server.mcpserver"):
        sys.modules[m] = mock.MagicMock()
    sys.modules["mcp.server.mcpserver.exceptions"] = types.SimpleNamespace(ToolError=Exception)  # server.py nhập ToolError
    if env_file is None:
        os.environ.pop(ENV, None)
    else:
        os.environ[ENV] = env_file
    spec = importlib.util.spec_from_file_location("server_under_test", SERVER)
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod


class Doc:
    def __init__(self, name):
        self.FullName = name
        self.Name = os.path.basename(name)
        self._o = self


class App:
    def __init__(self, active, names):
        self.ActiveDocument = Doc(active)
        self.Documents = types.SimpleNamespace(Count=len(names), Item=lambda i: Doc(names[i]))


def raises(f, text):
    try:
        f()
    except RuntimeError as e:
        return text in str(e)
    return False


tmp = tempfile.mkdtemp()
pin = os.path.join(tmp, "cad_claude_doc_1234.txt")


def write(text, age=0.0, bump=True):
    """ghi tệp; bump=False: giữ nguyên mtime/size cũ để mô phỏng 'không đổi'"""
    with open(pin, "w", encoding="utf-8") as f:
        f.write(text)
    t = time.time() - age
    os.utime(pin, (t, t))


A, B = r"C:\x\A.dwg", r"C:\x\B.dwg"
app = App(r"C:\x\Active.dwg", [A, B, r"C:\x\Active.dwg"])

# --- không có biến: ghim một lần (như cũ). gettempdir() trỏ vào thư mục tạm của test: không đụng tệp ghim thật ---
default_pin = os.path.join(tmp, "cad_claude_doc.txt")


def load_default():
    with mock.patch.object(tempfile, "gettempdir", lambda: tmp):
        return load(None)


with open(default_pin, "w", encoding="utf-8") as f:
    f.write(A)
s = load_default()
check("không biến: PIN_FILE là tệp toàn cục cũ", s.PIN_FILE == default_pin and s._PIN_ENV is None, s.PIN_FILE)
check("không biến: lần đầu đọc A", s._pin_name(app) == A)
with open(default_pin, "w", encoding="utf-8") as f:
    f.write(B)
check("không biến: ghi lần 2 bị bỏ qua (giữ A)", s._pin_name(app) == A)
s = load_default()  # tệp cũ hơn 10 phút -> ActiveDocument
t = time.time() - 700
os.utime(default_pin, (t, t))
check("không biến: tệp cũ > 600 s -> ActiveDocument", s._pin_name(app) == r"C:\x\Active.dwg")
os.remove(default_pin)
s = load_default()
check("không biến: không có tệp -> ActiveDocument", s._pin_name(app) == r"C:\x\Active.dwg")

# --- có biến: đọc lại khi (mtime_ns, size) đổi ---
write(A)
s = load(pin)
check("biến: PIN_FILE theo biến", s.PIN_FILE == pin and s._PIN_ENV == pin, s.PIN_FILE)
check("biến: A", s._pin_name(app) == A)
write(B + "x")  # đổi cả cỡ lẫn mtime
check("biến: ghi B -> B (đọc lại)", s._pin_name(app) == B + "x")
write(B)
check("biến: B (cỡ khác)", s._pin_name(app) == B)
write(A)  # cùng cỡ với B nhưng mtime_ns mới
time.sleep(0.02)
write(B)
write(A)
check("biến: cùng cỡ, mtime mới -> đọc lại", s._pin_name(app) == A)

# --- không đổi sau 601 s: vẫn A (không còn cửa sổ 600 s) ---
write(A, age=3.0)
s = load(pin)
check("biến: A trước khi tua giờ", s._pin_name(app) == A)
real = time.time
with mock.patch.object(time, "time", lambda: real() + 601):
    check("biến: tệp không đổi sau 601 s vẫn là A", s._pin_name(app) == A)
t = time.time() - 700
os.utime(pin, (t, t))  # mtime đổi (cũ 700 s) nhưng nội dung A -> đọc lại vẫn A
check("biến: mtime cũ 700 s, nội dung A -> A", s._pin_name(app) == A)

# --- rỗng / mất / không đọc được: lỗi, KHÔNG rơi về ActiveDocument ---
write("")
s = load(pin)
check("biến + tệp rỗng: lỗi", raises(lambda: s._pin_name(app), "không có bản vẽ làm việc"))
check("biến + tệp rỗng, hỏi lại: vẫn lỗi (không ActiveDocument)", raises(lambda: s._pin_name(app), "không có bản vẽ làm việc"))
write("   \n")
check("biến + tệp chỉ khoảng trắng: lỗi", raises(lambda: s._pin_name(app), "không có bản vẽ làm việc"))
write(A)
check("biến: ghi A sau khi rỗng -> A", s._pin_name(app) == A)
os.remove(pin)
check("biến + tệp mất: lỗi (dù trước đó là A)", raises(lambda: s._pin_name(app), "không có bản vẽ làm việc"))
s = load(os.path.join(tmp, "khong_co_tep.txt"))
check("biến + chưa có tệp: lỗi", raises(lambda: s._pin_name(app), "không có bản vẽ làm việc"))
with open(pin, "wb") as f:
    f.write(b"\xff\xfe\xfa\xfb\x80\x81")  # không phải UTF-8
s = load(pin)
check("biến + tệp không đọc được (không phải UTF-8): lỗi", raises(lambda: s._pin_name(app), "không có bản vẽ làm việc"))
os.remove(pin)
os.mkdir(pin)  # tệp là thư mục: open() ném OSError
check("biến + 'tệp' là thư mục: lỗi", raises(lambda: s._pin_name(app), "không có bản vẽ làm việc"))
os.rmdir(pin)

# --- _doc(): trả đúng bản vẽ đã ghim; đóng rồi thì lỗi; tên rỗng thì lỗi ---
write(B)
s = load(pin)
s._app = lambda: app
check("_doc() chọn đúng bản vẽ ghim", s._doc()._o.FullName == B)
write(A)
check("_doc() theo tệp mới (lượt 2 trên tiến trình giữ ấm)", s._doc()._o.FullName == A)
write(r"C:\x\DaDong.dwg")
check("_doc() bản vẽ đã đóng: lỗi", raises(s._doc, "đã bị đóng"))
write("")
check("_doc() tệp rỗng: lỗi 'không có bản vẽ làm việc'", raises(s._doc, "không có bản vẽ làm việc"))

print("%d/%d PASS" % (n - fail, n))
sys.exit(fail)
