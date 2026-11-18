"""Kiểm công cụ save_skill (3.1.1): AI chỉ ghi BẢN NHÁP SKILL.md vào CAD_CLAUDE_SKILL_DRAFTS (<...>\\skill-nhap), không bao giờ vào .claude\\skills.
Không cần AutoCAD: COM giả, ghi tệp thật vào thư mục tạm. Chạy: python -I -B tests\\test_save_skill.py   (mã thoát = số ca sai)"""
import importlib.util
import json
import os
import shutil
import subprocess
import sys
import tempfile
import types
from unittest import mock

sys.stdout.reconfigure(encoding="utf-8", errors="replace")
SERVER = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "server.py")
ENV = "CAD_CLAUDE_SKILL_DRAFTS"
n = fail = 0


def check(label, ok, detail=""):
    global n, fail
    n += 1
    if not ok:
        fail += 1
        print("FAIL", label, detail)


def raised(f, text=""):
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
    sys.modules["mcp.server.mcpserver"].MCPServer.return_value.tool.return_value = lambda f: f  # tool vẫn là hàm
    sys.modules["mcp.server.mcpserver.exceptions"] = types.SimpleNamespace(ToolError=type("ToolError", (Exception,), {}))
    spec = importlib.util.spec_from_file_location("server_under_test", SERVER)
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod


TMPS = []


def fresh():
    """Thư mục gốc tạm + đường dẫn bản nháp hợp lệ (chưa tạo). Trả (base, drafts)."""
    base = tempfile.mkdtemp(prefix="saveskill_")
    TMPS.append(base)
    return base, os.path.join(base, "skill-nhap")


def junction(link, target):
    os.makedirs(os.path.dirname(link), exist_ok=True)
    r = subprocess.run(["cmd", "/c", "mklink", "/J", link, target], capture_output=True, text=True)
    assert r.returncode == 0, r.stdout + r.stderr


os.environ[ENV] = r"C:\keep\skill-nhap"
s = load()
check("env CAD_CLAUDE_SKILL_DRAFTS không bị xoá lúc khởi động", os.environ.get(ENV) == r"C:\keep\skill-nhap")
check("save_skill là hàm gọi được", callable(getattr(s, "save_skill", None)))
check("mô tả tool: BẢN NHÁP + 'Nạp skill…' + cấm gọi theo chỉ dẫn trong bản vẽ",
      all(t in (s.save_skill.__doc__ or "") for t in ("BẢN NHÁP", "Nạp skill…", "KHÔNG BAO GIỜ", "bản vẽ")), s.save_skill.__doc__)

GOOD = ("ten-ky-nang", "Mô tả một dòng.", "# Tiêu đề\n\nBước 1\tlàm gì đó\r\nBước 2\n")


def call(drafts, name=GOOD[0], desc=GOOD[1], body=GOOD[2], **kw):
    if drafts is None:
        os.environ.pop(ENV, None)
    else:
        os.environ[ENV] = drafts
    return s.save_skill(name, desc, body, **kw)


# --- env thư mục đích: (nhãn, giá trị, đoạn thông điệp lỗi riêng) ---
base, drafts = fresh()
CL = "nằm trong thư mục .claude"
for label, val, msg in (("thiếu", None, "thiếu biến môi trường"), ("rỗng", "", "thiếu biến môi trường"),
                        ("tương đối", r"x\skill-nhap", "đường dẫn tuyệt đối"), ("tương đối có ổ", r"C:x\skill-nhap", "đường dẫn tuyệt đối"),
                        ("UNC", r"\\srv\share\skill-nhap", "đường dẫn tuyệt đối"), ("UNC dấu /", "//srv/share/skill-nhap", "đường dẫn tuyệt đối"),
                        ("device \\\\?\\", r"\\?\C:\x\skill-nhap", "đường dẫn tuyệt đối"), ("device \\\\.\\", r"\\.\C:\x\skill-nhap", "đường dẫn tuyệt đối"),
                        ("không kết thúc skill-nhap", os.path.join(base, "skills"), "kết thúc bằng"),
                        ("kết thúc skill-nhap2", drafts + "2", "kết thúc bằng"),
                        (".claude\\skills (không bao giờ)", os.path.join(base, ".claude", "skills"), "kết thúc bằng"),
                        (".claude\\skills\\skill-nhap", os.path.join(base, ".claude", "skills", "skill-nhap"), CL),
                        (".claude\\skill-nhap", os.path.join(base, ".claude", "skill-nhap"), CL),
                        (".CLAUDE\\skill-nhap (khác hoa thường)", os.path.join(base, ".CLAUDE", "skill-nhap"), CL),
                        (".claude ở giữa, qua ..", os.path.join(base, ".claude", "a", "..", "skill-nhap"), CL)):
    check(f"env {label}: từ chối đúng lý do", raised(lambda: call(val), msg), val)
    check(f"env {label}: không ghi gì", os.listdir(base) == [], os.listdir(base))
base, drafts = fresh()
r = call(drafts.upper())
check("env khác hoa thường (SKILL-NHAP): nhận", "Đã lưu" in r, r)

# --- đầu vào: (nhãn, tham số, đoạn thông điệp lỗi riêng của nhóm) ---
HID = (0xFE00, 0xFE0F, 0xE0100, 0xE01EF, 0x115F, 0x1160, 0x3164, 0xFFA0)  # ký tự ẩn ngoài nhóm Cf
N_, D_, CF_, HD_ = "name phải khớp", "description phải", "Unicode Cf", "ký tự ẩn"
SH_ = "cú pháp chèn lệnh shell"
BAD = [
    ("name ../x", dict(name="../x"), N_), ("name A", dict(name="A"), N_), ("name a b", dict(name="a b"), N_),
    ("name 65 ký tự", dict(name="a" * 65), N_), ("name rỗng", dict(name=""), N_), ("name -a", dict(name="-a"), N_),
    ("name xuống dòng cuối", dict(name="ab\n"), N_), ("name a\\b", dict(name="a\\b"), N_), ("name bidi", dict(name="a\u202a"), N_),
    ("name Cf U+200B", dict(name="a\u200b"), N_), ("name bắt đầu bằng số", dict(name="1abc"), N_), ("name toàn số", dict(name="123"), N_),
    ("name a--b", dict(name="a--b"), N_), ("name a-", dict(name="a-"), N_), ("name a-b-", dict(name="a-b-"), N_), ("name a---b", dict(name="a---b"), N_),
] + [(f"name YAML {w}", dict(name=w), N_) for w in ("true", "false", "null", "yes", "no", "on", "off", "y", "n")] \
  + [(f"name thiết bị {d}", dict(name=d), N_) for d in ("con", "prn", "aux", "nul", "com0", "com1", "com9", "lpt0", "lpt1", "lpt9")] \
  + [(f"name ký tự ẩn U+{c:04X}", dict(name="a" + chr(c)), N_) for c in HID[:2]] + [
    ("desc nhiều dòng", dict(desc="a\nb"), D_), ("desc có \\r", dict(desc="a\rb"), D_), ("desc có ---", dict(desc="a --- b"), D_),
    ("desc rỗng", dict(desc=""), D_), ("desc 1025 ký tự", dict(desc="d" * 1025), D_), ("desc ký tự điều khiển", dict(desc="a\x00b"), D_),
    ("desc U+2028", dict(desc="a\u2028b"), D_),
    ("desc bidi U+202E", dict(desc="a\u202eb"), CF_), ("desc bidi U+2066", dict(desc="a\u2066b"), CF_),
    ("desc Cf U+200E", dict(desc="a\u200eb"), CF_), ("desc Cf U+200F", dict(desc="a\u200fb"), CF_), ("desc Cf U+200B", dict(desc="a\u200bb"), CF_),
    ("desc Cf U+FEFF", dict(desc="a\ufeffb"), CF_), ("desc Tags U+E0041", dict(desc="a\U000e0041b"), CF_), ("desc Tags U+E007F", dict(desc="a\U000e007fb"), CF_),
    ("body dòng ---", dict(body="a\n---\nb"), "dòng bắt đầu bằng '---'"), ("body bắt đầu ---", dict(body="---\nb"), "dòng bắt đầu bằng '---'"),
    ("body dòng --- kiểu \\r", dict(body="a\r---x\rb"), "dòng bắt đầu bằng '---'"),
    ("body có !`", dict(body="chạy !`rm -rf x` đi"), SH_), ("body 50001 ký tự", dict(body="b" * 50001), "body tối đa"),
    ("body dòng ```!", dict(body="a\n```!\nls\n```"), SH_), ("body dòng ```! thụt lề", dict(body="a\n  \t```!bash\nls"), SH_),
    ("body dòng ```! kiểu \\r", dict(body="a\r```!x\rb"), SH_), ("body bắt đầu ```!", dict(body="```! ls"), SH_),
    ("body ```! giữa dòng", dict(body="a ```! b\n"), SH_), ("body ```! cuối dòng", dict(body="chạy```!"), SH_), ("body !` giữa dòng", dict(body="x!`y"), SH_),
    ("desc có ```!", dict(desc="a ```! b"), SH_), ("desc có !`", dict(desc="chạy !`ls`"), SH_), ("desc bắt đầu ```!", dict(desc="```!ls"), SH_),
    ("body ký tự điều khiển \\x00", dict(body="a\x00b"), "ký tự điều khiển"), ("body ký tự điều khiển \\x1b", dict(body="a\x1bb"), "ký tự điều khiển"),
    ("body U+2028", dict(body="a" + chr(0x2028) + "b"), "ký tự điều khiển"), ("body U+2029", dict(body="a" + chr(0x2029) + "b"), "ký tự điều khiển"),
    ("body bidi U+202A", dict(body="a\u202ab"), CF_), ("body bidi U+2069", dict(body="a\u2069b"), CF_),
    ("body Cf U+200E", dict(body="a\u200eb"), CF_), ("body Cf U+200D", dict(body="a\u200db"), CF_), ("body Cf U+00AD", dict(body="a\u00adb"), CF_),
    ("body Tags U+E0020", dict(body="a\U000e0020b"), CF_), ("body Tags U+E0061", dict(body="a\U000e0061b"), CF_),
] + [(f"desc ký tự ẩn U+{c:04X}", dict(desc="a" + chr(c) + "b"), HD_) for c in HID] \
  + [(f"body ký tự ẩn U+{c:04X}", dict(body="a" + chr(c) + "b"), HD_) for c in HID]
for label, kw, msg in BAD:
    base, drafts = fresh()
    check(f"đầu vào {label}: từ chối đúng lý do", raised(lambda: call(drafts, **kw), msg))
    check(f"đầu vào {label}: không ghi gì", os.listdir(base) == [], os.listdir(base))
base, drafts = fresh()
check("body đúng 50000 ký tự: nhận", "Đã lưu" in call(drafts, body="b" * 50000))
check("desc đúng 1024 ký tự: nhận", "Đã lưu" in call(drafts, name="d1024", desc="d" * 1024))
check("name đúng 64 ký tự: nhận", "Đã lưu" in call(drafts, name="a" + "0-" * 31 + "z"))
check("name con-x / com10 / comx (không phải tên thiết bị): nhận", all("Đã lưu" in call(drafts, name=x) for x in ("con-x", "com10", "comx")))
check("name yes-sir / nope / a1 (không phải từ YAML, bắt đầu bằng chữ): nhận", all("Đã lưu" in call(drafts, name=x) for x in ("yes-sir", "nope", "a1")))
check("name a-b-c / a1-2b (gạch nối đơn giữa các đoạn): nhận", all("Đã lưu" in call(drafts, name=x) for x in ("a-b-c", "a1-2b")))
check("body khối ```python và ``` (không có '!' ngay sau): nhận", "Đã lưu" in call(drafts, name="ok-3", body="```python\nx = 1\n```\na ``` b! `c`\n"))
check("desc có ``` và ! tách rời: nhận", "Đã lưu" in call(drafts, name="ok-4", desc="dùng ``` và ! và `x`"))
call(drafts, "crlf-1", body="a\r\nb\rc\nd")  # I2: \r\n và \r lẻ -> \n trước khi kiểm và ghi
check("body \\r\\n và \\r lẻ: ghi ra chỉ có \\n",
      open(os.path.join(drafts, "crlf-1", "SKILL.md"), "rb").read() == b'---\nname: crlf-1\ndescription: "M\xc3\xb4 t\xe1\xba\xa3 m\xe1\xbb\x99t d\xc3\xb2ng."\n---\n\na\nb\nc\nd\n')
check("body \\r\\n + dòng --- sau chuẩn hoá: vẫn từ chối", raised(lambda: call(drafts, "crlf-2", body="a\r\n---\r\nb"), "dòng bắt đầu bằng '---'"))
check("body 75000 ký tự trước chuẩn hoá (\\r\\n) = 50000 sau: nhận (độ dài tính sau chuẩn hoá)", "Đã lưu" in call(drafts, "crlf-3", body="x\r\n" * 25000))
check("body 50001 ký tự sau chuẩn hoá: từ chối", raised(lambda: call(drafts, "crlf-4", body="x\r\n" * 25000 + "y"), "body tối đa"))
check("body có '---' giữa dòng và '!' đứng riêng: nhận", "Đã lưu" in call(drafts, name="ok-2", body="a --- b\nHello! `x`\n"))

# --- ghi mới ---
base, drafts = fresh()
name, desc, body = "cua-so", "Vẽ cửa sổ", "# Cửa sổ\n\nDùng `create_block`.\tHết\r\n"
p = os.path.join(drafts, name, "SKILL.md")
r = call(drafts, name, desc, body)
check("ghi mới: văn bản trả về", r == f"Đã lưu BẢN NHÁP skill {name} tại {os.path.normpath(p)}. Người dùng phải bấm 'Nạp skill…' trên palette để xem và bật; "
      "skill chưa có hiệu lực cho tới khi được duyệt.", r)
raw = open(p, "rb").read()
check("ghi mới: byte chính xác, UTF-8 không BOM", raw == f'---\nname: {name}\ndescription: "{desc}"\n---\n\n{body.replace(chr(13) + chr(10), chr(10))}\n'.encode("utf-8"), raw)
check("ghi mới: không còn \\r nào trong tệp", b"\r" not in raw)
check("ghi mới: không BOM", not raw.startswith(b"\xef\xbb\xbf"))
check("ghi mới: thư mục skill chỉ có SKILL.md", os.listdir(os.path.join(drafts, name)) == ["SKILL.md"])

# --- trùng tên ---
before = open(p, "rb").read()
check("trùng tên không overwrite: lỗi 'đã có skill'", raised(lambda: call(drafts, name, "khác", "mới"), f"đã có skill {name}"))
check("trùng tên: tệp cũ nguyên vẹn", open(p, "rb").read() == before)
r = call(drafts, name, "Mô tả mới", "Nội dung mới", overwrite=True)
check("overwrite=True: thay nội dung", open(p, "rb").read() == '---\nname: cua-so\ndescription: "Mô tả mới"\n---\n\nNội dung mới\n'.encode("utf-8"))
check("overwrite=True: không để lại tệp tạm", os.listdir(os.path.join(drafts, name)) == ["SKILL.md"], os.listdir(os.path.join(drafts, name)))
check("overwrite=True: chưa có bản nháp vẫn tạo", "Đã lưu" in call(drafts, "moi-tinh", "d", "b", overwrite=True))
# M6: tệp tạm .SKILL-*.tmp sót lại (lần ghi trước chết giữa chừng) bị dọn khi bắt đầu ghi đè; tệp khác không đụng
stale = [os.path.join(drafts, name, x) for x in (".SKILL-abc123.tmp", ".SKILL-zzz.tmp")]
for sp in stale:
    open(sp, "wb").write(b"rac")
keep = os.path.join(drafts, name, "ghi-chu.txt")
open(keep, "wb").write(b"giu")
call(drafts, name, "d3", "b3", overwrite=True)
check("overwrite=True: dọn tệp .SKILL-*.tmp sót", not any(os.path.exists(sp) for sp in stale), os.listdir(os.path.join(drafts, name)))
check("overwrite=True: tệp khác trong thư mục skill không bị xoá", os.path.isfile(keep) and os.path.isfile(p))
os.remove(keep)

# --- overwrite: fsync trước os.replace; os.replace lỗi thì tệp cũ nguyên byte, không còn tệp tạm ---
ORDER = []
real_fsync, real_replace = os.fsync, os.replace
with mock.patch.object(os, "fsync", lambda fd: (ORDER.append("fsync"), real_fsync(fd))[1]), \
        mock.patch.object(os, "replace", lambda a, b: (ORDER.append("replace"), real_replace(a, b))[1]):
    call(drafts, name, "d2", "b2", overwrite=True)
check("overwrite=True: fsync rồi mới os.replace", ORDER == ["fsync", "replace"], ORDER)
before = open(p, "rb").read()


def bad_replace(a, b):
    raise OSError("giả lập replace")


with mock.patch.object(os, "replace", bad_replace):
    check("overwrite=True, os.replace lỗi: báo lỗi", raised(lambda: call(drafts, name, "x", "y", overwrite=True), "giả lập replace"))
check("os.replace lỗi: SKILL.md cũ nguyên byte", open(p, "rb").read() == before)
check("os.replace lỗi: không còn tệp tạm .SKILL-*.tmp", os.listdir(os.path.join(drafts, name)) == ["SKILL.md"], os.listdir(os.path.join(drafts, name)))

# --- <name> là TỆP, không phải thư mục: lỗi, tệp đó không bị xoá ---
base, drafts = fresh()
os.makedirs(drafts)
fp = os.path.join(drafts, "la-tep")
open(fp, "wb").write(b"giu nguyen")
for ow in (False, True):
    check(f"<name> là tệp (overwrite={ow}): từ chối", raised(lambda: call(drafts, "la-tep", overwrite=ow), "không phải thư mục"))
    check(f"<name> là tệp (overwrite={ow}): tệp còn nguyên", os.path.isfile(fp) and open(fp, "rb").read() == b"giu nguyen")

# --- trần 20 bản nháp ---
base, drafts = fresh()
for i in range(20):
    call(drafts, "nhap-%02d" % i, "d", "b")
check("20 bản nháp: tên mới bị từ chối", raised(lambda: call(drafts, "nhap-20"), "đã có 20 bản nháp; người dùng hãy duyệt hoặc xoá bớt"))
check("20 bản nháp: tên mới bị từ chối cả khi overwrite=True", raised(lambda: call(drafts, "nhap-20", overwrite=True), "đã có 20 bản nháp"))
check("20 bản nháp: không tạo thêm", len(os.listdir(drafts)) == 20, len(os.listdir(drafts)))
check("20 bản nháp: overwrite bản cùng tên vẫn được", "Đã lưu" in call(drafts, "nhap-07", "mới", "mới", overwrite=True))
check("20 bản nháp: trùng tên không overwrite vẫn báo 'đã có skill'", raised(lambda: call(drafts, "nhap-07"), "đã có skill nhap-07"))
base, drafts = fresh()
for i in range(19):
    call(drafts, "nhap-%02d" % i, "d", "b")
check("19 bản nháp: thêm bản thứ 20 được", "Đã lưu" in call(drafts, "nhap-19"))

# --- reparse point / đích resolve vào .claude ---
base, drafts = fresh()
other = os.path.join(base, "other")
os.makedirs(other)
junction(drafts, other)
check("skill-nhap là junction: từ chối", raised(lambda: call(drafts), "reparse"))
check("skill-nhap là junction: đích không bị ghi", os.listdir(other) == [], os.listdir(other))
base, drafts = fresh()
claude = os.path.join(base, ".claude")
os.makedirs(claude)
junction(os.path.join(base, "lnk"), claude)
check("đường dẫn qua junction resolve vào .claude\\: từ chối", raised(lambda: call(os.path.join(base, "lnk", "skill-nhap")), CL))
check("đường dẫn qua junction resolve vào .claude\\: .claude không bị ghi", os.listdir(claude) == [], os.listdir(claude))
for ow in (False, True):
    base, drafts = fresh()
    other = os.path.join(base, "other")
    os.makedirs(other)
    os.makedirs(drafts)
    junction(os.path.join(drafts, "tt"), other)
    check(f"thư mục con <name> là junction (overwrite={ow}): từ chối", raised(lambda: call(drafts, "tt", overwrite=ow), "reparse"))
    check(f"thư mục con <name> là junction (overwrite={ow}): đích không bị ghi", os.listdir(other) == [], os.listdir(other))

# --- realpath thoát khỏi thư mục nháp: xoá tệp vừa ghi ---
base, drafts = fresh()
real = os.path.realpath
with mock.patch.object(os.path, "realpath", lambda q: r"C:\elsewhere\SKILL.md" if q.endswith("SKILL.md") else real(q)):
    check("realpath tệp ngoài thư mục nháp: lỗi", raised(lambda: call(drafts, "thoat"), "ngoài"))
check("realpath ngoài: tệp vừa ghi bị xoá", not os.path.exists(os.path.join(drafts, "thoat", "SKILL.md")))
check("realpath ngoài: thư mục con do lần gọi này tạo bị xoá", not os.path.exists(os.path.join(drafts, "thoat")))

# --- description = json.dumps (chuỗi YAML ngoặc kép): ': ', '[', '#', '!', '"', '\' vẫn parse đúng ---
base, drafts = fresh()
for i, desc in enumerate(("Dùng khi: vẽ cửa", "[danh sách] x", "# không phải chú thích", "!tag x", 'nói "xin chào"', "đường C:\\a\\b",
                          'kết thúc bằng \\', '"đầu ngoặc"', "a: b #c {d} &e *f |g >h 'i' %j @k `l`")):
    nm = "yaml-%d" % i
    call(drafts, nm, desc)
    lines = open(os.path.join(drafts, nm, "SKILL.md"), encoding="utf-8").read().split("\n")
    check(f"YAML mô tả {desc!r}: frontmatter đúng cấu trúc", lines[0] == "---" and lines[1] == "name: " + nm and lines[3] == "---" and lines[4] == "", lines[:5])
    check(f"YAML mô tả {desc!r}: đúng json.dumps(ensure_ascii=False)", lines[2] == "description: " + json.dumps(desc, ensure_ascii=False), lines[2])
    check(f"YAML mô tả {desc!r}: giải mã ra đúng nguyên văn", lines[2].startswith("description: \"") and json.loads(lines[2][len("description: "):]) == desc, lines[2])

# --- lỗi giữa chừng (overwrite=False): thư mục con do lần gọi này tạo bị xoá; thư mục đã có thì giữ ---
base, drafts = fresh()
real_open = open


def boom(p, *a, **k):
    if str(p).endswith("SKILL.md"):
        raise PermissionError("giả lập")
    return real_open(p, *a, **k)


with mock.patch("builtins.open", boom):
    check("mở tệp lỗi: báo lỗi", raised(lambda: call(drafts, "loi-1"), "giả lập"))
check("mở tệp lỗi: thư mục con rỗng do lần gọi này tạo bị xoá", not os.path.exists(os.path.join(drafts, "loi-1")), os.listdir(drafts))
call(drafts, "con-lai", "d", "b")
with mock.patch("builtins.open", boom):
    check("mở lỗi trên thư mục đã có: báo lỗi", raised(lambda: call(drafts, "con-lai", "x", "y"), "giả lập"))
check("thư mục đã có SKILL.md: không bị xoá khi lần gọi khác lỗi", os.path.isfile(os.path.join(drafts, "con-lai", "SKILL.md")))

for t in TMPS:
    shutil.rmtree(t, ignore_errors=True)
print("%d/%d PASS" % (n - fail, n))
sys.exit(fail)
