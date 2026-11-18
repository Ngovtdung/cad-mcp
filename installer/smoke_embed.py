# -*- coding: utf-8 -*-
"""Kiểm Python nhúng của một bản cài: import đủ module server.py cần, sys.path và pythoncom/pywintypes nằm trong python nhúng,
cache COM (gen_py) trong đó, server MCP trả lời initialize + tools/list qua stdio, và xls2tsv.py đọc được .xls bằng xlrd vendored.
KHÔNG gọi tool nào -> không chạm AutoCAD. PATH chỉ còn System32, bỏ PYTHON*/ANTHROPIC_*/CLAUDE_*/DEEPSEEK_* -> không cần Python hệ thống.
    python -I smoke_embed.py <thư_mục_cài (chứa server.py, xls2tsv.py, vendor\\, python\\python.exe)>"""
import json, os, re, subprocess, sys, threading
from pathlib import Path

XLS_FIXTURE = Path(__file__).resolve().parent.parent / "tests" / "fixtures-xls" / "book.xls"   # chỉ để kiểm, không đóng gói
CODE = r"""
import os, sys
import pythoncom, pywintypes, win32com, win32com.client, win32gui, win32ui
from PIL import Image
from mcp.server.mcpserver import Image as McpImage, MCPServer
py = os.path.dirname(os.path.abspath(sys.executable)).lower()
bad = [p for p in sys.path if p and not os.path.abspath(p).lower().startswith(py)]
assert not bad, "sys.path ngoai python nhung: %r" % bad
for m in (pythoncom, pywintypes):
    assert os.path.abspath(m.__file__).lower().startswith(py), m.__file__
win32com.client.gencache.EnsureDispatch("Scripting.Dictionary")
assert os.path.abspath(win32com.__gen_path__).lower().startswith(py), win32com.__gen_path__
print("IMPORT OK", sys.version.split()[0])
"""
NEED_TOOLS = {"info", "command", "list_entities"}
SECRET_ENV = re.compile(r"(?i)^(PYTHON|ANTHROPIC_|CLAUDE_|DEEPSEEK_)")


def _env():
    root = os.environ.get("SystemRoot", r"C:\Windows")
    env = {k: v for k, v in os.environ.items() if not SECRET_ENV.match(k)}
    env["PATH"] = rf"{root}\System32;{root}"
    return env


def check_xls(app):
    py = app / "python" / "python.exe"
    env = {k: os.environ[k] for k in ("SystemRoot", "windir", "TEMP", "TMP") if k in os.environ}   # như Attach.RunXls
    r = subprocess.run([str(py), "-I", "-S", "-B", "-X", "utf8", str(app / "xls2tsv.py")], input=XLS_FIXTURE.read_bytes(),
                       capture_output=True, env=env, cwd=str(app), timeout=60)
    first = r.stdout.split(b"\n", 1)[0]
    if r.returncode or not first.startswith(b"XLS2TSV 1 "):
        raise RuntimeError(f"xls2tsv lỗi mã {r.returncode}: {r.stderr[-500:]!r}")
    if (app / "python" / "Lib" / "site-packages" / "xlrd").exists():
        raise RuntimeError("xlrd không được nằm trong site-packages (phải dùng vendor\\)")
    print("XLS OK", first.decode("ascii", "replace"))


def check(app):
    app = Path(app).resolve()
    py = app / "python" / "python.exe"
    if not py.is_file():
        raise RuntimeError(f"không có {py}")
    r = subprocess.run([str(py), "-I", "-c", CODE], env=_env(), cwd=os.environ["TEMP"], capture_output=True,
                       text=True, encoding="utf-8", errors="replace", timeout=180)
    print(r.stdout.strip())
    if r.returncode or "IMPORT OK" not in r.stdout:
        raise RuntimeError("import lỗi: " + r.stderr.strip()[-1500:])
    p = subprocess.Popen([str(py), "-I", str(app / "server.py")], env=_env(), cwd=os.environ["TEMP"],
                         stdin=subprocess.PIPE, stdout=subprocess.PIPE, stderr=subprocess.DEVNULL)
    timer = threading.Timer(90, p.kill)
    timer.start()
    try:
        def send(o):
            p.stdin.write((json.dumps(o) + "\n").encode())
            p.stdin.flush()

        def recv(i):
            for line in p.stdout:
                try:
                    m = json.loads(line)
                except ValueError:
                    continue
                if m.get("id") == i:
                    return m
            raise RuntimeError("server MCP đóng stdout hoặc quá 90 s")

        send({"jsonrpc": "2.0", "id": 1, "method": "initialize", "params": {
            "protocolVersion": "2025-06-18", "capabilities": {}, "clientInfo": {"name": "smoke", "version": "0"}}})
        recv(1)
        send({"jsonrpc": "2.0", "method": "notifications/initialized"})
        send({"jsonrpc": "2.0", "id": 2, "method": "tools/list"})
        names = {t["name"] for t in recv(2)["result"]["tools"]}
    finally:
        timer.cancel()
        p.kill()
        p.wait()
    if not NEED_TOOLS <= names:
        raise RuntimeError(f"thiếu tool: {sorted(NEED_TOOLS - names)}")
    print(f"MCP OK: {len(names)} tool")
    check_xls(app)


if __name__ == "__main__":
    if len(sys.argv) != 2:
        sys.exit(__doc__)
    try:
        check(sys.argv[1])
    except RuntimeError as e:
        sys.exit(f"SMOKE FAIL: {e}")
