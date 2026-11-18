# -*- coding: utf-8 -*-
"""Python nhúng + vendor xlrd cho bản cài (spec §3). Nguồn DUY NHẤT: python.org (zip ghim SHA256 + chữ ký PSF) và PyPI (wheel --require-hashes).
    python -I build_python.py --lock               tải wheel theo requirements-embed.txt, ghi requirements-embed.lock
    python -I build_python.py --fetch              tải zip Python nhúng, kiểm chữ ký PSF, in SHA256 (để ghim PY_SHA256)
    python -I build_python.py <thư_mục_mới>        dựng Python nhúng (pack_setup.py gọi build())
    python -I build_python.py --vendor <thư_mục>   dựng vendor xlrd (pack_setup.py gọi build_vendor())"""
import hashlib, os, re, shutil, subprocess, sys, tempfile, urllib.request, zipfile
from pathlib import Path

INST = Path(__file__).resolve().parent
ROOT = INST.parent                                   # C:\Autodesk\cad-mcp
CACHE = Path(r"C:\Autodesk\dist\_cache-setup")
WHEELS = CACHE / "wheels"
REQ = INST / "requirements-embed.txt"
LOCK = INST / "requirements-embed.lock"
XLS_REQ = ROOT / "requirements-xls.txt"             # xlrd==2.0.2 + hash (spec-attachments §2.2c)
PY_VER = "3.13.17"
PY_URL = f"https://www.python.org/ftp/python/{PY_VER}/python-{PY_VER}-embed-amd64.zip"
PY_SHA256 = "9e35319bffedab8b36ad44d8f377cb4f1733bcec257a0b6d04d94f01486ddb38"
PSF = "Python Software Foundation"
PS = Path(os.environ.get("SystemRoot", r"C:\Windows")) / "System32" / "WindowsPowerShell" / "v1.0" / "powershell.exe"
# ._pth: chỉ các đường dẫn này vào sys.path; PYTHONPATH/PYTHONHOME/registry bị bỏ qua. "import site" -> pywin32.pth -> pywin32_bootstrap
PTH = "python313.zip\n.\nLib\\site-packages\nimport site\n"
VENDOR_FILES = {f"xlrd/{n}.py" for n in ("__init__", "biffh", "book", "compdoc", "formatting", "formula", "info", "sheet",
                                          "timemachine", "xldate")} | \
               {f"xlrd-2.0.2.dist-info/{n}" for n in ("LICENSE", "METADATA", "WHEEL", "top_level.txt")}
SECRET_ENV = re.compile(r"(?i)^(ANTHROPIC_|CLAUDE_|DEEPSEEK_)")


def child_env(**extra):
    """Môi trường cho tiến trình con: bỏ biến bí mật."""
    # PSModulePath bỏ đi: nếu chạy từ pwsh 7, powershell.exe 5.1 con kế thừa nó và không nạp được Microsoft.PowerShell.Security
    env = {k: v for k, v in os.environ.items() if not SECRET_ENV.match(k) and k.lower() != "psmodulepath"}
    env.update(extra)
    return env


def sha256(p):
    h = hashlib.sha256()
    with open(p, "rb") as f:
        for b in iter(lambda: f.read(1 << 20), b""):
            h.update(b)
    return h.hexdigest()


def pip(*args):
    subprocess.run([sys.executable, "-I", "-m", "pip", "--isolated", *args, "--disable-pip-version-check"], check=True, env=child_env())


def check_sig(path, signer=None, microsoft_toolchain=False):
    """signer: SimpleName nghiêm; microsoft_toolchain=True: chế độ org+issuer Microsoft (chỉ cho dotnet.exe)."""
    mode = ["-MicrosoftToolchain"] if microsoft_toolchain else ["-Signer", signer]
    r = subprocess.run([str(PS), "-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass", "-File", str(INST / "check-sig.ps1"),
                        "-Path", str(path), *mode], capture_output=True, text=True, timeout=120, env=child_env())
    if r.returncode:
        raise SystemExit(r.stdout.strip() or f"chữ ký hỏng: {path}")


def fetch_python():
    z = CACHE / f"python-{PY_VER}-embed-amd64.zip"
    if not z.is_file():
        CACHE.mkdir(parents=True, exist_ok=True)
        tmp = z.with_suffix(".part")
        urllib.request.urlretrieve(PY_URL, tmp)
        tmp.replace(z)
    return z


def check_psf(d):
    for n in ("python.exe", "python313.dll"):
        check_sig(Path(d) / n, PSF)


def fetch_and_print():
    z = fetch_python()
    with tempfile.TemporaryDirectory(dir=CACHE) as t:
        with zipfile.ZipFile(z) as f:
            f.extractall(t)
        try:
            check_psf(t)
        except SystemExit:
            z.unlink()
            raise
    print(sha256(z))


def lock():
    if WHEELS.exists():
        shutil.rmtree(WHEELS)
    pip("download", "--only-binary=:all:", "--index-url", "https://pypi.org/simple", "-d", str(WHEELS), "-r", str(REQ))
    others = [p.name for p in WHEELS.iterdir() if p.suffix != ".whl"]
    if others:
        raise SystemExit(f"có tệp không phải wheel: {others}")
    lines = [f"{w.name.split('-')[0]}=={w.name.split('-')[1]} --hash=sha256:{sha256(w)}" for w in sorted(WHEELS.glob("*.whl"))]
    if any(l.lower().startswith("xlrd==") for l in lines):
        raise SystemExit("xlrd không được vào site-packages")
    LOCK.write_text("# Sinh bởi build_python.py --lock — không sửa tay.\n" + "\n".join(lines) + "\n", encoding="utf-8")
    print(f"{LOCK}: {len(lines)} gói")


def _pip_target(req, target):
    pip("download", "--only-binary=:all:", "--no-deps", "--require-hashes", "-d", str(WHEELS), "-r", str(req))
    # --target: không chạy pywin32_postinstall (không đăng ký COM, không chép vào System32); --no-compile: không .pyc mang đường dẫn máy build
    pip("install", "--no-index", "--find-links", str(WHEELS), "--only-binary=:all:", "--no-deps", "--require-hashes",
        "--no-compile", "--no-warn-script-location", "--target", str(target), "-r", str(req))
    shutil.rmtree(Path(target) / "bin", ignore_errors=True)   # exe console-script của pip nhúng đường dẫn python máy build


def build(dest):
    dest = Path(dest)
    if dest.exists():
        raise SystemExit(f"{dest} đã tồn tại")
    z = fetch_python()
    if sha256(z) != PY_SHA256:
        z.unlink()
        raise SystemExit(f"SHA256 {z.name} sai — đã xoá khỏi cache")
    with zipfile.ZipFile(z) as f:
        f.extractall(dest)
    check_psf(dest)
    pth = dest / "python313._pth"
    if not pth.is_file():
        raise SystemExit("gói nhúng không có python313._pth")
    pth.write_text(PTH, encoding="ascii")
    site = dest / "Lib" / "site-packages"
    _pip_target(LOCK, site)
    if (site / "xlrd").exists():
        raise SystemExit("xlrd lọt vào site-packages")
    gen = site / "win32com" / "gen_py"                       # cache COM nằm trong bản cài (win32com\__init__.py:95-113)
    gen.mkdir(exist_ok=True)
    (gen / "__init__.py").write_text("", encoding="ascii")
    print(f"Python nhúng {PY_VER} -> {dest}")


def build_vendor(dest):
    dest = Path(dest)
    if dest.exists():
        raise SystemExit(f"{dest} đã tồn tại")
    _pip_target(XLS_REQ, dest)
    for n in ("RECORD", "INSTALLER", "REQUESTED", "direct_url.json"):
        (dest / "xlrd-2.0.2.dist-info" / n).unlink(missing_ok=True)
    got = {p.relative_to(dest).as_posix() for p in dest.rglob("*") if p.is_file()}
    if got != VENDOR_FILES:
        raise SystemExit(f"vendor sai: thừa {sorted(got - VENDOR_FILES)}, thiếu {sorted(VENDOR_FILES - got)}")
    print(f"vendor xlrd -> {dest}")


if __name__ == "__main__":
    a = sys.argv[1:]
    if a == ["--lock"]:
        lock()
    elif a == ["--fetch"]:
        fetch_and_print()
    elif len(a) == 2 and a[0] == "--vendor":
        build_vendor(a[1])
    elif len(a) == 1 and not a[0].startswith("-"):
        build(a[0])
    else:
        sys.exit(__doc__)
