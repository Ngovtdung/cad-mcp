# -*- coding: utf-8 -*-
"""Đóng gói Setup.exe cho "Claude AI cho AutoCAD" rồi tự kiểm bằng cài/quét/gỡ thử im lặng (/TESTROOT). Spec §5.
Nguồn gói CHỈ: allowlist trong C:\\Autodesk\\cad-mcp + C:\\Autodesk\\dist\\_cache-setup (ghim hash).
Tệp bí mật chỉ được ĐỌC vào bộ nhớ để so khớp; không in, không ghi, không chép. Thông báo lỗi chỉ có "tệp: tên luật" và SỐ đếm;
ngoại lệ không lường trước chỉ in TÊN kiểu.
    python -I -B C:\\Autodesk\\cad-mcp\\installer\\pack_setup.py [--force]
    python -I -B C:\\Autodesk\\cad-mcp\\installer\\pack_setup.py --selftest
"""
import hashlib, json, os, re, shutil, subprocess, sys, time, unicodedata, winreg, zipfile
from pathlib import Path

INST = Path(__file__).resolve().parent
ROOT = INST.parent                                   # C:\Autodesk\cad-mcp
sys.path.insert(0, str(INST))                        # -I bỏ thư mục script khỏi sys.path
import build_python, smoke_embed                     # noqa: E402

DIST = Path(r"C:\Autodesk\dist")
DOTNET = Path(r"C:\Autodesk\tools\dotnet\dotnet.exe")
ISCC = Path(r"C:\Autodesk\tools\InnoSetup6\ISCC.exe")
INNO_EXE = Path(r"C:\Autodesk\tools\_dl\innosetup-6.7.4.exe")             # bộ cài gốc (Authenticode Valid, Pyrsys B.V.); kiểm hash nếu còn
INNO_SHA256 = "a549929319f9120716677c685535ca4671e4b937b5fdc87e2585f32d8eb30e8c"
# *.e32/*.e64 không ký (NotSigned) nhưng thành stub Setup + unins000.exe -> ghim hash; đã so trùng byte với bản cài mới từ INNO_EXE
INNO_STUBS = {"Setup.e32": "0c94d43e2c0f4d8371a2e1813b5446f3bc85ece3211dc4a6676159b534bbc483",
              "SetupCustomStyle.e32": "6848503904920ba8dc5e0d5af45091c081d1f03ac85d454db0ad31e522c92b31",
              "SetupLdr.e32": "998c58046b0808e75a030859962e1fe1ea600e8a5cf4646b09aadce77342c1a9",
              "SetupLdr.e64": "76238ce87139bd7d888b72b5694a4e1d5f3f73e8c9a62cc8a2a0e124e455e4f3"}
VI_ISL_SHA256 ="826bc509b612d96ccfec89b81956640a3b56b547d194404e5ff95ae1a85babf0"
ISL_EMAIL_OK = {"memecoder17@gmail.com"}                                 # e-mail người dịch trong Vietnamese.isl (Task 0 bước 4)
APP_GUID = "{E16EAF7B-C860-4636-BC5D-8CA9A1004C51}"
UNINST = r"Software\Microsoft\Windows\CurrentVersion\Uninstall"
HOME = Path(os.environ["USERPROFILE"])
CUSTOMER = ["acme client", "âcme client", "acmeclient"]    # Q-name: tên công ty/khách hàng CẤM trong gói (mọi dạng hoa/thường/dấu: xem _forms)
THIRD = ("app/python/", "app/vendor/")               # bên thứ ba: chỉ R-SECRET / R-PATH / R-NAME / luật tên chung
GENERATED = {"mcp.json", "unins000.exe", "unins000.dat"}   # sinh lúc cài (chứa đường dẫn cài) -> bỏ R-PATH, vẫn quét luật khác
UNINS = {"unins000.exe", "unins000.dat"}             # D1: mã của Inno (bên thứ ba) -> chỉ needle, không áp luật regex của "tệp của mình"
MAX_JSON = 32 * 1024 * 1024

R_TOKEN = [re.compile(p) for p in (
    rb"sk-ant-[A-Za-z0-9_\-]{8,}", rb"sk-[A-Za-z0-9]{16,}", rb"eyJ[A-Za-z0-9_\-]{10,}\.[A-Za-z0-9_\-]{10,}",
    rb'(?i)"(accessToken|refreshToken|api_?key|apiKey|primaryApiKey)"\s*:\s*"[^"]{8,}')]
R_EMAIL = re.compile(rb"[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}")
R_USERPATH = re.compile(rb"(?i)[a-z]:\\users\\[^\\\s\"]+\\")
R_DWGCODE = re.compile(rb"\bSV-?\d{3,}")             # Q-name: mã bản vẽ khách hàng
# D4: ảnh/biểu tượng là dữ liệu nhị phân nén -> chuỗi giống e-mail xuất hiện ngẫu nhiên; R-EMAIL KHÔNG áp cho *.png / *.ico (chỉ hai đuôi này,
# không nới luật nào khác). Bù lại: PNG chỉ được có các chunk dưới đây (không tEXt/iTXt/eXIf... mang tên/đường dẫn/GPS); mọi ảnh còn do người xem xét.
EMAIL_EXEMPT = (".png", ".ico")
PNG_CHUNKS = {b"IHDR", b"PLTE", b"IDAT", b"IEND", b"tRNS", b"sRGB", b"gAMA", b"pHYs", b"cHRM"}
FORBID_SEG_ANY = {".claude", ".agents", "skills", "khungten", "__pycache__", "cad-mcp-deepseek", "logs"}   # M1: nhật ký hiệu năng (perf.jsonl) không bao giờ vào gói
FORBID_SEG_OWN = {"tests", "docs_src"}
FORBID_FILE_ANY = re.compile(r"(?i)^(skill\.md|.*_helper\.py|e2p\.lsp|ask\.py|claude\.lsp|.*\.pyc|.*\.bak|.*\.log|goi_y\.txt|"
                             r"config\.json|\.credentials\.json|\.claude\.json.*|\.env|cad_claude_session\.txt|cad_claude_doc(_[0-9]+){0,2}\.txt|.*\.jsonl)$")
FORBID_FILE_OWN = re.compile(r"(?i)^(_.*|mcp\.json|.*\.(pem|pfx|p12|key))$")


class PackError(Exception):
    """Lỗi do chính pack_setup nêu (thông điệp đã được soạn, không chứa bí mật). Ngoại lệ khác chỉ in tên kiểu."""


def die(msg):
    print("LỖI: " + msg)
    sys.exit(1)


def sha256(p):
    return build_python.sha256(p)


def has_bom(p):
    """UTF-8 BOM (EF BB BF): .iss và truoc-khi-cai.txt bắt buộc có (công cụ ghi tệp dễ làm mất U+FEFF)."""
    return Path(p).read_bytes().startswith(b"\xef\xbb\xbf")


def fresh(d):
    d = Path(d)
    if d.parent != DIST:                              # chỉ xoá/tạo thư mục con trực tiếp của DIST
        die(f"từ chối thao tác ngoài {DIST}: {d}")
    if d.exists():
        shutil.rmtree(d)
    d.mkdir(parents=True)
    return d


def tree(d):
    d = Path(d)
    return {p.relative_to(d).as_posix(): sha256(p) for p in sorted(d.rglob("*")) if p.is_file()}


def run(args, timeout):
    r = subprocess.run([str(a) for a in args], capture_output=True, text=True, encoding="utf-8", errors="replace", timeout=timeout,
                       env=build_python.child_env(DOTNET_CLI_TELEMETRY_OPTOUT="1", DOTNET_NOLOGO="1",
                                                  DOTNET_GENERATE_ASPNET_CERTIFICATE="false"))
    if r.returncode:
        # đuôi output có thể chứa đường dẫn có tên user: ghi tệp cục bộ để tự đọc, KHÔNG in
        (DIST / "_last-tool-error.txt").write_text((r.stdout + r.stderr)[-6000:], encoding="utf-8")
        raise PackError(f"{Path(str(args[0])).name} lỗi mã {r.returncode} (chi tiết: dist\\_last-tool-error.txt)")
    return r.stdout


def reg_get(sub, name):
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, sub) as k:
            return winreg.QueryValueEx(k, name)[0]
    except OSError:
        return None


def reg_exists(sub):
    try:
        winreg.OpenKey(winreg.HKEY_CURRENT_USER, sub).Close()
        return True
    except OSError:
        return False


def real_state():
    """Trạng thái THẬT không được đổi: Home thật, cây bundle thật, khoá gỡ thật."""
    b = Path(os.environ["APPDATA"]) / "Autodesk" / "ApplicationPlugins" / "ClaudeCAD.bundle"
    return reg_get(r"Software\ClaudeCad", "Home"), (tree(b) if b.exists() else None), reg_exists(f"{UNINST}\\{APP_GUID}_is1")


# ---------- needle (chỉ trong bộ nhớ) ----------

class Needles:
    def __init__(self, secret, path, name, host=None, host_exact=()):
        # regex bytes đã biên dịch (hoặc None). host = COMPUTERNAME tách riêng: chỉ nó được miễn cho tệp bên thứ ba ĐÃ CHỨNG MINH
        # trùng byte với bản gốc (xem upstream_hashes / leaks)
        self.secret, self.path, self.name, self.host = secret, path, name, host
        self.host_exact = tuple(host_exact)   # dạng NGUYÊN CHỮ (đúng hoa/thường, + IN HOA) của COMPUTERNAME, UTF-8 và UTF-16LE: chỉ để đếm trong unins000.*


def _alt(items):
    items = sorted({i for i in items if i}, key=len, reverse=True)
    return re.compile(b"|".join(re.escape(i) for i in items)) if items else None


def _forms(s):
    """D3: needle được 'gấp' giống hệt khô cỏ (haystack = bytes.lower(), chỉ hạ chữ ASCII): với mỗi dạng chữ (nguyên, lower, upper, title,
    NFC/NFD) và mỗi mã hoá (UTF-8, UTF-16LE) lấy encode(...).lower(). Nhờ vậy 'ÂCME CLIENT' (Â = C3 82, không bị bytes.lower() hạ) vẫn khớp."""
    out = set()
    for v in (s, s.lower(), s.upper(), s.title()):
        for n in {v, unicodedata.normalize("NFC", v), unicodedata.normalize("NFD", v)}:
            out |= {n.encode("utf-8").lower(), n.encode("utf-16-le").lower()}
    return out


def build_needles(secrets, names, paths, short_names, host=()):
    sec = set().union(*[_forms(s) for s in secrets]) if secrets else set()
    pth = set()
    for s in list(paths) + [rf"\users\{n}" + "\\" for n in short_names]:
        for v in (s, s.replace("\\", "/"), s.replace("\\", "\\\\")):
            pth |= _forms(v)
    nm = set().union(*[_forms(s) for s in list(names) + CUSTOMER])
    hx = {v.encode(e) for h in host for v in {h, h.upper()} for e in ("utf-8", "utf-16-le")}
    return Needles(_alt(sec), _alt(pth), _alt(nm), _alt(set().union(*[_forms(s) for s in host])) if host else None, hx)


def host_exact(b, nd):
    """Số lần COMPUTERNAME xuất hiện NGUYÊN CHỮ (đúng hoa/thường hoặc IN HOA; UTF-8 + UTF-16LE) trong b."""
    return sum(b.count(x) for x in nd.host_exact)


def _json(p):
    try:
        p = Path(p)
        if p.stat().st_size > MAX_JSON:
            return None
        return json.loads(p.read_text(encoding="utf-8-sig"))
    except (OSError, ValueError):
        return None


def _strings(o):
    """Mọi chuỗi GIÁ TRỊ (không lấy tên khoá) trong o."""
    if isinstance(o, str):
        yield o
    elif isinstance(o, dict):
        for v in o.values():
            yield from _strings(v)
    elif isinstance(o, list):
        for v in o:
            yield from _strings(v)


def _under_keys(o, keys):
    """D2: mọi chuỗi nằm dưới một khoá có tên trong keys, ở BẤT KỲ độ sâu nào (vd mcpServers lồng trong projects.*)."""
    if isinstance(o, dict):
        for k, v in o.items():
            if k in keys:
                yield from _strings(v)
            else:
                yield from _under_keys(v, keys)
    elif isinstance(o, list):
        for v in o:
            yield from _under_keys(v, keys)


def _account(d):
    if not isinstance(d, dict):
        return []
    out = [d.get("primaryApiKey")]
    oa = d.get("oauthAccount")
    if isinstance(oa, dict):
        out += [oa.get(k) for k in ("emailAddress", "organizationName", "organizationUuid", "accountUuid", "displayName")]
    return [x for x in out if isinstance(x, str)]


def _os_names():
    out = [os.environ.get("USERNAME", ""), os.environ.get("COMPUTERNAME", "")]
    try:
        import win32api
        out.append(win32api.GetUserNameEx(3))
    except ImportError:                               # thiếu pywin32 -> needle họ tên đầy đủ sẽ trống: không được lặng lẽ bỏ qua
        raise PackError("thiếu pywin32 — không thu được tên đầy đủ để dò; chạy bằng python có pywin32 (không phải py)")
    except Exception:
        pass
    try:
        import win32net
        out.append(win32net.NetUserGetInfo(None, os.environ["USERNAME"], 2)["full_name"])
    except ImportError:
        raise PackError("thiếu pywin32 (win32net)")
    except Exception:
        pass
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"Software\Microsoft\Office\Common\UserInfo") as k:
            out.append(winreg.QueryValueEx(k, "UserName")[0])
    except OSError:
        pass
    return [x.strip() for x in out if isinstance(x, str) and x.strip()]


SOURCES = ("credentials", "accounts_cfg", "claude_json_account", "claude_json_mcp_env_headers", "settings_env_apikeyhelper",
           "deepseek", "environ", "session_keys")


def gather(home, config_dir, extra_dirs, deepseek_cfg, environ):
    """Thu chuỗi bí mật theo nguồn -> {nguồn: set(chuỗi)}. Chỉ đọc vào bộ nhớ.
    credentials: mọi chuỗi trong mọi .credentials.json dưới ~\\.claude* và $CLAUDE_CONFIG_DIR
    accounts_cfg: token >=16 ký tự trong .claude*\\**\\accounts\\*.cfg
    claude_json_account: primaryApiKey, oauthAccount.{emailAddress,organizationName,organizationUuid,accountUuid,displayName}
    claude_json_mcp_env_headers: mọi chuỗi dưới khoá mcpServers / env / headers trong mỗi .claude.json*
    settings_env_apikeyhelper: env + apiKeyHelper trong mỗi settings*.json dưới ~\\.claude* và extra_dirs (C:\\Autodesk\\cad-mcp\\.claude)
    session_keys: *.key dưới ~\\.claude* (gồm .claude-switch\\cfg\\*\\sessions\\*.key): toàn bộ nội dung đã strip + token >=16 ký tự
    deepseek: api_key; environ: ANTHROPIC_*, DEEPSEEK_*, CLAUDE_CODE_OAUTH_TOKEN.
    .credentials.json / *.key không đọc được, không phân tích được hoặc > 32 MB -> PackError (không âm thầm bỏ qua)."""
    src = {n: set() for n in SOURCES}
    dirs = [d for d in Path(home).glob(".claude*") if d.is_dir()]
    if config_dir and Path(config_dir).is_dir():
        dirs.append(Path(config_dir))
    bad = 0                                          # R3/R4: tệp bí mật không dùng được -> dừng, không fail-open
    for d in dirs:
        for p in d.rglob(".credentials.json"):
            j = _json(p)
            if j is None:
                bad += 1
            src["credentials"] |= set(_strings(j))
        for p in d.rglob("*.key"):
            try:
                if p.stat().st_size > MAX_JSON:
                    raise OSError
                txt = p.read_text(encoding="utf-8", errors="ignore").strip()
            except OSError:
                bad += 1
                continue
            if txt:
                src["session_keys"].add(txt)
                src["session_keys"] |= set(re.findall(r"[A-Za-z0-9+/=_.\-]{16,}", txt))
        for p in d.rglob("*.cfg"):
            if p.parent.name.lower() == "accounts":
                try:
                    for tok in re.findall(r"[A-Za-z0-9+/=_.\-]{16,}", p.read_text(encoding="utf-8", errors="ignore")):
                        src["accounts_cfg"].add(tok)
                        if "=" in tok:               # "KHOÁ=giá trị": thêm cả phần giá trị đứng riêng (token nằm trong tệp khác không kèm "KHOÁ=")
                            src["accounts_cfg"].add(tok.split("=", 1)[1])
                except OSError:
                    pass
    cj = [p for d in dirs for p in d.rglob(".claude.json*")] + list(Path(home).glob(".claude.json*"))
    for p in cj:
        j = _json(p)
        src["claude_json_account"] |= set(_account(j))
        src["claude_json_mcp_env_headers"] |= set(_under_keys(j, {"mcpServers", "env", "headers"}))
    for d in dirs + [Path(x) for x in extra_dirs if Path(x).is_dir()]:
        for p in d.rglob("settings*.json"):
            j = _json(p)
            if isinstance(j, dict):
                src["settings_env_apikeyhelper"] |= set(_strings(j.get("env")))
                if isinstance(j.get("apiKeyHelper"), str):
                    src["settings_env_apikeyhelper"].add(j["apiKeyHelper"])
    ds = _json(deepseek_cfg) if deepseek_cfg else None
    if isinstance(ds, dict) and isinstance(ds.get("api_key"), str):
        src["deepseek"].add(ds["api_key"])
    src["environ"] |= {v for k, v in environ.items() if re.match(r"(?i)^(ANTHROPIC_|DEEPSEEK_)|^CLAUDE_CODE_OAUTH_TOKEN$", k)}
    if bad:
        raise PackError(f"{bad} tệp .credentials.json/*.key không đọc được, không phân tích được hoặc > 32 MB — sửa/xoá rồi chạy lại")
    return src


def _usable(vals):
    return {s.strip() for s in vals if isinstance(s, str) and len(s.strip()) >= 8}   # < 8 ký tự: bỏ (khớp vô nghĩa)


# Định danh sản phẩm công khai, có sẵn trong chính tệp phát hành của mình (vd URL trong claude-cli.ps1) -> không phải bí mật (S2, spec §5.3).
# So NGUYÊN chuỗi (đã strip, hạ chữ), KHÔNG so chuỗi con. ponytail: danh sách duyệt tay; thêm mục mới phải qua rà bảo mật (fable-expert).
PUBLIC_IDS = {"claude-code"}


def _usable(vals):
    return {s.strip() for s in vals if isinstance(s, str) and len(s.strip()) >= 8}   # < 8 ký tự: bỏ (khớp vô nghĩa)


def _secretish(vals):
    return {s for s in _usable(vals) if s.lower() not in PUBLIC_IDS}


def collect_needles(src=None):
    """Needle thật của máy build (src: chỉ để selftest tiêm nguồn giả). KHÔNG BAO GIỜ in/ghi/chép giá trị; chỉ in SỐ chuỗi mỗi nguồn.
    D2: không thu được bí mật nào từ .credentials.json -> dừng (tập rỗng không được âm thầm vô hiệu R-SECRET)."""
    if src is None:
        src = gather(HOME, os.environ.get("CLAUDE_CONFIG_DIR"), [ROOT / ".claude"], r"C:\Autodesk\cad-mcp-deepseek\config.json", os.environ)
    for n in SOURCES:
        print(f"needle nguồn {n}: {len(_usable(src[n]))}")
    if not _secretish(src["credentials"]):           # tính SAU khi lọc PUBLIC_IDS
        raise PackError("không thu được bí mật nào từ .credentials.json — R-SECRET sẽ vô hiệu; đăng nhập Claude trên máy build rồi chạy lại")
    allv = set().union(*src.values())
    keep = _secretish(allv)
    print(f"needle loại (PUBLIC_IDS): {len(_usable(allv)) - len(keep)}")     # chỉ in SỐ
    host = os.environ.get("COMPUTERNAME", "").strip()
    names = [n for n in _os_names() if n.lower() != host.lower()]
    print(f"needle tên: {len(names)}")                # chỉ in SỐ
    long_names = [n for n in names if len(n) >= 6]
    short_names = [n for n in names if len(n) < 6] + ([host] if host and len(host) < 6 else [])
    nd = build_needles(keep, long_names, [os.environ["USERPROFILE"], r"C:\Autodesk"], short_names,
                       host=[host] if len(host) >= 6 else ())
    return nd, len({s.strip() for s in allv if isinstance(s, str)}) - len(keep) + len(short_names)


# ---------- chứng minh tệp bên thứ ba trùng byte với bản gốc ----------

# Tệp do chính pack/pip sinh ra, không có trong zip/wheel gốc (quy ước đã duyệt). Ngoài tập này mà không trùng gốc -> hỏng.
GEN_THIRD = re.compile(r"app/python/python313\._pth|app/python/Lib/site-packages/win32com/gen_py/__init__\.py|"
                       r"app/python/Lib/site-packages/[^/]+\.dist-info/(RECORD|INSTALLER|REQUESTED)")


def upstream_hashes():
    """SHA256 của mọi tệp trong zip Python nhúng (đã khớp PY_SHA256) và trong các wheel mà hash của CHÍNH wheel nằm trong lock
    (requirements-embed.lock + requirements-xls.txt) ngay lúc này. Thiếu/sai -> PackError."""
    z = build_python.CACHE / f"python-{build_python.PY_VER}-embed-amd64.zip"
    if not z.is_file() or sha256(z) != build_python.PY_SHA256:
        raise PackError("zip Python nhúng thiếu hoặc SHA256 khác PY_SHA256")
    out = set()
    with zipfile.ZipFile(z) as f:
        out |= {hashlib.sha256(f.read(i)).hexdigest() for i in f.infolist() if not i.is_dir()}
    want = set(re.findall(r"--hash=sha256:([0-9a-f]{64})", build_python.LOCK.read_text(encoding="utf-8") + build_python.XLS_REQ.read_text(encoding="utf-8")))
    seen = set()
    for w in sorted(build_python.WHEELS.glob("*.whl")):
        h = sha256(w)
        if h in want:
            seen.add(h)
            with zipfile.ZipFile(w) as f:
                out |= {hashlib.sha256(f.read(i)).hexdigest() for i in f.infolist() if not i.is_dir()}
    if seen != want:
        raise PackError(f"wheel không khớp lock: thiếu {len(want - seen)}/{len(want)}")
    return out


def third_party_proof(root, upstream, prefix=""):
    """(số tệp bên thứ ba, danh sách tệp KHÔNG trùng bản gốc). Chạy trên MỌI tệp dưới app/python và app/vendor. prefix: "" khi root là staging, "app/" khi root là cây {app} đã cài."""
    n, non = 0, []
    for p in sorted(Path(root).rglob("*")):
        if p.is_file():
            rel = prefix + p.relative_to(root).as_posix()
            if rel.startswith(THIRD):
                n += 1
                if hashlib.sha256(p.read_bytes()).hexdigest() not in upstream:
                    non.append(rel)
    return n, non


def report_proof(root, upstream, prefix=""):
    n, non = third_party_proof(root, upstream, prefix)
    meta = [x for x in non if x.rsplit("/", 1)[-1] in ("RECORD", "INSTALLER", "REQUESTED")]      # gom cho gọn; tệp còn lại in tên
    print(f"bên thứ ba: {n} tệp, trùng bản gốc {n - len(non)}, KHÔNG trùng {len(non)} (dist-info RECORD/INSTALLER/REQUESTED: {len(meta)}; "
          f"khác: " + (", ".join(x for x in non if x not in meta) or "không") + ")")
    odd = [x for x in non if not GEN_THIRD.fullmatch(x)]
    if odd:
        raise PackError(f"tệp bên thứ ba lạ (không trùng bản gốc, ngoài tập sinh sẵn): {len(odd)}: " + ", ".join(odd[:10]))


# ---------- quét ----------

def _variants(b):
    low = b.lower()
    return low, low.replace(b"\0", b"")


def png_clean(b):
    """D4: PNG hợp lệ, bắt đầu IHDR, kết thúc IEND, không dữ liệu thừa, và chỉ gồm các chunk trong PNG_CHUNKS."""
    if b[:8] != b"\x89PNG\r\n\x1a\n":
        return False
    i, first, last = 8, None, None
    while i + 12 <= len(b):
        n, t = int.from_bytes(b[i:i + 4], "big"), b[i + 4:i + 8]
        if t not in PNG_CHUNKS or i + 12 + n > len(b):
            return False
        first = first or t
        last = t
        i += 12 + n
        if t == b"IEND":
            break
    return first == b"IHDR" and last == b"IEND" and i == len(b)


def leaks(root, nd, extra=(), installed=False, prefix="", upstream=frozenset(), exempt=None):
    """Danh sách 'tệp: luật'. extra = (đường dẫn tuyệt đối, là_isl) cho .iss/.isl.
    installed=True: cây đã cài (cho phép tệp sinh lúc cài). prefix (D1): 'app/' hoặc 'bundle/' — đặt lại tiền tố đường dẫn tương đối để
    luật bên thứ ba (THIRD) và GENERATED so khớp đúng khi quét thư mục con riêng (cây {app} đã cài không có tiền tố 'app/').
    upstream: tập SHA256 tệp gốc (upstream_hashes). CHỈ needle COMPUTERNAME (nd.host) được miễn, CHỈ cho tệp dưới THIRD có SHA256 thuộc upstream
    (tên máy có thể là từ tiếng Anh thường gặp trong thư viện); needle R-SECRET (S1: lấy mọi chuỗi từ cấu hình Claude, có thể trùng chữ trong thư viện) được miễn
    theo ĐÚNG điều kiện đó, không miễn unins000.*; exempt (list, tuỳ chọn) nhận đường dẫn các tệp được miễn R-SECRET; mọi tệp khác (của mình, DLL, .iss/.isl, mcp.json, tệp bên thứ ba bị sửa) dùng đủ needle."""
    root = Path(root)
    items = [(p, prefix + p.relative_to(root).as_posix(), p.relative_to(root).as_posix(), False) for p in sorted(root.rglob("*")) if p.is_file()]
    items += [(Path(p), Path(p).name, Path(p).name, isl) for p, isl in extra]
    bad = []
    for p, rel, sub, isl in items:
        b, why = p.read_bytes(), []
        low, nul = _variants(b)
        raw_nul = b.replace(b"\0", b"")
        gen = installed and sub in GENERATED
        third = rel.startswith(THIRD) or (installed and sub in UNINS)
        if nd.secret and (nd.secret.search(low) or nd.secret.search(nul)):
            if rel.startswith(THIRD) and hashlib.sha256(b).hexdigest() in upstream:   # S1 (spec §5.3): trùng byte bản gốc đã ghim hash -> không thể mang bí mật máy build
                if exempt is not None:
                    exempt.append(rel)
            else:
                why.append("R-SECRET")
        if not gen and nd.path and (nd.path.search(low) or nd.path.search(nul)):
            why.append("R-PATH")
        if nd.name and (nd.name.search(low) or nd.name.search(nul)):
            why.append("R-NAME")
        if nd.host and (nd.host.search(low) or nd.host.search(nul)):
            ok = rel.startswith(THIRD) and hashlib.sha256(b).hexdigest() in upstream
            if not ok and installed and sub in UNINS:
                # unins000.* do Inno sinh/ chép lúc cài, chỉ xét ở cây ĐÃ CÀI. .exe = stub Inno: từ thường gặp (hạ chữ) được miễn, khớp NGUYÊN CHỮ thì hỏng.
                # .dat: Inno ghi tên máy cài đúng 1 lần nguyên chữ -> miễn khi số khớp nguyên chữ == 1, số khác (0, 2, ...) -> hỏng (fail closed).
                n = host_exact(b, nd)
                ok = (n == 0) if sub == "unins000.exe" else (n == 1)
            if not ok:
                why.append("R-NAME")
        if rel.lower().endswith(".png") and not png_clean(b):
            why.append("R-PNG")
        if not third:
            if R_DWGCODE.search(b) or R_DWGCODE.search(raw_nul):
                why.append("R-NAME")
            if any(r.search(b) or r.search(raw_nul) for r in R_TOKEN):
                why.append("R-TOKEN")
            if not rel.lower().endswith(EMAIL_EXEMPT):
                mails = {m.lower() for m in R_EMAIL.findall(b) + R_EMAIL.findall(raw_nul)}
                if isl:
                    mails -= {e.lower().encode() for e in ISL_EMAIL_OK}
                if mails:
                    why.append("R-EMAIL")
            if R_USERPATH.search(b) or R_USERPATH.search(raw_nul):
                why.append("R-USERPATH")
        segs = [s.lower() for s in rel.split("/")]
        name = segs[-1]
        if any(s in FORBID_SEG_ANY for s in segs[:-1]) or (not third and any(s in FORBID_SEG_OWN for s in segs[:-1])):
            why.append("R-NAMEFILE")
        elif not (gen and name in GENERATED) and (FORBID_FILE_ANY.match(name) or (not third and FORBID_FILE_OWN.match(name))):
            why.append("R-NAMEFILE")
        if why:
            bad.append(f"{rel}: {', '.join(sorted(set(why)))}")
    return bad


def _print_exempt(where, ex):
    """S1: in SỐ và đường dẫn tệp bên thứ ba được miễn R-SECRET (trùng byte bản gốc). Không in chuỗi khớp."""
    print(f"R-SECRET miễn (bên thứ ba trùng bản gốc) {where}: {len(ex)} tệp" + "".join("\n  " + x for x in ex))


# ---------- staging ----------

def own_files():
    html = (INST / "docs" / "HDSD.html").read_text(encoding="utf-8")
    srcs = re.findall(r'src="([^"]*)"', html)
    wrong = [s for s in srcs if not re.fullmatch(r"img/[A-Za-z0-9_\-]+\.png", s)]
    if wrong:
        die(f"HDSD.html có src không hợp lệ: {wrong}")
    out = INST / "bin" / "Release"
    f = [(ROOT / "server.py", "app/server.py"),
         (ROOT / "xls2tsv.py", "app/xls2tsv.py"),
         (ROOT / "app_icon.ico", "app/app_icon.ico"),
         (INST / "docs" / "HDSD.html", "app/docs/HDSD.html"),
         (ROOT / "PackageContents.xml", "bundle/PackageContents.xml"),
         (out / "net46" / "ClaudePalette.dll", "bundle/Contents/net46/ClaudePalette.dll"),
         (out / "net8.0-windows" / "ClaudePalette.dll", "bundle/Contents/net8/ClaudePalette.dll"),
         (out / "net10.0-windows" / "ClaudePalette.dll", "bundle/Contents/net10/ClaudePalette.dll"),
         (INST / "docs" / "truoc-khi-cai.txt", "setup/truoc-khi-cai.txt"),
         (INST / "claude-cli.ps1", "setup/claude-cli.ps1")]
    f += [(ROOT / "app_icon.ico", f"bundle/Contents/{d}icon.ico") for d in ("", "net46/", "net8/", "net10/")]
    f += [(INST / "docs" / s, f"app/docs/{s}") for s in sorted(set(srcs))]
    for s, _ in f:                                    # U-source: mọi nguồn nằm dưới C:\Autodesk\cad-mcp
        if not s.resolve().is_relative_to(ROOT):
            die(f"nguồn ngoài {ROOT}: {s}")
    return f


def stage_all(stage):
    own = own_files()
    for s, rel in own:
        (stage / rel).parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(s, stage / rel)
    build_python.build(stage / "app" / "python")
    build_python.build_vendor(stage / "app" / "vendor")
    got = {k for k in tree(stage) if not k.startswith(THIRD)}
    want = {rel for _, rel in own}
    if got != want:                                   # U-skill: không tệp nào ngoài allowlist
        raise PackError(f"staging khác allowlist: thừa {sorted(got - want)}, thiếu {sorted(want - got)}")


# ---------- kiểm sau đóng gói ----------

def verify(exe, vroot, want_app, want_bundle, snap0, nd, fail, upstream):
    test = fresh(vroot) / "thử nghiệm"               # có dấu + dấu cách: thử luôn mã hoá mcp.json
    test.mkdir()
    (test / ".cad-mcp-testroot").write_text("", encoding="ascii")
    (test / "sentinel.txt").write_text("gỡ cài đặt không được xoá tệp này", encoding="utf-8")
    app = test / "cad-mcp"
    bundle = test / "ApplicationPlugins" / "ClaudeCAD.bundle"
    tkey, tuninst = r"Software\ClaudeCad-Test", f"{UNINST}\\{APP_GUID}-TEST_is1"
    r = subprocess.run([str(exe), "/VERYSILENT", "/SUPPRESSMSGBOXES", "/NORESTART", "/SP-", "/ACCEPTRISK=yes",
                        f"/TESTROOT={test}", f"/DIR={app}", f"/LOG={vroot / 'install.log'}"], timeout=900, env=build_python.child_env())
    if r.returncode:
        fail.append(f"cài thử lỗi mã {r.returncode}")
        return
    got = tree(app)
    if set(got) != set(want_app) | GENERATED:
        fail.append(f"tệp cài khác staging: thừa {sorted(set(got) - set(want_app) - GENERATED)[:10]}, "
                    f"thiếu {sorted((set(want_app) | GENERATED) - set(got))[:10]}")
    elif any(got[k] != v for k, v in want_app.items()):
        fail.append("nội dung tệp cài khác staging")
    if (tree(bundle) if bundle.exists() else {}) != want_bundle:
        fail.append("bundle cài khác staging")
    if reg_get(tkey, "Home") != str(app):
        fail.append(r"HKCU\Software\ClaudeCad-Test\Home sai")
    if not reg_exists(tuninst):
        fail.append("thiếu khoá gỡ -TEST_is1")
    try:
        cfg = json.loads((app / "mcp.json").read_bytes().decode("ascii"))["mcpServers"]["autocad"]
    except Exception as e:
        cfg = f"không đọc được: {type(e).__name__}"
    if cfg != {"command": str(app / "python" / "python.exe"), "args": ["-I", str(app / "server.py")]}:
        fail.append("mcp.json sai")
    # U-secret 3: quét bản ĐÃ CÀI, trước smoke. D1: prefix để luật bên thứ ba/GENERATED áp đúng
    for u in sorted(UNINS):
        if (app / u).is_file():
            print(f"{u}: COMPUTERNAME nguyên chữ {host_exact((app / u).read_bytes(), nd)} lần")
    ex = []
    fail += ["cây đã cài lộ: " + x for x in leaks(app, nd, installed=True, prefix="app/", upstream=upstream, exempt=ex)]
    _print_exempt("cây đã cài", ex)
    fail += ["bundle đã cài lộ: " + x for x in leaks(bundle, nd, prefix="bundle/")]
    try:
        smoke_embed.check(app)
    except Exception as e:
        fail.append(f"smoke: {type(e).__name__}")
    (app / "goi_y.txt").write_text("; palette tự tạo", encoding="utf-8")
    subprocess.run([str(app / "unins000.exe"), "/VERYSILENT", "/SUPPRESSMSGBOXES", "/NORESTART"], timeout=300, env=build_python.child_env())
    for _ in range(240):                              # unins000 tự chép sang %TEMP% rồi thoát -> chờ
        if (app.exists() and set(tree(app)) == {"goi_y.txt"} and not bundle.exists()
                and not reg_exists(tkey) and not reg_exists(tuninst)):   # khoá -Test chỉ bị xoá ở usPostUninstall, SAU khi xoá tệp (spec §5.3.1)
            break
        time.sleep(0.5)
    left = set(tree(app)) if app.exists() else set()
    if left != {"goi_y.txt"}:
        fail.append(f"sau gỡ phải chỉ còn goi_y.txt, có: {sorted(left)[:10]}")
    if bundle.exists():
        fail.append("gỡ còn sót bundle thử")
    if not (test / "sentinel.txt").is_file():
        fail.append("GỠ ĐÃ XOÁ TỆP NGOÀI THƯ MỤC CÀI")
    if reg_exists(tkey):
        fail.append(r"gỡ không xoá Software\ClaudeCad-Test")
    if reg_exists(tuninst):
        fail.append("gỡ không xoá khoá -TEST_is1")
    if real_state() != snap0:
        fail.append("TRẠNG THÁI THẬT BỊ ĐỔI (Home thật / bundle thật / khoá gỡ thật)")


# ---------- tự kiểm bộ quét ----------

def _png(*chunks):
    import struct, zlib
    out = b"\x89PNG\r\n\x1a\n"
    for t, d in chunks:
        out += struct.pack(">I", len(d)) + t + d + struct.pack(">I", zlib.crc32(t + d) & 0xFFFFFFFF)
    return out


def _png_ok(extra=()):
    import struct, zlib
    ihdr = struct.pack(">IIBBBBB", 1, 1, 8, 2, 0, 0, 0)
    return _png((b"IHDR", ihdr), *extra, (b"IDAT", zlib.compress(b"\x00\x01\x02\x03")), (b"IEND", b""))


def selftest():
    global ISL_EMAIL_OK
    DIST.mkdir(parents=True, exist_ok=True)
    t = fresh(DIST / "_selftest-setup")
    checks = []                                       # (tên ca, đạt?) — D1/D2/D4 ngoài bảng quét cây
    try:
        sec = "S3cretTok_ABCDEFGH1234"
        nd = build_needles({sec}, ["DESKTOP-TEST9"], [r"C:\Users\HOME", r"C:\Autodesk"], ["home"])
        files = {
            "app/ok.txt": "bình thường, Home, home".encode(),
            "app/sec.txt": ("x " + sec).encode(),                         # R-SECRET
            "app/sec16.dll": ("x" + sec).encode("utf-16-le"),             # R-SECRET qua UTF-16LE
            "app/python/third.py": sec.encode(),                          # R-SECRET cả ở bên thứ ba
            "app/p1.txt": rb"C:\Users\HOME\x",                            # R-PATH (+ R-USERPATH)
            "app/p2.txt": b"c:/autodesk/x",                               # R-PATH dạng /
            "app/python/p3.txt": rb"d:\users\home\y",                     # R-PATH tên ngắn dạng đường dẫn
            "app/n1.txt": "máy DESKTOP-TEST9".encode(),                    # R-NAME
            "app/n2.txt": "Công ty Acme Client".encode(),                    # R-NAME khách hàng
            "app/n3.txt": b"ban ve SV-00123",                             # R-NAME mã bản vẽ
            "app/n4.txt": b"ban ve SV-123",                               # R-NAME mã bản vẽ 3 chữ số (Q-name)
            "app/n5.txt": b"ban ve SV456 x",                              # R-NAME mã bản vẽ không gạch (Q-name)
            "app/n6.txt": b"ACME CLIENT",                                   # R-NAME dạng không dấu IN HOA
            "app/n7.txt": "CÔNG TY ÂCME CLIENT".encode(),                    # R-NAME IN HOA có dấu UTF-8 (D3)
            "app/n8.dll": "CÔNG TY ÂCME CLIENT".encode("utf-16-le"),         # R-NAME IN HOA có dấu UTF-16LE (D3)
            "app/n9.txt": "Âcme Client".encode(),                            # R-NAME dạng Title
            "app/n10.txt": "DESKTOP-test9 và desktop-TEST9".encode(),     # R-NAME tên máy lẫn hoa/thường
            "app/ok2.txt": b"SV-12 va SVG123 va XSV-1234x",               # KHÔNG hit: 2 chữ số / chữ thừa / không có \b
            "app/t1.txt": b"sk-ant-oat01-" + b"A" * 40,                   # R-TOKEN
            "app/t2.txt": b"sk-" + b"b" * 20,                             # R-TOKEN
            "app/t3.txt": b"eyJhbGciOiJIUzI1.eyJzdWIiOiIxMjM0",            # R-TOKEN JWT
            "app/t4.txt": b'{"accessToken": "abcdefghij"}',               # R-TOKEN json
            "app/t5.dll": "sk-ant-api03-ZZZZZZZZZZ".encode("utf-16-le"),  # R-TOKEN qua NUL-strip
            "app/e1.txt": b"lien he ai.do@example.com",                   # R-EMAIL
            "app/e2.dll": "ai.do@example.com".encode("utf-16-le"),        # R-EMAIL qua NUL-strip
            "app/img/ok.png": _png_ok(),                                  # ảnh sạch: không hit
           "app/img/text.png": _png_ok([(b"tEXt", b"Author\0someone")]),  # R-PNG: chunk tEXt
            "app/img/exif.png": _png_ok([(b"eXIf", b"MM\0*")]),            # R-PNG: chunk eXIf
            "app/img/trail.png": _png_ok() + b"tail",                      # R-PNG: dữ liệu thừa sau IEND
            "app/img/bad.png": b"not a png",                               # R-PNG: không phải PNG
            "app/ico/mail.ico": b"\0\0\1\0" + b"zz ai.do@example.com zz",   # .ico: miễn R-EMAIL
            "app/u1.txt": rb"D:\Users\bob\x",                             # R-USERPATH
            "app/SKILL.md": b"x",                                         # R-NAMEFILE
            "app/x_helper.py": b"x",
            "app/.claude/a.txt": b"x",
            "app/skills/b.txt": b"x",
            "app/KhungTen/c.txt": b"x",
            "app/E2P.lsp": b"x",
            "app/_scratch.py": b"x",
            "app/a.bak": b"x",
            "app/tests/t.py": b"x",
            "app/__pycache__/m.pyc": b"x",
            "app/goi_y.txt": b"x",
            "app/logs/perf.txt": b"x",                                     # M1: thư mục logs
            "app/perf.jsonl": b"x",                                        # M1: *.jsonl
            "app/data/perf.1.jsonl": b"x",
            "app/cad_claude_doc_1234.txt": b"x",                           # tệp ghim bản vẽ theo pid
            "app/cad_claude_doc_1234_2.txt": b"x",                         # 3.1: tệp ghim theo pid + số thứ tự
            "app/python/logs/x.txt": b"x",                                 # M1: cả bên thứ ba
            "app/mcp.json": b"{}",
            "app/.credentials.json": b"{}",
            "app/python/jsonschema/tests/t.py": b"x",                     # bên thứ ba: tests hợp lệ
            "app/python/cacert.pem": b"-----BEGIN CERTIFICATE-----",      # bên thứ ba: .pem hợp lệ
            "app/vendor/xlrd-2.0.2.dist-info/METADATA": b"Author-email: a@b.org",   # bên thứ ba: e-mail hợp lệ
        }
        # PNG có chuỗi giống e-mail nằm trong dữ liệu IDAT (không phải chunk lạ): R-EMAIL không áp cho .png
        import struct, zlib
        files["app/img/mail.png"] = _png((b"IHDR", struct.pack(">IIBBBBB", 1, 1, 8, 2, 0, 0, 0)), (b"IDAT", b"zz ai.do@example.com zz"), (b"IEND", b""))
        for rel, data in files.items():
            (t / rel).parent.mkdir(parents=True, exist_ok=True)
            (t / rel).write_bytes(data)
        isl = DIST / "selftest-setup.isl"            # NGOÀI thư mục quét t
        isl.write_bytes(b"; nguoi dich: dich.gia@example.org\n[Messages]\n")
        saved, ISL_EMAIL_OK = ISL_EMAIL_OK, {"dich.gia@example.org"}
        try:
            got = {x.split(": ")[0]: x.split(": ")[1] for x in leaks(t, nd, extra=[(isl, True)])}
        finally:
            ISL_EMAIL_OK = saved
        want = {"app/sec.txt": "R-SECRET", "app/sec16.dll": "R-SECRET", "app/python/third.py": "R-SECRET",
                "app/p1.txt": "R-PATH, R-USERPATH", "app/p2.txt": "R-PATH", "app/python/p3.txt": "R-PATH",
                **{f"app/n{i}.{x}": "R-NAME" for i, x in ((1, "txt"), (2, "txt"), (3, "txt"), (4, "txt"), (5, "txt"), (6, "txt"),
                                                         (7, "txt"), (8, "dll"), (9, "txt"), (10, "txt"))},
                "app/t1.txt": "R-TOKEN", "app/t2.txt": "R-TOKEN", "app/t3.txt": "R-TOKEN", "app/t4.txt": "R-TOKEN",
                "app/t5.dll": "R-TOKEN", "app/e1.txt": "R-EMAIL", "app/e2.dll": "R-EMAIL", "app/u1.txt": "R-USERPATH",
                **{k: "R-PNG" for k in ("app/img/text.png", "app/img/exif.png", "app/img/trail.png", "app/img/bad.png")},
                **{k: "R-NAMEFILE" for k in ("app/SKILL.md", "app/x_helper.py", "app/.claude/a.txt", "app/skills/b.txt",
                                             "app/KhungTen/c.txt", "app/E2P.lsp", "app/_scratch.py", "app/a.bak", "app/tests/t.py",
                                             "app/__pycache__/m.pyc", "app/goi_y.txt", "app/mcp.json", "app/.credentials.json",
                                             "app/logs/perf.txt", "app/perf.jsonl", "app/data/perf.1.jsonl",
                                             "app/cad_claude_doc_1234.txt", "app/cad_claude_doc_1234_2.txt", "app/python/logs/x.txt")}}
        for k in sorted(set(got) | set(want)):
            checks.append((f"quét {k}", got.get(k) == want.get(k)))

        # D1: cây {app} ĐÃ CÀI (không có tiền tố 'app/') quét với prefix="app/": bên thứ ba + tệp sinh lúc cài không bị coi là tệp của mình
        inst = t.parent / "_selftest-setup-inst"
        shutil.rmtree(inst, ignore_errors=True)
        try:
            lay = {"python/Lib/site-packages/x/__init__.py": b"x = 1  # a@b.org",
                   "python/Lib/site-packages/x/_priv.py": b"x",
                   "python/Lib/site-packages/y/tests/t.py": b"x",
                   "mcp.json": rb'{"command": "C:\\Autodesk\\x\\python.exe"}',
                   "unins000.exe": b"MZ sk-" + b"q" * 20 + b" a@b.org",           # bên thứ ba: luật regex không áp
                   "unins000.dat": b"a@b.org " + rb"c:\users\bob\x\x",
                   "server.py": b"print(1)"}
            for rel, data in lay.items():
                (inst / rel).parent.mkdir(parents=True, exist_ok=True)
                (inst / rel).write_bytes(data)
            checks.append(("D1 cây đã cài, prefix=app/ -> 0 hit", leaks(inst, nd, installed=True, prefix="app/") == []))
            checks.append(("D1 không prefix -> tệp bên thứ ba bị coi là của mình (chứng minh cần prefix)",
                           leaks(inst, nd, installed=True) != []))
            (inst / "unins000.dat").write_bytes(sec.encode())                       # needle vẫn áp cho unins000
            checks.append(("D1 unins000 vẫn bị needle", leaks(inst, nd, installed=True, prefix="app/") == ["app/unins000.dat: R-SECRET"]))
            (inst / "unins000.dat").write_bytes(b"x")
            (inst / "python" / "Lib" / "site-packages" / "x" / "__init__.py").write_bytes(sec.encode())
            checks.append(("D1 third-party vẫn bị needle", leaks(inst, nd, installed=True, prefix="app/")
                           == ["app/python/Lib/site-packages/x/__init__.py: R-SECRET"]))
            (inst / "python" / "Lib" / "site-packages" / "x" / "__init__.py").write_bytes(b"x")
            (inst / "extra.txt").write_bytes(b"a@b.org")                             # tệp của mình trong cây đã cài: vẫn bị quét
            checks.append(("D1 tệp của mình trong cây cài vẫn bị quét", leaks(inst, nd, installed=True, prefix="app/") == ["app/extra.txt: R-EMAIL"]))
            br = t.parent / "_selftest-setup-bun"
            shutil.rmtree(br, ignore_errors=True)
            (br / "Contents" / "net46").mkdir(parents=True)
            (br / "Contents" / "net46" / "ClaudePalette.dll").write_bytes("a@b.org".encode("utf-16-le"))
            checks.append(("D1 bundle prefix=bundle/", leaks(br, nd, prefix="bundle/") == ["bundle/Contents/net46/ClaudePalette.dll: R-EMAIL"]))
            shutil.rmtree(br, ignore_errors=True)
        finally:
            shutil.rmtree(inst, ignore_errors=True)

        # D2: gather() — mọi nguồn, đếm theo nguồn, và dừng khi không có bí mật từ .credentials.json
        fh = t.parent / "_selftest-setup-home"
        shutil.rmtree(fh, ignore_errors=True)
        try:
            (fh / ".claude").mkdir(parents=True)
            (fh / ".claude-switch" / "cfg" / "acc-1").mkdir(parents=True)
            (fh / ".claude-switch" / "accounts").mkdir(parents=True)
            (fh / ".claude-switch" / "cfg" / "acc-1" / "sessions").mkdir(parents=True)
            dev = fh / "dev" / ".claude"
            dev.mkdir(parents=True)
            (fh / ".claude" / ".credentials.json").write_text(json.dumps({"claudeAiOauth": {"accessToken": "CRED-AAAA-111111", "n": 5}}), "utf-8")
            (fh / ".claude-switch" / "cfg" / "acc-1" / ".credentials.json").write_text(json.dumps({"k": "CRED-BBBB-222222"}), "utf-8")
            (fh / ".claude-switch" / "cfg" / "acc-1" / "sessions" / "s1.key").write_text("  SESSKEY-RRRR-777777 còn lại có dấu cách\n", "utf-8")
            (fh / ".claude-switch" / "accounts" / "acc-1.cfg").write_text("TOKEN=ACCT-CCCC-333333-XYZ0\n", "utf-8")
            (fh / ".claude.json").write_text(json.dumps({
                "primaryApiKey": "PRIM-DDDD-444444",
                "oauthAccount": {"emailAddress": "u@ex.test", "organizationName": "Org Ten Dai", "organizationUuid": "ORGU-EEEE-555555", "accountUuid": "ACCU-FFFF-666666"},
                "mcpServers": {"s": {"env": {"K": "MCPENV-GGGG-777777"}, "headers": {"Authorization": "HDRS-HHHH-888888"}}},
                "projects": {"p": {"mcpServers": {"q": {"args": ["ARGS-IIII-999999"], "env": {"Z": "NESTENV-JJJJ-000000"}}}}}}), "utf-8")
            (fh / ".claude" / "settings.json").write_text(json.dumps({"env": {"E": "SETENV-KKKK-111111"}, "apiKeyHelper": "HELPER-LLLL-222222", "other": "NOTNEEDLE-MMMM"}), "utf-8")
            (dev / "settings.local.json").write_text(json.dumps({"env": {"E": "DEVENV-NNNN-333333"}}), "utf-8")
            ds = fh / "ds.json"
            ds.write_text(json.dumps({"api_key": "DEEP-OOOO-444444"}), "utf-8")
            src = gather(fh, None, [dev], ds, {"ANTHROPIC_API_KEY": "ENVV-PPPP-555555", "PATH": "NOTNEEDLE-QQQQ"})
            allv = set().union(*src.values())
            for need in ("CRED-AAAA-111111", "CRED-BBBB-222222", "PRIM-DDDD-444444", "u@ex.test", "Org Ten Dai", "ORGU-EEEE-555555", "ACCU-FFFF-666666",
                         "MCPENV-GGGG-777777", "HDRS-HHHH-888888", "ARGS-IIII-999999", "NESTENV-JJJJ-000000", "SETENV-KKKK-111111",
                         "HELPER-LLLL-222222", "DEVENV-NNNN-333333", "DEEP-OOOO-444444", "ENVV-PPPP-555555"):
                checks.append((f"D2 nguồn có '{need[:4]}…'", need in allv))
            checks.append(("D2 token accounts\\*.cfg (cả phần sau KHOÁ=)", "ACCT-CCCC-333333-XYZ0" in src["accounts_cfg"]))
            checks.append(("D2 không lấy chuỗi ngoài nguồn", "NOTNEEDLE-MMMM" not in allv and "NOTNEEDLE-QQQQ" not in allv))
            checks.append(("D2 đếm theo nguồn (credentials=2)", len(_usable(src["credentials"])) == 2))
            checks.append(("R4 session_keys: toàn bộ nội dung (đã strip) + token >=16", {"SESSKEY-RRRR-777777 còn lại có dấu cách", "SESSKEY-RRRR-777777"} <= src["session_keys"]))
            checks.append(("R4 needle session key bắt được", bool(build_needles(_usable(src["session_keys"]), [], [], []).secret.search(b"x sesskey-rrrr-777777 y"))))
            nd2 = build_needles(_usable(allv), [], [], [])
            checks.append(("D2 needle mcp env bắt được", bool(nd2.secret.search(b"x mcpenv-gggg-777777 y"))))
            checks.append(("D2 needle header bắt được", bool(nd2.secret.search(b"hdrs-hhhh-888888"))))
            checks.append(("D2 needle organizationUuid bắt được", bool(nd2.secret.search(b"orgu-eeee-555555"))))
            (fh / ".claude" / ".credentials.json").unlink()
            (fh / ".claude-switch" / "cfg" / "acc-1" / ".credentials.json").unlink()
            src = gather(fh, None, [dev], ds, {})
            checks.append(("D2 thiếu .credentials.json -> rỗng (collect_needles sẽ dừng)", not _usable(src["credentials"])))
            try:
                collect_needles(src)
                stopped = False
            except PackError:
                stopped = True
            checks.append(("D2 collect_needles dừng khi không có bí mật .credentials.json", stopped))
            (fh / ".claude" / ".credentials.json").write_text("{không phải json", "utf-8")     # R3: không phân tích được -> dừng
            try:
                gather(fh, None, [dev], ds, {})
                stopped = False
            except PackError:
                stopped = True
            checks.append(("R3 .credentials.json hỏng -> gather dừng", stopped))
            (fh / ".claude" / ".credentials.json").unlink()
            saved_mod = sys.modules.get("win32api")
            sys.modules["win32api"] = None                                                       # R2: import win32api -> ImportError
            try:
                _os_names()
                stopped = False
            except PackError:
                stopped = True
            finally:
                if saved_mod is None:
                    del sys.modules["win32api"]
                else:
                    sys.modules["win32api"] = saved_mod
            checks.append(("R2 thiếu pywin32 -> _os_names dừng", stopped))
        finally:
            shutil.rmtree(fh, ignore_errors=True)

        # COMPUTERNAME (nd.host): miễn CHỈ cho tệp bên thứ ba trùng byte bản gốc
        nd3 = build_needles({sec}, [], [], [], host=["Workstn"])
        ident = b"the workstn word in a library"
        up = {hashlib.sha256(ident).hexdigest()}
        h3 = t.parent / "_selftest-setup-host"
        shutil.rmtree(h3, ignore_errors=True)
        try:
            lay = {"app/python/ident.py": ident,                              # trùng gốc -> miễn
                   "app/vendor/ident2.py": ident,                             # trùng gốc (vendor) -> miễn
                   "app/python/mod.py": ident + b" # edited",                 # bản sửa -> đủ needle
                   "app/python/Lib/WORKSTN.py": b"WORKSTN",                   # không trùng gốc, IN HOA -> đủ needle
                   "app/own.txt": ident,                                      # tệp của mình trùng byte với gốc -> VẪN hit
                   "app/own.dll": "Workstn".encode("utf-16-le"),              # DLL: hit
                   "app/python/clean.py": b"nothing"}
            for rel, data in lay.items():
                (h3 / rel).parent.mkdir(parents=True, exist_ok=True)
                (h3 / rel).write_bytes(data)
            got = {x.split(": ")[0]: x.split(": ")[1] for x in leaks(h3, nd3, upstream=up)}
            checks.append(("HOST tệp trùng gốc được miễn, bản sửa/IN HOA/tệp của mình/DLL thì không",
                           got == {"app/python/mod.py": "R-NAME", "app/python/Lib/WORKSTN.py": "R-NAME", "app/own.txt": "R-NAME", "app/own.dll": "R-NAME"}))
            got = {x.split(": ")[0] for x in leaks(h3, nd3)}
            checks.append(("HOST không có upstream -> tệp bên thứ ba cũng hit", {"app/python/ident.py", "app/vendor/ident2.py"} <= got))
            (h3 / "mcp.json").write_bytes(b'{"command": "workstn"}')
            checks.append(("HOST mcp.json sinh ra vẫn bị quét", "app/mcp.json: R-NAME" in leaks(h3, nd3, installed=True, prefix="app/", upstream=up)))
            (h3 / "app" / "python" / "python313._pth").write_bytes(b"python313.zip\n.\n")
            (h3 / "app" / "python" / "Lib" / "site-packages" / "p-1.dist-info").mkdir(parents=True)
            (h3 / "app" / "python" / "Lib" / "site-packages" / "p-1.dist-info" / "RECORD").write_bytes(b"x")
            n, non = third_party_proof(h3, up, "")
            checks.append(("HOST proof đếm tệp không trùng gốc", n == 7 and sorted(non) == [
                "app/python/Lib/WORKSTN.py", "app/python/Lib/site-packages/p-1.dist-info/RECORD", "app/python/clean.py", "app/python/mod.py", "app/python/python313._pth"]))
            checks.append(("HOST tập sinh sẵn nhận đúng, tệp lạ thì không",
                           bool(GEN_THIRD.fullmatch("app/python/python313._pth")) and bool(GEN_THIRD.fullmatch("app/python/Lib/site-packages/p-1.dist-info/RECORD"))
                           and not GEN_THIRD.fullmatch("app/python/Lib/WORKSTN.py")))
        finally:
            shutil.rmtree(h3, ignore_errors=True)

        # S1 (spec §5.3): R-SECRET miễn CHỈ cho tệp dưới THIRD có sha256 thuộc upstream
        s1c = b"v = '" + sec.encode() + b"'\n"
        pathc = rb"c:\users\home\z " + sec.encode()
        hs = lambda b: hashlib.sha256(b).hexdigest()
        up1 = {hs(s1c), hs(pathc), hs(sec.encode())}
        s1 = t.parent / "_selftest-setup-s1"
        shutil.rmtree(s1, ignore_errors=True)
        try:
            lay = {"app/python/x.py": s1c,                                    # (a) trùng gốc -> không R-SECRET
                   "app/vendor/x2.py": s1c,
                   "app/python/y.py": s1c + b"#",                             # (b) thêm 1 byte -> R-SECRET
                   "app/own.py": s1c,                                         # tệp của mình trùng hash upstream -> VẪN R-SECRET
                   "app/python/p.py": pathc}                                  # trùng gốc nhưng có đường dẫn user -> R-PATH, không R-SECRET
            for rel, data in lay.items():
                (s1 / rel).parent.mkdir(parents=True, exist_ok=True)
                (s1 / rel).write_bytes(data)
            ex = []
            got = {x.split(": ")[0]: x.split(": ")[1] for x in leaks(s1, nd, upstream=up1, exempt=ex)}
            checks.append(("S1 trùng gốc được miễn R-SECRET; +1 byte / tệp của mình thì không; R-PATH vẫn áp",
                           got == {"app/python/y.py": "R-SECRET", "app/own.py": "R-SECRET", "app/python/p.py": "R-PATH"}))
            checks.append(("S1 ghi lại đúng các tệp được miễn", sorted(ex) == ["app/python/p.py", "app/python/x.py", "app/vendor/x2.py"]))
            got = {x.split(": ")[0] for x in leaks(s1, nd)}
            checks.append(("S1 không truyền upstream -> tệp bên thứ ba vẫn R-SECRET", {"app/python/x.py", "app/vendor/x2.py", "app/python/y.py"} <= got))
        finally:
            shutil.rmtree(s1, ignore_errors=True)
        s1u = t.parent / "_selftest-setup-s1u"
        shutil.rmtree(s1u, ignore_errors=True)
        s1u.mkdir()
        try:
            (s1u / "unins000.dat").write_bytes(sec.encode())
            checks.append(("S1 unins000.dat có hash ∈ upstream vẫn R-SECRET",
                           leaks(s1u, nd, installed=True, prefix="app/", upstream={hs(sec.encode())}) == ["app/unins000.dat: R-SECRET"]))
        finally:
            shutil.rmtree(s1u, ignore_errors=True)

        # upstream_hashes: wheel lạ (hash không có trong lock) KHÔNG góp hash
        fk = t.parent / "_selftest-setup-up"
        shutil.rmtree(fk, ignore_errors=True)
        (fk / "wheels").mkdir(parents=True)
        bp = ("CACHE", "WHEELS", "LOCK", "XLS_REQ", "PY_VER", "PY_SHA256")
        saved_bp = {k: getattr(build_python, k) for k in bp}
        try:
            def mkzip(p, name, data):
                with zipfile.ZipFile(p, "w") as f:
                    f.writestr(name, data)
            mkzip(fk / "python-9.9.9-embed-amd64.zip", "py.txt", b"embedded")
            mkzip(fk / "wheels" / "good-1-py3-none-any.whl", "good.py", b"good")
            mkzip(fk / "wheels" / "stray-1-py3-none-any.whl", "stray.py", b"stray")
            (fk / "lock.txt").write_text("good==1 --hash=sha256:" + sha256(fk / "wheels" / "good-1-py3-none-any.whl") + "\n", encoding="utf-8")
            (fk / "xls.txt").write_text("# none\n", encoding="utf-8")
            build_python.CACHE, build_python.WHEELS = fk, fk / "wheels"
            build_python.LOCK, build_python.XLS_REQ = fk / "lock.txt", fk / "xls.txt"
            build_python.PY_VER, build_python.PY_SHA256 = "9.9.9", sha256(fk / "python-9.9.9-embed-amd64.zip")
            u = upstream_hashes()
            checks.append(("upstream: zip nhúng + wheel đúng lock góp hash; wheel lạ thì không",
                           hs(b"embedded") in u and hs(b"good") in u and hs(b"stray") not in u))
        finally:
            for k, v in saved_bp.items():
                setattr(build_python, k, v)
            shutil.rmtree(fk, ignore_errors=True)

        # S2 (spec §5.3): PUBLIC_IDS loại khỏi needle theo NGUYÊN chuỗi (không phân biệt hoa/thường, đã strip), không loại chuỗi dài hơn
        checks.append(("S2 PUBLIC_IDS == {'claude-code'}", PUBLIC_IDS == {"claude-code"}))
        src2 = {n: set() for n in SOURCES}
        src2["credentials"] = {sec}
        src2["claude_json_mcp_env_headers"] = {"claude-code", "Claude-Code ", "my-claude-code-key1", "sk-ant-api03-FAKEFAKEFAKE"}
        nd_s2, _ = collect_needles(src2)
        checks.append(("S2 'claude-code' / 'Claude-Code ' không còn là needle (URL công khai không bị bắt)",
                       not nd_s2.secret.search(b"https://x/claude-code-releases/y") and not nd_s2.secret.search(b"claude-code")))
        checks.append(("S2 chuỗi dài hơn chứa claude-code vẫn là needle", bool(nd_s2.secret.search(b"a my-claude-code-key1 b"))))
        checks.append(("S2 token sk-ant giả vẫn là needle", bool(nd_s2.secret.search(b"x sk-ant-api03-fakefakefake y"))))
        src3 = {n: set() for n in SOURCES}
        src3["credentials"] = {"Claude-Code "}
        try:
            collect_needles(src3)
            stopped = False
        except PackError:
            stopped = True
        checks.append(("S2 credentials chỉ còn claude-code -> chốt 'không thu được bí mật' dừng", stopped))

        # unins000.* (cây đã cài): COMPUTERNAME
        ui = t.parent / "_selftest-setup-unins"
        shutil.rmtree(ui, ignore_errors=True)
        ui.mkdir()
        try:
            def un(exe=b"", dat=b"", inst=True, other=None):
                for f in ui.iterdir():
                    f.unlink()
                (ui / "unins000.exe").write_bytes(exe)
                (ui / "unins000.dat").write_bytes(dat)
                if other:
                    (ui / "other.txt").write_bytes(other)
                return sorted(leaks(ui, nd3, installed=inst, prefix="app/"))
            checks.append(("UNINS exe: chỉ khớp hạ chữ -> miễn", un(exe=b"a workstn b") == []))
            checks.append(("UNINS exe: dạng title (Workstn) là nguyên chữ -> hỏng", un(exe=b"x Workstn x") == ["app/unins000.exe: R-NAME"]))
            checks.append(("UNINS exe: IN HOA UTF-8 -> hỏng", un(exe=b"x WORKSTN") == ["app/unins000.exe: R-NAME"]))
            checks.append(("UNINS exe: IN HOA UTF-16LE -> hỏng", un(exe="x WORKSTN".encode("utf-16-le")) == ["app/unins000.exe: R-NAME"]))
            checks.append(("UNINS dat: nguyên chữ đúng 1 lần -> miễn", un(dat=b"x WORKSTN y") == []))
            checks.append(("UNINS dat: 0 lần nguyên chữ (chỉ hạ chữ) -> hỏng", un(dat=b"x workstn y") == ["app/unins000.dat: R-NAME"]))
            checks.append(("UNINS dat: 2 lần -> hỏng", un(dat=b"WORKSTN WORKSTN") == ["app/unins000.dat: R-NAME"]))
            checks.append(("UNINS dat: 1 lần UTF-8 + 1 lần UTF-16LE (=2) -> hỏng", un(dat=b"WORKSTN " + "WORKSTN".encode("utf-16-le")) == ["app/unins000.dat: R-NAME"]))
            checks.append(("UNINS chỉ ở cây đã cài: staging (installed=False) thì hạ chữ cũng hỏng",
                           un(exe=b"a workstn b", inst=False) == ["app/unins000.exe: R-NAME"]))
            checks.append(("UNINS miễn chỉ riêng needle host: tệp khác trong cây cài vẫn hit",
                           un(exe=b"a workstn b", dat=b"WORKSTN", other=b"workstn") == ["app/other.txt: R-NAME"]))
            checks.append(("UNINS needle khác vẫn áp cho unins000.exe",
                           un(exe=b"a workstn " + sec.encode()) == ["app/unins000.exe: R-SECRET"]))
        finally:
            shutil.rmtree(ui, ignore_errors=True)

        bom = t.parent / "_selftest-setup-bom.txt"
        bom.write_bytes(b"\xef\xbb\xbfabc")
        nobom = t.parent / "_selftest-setup-nobom.txt"
        nobom.write_bytes(b"abc")
        checks.append(("BOM có -> đạt, thiếu -> hỏng", has_bom(bom) and not has_bom(nobom)))
        bom.unlink()
        nobom.unlink()

        bad = [n for n, ok in checks if not ok]
        if bad:
            die(f"SELFTEST SAI ({len(bad)}/{len(checks)} ca):\n  " + "\n  ".join(bad))
        print(f"SELFTEST OK ({len(checks)} ca)")
    finally:
        shutil.rmtree(t, ignore_errors=True)
        (DIST / "selftest-setup.isl").unlink(missing_ok=True)


def stage_scan():
    """Thử: dựng staging + chứng minh + quét, KHÔNG tạo exe. Chỉ in số đếm và 'tệp: luật'."""
    try:
        nd, _ = collect_needles()
        up = upstream_hashes()
        stage = fresh(DIST / "_stage-scan")
        try:
            stage_all(stage)
            report_proof(stage, up)
            ex = []
            bad = leaks(stage, nd, extra=[(INST / "cad-mcp.iss", False), (INST / "Vietnamese.isl", True)], upstream=up, exempt=ex)
            _print_exempt("staging", ex)
            files = sum(1 for p in stage.rglob("*") if p.is_file())
            by = {}
            for x in bad:
                for r in x.split(": ")[1].split(", "):
                    by[r] = by.get(r, 0) + 1
            print(f"STAGE-SCAN: {files} tệp quét, {len(bad)} tệp có hit, theo luật {by or 0}")
            for x in bad[:20]:
                print("  " + x)
            print(f"R-NAME bên thứ ba: {sum(1 for x in bad if x.startswith(THIRD) and 'R-NAME' in x)}")
        finally:
            shutil.rmtree(stage, ignore_errors=True)
    except PackError as e:
        die(str(e))
    except Exception as e:
        die(type(e).__name__)


def check_inno():
    """R1: chữ ký Pyrsys B.V. cho ISCC/ISCmplr/ISPP/islzma*.dll; hash ghim cho *.e32/*.e64; hash bộ cài gốc nếu còn. Sai -> dừng."""
    d = ISCC.parent
    lz = sorted(d.glob("islzma*.dll"))
    if not lz:
        die("Inno Setup: thiếu islzma*.dll")
    try:
        for p in (ISCC, d / "ISCmplr.dll", d / "ISPP.dll", *lz):
            build_python.check_sig(p, "Pyrsys B.V.")
    except SystemExit as e:
        die(f"Inno Setup: {e}")
    got = {p.name: sha256(p) for p in sorted([*d.glob("*.e32"), *d.glob("*.e64")])}
    if got != INNO_STUBS:
        die("Inno Setup: *.e32/*.e64 khác bản đã ghim: " + ", ".join(sorted(k for k in set(got) | set(INNO_STUBS) if got.get(k) != INNO_STUBS.get(k))))
    if INNO_EXE.is_file() and sha256(INNO_EXE) != INNO_SHA256:
        die("innosetup-6.7.4.exe khác bản đã ghim")


def main():
    if "--selftest" in sys.argv:
        return selftest()
    if "--stage-scan" in sys.argv:
        return stage_scan()
    force = "--force" in sys.argv
    pc = (ROOT / "PackageContents.xml").read_text(encoding="utf-8")
    m = re.search(r'AppVersion="([^"]+)"', pc)
    ver = m.group(1) if m else ""
    if not re.fullmatch(r"\d+\.\d+\.\d+", ver):
        die("AppVersion phải dạng x.y.z")
    if re.findall(r'<ComponentEntry\b[^>]*?\sVersion="([^"]+)"', pc) != [ver] * 3:
        die("phải có đúng 3 ComponentEntry với Version = AppVersion")
    exe = DIST / f"cad-mcp-setup-{ver}.exe"
    sha, man = Path(str(exe) + ".sha256"), DIST / f"cad-mcp-setup-{ver}.files.txt"
    if exe.exists() and not force:
        die(f"{exe} đã có — tăng AppVersion hoặc chạy với --force")
    need = [DOTNET, ISCC, INST / "Vietnamese.isl", build_python.LOCK, build_python.XLS_REQ, INST / "packages.lock.json",
            INST / "nuget.config", INST / "cad-mcp.iss", INST / "claude-cli.ps1", INST / "check-sig.ps1",
            INST / "docs" / "truoc-khi-cai.txt", INST / "docs" / "HDSD.html"]
    missing = [str(p) for p in need if not p.is_file()]
    if missing:
        die("thiếu: " + "; ".join(missing))
    try:
        build_python.check_sig(DOTNET, microsoft_toolchain=True)     # dotnet.exe: Valid + O=Microsoft Corporation + Microsoft Code Signing PCA
    except SystemExit as e:
        die(f"dotnet.exe: {e}")
    check_inno()
    if sha256(INST / "Vietnamese.isl") != VI_ISL_SHA256:
        die("Vietnamese.isl khác bản đã ghim")
    for p in (INST / "cad-mcp.iss", INST / "docs" / "truoc-khi-cai.txt"):
        if not has_bom(p):
            die(f"{p.name} phải là UTF-8 có BOM")
    if not (INST / "claude-cli.ps1").read_bytes().isascii():
        die("claude-cli.ps1 phải thuần ASCII")
    try:
        nd, skipped = collect_needles()
    except PackError as e:
        die(str(e))
    except Exception as e:                            # chỉ TÊN kiểu: traceback sẽ in đường dẫn người dùng
        die(type(e).__name__)
    print(f"needle: {skipped} chuỗi ngắn chỉ dùng dạng đường dẫn/bỏ qua")      # chỉ in SỐ
    snap0 = real_state()
    DIST.mkdir(parents=True, exist_ok=True)
    stage_root, vroot = DIST / f"_stage-setup-{ver}", DIST / f"_verify-setup-{ver}"
    fail, h = [], ""
    try:
        run([DOTNET, "build", INST / "ClaudePalette.csproj", "-c", "Release", "--no-incremental", "-p:RestoreLockedMode=true", "-nologo",
             f"-p:Version={ver}"], 900)
        up = upstream_hashes()
        stage = fresh(stage_root)
        stage_all(stage)
        report_proof(stage, up)
        ex = []
        bad = leaks(stage, nd, extra=[(INST / "cad-mcp.iss", False), (INST / "Vietnamese.isl", True)], upstream=up, exempt=ex)
        _print_exempt("staging", ex)
        if bad:
            raise PackError("staging lộ: " + "; ".join(bad[:30]))
        want_app, want_bundle = tree(stage / "app"), tree(stage / "bundle")
        rows = {**{"app/" + k: v for k, v in want_app.items()}, **{"bundle/" + k: v for k, v in want_bundle.items()}}
        man.write_text("".join(f"{v}  {k}\n" for k, v in sorted(rows.items())), encoding="utf-8")
        for f in (exe, sha):
            f.unlink(missing_ok=True)
        run([ISCC, "/Qp", f"/DAppVer={ver}", f"/DStage={stage}", f"/DOutDir={DIST}", INST / "cad-mcp.iss"], 1800)
        if not exe.is_file():
            raise PackError("ISCC không tạo " + exe.name)
        b = exe.read_bytes()                          # byte thô của exe (phần lớn đã nén; bản cài thật được quét ở verify)
        low, nul = _variants(b)
        if any(r and (r.search(low) or r.search(nul)) for r in (nd.secret, nd.path, nd.name, nd.host)):
            raise PackError("exe thô chứa bí mật/đường dẫn/tên")
        h = sha256(exe)
        sha.write_text(f"{h}  {exe.name}\n", encoding="ascii")
        verify(exe, vroot, want_app, want_bundle, snap0, nd, fail, up)
    except BaseException as e:
        fail.append(str(e)[:1500] if isinstance(e, PackError) else f"{type(e).__name__}")
    if fail:
        log = vroot / "install.log"
        if log.is_file():
            lb = log.read_bytes()
            lo, nu = _variants(lb)
            if not (nd.secret and (nd.secret.search(lo) or nd.secret.search(nu))):   # log có bí mật -> không chép
                shutil.copy2(log, DIST / "_last-verify-setup.log")
        for f in (exe, sha, man):
            f.unlink(missing_ok=True)
        for d in (stage_root, vroot):
            shutil.rmtree(d, ignore_errors=True)
        for x in (exe, sha, man, stage_root, vroot):
            if x.exists():
                print("CÒN SÓT: " + str(x))
        die("ĐÓNG GÓI/KIỂM TRA THẤT BẠI: " + " | ".join(fail))
    shutil.rmtree(stage_root)
    shutil.rmtree(vroot)
    sdk = run([DOTNET, "--version"], 60).strip()
    print(f"OK {exe} ({exe.stat().st_size} byte)\nsha256 {h}\nPython nhúng {build_python.PY_VER} · .NET SDK {sdk} · {ISCC}")


if __name__ == "__main__":
    for _s in (sys.stdout, sys.stderr):
        try:
            _s.reconfigure(encoding="utf-8", errors="replace")
        except Exception:
            pass
    main()
