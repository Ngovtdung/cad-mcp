"""Mọi tiến trình claude của palette phải tạo qua ClaudeUi.NewPsi (spec-3.1.0 Task 5). python -I -B tests\\test_spawn_sites.py"""
import os, re, sys
sys.stdout.reconfigure(encoding="utf-8", errors="replace")  # console cp1252 không in được tiếng Việt
src =open(os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "ClaudePalette.cs"), encoding="utf-8-sig").read()
fail = 0
cases = 0


def check(label, ok):
    global fail, cases
    cases += 1
    if not ok:
        fail += 1
        print("FAIL", label)


sites = [m.start() for m in re.finditer(r"new ProcessStartInfo\(", src)]
fn = src.find("static ProcessStartInfo NewPsi(")
end = src.find("\n        }\n", fn)
body = src[fn:end] if fn >= 0 else ""
check("đúng 1 chỗ new ProcessStartInfo, nằm trong NewPsi", len(sites) == 1 and fn >= 0 and fn < sites[0] < end)
check("NewPsi gọi Attach.AuthEnv", "Attach.AuthEnv(" in body)
check("NewPsi đọc chế độ qua AuthModeValue", "AuthModeValue()" in body)
check("không lùi về ANTHROPIC_API_KEY của Windows", 'GetEnvironmentVariable("ANTHROPIC_API_KEY")' not in src)
check("auth login qua NewPsi", re.search(r'NewPsi\(exe, "auth login", out', src) is not None)
check("auth status qua NewPsi", re.search(r'NewPsi\(exe, "auth status --json", out', src) is not None)
check("chat qua NewPsi", re.search(r"NewPsi\(exe, args, out", src) is not None)
check("Live.Key có AuthStamp", re.search(r"string key = exe \+ .*AuthStamp\(\)", src) is not None)
check("Live chỉ mang chế độ (Api), không mang key", "public bool Api;" in src and "L.Secret" not in src)
check("không còn tệp session", "SessionFile" not in src and "cad_claude_session" not in src)


def body_of(sig):
    i = src.find(sig)
    return src[i:src.find("\n        }\n", i)] if i >= 0 else ""


# 3.1.0 review M2: key không nằm lại trong Process.StartInfo của tiến trình ấm
check("DropKey xoá ANTHROPIC_API_KEY khỏi psi", 'Remove("ANTHROPIC_API_KEY")' in body_of("static void DropKey("))
check("Spawn xoá key trong finally sau Process.Start", "finally { DropKey(psi); }" in body_of("Live Spawn("))
# M4: chế độ chỉ đọc 1 lần (NewPsi); auth login / status dừng khi NewPsi trả api; AuthMode sai kiểu = giá trị lạ
check("auth login dừng khi api", re.search(r'NewPsi\(exe, "auth login", out api\);\s*if \(api\) \{ DropKey\(psi\);', src) is not None)
check("auth status dừng khi api", re.search(r'NewPsi\(exe, "auth status --json", out api\);\s*if \(api\) \{ DropKey\(psi\);', src) is not None)
check("Login không đọc chế độ lần hai", "AuthModeValue()" not in body_of("void Login()"))
check("AccountInfo không đọc chế độ lần hai", "AuthModeValue()" not in body_of("static string AccountInfo()"))
am = body_of("static string AuthModeValue()")
check("AuthModeValue đọc giá trị thô, sai kiểu = lạ", 'GetValue("AuthMode")' in am and "Load(" not in am)

# spec-3.1.0 §9 S1: ảnh lỗi phải phân biệt được phiên bản palette; mọi nhánh trả về của AccountInfo đều có tiền tố
ai = body_of("static string AccountInfo()")
check("AccountInfo ghi phiên bản palette", "GetName().Version" in ai)
check("AccountInfo: mọi return đi qua tiền tố phiên bản", ai.count("return ") > 0 and ai.count("return ") == ai.count("return v + "))

# spec-3.1.0 §9 S3: bộ lọc xác thực của claude-cli.ps1 (status) phải trùng tập literal của Attachments.IsAuthVar
here = os.path.dirname(os.path.abspath(__file__))
cs = open(os.path.join(here, "..", "Attachments.cs"), encoding="utf-8-sig").read()
ps = open(os.path.join(here, "..", "installer", "claude-cli.ps1"), encoding="ascii").read()


def lits(text, pat):
    m = re.search(pat, text)
    return set(re.findall(r"""["']([^"']+)["']""", m.group(1))) if m else set()


for nm in ("AuthPrefixes", "AuthNames"):
    c = lits(cs, r"string\[\] " + nm + r" = \{([^}]*)\}")
    p = lits(ps, r"\$" + nm + r" = @\(([^)]*)\)")
    check("claude-cli.ps1 $" + nm + " trùng Attachments.cs (" + str(len(c)) + " literal)", len(c) > 0 and c == p)
sb = ps[ps.find("if ($Mode -eq 'status')"):ps.find("elseif ($Mode -eq 'verify')")]
check("status lọc EnvironmentVariables bằng Test-AuthVar trước Start()", "Test-AuthVar" in sb and sb.find("Test-AuthVar") < sb.find("$p.Start()"))
check("status thật sự Remove khoá khỏi .EnvironmentVariables", ".EnvironmentVariables.Remove(" in sb and sb.find(".EnvironmentVariables.Remove(") < sb.find("$p.Start()"))
check("claude-cli.ps1 thuần ASCII", ps.isascii())

print(f"{cases - fail}/{cases} PASS")
sys.exit(fail)
