"""Server MCP giả (chỉ stdlib) cho tests\SmokeWarm.exe probe: báo SHA-256 (12 ký tự đầu) của biến CAD_CLAUDE_PIN_FILE mà server nhận được,
để biết biến môi trường có tới server MCP con qua claude.exe không. Không in chính giá trị (là đường dẫn có tên người dùng)."""
import hashlib
import json
import os
import sys

for raw in sys.stdin:
    try:
        m = json.loads(raw)
    except ValueError:
        continue
    i, method = m.get("id"), m.get("method")
    if i is None:
        continue  # notification
    if method == "initialize":
        res = {"protocolVersion": (m.get("params") or {}).get("protocolVersion", "2025-06-18"),
               "capabilities": {"tools": {}}, "serverInfo": {"name": "autocad", "version": "0"}}
    elif method == "tools/list":
        res = {"tools": [{"name": "env_probe", "description": "Báo biến môi trường ghim bản vẽ (kiểm thử).",
                          "inputSchema": {"type": "object", "properties": {}}}]}
    elif method == "tools/call":
        v = os.environ.get("CAD_CLAUDE_PIN_FILE")
        t = "PROBE:" + (hashlib.sha256(v.encode("utf-8")).hexdigest()[:12] if v else "unset")
        res = {"content": [{"type": "text", "text": t}]}
    else:
        res = {}
    sys.stdout.write(json.dumps({"jsonrpc": "2.0", "id": i, "result": res}) + "\n")
    sys.stdout.flush()
