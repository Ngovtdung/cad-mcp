// Thử tay (có gọi claude thật, KHÔNG cần AutoCAD): tiến trình giữ ấm + Job Object, dùng đúng bộ args / ProcJob / BuildUserLine của palette.
// Build: tests\smoke-warm.cmd.  Chạy:
//   SmokeWarm.exe smoke                 2 tin nhắn trên 1 tiến trình, MCP rỗng -> in ms của lượt 1 và lượt 2
//   SmokeWarm.exe probe <python> [env]  CAD_CLAUDE_PIN_FILE có tới server MCP con qua claude.exe không (tuỳ chọn "env": thêm khối env vào cấu hình MCP)
//   SmokeWarm.exe hold <tệp pid>        spawn claude (chưa gửi gì, không tốn lượt) vào Job rồi đứng chờ: giết tiến trình này, claude.exe phải chết ≤ 5 s
// Không in đường dẫn / tài khoản: chỉ số ms, đúng/sai.
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading;
using System.Web.Script.Serialization;
using ClaudeCad;

static class SmokeWarm
{
    static string Tmp = Path.Combine(Path.GetTempPath(), "cad-mcp-smoke");

    static int Main(string[] a)
    {
        Console.OutputEncoding = Encoding.UTF8;
        Directory.CreateDirectory(Tmp);
        if (a.Length == 0) { Console.WriteLine("smoke | probe <python> [env] | hold <tệp pid>"); return 2; }
        string exe = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), ".local", "bin", "claude.exe");
        if (!File.Exists(exe)) { Console.WriteLine("không thấy claude.exe"); return 2; }
        string pin = Path.Combine(Path.GetTempPath(), "cad_claude_doc_" + Process.GetCurrentProcess().Id + ".txt");
        string cfg = Path.Combine(Tmp, "empty-mcp.json");
        if (a[0] == "probe")
        {
            string stub = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "envprobe_stub.py");
            bool envBlock = a.Length > 2 && a[2] == "env";
            cfg = Path.Combine(Tmp, "probe-mcp.json");
            File.WriteAllText(cfg, "{\"mcpServers\":{\"autocad\":{\"command\":" + Attach.Json(a[1]) + ",\"args\":[\"-I\",\"-S\",\"-B\"," + Attach.Json(stub) + "]"
                + (envBlock ? ",\"env\":{\"CAD_CLAUDE_PIN_FILE\":\"${CAD_CLAUDE_PIN_FILE:-}\"}" : "") + "}}}");
        }
        else File.WriteAllText(cfg, "{\"mcpServers\":{}}");
        string args = Attach.ClaudeArgs(cfg, null, null, null);
        var job = ProcJob.TryCreate();
        var psi = new ProcessStartInfo(exe, args)
        {
            UseShellExecute = false, CreateNoWindow = true, WorkingDirectory = Tmp,
            RedirectStandardInput = true, RedirectStandardOutput = true, RedirectStandardError = true,
            StandardOutputEncoding = Encoding.UTF8, StandardErrorEncoding = Encoding.UTF8
        };
        psi.EnvironmentVariables["ENABLE_TOOL_SEARCH"] = "false";
        psi.EnvironmentVariables["CAD_CLAUDE_PIN_FILE"] = pin;
        psi.EnvironmentVariables["CAD_CLAUDE_ACAD_PID"] = Process.GetCurrentProcess().Id.ToString();  // 3.1 N4: có PIN_FILE thì server.py đòi pid
        Attach.NoStdinBom();
        var sw = Stopwatch.StartNew();
        var p = Process.Start(psi);
        bool inJob = job != null && job.Assign(p);
        Console.WriteLine("job: " + (job != null) + ", gán: " + inJob + ", spawn " + sw.ElapsedMilliseconds + " ms");
        p.ErrorDataReceived += (s, e) => { };
        p.BeginErrorReadLine();
        if (a[0] == "hold")
        {
            File.WriteAllText(a[1], p.Id.ToString());
            Console.WriteLine("hold: claude pid đã ghi vào tệp; chờ bị giết...");
            Thread.Sleep(Timeout.Infinite);
        }
        var js = new JavaScriptSerializer { MaxJsonLength = int.MaxValue };
        Func<string, long, string> turn = (prompt, t0) =>
        {
            var line = Attach.BuildUserLine(prompt, new List<Attachment>());
            long w0 = sw.ElapsedMilliseconds;
            byte[] b = new UTF8Encoding(false).GetBytes(line);
            p.StandardInput.BaseStream.Write(b, 0, b.Length);
            p.StandardInput.BaseStream.Flush();
            string l, text = "", probe = null;
            long firstEv = -1;
            while ((l = p.StandardOutput.ReadLine()) != null)
            {
                if (firstEv < 0) firstEv = sw.ElapsedMilliseconds - w0;
                var ev = js.DeserializeObject(l) as Dictionary<string, object>;
                if (ev == null) continue;
                object ty;
                ev.TryGetValue("type", out ty);
                if ("result".Equals(ty))
                {
                    object r; ev.TryGetValue("result", out r);
                    Console.WriteLine("  result sau " + (sw.ElapsedMilliseconds - w0) + " ms (từ lúc gửi), sự kiện đầu sau " + firstEv + " ms, is_error=" + (ev.ContainsKey("is_error") ? ev["is_error"] : "?"));
                    return probe ?? Convert.ToString(r);
                }
                if ("user".Equals(ty) && l.Contains("PROBE:"))
                {
                    int i = l.IndexOf("PROBE:", StringComparison.Ordinal);
                    probe = l.Substring(i + 6, Math.Min(16, l.Length - i - 6)).TrimEnd('"', '\\', ' ');
                }
            }
            return "(stdout đóng)";
        };
        int rc = 0;
        if (a[0] == "smoke")
        {
            Console.WriteLine("lượt 1:");
            string r1 = turn("Trả lời đúng một chữ: ok", 0);
            Console.WriteLine("lượt 2 (cùng tiến trình):");
            long t2 = sw.ElapsedMilliseconds;
            string r2 = turn("Trả lời đúng một chữ: xong", 0);
            Console.WriteLine("trả lời 1: " + r1.Trim() + " | trả lời 2: " + r2.Trim() + " | claude còn sống: " + !p.HasExited);
        }
        else if (a[0] == "probe")
        {
            string got = turn("Gọi tool mcp__autocad__env_probe đúng một lần (không tham số), rồi trả lời một chữ: xong", 0);
            string want = Hash(pin);
            Console.WriteLine("probe nhận được: " + (got == null ? "(không)" : got.Length > 12 ? got.Substring(0, 12) : got) + "; envvar tới được server MCP con: " + (got != null && got.StartsWith(want) ));
            rc = got != null && got.StartsWith(want) ? 0 : 1;
        }
        if (job != null) job.Kill(); else { try { p.Kill(); } catch (Exception) { } }
        Console.WriteLine("đã kill job; claude còn sống: " + !p.WaitForExit(3000));
        try { File.Delete(pin); } catch (Exception) { }
        return rc;
    }

    static string Hash(string s)
    {
        using (var h = SHA256.Create()) return BitConverter.ToString(h.ComputeHash(Encoding.UTF8.GetBytes(s))).Replace("-", "").ToLowerInvariant().Substring(0, 12);
    }
}
