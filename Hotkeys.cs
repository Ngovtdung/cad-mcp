// Phím tắt của palette Claude (spec docs\spec-hotkeys.md): đọc/ghi/kiểm tổ hợp phím. Từ 3.1 (Q2) cả hai phím xử lý trong Commands.OnPreTranslate (chỉ khi AutoCAD nhận phím).
// Không tham chiếu AutoCAD -> kiểm ngoài AutoCAD bằng tests\run-tests.cmd.
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Windows.Interop;

namespace ClaudeCad
{
    public struct HotKey
    {
        public const int Alt = 1, Ctrl = 2, Shift = 4, Win = 8;  // = MOD_* của RegisterHotKey = System.Windows.Input.ModifierKeys
        public readonly int Mods, Vk;                              // Vk = 0: tắt
        public static readonly HotKey None = new HotKey(0, 0);
        public HotKey(int mods, int vk) { Mods = mods & 15; Vk = vk; }
        public bool IsNone { get { return Vk == 0; } }

        static readonly string[] ModNames = { "Ctrl", "Alt", "Shift", "Win" };
        static readonly int[] ModBits = { Ctrl, Alt, Shift, Win };
        static readonly Dictionary<int, string> Names = BuildNames();
        static Dictionary<int, string> BuildNames()
        {
            var d = new Dictionary<int, string>();
            for (int i = 0; i < 26; i++) d[0x41 + i] = ((char)('A' + i)).ToString();
            for (int i = 0; i < 10; i++) { d[0x30 + i] = i.ToString(); d[0x60 + i] = "Num" + i; }
            for (int i = 1; i <= 24; i++) d[0x6F + i] = "F" + i;
            d[0x20] = "Space"; d[0x21] = "PageUp"; d[0x22] = "PageDown"; d[0x23] = "End"; d[0x24] = "Home";
            d[0x25] = "Left"; d[0x26] = "Up"; d[0x27] = "Right"; d[0x28] = "Down"; d[0x2D] = "Insert"; d[0x13] = "Pause";
            d[0xBA] = ";"; d[0xBB] = "="; d[0xBC] = ","; d[0xBD] = "-"; d[0xBE] = "."; d[0xBF] = "/";   // dấu theo vị trí phím US
            d[0xC0] = "`"; d[0xDB] = "["; d[0xDC] = "\\"; d[0xDD] = "]"; d[0xDE] = "'";
            return d;
        }

        // AutoCAD 2022 acad.CUIX (AcceleratorRoot.cui, kiểm 2026-11-14) + Windows + IME Pinyin + UniKey. Spec §2, §4.
        public static readonly string[] TakenList = {
            "Ctrl+A", "Ctrl+B", "Ctrl+C", "Ctrl+D", "Ctrl+E", "Ctrl+F", "Ctrl+G", "Ctrl+H", "Ctrl+I", "Ctrl+K", "Ctrl+L", "Ctrl+N", "Ctrl+O",
            "Ctrl+P", "Ctrl+Q", "Ctrl+R", "Ctrl+S", "Ctrl+T", "Ctrl+U", "Ctrl+V", "Ctrl+W", "Ctrl+X", "Ctrl+Y", "Ctrl+Z", "Ctrl+0", "Ctrl+1", "Ctrl+2", "Ctrl+3", "Ctrl+4", "Ctrl+6", "Ctrl+7",
            "Ctrl+8", "Ctrl+9", "Ctrl+F2", "Ctrl+Home", "Ctrl+Shift+C", "Ctrl+Shift+H", "Ctrl+Shift+I", "Ctrl+Shift+L", "Ctrl+Shift+P",
            "Ctrl+Shift+S", "Ctrl+Shift+V", "Alt+F4", "Alt+F8", "Alt+F11", "Ctrl+Space", "Alt+Space", "Alt+Z" };

        public override string ToString()
        {
            if (IsNone) return "";
            string s = "", n;
            for (int i = 0; i < 4; i++) if ((Mods & ModBits[i]) != 0) s += ModNames[i] + "+";
            return s + (Names.TryGetValue(Vk, out n) ? n : "0x" + Vk.ToString("X2"));
        }

        // "" / null -> tắt (true). Bổ trợ theo thứ tự bất kỳ, không phân biệt hoa thường.
        public static bool TryParse(string s, out HotKey k)
        {
            k = None;
            if (string.IsNullOrWhiteSpace(s)) return true;
            string[] parts = s.Split('+');
            int mods = 0;
            for (int i = 0; i < parts.Length - 1; i++)
            {
                string p = parts[i].Trim();
                int j = Array.FindIndex(ModNames, m => string.Equals(m, p, StringComparison.OrdinalIgnoreCase));
                if (j < 0) return false;
                mods |= ModBits[j];
            }
            string key = parts[parts.Length - 1].Trim();
            foreach (var e in Names)
                if (string.Equals(e.Value, key, StringComparison.OrdinalIgnoreCase)) { k = new HotKey(mods, e.Key); return true; }
            return false;
        }

        // null = dùng được. global = phím chụp vùng (RegisterHotKey toàn hệ thống). Spec §4.
        public static string Problem(HotKey k, bool global)
        {
            if (k.IsNone) return null;
            if (!Names.ContainsKey(k.Vk)) return "Phím này không dùng làm phím tắt được.";
            bool f13 = k.Vk >= 0x7C && k.Vk <= 0x87;
            bool ch = (k.Vk >= 0x30 && k.Vk <= 0x5A) || k.Vk >= 0xBA;  // chữ, số, dấu
            int m = k.Mods;
            if ((m & (Ctrl | Alt | Win)) == 0 && !f13) return "Cần có Ctrl, Alt hoặc Win (riêng F13–F24 bấm một mình được).";
            if (ch && (m & (Ctrl | Alt)) == (Ctrl | Alt) && (m & Win) == 0)
                return "Ctrl+Alt+chữ/số trùng phím AltGr của nhiều bàn phím (và phím chụp của Zalo) - chọn tổ hợp khác.";
            if (global && ch && (m & Win) == 0 && Bits(m) < 2)
                return "Phím toàn hệ thống cần 2 phím bổ trợ (vd Ctrl+Shift+…) để không chặn phím của chương trình khác.";
            if (Array.IndexOf(TakenList, k.ToString()) >= 0) return k + " đã là phím của AutoCAD/Windows/UniKey.";
            return null;
        }
        static int Bits(int m) { int n = 0; for (; m != 0; m &= m - 1) n++; return n; }

        public bool Matches(int vk, int mods) { return !IsNone && vk == Vk && mods == Mods; }
    }

    // ponytail: không còn được gọi từ 3.1 (Q2: bỏ RegisterHotKey toàn hệ thống; Hotkeys.Start không còn ai gọi, Stop chỉ còn là no-op). Giữ vì tests\AttachTests.cs còn kiểm lớp này;
    // xoá cùng khối "GlobalHotkey" trong AttachTests.cs + Hotkeys.Start/Stop/global (và lời gọi Stop trong Commands.StopHotkeys) khi dọn.
    // Phím chụp vùng toàn hệ thống: cửa sổ chỉ-nhận-thông-điệp (cha HWND_MESSAGE) trên luồng UI chính của AutoCAD.
    public sealed class GlobalHotkey : IDisposable
    {
        internal const int Id = 0x0C1A;  // id của ứng dụng phải < 0xC000
        const int WmHotkey = 0x0312, ModNoRepeat = 0x4000, ErrTaken = 1409;
        readonly HwndSource src;
        readonly Action onPress;
        bool registered;

        public GlobalHotkey(Action onPress)
        {
            this.onPress = onPress;
            src = new HwndSource(new HwndSourceParameters("ClaudeCadHotkey") { ParentWindow = new IntPtr(-3), WindowStyle = 0 });
            src.AddHook(WndProc);
        }

        internal IntPtr Handle { get { return src.Handle; } }

        // null = OK (hoặc k tắt); khác null = câu báo lỗi. Luôn huỷ đăng ký cũ trước.
        public string Register(HotKey k)
        {
            Unregister();
            if (k.IsNone) return null;
            if (RegisterHotKey(src.Handle, Id, k.Mods | ModNoRepeat, k.Vk)) { registered = true; return null; }
            int e = Marshal.GetLastWin32Error();
            return e == ErrTaken
                ? k + " đang được chương trình khác dùng (có thể là một AutoCAD khác đang mở) - chọn phím khác trong Cài đặt."
                : "Không đăng ký được " + k + " (mã lỗi " + e + ").";
        }

        public void Unregister()
        {
            if (!registered) return;
            UnregisterHotKey(src.Handle, Id);
            registered = false;
        }

        IntPtr WndProc(IntPtr h, int msg, IntPtr w, IntPtr l, ref bool handled)
        {
            if (msg == WmHotkey && w.ToInt64() == Id)
            {
                handled = true;
                try { onPress(); } catch (Exception) { }  // không để lỗi thoát ra vòng thông điệp của AutoCAD
            }
            return IntPtr.Zero;
        }

        public void Dispose() { Unregister(); src.Dispose(); }

        [DllImport("user32.dll", SetLastError = true)] static extern bool RegisterHotKey(IntPtr h, int id, int mods, int vk);
        [DllImport("user32.dll", SetLastError = true)] static extern bool UnregisterHotKey(IntPtr h, int id);
    }

    // Trạng thái phím tắt dùng chung (mọi thứ chạy trên luồng UI chính -> không khoá). Spec §5.
    public static class Hotkeys
    {
        public const string DefaultCapture = "Ctrl+Shift+X", DefaultToggle = "Ctrl+5";
        public static HotKey CaptureKey, ToggleKey;
        public static string CaptureError;  // null = phím chụp vùng đang chạy (hoặc tắt)
        public static string ToggleError;   // null = giá trị phím ẩn/hiện hợp lệ (hoặc tắt)
        public static string StartError;    // lỗi khởi tạo (Commands.Initialize): giữ qua các lần Apply sau
        static string parseError;
        static GlobalHotkey global;          // null ngoài AutoCAD (test, harness): không đăng ký thật
        static bool suspended;

        public static bool Suspended { get { return suspended; } }

        public static void Start(Action onCapture) { if (global == null) global = new GlobalHotkey(onCapture); }

        public static void Stop()
        {
            if (global == null) return;
            global.Dispose();
            global = null;
        }

        public static void Apply(string capture, string toggle)
        {
            HotKey c, t;
            parseError = null;
            ToggleError = null;
            // giá trị đọc từ registry có thể bị sửa tay: kiểm như lúc đặt trong Cài đặt (spec §4)
            if (!HotKey.TryParse(capture, out c)) { parseError = "giá trị lưu trong registry không hợp lệ: \"" + capture + "\"."; c = HotKey.None; }
            else if (HotKey.Problem(c, true) != null) { parseError = "giá trị lưu trong registry không dùng được: \"" + capture + "\" - " + HotKey.Problem(c, true); c = HotKey.None; }
            if (!HotKey.TryParse(toggle, out t)) { ToggleError = "giá trị lưu trong registry không hợp lệ: \"" + toggle + "\"."; t = HotKey.None; }
            else if (HotKey.Problem(t, false) != null) { ToggleError = "giá trị lưu trong registry không dùng được: \"" + toggle + "\" - " + HotKey.Problem(t, false); t = HotKey.None; }
            else if (!t.IsNone && t.Matches(c.Vk, c.Mods)) { ToggleError = t + " trùng phím chụp vùng."; t = HotKey.None; }
            CaptureKey = c;
            ToggleKey = t;
            CaptureError = parseError ?? StartError;
            Register();
        }

        // true: ô bấm phím trong Cài đặt đang có focus -> nhả phím chụp vùng, bỏ qua phím ẩn/hiện
        public static void Suspend(bool on) { suspended = on; Register(); }

        static void Register()
        {
            if (global == null) return;
            if (suspended) { global.Unregister(); return; }
            CaptureError = parseError ?? StartError ?? global.Register(CaptureKey);
        }

        public static int CurrentMods()
        {
            int m = 0;
            if (Native.GetKeyState(0x11) < 0) m |= HotKey.Ctrl;
            if (Native.GetKeyState(0x12) < 0) m |= HotKey.Alt;
            if (Native.GetKeyState(0x10) < 0) m |= HotKey.Shift;
            if (Native.GetKeyState(0x5B) < 0 || Native.GetKeyState(0x5C) < 0) m |= HotKey.Win;
            return m;
        }
    }

    internal static class Native
    {
        [DllImport("user32.dll")] internal static extern short GetKeyState(int vk);
        [DllImport("user32.dll")] internal static extern IntPtr GetForegroundWindow();
        [DllImport("user32.dll")] internal static extern bool SetForegroundWindow(IntPtr h);
        [DllImport("user32.dll")] internal static extern bool IsIconic(IntPtr h);
        [DllImport("user32.dll")] internal static extern bool IsWindowEnabled(IntPtr h);
        [DllImport("user32.dll")] internal static extern bool ShowWindow(IntPtr h, int cmd);
        [DllImport("user32.dll")] static extern uint GetWindowThreadProcessId(IntPtr h, out uint pid);

        internal static bool IsOwnProcess(IntPtr h)
        {
            uint pid;
            GetWindowThreadProcessId(h, out pid);
            return pid == (uint)System.Diagnostics.Process.GetCurrentProcess().Id;
        }
    }
}
