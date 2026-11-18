// Palette Claude cho AutoCAD, giao diện theo palette Properties: gợi ý lệnh (như Claude trong Excel), hội thoại, cài đặt. Giao diện: ClaudePalette.xaml.
// Chạy Claude Code headless, dùng MCP autocad (server.py) để đọc/sửa bản vẽ.
// Build: build.cmd (máy dev; installer\ClaudePalette.csproj: net46 cho AutoCAD 2018-2024, net8 cho 2025-2026, net10 cho 2027+). Phát hành: installer\pack_setup.py -> Setup.exe.
// Lệnh: CLAUDE (ẩn/hiện palette), CLAUDENEW (hội thoại mới), CLAUDECHUP (chụp vùng). Phím tắt: Hotkeys.cs.
using System;
using System.Collections;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
#if !NET
using System.Web.Script.Serialization;
#endif
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Markup;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Threading;
using Microsoft.Win32;
using Acad = Autodesk.AutoCAD.ApplicationServices.Application;
using Acw = Autodesk.AutoCAD.Windows;
using Adw = Autodesk.Windows;

[assembly: Autodesk.AutoCAD.Runtime.CommandClass(typeof(ClaudeCad.Commands))]
[assembly: Autodesk.AutoCAD.Runtime.ExtensionApplication(typeof(ClaudeCad.Commands))]

namespace ClaudeCad
{
    public class Commands : Autodesk.AutoCAD.Runtime.IExtensionApplication
    {
        static Acw.PaletteSet ps;
        static ContentControl host;  // nội dung palette = Page của bản vẽ đang mở
        // ponytail: mỗi bản vẽ một ClaudeUi + một claude.exe (rảnh 10 phút tự tắt); nhiều bản vẽ cùng chạy ~N×300 MB. Thêm giới hạn khi có phản ánh
        static readonly Dictionary<object, ClaudeUi> uis = new Dictionary<object, ClaudeUi>();
        static readonly object NoDoc = new object();  // khi không có bản vẽ nào
        static int serial;

        // UI của bản vẽ đang mở (tạo khi cần). Chỉ gọi từ luồng UI của AutoCAD
        static ClaudeUi Ui
        {
            get
            {
                var d = Acad.DocumentManager.MdiActiveDocument;
                object k = d != null ? (object)d : NoDoc;
                ClaudeUi u;
                if (!uis.TryGetValue(k, out u))
                {
                    Func<string> name;
                    if (d != null) name = () => d.Name; else name = () => null;
                    u = new ClaudeUi(name, ++serial);
                    uis[k] = u;
                }
                return u;
            }
        }

        // đổi chế độ / key / đăng nhập xong: tiến trình rảnh của MỌI bản vẽ không còn khớp -> kill (tiến trình đang chạy lượt thì giữ)
        internal static void KillIdleAll() { foreach (var u in uis.Values.ToList()) { try { u.KillIdle(); } catch (Exception) { } } }

        static void Swap()
        {
            if (host == null) return;
            var u = Ui;
            if (ReferenceEquals(host.Content, u.Page)) return;
            SetKeepFocus(false);  // không mang bàn phím theo khi đổi bản vẽ
            host.Content = u.Page;
            u.SyncSettings();
        }

        static void OnDocCurrent(object s, Autodesk.AutoCAD.ApplicationServices.DocumentCollectionEventArgs e) { try { Swap(); } catch (Exception) { } }

        static void OnDocDestroyed(object s, Autodesk.AutoCAD.ApplicationServices.DocumentDestroyedEventArgs e) { try { Swap(); } catch (Exception) { } }

        static void OnDocClosing(object s, Autodesk.AutoCAD.ApplicationServices.DocumentCollectionEventArgs e)
        {
            try
            {
                ClaudeUi u;
                if (e.Document == null || !uis.TryGetValue(e.Document, out u)) return;
                uis.Remove(e.Document);
                if (host != null && ReferenceEquals(host.Content, u.Page)) host.Content = null;
                u.Shutdown();  // kill claude + python của bản vẽ này, xoá tệp ghim, dừng worker
            }
            catch (Exception) { }
        }

        public void Initialize()
        {
            // palette đang mở lúc thoát AutoCAD -> tự mở lại lần sau (như palette có sẵn của CAD)
            Acad.BeginQuit += (s, e) => ClaudeUi.Save("Visible", ps != null && ps.Visible ? "1" : "0");
            if (ClaudeUi.Load("Visible", "0") == "1") Acad.Idle += ShowOnIdle;
            // nút trên Ribbon (tab Add-ins, cạnh D2P). Đổi workspace thì Ribbon dựng lại -> thêm lại
            if (Adw.ComponentManager.Ribbon != null) AddRibbon();
            else Adw.ComponentManager.ItemInitialized += OnRibbonReady;
            Acad.SystemVariableChanged += (s, e) => { if (e.Name == "WSCURRENT") AddRibbon(); };
            // phím tắt (spec-hotkeys §5, 3.1 Q2): cả phím chụp vùng lẫn ẩn/hiện xử lý trong OnPreTranslate -> chỉ chạy khi AutoCAD đang nhận phím
            // BeginQuit chạy trước hộp "Lưu thay đổi?" (huỷ thoát vẫn còn phím) -> gỡ ở QuitWillStart. Mỗi bước một try: lỗi bước này không chặn bước kia
            try { Acad.QuitWillStart += (s, e) => { StopHotkeys(); ShutdownUi(); }; } catch (Exception) { }  // I1: kill claude ở đây và ở Terminate, KHÔNG ở BeginQuit (còn huỷ thoát được)
            try { Acad.PreTranslateMessage += OnPreTranslate; } catch (Exception) { }
            try { ClaudeUi.ApplyHotkeys(); }
            catch (Exception ex) { Hotkeys.StartError = "không đọc được phím tắt: " + ex.Message; }
        }

        public void Terminate() { StopHotkeys(); ShutdownUi(); }

        static void ShutdownUi()
        {
            foreach (var u in uis.Values.ToList()) { try { u.Shutdown(); } catch (Exception) { } }
            uis.Clear();
        }

        static void StopHotkeys()
        {
            try { Acad.PreTranslateMessage -= OnPreTranslate; } catch (Exception) { }
            try { Hotkeys.Stop(); } catch (Exception) { }
        }

        internal static bool PaletteVisible { get { return ps != null && ps.Visible; } }

        // Phím ẩn/hiện palette và phím chụp vùng: chỉ trong AutoCAD (như Ctrl+1). Spec §6.3, 3.1 Q2. Chạy cho MỌI thông điệp -> lọc rẻ trước.
        static void OnPreTranslate(object s, Autodesk.AutoCAD.ApplicationServices.PreTranslateMessageEventArgs e)
        {
            try
            {
                var m = e.Message;
                if (m.message != 0x100 && m.message != 0x104) return;  // WM_KEYDOWN / WM_SYSKEYDOWN
                if ((m.lParam.ToInt64() & 0x40000000) != 0 || Hotkeys.Suspended) return;  // giữ phím (lặp) / đang đặt phím trong Cài đặt
                int vk = (int)m.wParam.ToInt64(), mods = Hotkeys.CurrentMods();
                if (Hotkeys.CaptureKey.Matches(vk, mods))
                {
                    if (!Native.IsWindowEnabled(Acad.MainWindow.Handle)) return;  // hộp thoại modal: để phím tới hộp thoại
                    e.Handled = true;
                    IntPtr fg = Native.GetForegroundWindow();
                    Dispatcher.CurrentDispatcher.BeginInvoke(new Action(() => StartCapture(fg)));
                    return;
                }
                if (!Hotkeys.ToggleKey.Matches(vk, mods)) return;
                if (!Native.IsWindowEnabled(Acad.MainWindow.Handle)) return;  // hộp thoại modal đang mở: để phím tới hộp thoại
                e.Handled = true;
                Dispatcher.CurrentDispatcher.BeginInvoke(new Action(TogglePanel));
            }
            catch (Exception) { }
        }

        static void TogglePanel()
        {
            try
            {
                if (ps != null && ps.Visible) { ReleaseFocus(); ps.Visible = false; }  // trả bàn phím về vùng vẽ rồi ẩn
                else ShowPanel();
            }
            catch (Exception) { }
        }

        // phím tắt (returnTo = cửa sổ đứng trước) hoặc lệnh CLAUDECHUP (IntPtr.Zero). Spec §6.1.
        static void StartCapture(IntPtr returnTo)
        {
            try
            {
                if (!Native.IsWindowEnabled(Acad.MainWindow.Handle)) { System.Media.SystemSounds.Beep.Play(); return; }  // AutoCAD đang mở hộp thoại modal
                EnsurePanel();
                Swap();
                Ui.CaptureRegion(returnTo);
            }
            catch (Exception ex) { try { Ui.Error("Không chụp được màn hình: " + ex.Message); } catch (Exception) { } }
        }

        // Bàn phím: chỉ giữ khi đang gõ trong ô chat. KeepFocus=true thường trực làm Enter/Space/Esc của lệnh CAD
        // chạy vào ô chat (vd COPY kẹt ở "Select objects", Esc không hủy được lệnh).
        internal static void SetKeepFocus(bool on)
        {
            if (ps != null && ps.KeepFocus != on) ps.KeepFocus = on;
        }

        // Chuột rời palette -> nhả bàn phím về vùng vẽ
        internal static void ReleaseFocus()
        {
            SetKeepFocus(false);
            try { Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView(); } catch { }
        }

        // Q2: ẩn palette khi chụp vùng, chụp xong hiện lại
        internal static void SetPaletteVisible(bool on)
        {
            if (ps != null && ps.Visible != on) ps.Visible = on;
        }

        static void OnRibbonReady(object s, Adw.RibbonItemEventArgs e)
        {
            if (Adw.ComponentManager.Ribbon == null) return;
            Adw.ComponentManager.ItemInitialized -= OnRibbonReady;
            AddRibbon();
        }

        static void AddRibbon()
        {
            var rc = Adw.ComponentManager.Ribbon;
            if (rc == null) return;
            var tab = rc.Tabs.FirstOrDefault(t => t.Id == "ID_TabAddins" || t.Title == "Add-ins")
                ?? rc.Tabs.FirstOrDefault(t => t.Id == "CLAUDE_TAB");
            if (tab == null)  // workspace không có tab Add-ins -> tab riêng
            {
                tab = new Adw.RibbonTab { Title = "Claude AI", Id = "CLAUDE_TAB" };
                rc.Tabs.Add(tab);
            }
            if (tab.Panels.Any(p => p.Source != null && p.Source.Id == "CLAUDE_PANEL")) return;
            var src = new Adw.RibbonPanelSource { Title = "Claude AI", Id = "CLAUDE_PANEL" };
            src.Items.Add(new Adw.RibbonButton
            {
                Text = "Claude AI", ShowText = true, ShowImage = true, Size = Adw.RibbonItemSize.Large, Orientation = Orientation.Vertical,
                LargeImage = IconFrame(32), Image = IconFrame(16), ToolTip = "Mở palette Claude (lệnh CLAUDE)", CommandHandler = new Cmd(ShowPanel)
            });
            var row = new Adw.RibbonRowPanel();
            row.Items.Add(SmallButton("Hội thoại mới", "", "#3DDC4A", () => { ShowPanel(); Ui.NewChat(); }));
            row.Items.Add(new Adw.RibbonRowBreak());
            row.Items.Add(SmallButton("Sửa gợi ý lệnh", "", "#4DA3FF", () => { ShowPanel(); Ui.EditSuggestions(); }));
            src.Items.Add(row);
            tab.Panels.Add(new Adw.RibbonPanel { Source = src });
        }

        static Adw.RibbonButton SmallButton(string text, string glyph, string color, Action run)
        {
            return new Adw.RibbonButton
            {
                Text = text, ShowText = true, ShowImage = true, Size = Adw.RibbonItemSize.Standard, Orientation = Orientation.Horizontal,
                Image = Glyph(glyph, color), CommandHandler = new Cmd(run)
            };
        }

        static ImageSource IconFrame(int size)
        {
            string ico = Path.Combine(ClaudeUi.DllDir, "icon.ico");
            if (!File.Exists(ico)) return null;
            var dec = new IconBitmapDecoder(new Uri(ico), BitmapCreateOptions.None, BitmapCacheOption.OnLoad);
            return dec.Frames.OrderBy(f => Math.Abs(f.PixelWidth - size)).First();
        }

        // icon 16px vẽ từ font Segoe MDL2 Assets
        static ImageSource Glyph(string ch, string color)
        {
            var tb = new TextBlock { Text = ch, FontFamily = new FontFamily("Segoe MDL2 Assets"), FontSize = 14, Foreground = ClaudeUi.BrushOf(color),
                TextAlignment = TextAlignment.Center, Width = 16, Height = 16 };
            tb.Measure(new Size(16, 16));
            tb.Arrange(new Rect(0, 0, 16, 16));
            var bmp = new RenderTargetBitmap(16, 16, 96, 96, PixelFormats.Pbgra32);
            bmp.Render(tb);
            return bmp;
        }

        class Cmd : ICommand
        {
            readonly Action run;
            public Cmd(Action run) { this.run = run; }
            public event EventHandler CanExecuteChanged { add { } remove { } }
            public bool CanExecute(object p) { return true; }
            public void Execute(object p) { run(); }
        }

        static void ShowOnIdle(object s, EventArgs e)
        {
            Acad.Idle -= ShowOnIdle;
            ShowPanel();
        }

        [Autodesk.AutoCAD.Runtime.CommandMethod("CLAUDE")]
        public void Claude() { TogglePanel(); }

        [Autodesk.AutoCAD.Runtime.CommandMethod("CLAUDECHUP")]
        public void ClaudeChup() { StartCapture(IntPtr.Zero); }

        [Autodesk.AutoCAD.Runtime.CommandMethod("CLAUDENEW")]
        public void ClaudeNew()
        {
            ShowPanel();
            Ui.NewChat();
        }

        static void EnsurePanel()
        {
            if (ps != null) return;
            host = new ContentControl();
            Acad.DocumentManager.DocumentBecameCurrent += OnDocCurrent;  // đăng ký một lần (EnsurePanel chỉ dựng ps một lần)
            Acad.DocumentManager.DocumentToBeDestroyed += OnDocClosing;
            Acad.DocumentManager.DocumentDestroyed += OnDocDestroyed;  // đóng bản vẽ cuối: không còn DocumentBecameCurrent -> tự đổi sang UI "không có bản vẽ"
            // GUID cố định -> AutoCAD nhớ vị trí dock/kích thước giữa các lần mở
            ps = new Acw.PaletteSet("CLAUDE AI", new Guid("6f0c2c3e-6a55-4b0e-9a3c-2b7f1f3c8d41"));
            ps.Style = Acw.PaletteSetStyles.ShowAutoHideButton | Acw.PaletteSetStyles.ShowCloseButton | Acw.PaletteSetStyles.ShowPropertiesMenu;
            ps.DockEnabled = Acw.DockSides.Left | Acw.DockSides.Right;
            ps.MinimumSize = new System.Drawing.Size(280, 420);
            ps.Size = new System.Drawing.Size(380, 760);
            ps.KeepFocus = false;  // chỉ bật khi ô chat có focus (xem SetKeepFocus / ReleaseFocus)
            string ico = Path.Combine(ClaudeUi.DllDir, "icon.ico");
            if (File.Exists(ico)) ps.Icon = new System.Drawing.Icon(ico);
            ps.AddVisual("Claude", host, true);
        }

        static void ShowPanel()
        {
            EnsurePanel();
            Swap();
            ps.Visible = true;
        }
    }

    // Phần không phụ thuộc AutoCAD/WPF, không trường static: tìm thư mục cài, tìm claude.exe, đọc JSON.
    // Kiểm bằng installer\test_core.ps1 trên DLL net46 và net8. C# 5 (csc của tests\ cũng biên dịch tệp này).
    public static class Core
    {
        // Chỉ nhận "X:\..." — loại tương đối, "\x" theo ổ hiện hành, "C:x", UNC "\\srv\share"
        public static bool IsDriveAbs(string p)
        {
            return p != null && p.Length >= 3 && ((p[0] >= 'A' && p[0] <= 'Z') || (p[0] >= 'a' && p[0] <= 'z'))
                && p[1] == ':' && p[2] == '\\';
        }

        // Home = thư mục chứa server.py + mcp.json. Setup (và build.cmd máy dev) ghi HKCU\Software\ClaudeCad\Home.
        // Nhận khi là "X:\..." và có mcp.json; không thì %LOCALAPPDATA%\cad-mcp (nơi Setup cài).
        public static string ResolveHome(string reg, string localAppData)
        {
            if (IsDriveAbs(reg))
            {
                try { if (File.Exists(Path.Combine(reg, "mcp.json"))) return reg.TrimEnd('\\'); }
                catch (ArgumentException) { }
            }
            return Path.Combine(localAppData, "cad-mcp");
        }

        // claude.exe: %USERPROFILE%\.local\bin (bộ cài chính thức) rồi các mục PATH dạng "X:\...".
        // Không bao giờ trả tên trần "claude": không chạy nhầm tệp trong thư mục hiện hành của AutoCAD. null = chưa cài.
        public static string FindClaude(string userProfile, string pathVar)
        {
            var dirs = new List<string>();
            if (!string.IsNullOrEmpty(userProfile)) dirs.Add(Path.Combine(userProfile, @".local\bin"));
            dirs.AddRange((pathVar ?? "").Split(';'));
            foreach (string d in dirs)
            {
                string t = d.Trim().Trim('"').Trim();
                if (!IsDriveAbs(t)) continue;
                try
                {
                    string p = Path.Combine(t, "claude.exe");
                    if (File.Exists(p)) return p;
                }
                catch (ArgumentException) { }  // mục PATH có ký tự sai (.NET Framework)
            }
            return null;
        }

        // Như JavaScriptSerializer.DeserializeObject: object -> Dictionary<string, object>, mảng -> IEnumerable,
        // chuỗi/số/bool/null; JSON hỏng -> ném lỗi. .NET 8 không có System.Web.Extensions -> System.Text.Json.
        public static object ParseJson(string s)
        {
#if NET
            using (var d = System.Text.Json.JsonDocument.Parse(s, new System.Text.Json.JsonDocumentOptions { MaxDepth = 256 })) return FromJson(d.RootElement);
#else
            return new JavaScriptSerializer { MaxJsonLength = int.MaxValue }.DeserializeObject(s);
#endif
        }

#if NET
        static object FromJson(System.Text.Json.JsonElement e)
        {
            switch (e.ValueKind)
            {
                case System.Text.Json.JsonValueKind.Object:
                    var o = new Dictionary<string, object>();
                    foreach (var p in e.EnumerateObject()) o[p.Name] = FromJson(p.Value);
                    return o;
                case System.Text.Json.JsonValueKind.Array:
                    var a = new List<object>();
                    foreach (var x in e.EnumerateArray()) a.Add(FromJson(x));
                    return a;
                case System.Text.Json.JsonValueKind.String:
                    return e.GetString();
                case System.Text.Json.JsonValueKind.Number:
                    long l;
                    if (e.TryGetInt64(out l)) return l;
                    return e.GetDouble();
                case System.Text.Json.JsonValueKind.True:
                    return true;
                case System.Text.Json.JsonValueKind.False:
                    return false;
                default:
                    return null;
            }
        }
#endif
    }

    public class ClaudeUi
    {
        // server.py, mcp.json, goi_y.txt, xls2tsv.py. Xem Core.ResolveHome
        public static readonly string Home = Core.ResolveHome(Load("Home", null), Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData));
        public static readonly string DllDir = Path.GetDirectoryName(typeof(ClaudeUi).Assembly.Location);
        static readonly string SuggestFile = Path.Combine(Home, "goi_y.txt");
        // Tìm lại mỗi lần dùng: cài Claude Code sau khi đã mở AutoCAD vẫn nhận
        internal static string ExeOverride = null;  // chỉ cho tests\ui_send_fail.ps1 (claude giả); sản phẩm không bao giờ đặt
        static string Exe
        {
            get { return ExeOverride ?? Core.FindClaude(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), Environment.GetEnvironmentVariable("PATH")); }
        }
        const string NoClaude = "Không tìm thấy claude.exe (Claude Code). Chạy lại Setup để cài, hoặc xem Hướng dẫn sử dụng, mục Xử lý sự cố.";
        const string Default = "Mặc định";
        const string DefaultSuggestions = @"; Mỗi dòng là một gợi ý lệnh cho Claude. Dòng bắt đầu bằng ; bị bỏ qua.
Ghi kích thước tất cả các cạnh của polyline đang chọn
Tính tổng chiều dài các đối tượng đang chọn
Tính diện tích từng polyline kín đang chọn và tổng diện tích
Đếm số lượng từng loại block trong bản vẽ
Liệt kê các layer và số đối tượng trên mỗi layer
Chuyển các text đang chọn sang layer TEXT màu vàng
Tìm và xoá các đối tượng trùng nhau (cùng loại, cùng vị trí)
Vẽ khung bản vẽ A3 tỉ lệ 1:100 tại gốc toạ độ
Chụp màn hình bản vẽ và cho biết đang có gì
Zoom toàn bộ bản vẽ rồi lưu file
";
        // màu lấy từ palette Properties của AutoCAD (xem ClaudePalette.xaml); viền xanh = màu nhấn #0696D7 của CAD
        static readonly Brush Dim = BrushOf("#A9B1BE"), Red = BrushOf("#FF8A80"), Bubble = BrushOf("#2E3440"), BubbleEdge = BrushOf("#0696D7");

        public readonly FrameworkElement Page;
        readonly TextBox input;
        readonly TextBlock hint, status, modelLabel, sugCount;
        readonly StackPanel log, suggestions;
        readonly ScrollViewer logScroll;
        readonly Button stopBtn;
        readonly ComboBox model, effort, withSel;
        readonly TextBox hkCapture, hkToggle;
        string shownToggleError;
        string shownHotkeyError;  // lỗi phím chụp vùng đã ghi vào hội thoại (không ghi lặp)
        List<string> allSuggestions = new List<string>();
        volatile Live live;  // tiến trình claude đang giữ (đang chạy hoặc ấm giữa 2 tin nhắn), null = không có (T4)
        Turn turn;           // lượt đang chạy (chỉ luồng UI), null = rảnh; sự kiện của lượt đã dừng/bị đè/xong bị bỏ qua
        string phase = "";   // giai đoạn hiện tại cho bộ đếm trên thanh trạng thái
        readonly DispatcherTimer tick, idleTimer;  // tick: 500 ms khi có lượt đang chạy; idleTimer: kill tiến trình ấm sau 10 phút rảnh
        readonly DedupSet dedup = new DedupSet();  // T6: PDF đã gửi đầy đủ trong phiên (SHA-256, chỉ trong RAM)
        readonly object writeGate = new object(), sessionGate = new object();
        const int IdleMinutes = 10;
        Process login;  // cửa sổ "claude auth login" đang mở (chỉ cho mở 1 cửa sổ)
        int steps;
        string last;
        bool filling;  // đang điền gợi ý vào ô nhập -> không lọc lại danh sách (giữ nguyên chỗ để bấm đúp)
        readonly WrapPanel chips;
        readonly UIElement dropHint;
        readonly List<Attachment> pending = new List<Attachment>();  // tệp chờ gửi + chip "đang đọc" (Kind = "loading")
        readonly Dictionary<Attachment, CancellationTokenSource> loading = new Dictionary<Attachment, CancellationTokenSource>();
        readonly BlockingCollection<Action> jobs = new BlockingCollection<Action>();
        Thread worker;   // 1 luồng nền STA cho mọi việc đọc/giải mã tệp (R5)
        int shotNo;      // đánh số anh-dan-N / chup-vung-N
        bool modal;      // đang mở hộp chọn tệp / chụp vùng -> MouseLeave không giật focus
        // 3.1.1 Agent Skills. Skill của người dùng: <Home>\.claude\skills. Bản nháp AI soạn (máy chủ ghi, CAD_CLAUDE_SKILL_DRAFTS): <Home>\skill-nhap, ngoài .claude
        static readonly string SkillsDir = Path.Combine(Home, ".claude", "skills");
        static readonly string DraftsDir = Path.Combine(Home, "skill-nhap");
        static readonly string UserProfile = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);  // ParentDirty bỏ qua .claude ở đây
        TextBlock skillStatus;        // chữ trạng thái hàng Skill ở Cài đặt (Attach.SkillStatusText)
        ContextMenu skillMenu;        // menu "Quản lý ▾"
        volatile bool skillsBlocked;  // system/init sai tool: cả phiên palette này chạy như 3.1.0
        bool skillWhyShown;           // dòng đỏ "skill bị tắt" chỉ hiện một lần mỗi phiên
        string lastDrafts = "";       // danh sách bản nháp đã báo (không lặp lại mỗi lượt)
        const int CaptureSettleMs = 300;  // chờ AutoCAD vẽ lại sau khi ẩn palette (nút hiệu chỉnh)

        readonly string PinFile;          // %TEMP%\cad_claude_doc_<pid>_<n>.txt
        readonly Func<string> docName;    // tên bản vẽ của UI này; null = ngoài AutoCAD (harness): tab đang mở như 3.0.x
        string sessionId;                 // F4: chỉ trong bộ nhớ, không còn tệp session (Q1: không resume qua khởi động lại AutoCAD)

        public ClaudeUi() : this(null, 0) { }

        public ClaudeUi(Func<string> docName, int serial)
        {
            this.docName = docName;
            PinFile = Path.Combine(Path.GetTempPath(), "cad_claude_doc_" + PidOfThis().ToString(CultureInfo.InvariantCulture)
                + "_" + serial.ToString(CultureInfo.InvariantCulture) + ".txt");
            try { File.Delete(PinFile); } catch (Exception) { }  // F4: không dùng tệp sót lại / đặt sẵn
            Page =(FrameworkElement)XamlReader.Load(typeof(ClaudeUi).Assembly.GetManifestResourceStream("ClaudePalette.xaml"));
            input = (TextBox)Page.FindName("Input");
            hint = (TextBlock)Page.FindName("Hint");
            status = (TextBlock)Page.FindName("Status");
            modelLabel = (TextBlock)Page.FindName("ModelLabel");
            sugCount = (TextBlock)Page.FindName("SugCount");
            log = (StackPanel)Page.FindName("Log");
            suggestions = (StackPanel)Page.FindName("Suggestions");
            logScroll = (ScrollViewer)Page.FindName("LogScroll");
            stopBtn = (Button)Page.FindName("StopBtn");
            chips = (WrapPanel)Page.FindName("Chips");
            dropHint = (UIElement)Page.FindName("DropHint");
            Attach.Home = Home;
            tick = new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(500) };
            tick.Tick += (s, e) =>
            {
                try
                {
                    var t = turn;
                    if (t == null) { tick.Stop(); return; }
                    status.Text = phase + " · " + Sec(t.Sw.ElapsedMilliseconds);
                }
                catch (Exception) { }  // D5: không để ngoại lệ thoát ra khỏi Tick (sập AutoCAD)
            };
            ThreadPool.QueueUserWorkItem(_ => CleanStalePins());
            idleTimer = new DispatcherTimer { Interval = TimeSpan.FromMinutes(IdleMinutes) };
            idleTimer.Tick += (s, e) =>
            {
                try { idleTimer.Stop(); if (turn == null) KillLive(); }
                catch (Exception) { }
            };

            // Enter gửi (đang chạy thì đè lệnh mới), Shift+Enter xuống dòng, Esc dừng, Ctrl+V / Shift+Insert: tệp/ảnh thành chip, chữ dán như cũ
            input.PreviewKeyDown += (s, e) =>
            {
                if (e.Key == Key.Enter && (Keyboard.Modifiers & ModifierKeys.Shift) == 0) { e.Handled = true; Send(input.Text); }
                else if (e.Key == Key.Escape) { e.Handled = true; Stop(true, "stopped"); }
                else if ((e.Key == Key.V && Keyboard.Modifiers == ModifierKeys.Control) || (e.Key == Key.Insert && Keyboard.Modifiers == ModifierKeys.Shift))
                {
                    IDataObject d;
                    try { d = Clipboard.GetDataObject(); }
                    catch (Exception ex) { AddText("Không đọc được clipboard: " + ex.Message, Red); e.Handled = true; return; }
                    if (SafeTake(d, true)) e.Handled = true;
                }
            };
            input.TextChanged += (s, e) =>
            {
                hint.Visibility = input.Text == "" ? Visibility.Visible : Visibility.Collapsed;
                if (!filling) FilterSuggestions();
            };
            // gõ trong ô chat: giữ bàn phím; chuột rời palette: trả bàn phím cho AutoCAD
            // (bỏ qua khi đang mở danh sách combo - danh sách là cửa sổ popup, chuột vào đó cũng tính là rời Page)
            input.GotKeyboardFocus += (s, e) => Commands.SetKeepFocus(true);
            Page.MouseLeave += (s, e) => LeaveFocus();
            // ô chữ trong hội thoại tự nuốt lăn chuột -> cuộn khung hội thoại ở đây
            logScroll.PreviewMouseWheel += (s, e) => { logScroll.ScrollToVerticalOffset(logScroll.VerticalOffset - e.Delta / 3.0); e.Handled = true; };
            ((Button)Page.FindName("SendBtn")).Click += (s, e) => Send(input.Text);
            ((Button)Page.FindName("NewBtn")).Click += (s, e) => NewChat();
            stopBtn.Click += (s, e) => Stop(true, "stopped");
            ((Button)Page.FindName("AttachBtn")).Click += (s, e) => PickFiles();
            ((Button)Page.FindName("CaptureBtn")).Click += (s, e) => CaptureRegion(IntPtr.Zero);
            // kéo tệp vào bất kỳ chỗ nào của palette: sự kiện Preview ở gốc chạy trước TextBox/ScrollViewer
            Page.PreviewDragEnter += (s, e) => OnDragOver(e);
            Page.PreviewDragOver += (s, e) => OnDragOver(e);
            Page.PreviewDragLeave += (s, e) => ShowDropHint(false);  // cả khi Esc huỷ kéo
            Page.PreviewDrop += (s, e) =>
            {
                ShowDropHint(false);
                try
                {
                    if (HasFiles(e.Data)) { e.Handled = true; SafeTake(e.Data, false); input.Focus(); }  // R-A2: thả = chuột ở trên palette, an toàn để lấy focus
                }
                catch (Exception ex) { AddText("Lỗi kéo thả: " + ex.Message, Red); }
            };

            model = Combo("Model", new[] { Default, "sonnet", "opus", "haiku", "fable" }, Default);
            effort = Combo("Effort", new[] { Default, "low", "medium", "high", "xhigh", "max" }, Default);
            withSel = Combo("WithSelection", new[] { "Có", "Không" }, "Có");
            authMode = (ComboBox)Page.FindName("AuthMode");
            authMode.ItemsSource = AuthLabels;
            SelectAuthLabel();
            authMode.SelectionChanged += (s, e) =>
            {
                if (syncingAuth || authMode.SelectedIndex < 0) return;
                Save("AuthMode", authMode.SelectedIndex == 1 ? ApiKeyMode : AccountMode);
                AuthChanged();
            };
            ((UIElement)Page.FindName("ApiKeySave")).MouseLeftButtonUp += (s, e) => SaveKey();
            ((UIElement)Page.FindName("ApiKeyPaste")).MouseLeftButtonUp += (s, e) => PasteKey();
            ((UIElement)Page.FindName("ApiKeyClear")).MouseLeftButtonUp += (s, e) => { Delete("ApiKeyDpapi"); AuthChanged(); };
            var keyBox = (PasswordBox)Page.FindName("ApiKeyBox");
            keyBox.GotKeyboardFocus += (s, e) => Commands.SetKeepFocus(true);  // như ô chat / ô phím tắt: không có thì AutoCAD nuốt phím gõ vào ô
            keyBox.KeyDown += (s, e) => { if (e.Key == Key.Enter) { e.Handled = true; SaveKey(); } };
            RefreshAuth();
            hkCapture = HotkeyBox("HotkeyCapture");
            hkToggle = HotkeyBox("HotkeyToggle");
            ((UIElement)Page.FindName("EditSuggestions")).MouseLeftButtonUp += (s, e) => EditSuggestions();
            ((UIElement)Page.FindName("ReloadSuggestions")).MouseLeftButtonUp += (s, e) => LoadSuggestions();
            // hàng Skill: chữ trạng thái + link "Quản lý ▾" mở menu (ba handler skill giữ nguyên, chỉ đổi nơi gọi)
            skillStatus = (TextBlock)Page.FindName("SkillStatus");
            skillMenu = (ContextMenu)Page.FindName("SkillMenu");
            var manage = (TextBlock)Page.FindName("SkillManage");
            manage.MouseLeftButtonUp += (s, e) => { skillMenu.PlacementTarget = manage; skillMenu.Placement = System.Windows.Controls.Primitives.PlacementMode.Bottom; skillMenu.IsOpen = true; e.Handled = true; };
            ((MenuItem)Page.FindName("MenuOpenSkills")).Click += (s, e) => OpenSkillFolder();
            ((MenuItem)Page.FindName("MenuNewSkill")).Click += (s, e) => NewSkillFile();
            ((MenuItem)Page.FindName("MenuImportSkill")).Click += (s, e) => ImportSkillUi();
            ((Expander)Page.FindName("SettingsSection")).Expanded += (s, e) => RefreshSkillStatus();
            // menu / combo đóng khi chuột đã ở ngoài palette (bấm vào bản vẽ, Esc): MouseLeave đã chạy lúc popup còn mở nên không ai nhả bàn phím -> nhả ở đây.
            // Mức Input: Click của MenuItem chạy trước (ImportSkillUi kịp đặt modal)
            Action leaveSoon = () => Page.Dispatcher.BeginInvoke(DispatcherPriority.Input, new Action(() => { if (!Page.IsMouseOver) LeaveFocus(); }));
            skillMenu.Closed += (s, e) => leaveSoon();
            foreach (var cb in new[] { model, effort, withSel, authMode }) cb.DropDownClosed += (s, e) => leaveSoon();

            UpdateLabels();
            LoadSuggestions();
            RefreshSkillStatus();
            Welcome();
            ShowHotkeys(true);  // lỗi đăng ký phím lúc khởi động hiện sau lời chào
        }

        // Chuột rời palette, hoặc combo / menu vừa đóng khi chuột đã ở ngoài: trả bàn phím cho AutoCAD (quyết định ở Attach.LeaveAction).
        // Combo và menu là popup giữ capture: MouseLeave chạy lúc chúng còn mở nên bị bỏ qua, nên khi đóng phải gọi lại ở đây
        void LeaveFocus()
        {
            bool open = (model != null && model.IsDropDownOpen) || (effort != null && effort.IsDropDownOpen)
                || (withSel != null && withSel.IsDropDownOpen) || (authMode != null && authMode.IsDropDownOpen) || (skillMenu != null && skillMenu.IsOpen);
            int a = Attach.LeaveAction(modal, Page.IsKeyboardFocusWithin, open);
            if (a == Attach.LeaveKeepOff) Commands.SetKeepFocus(false);
            else if (a == Attach.LeaveRelease) Commands.ReleaseFocus();
        }

        public static Brush BrushOf(string hex) { return (Brush)new BrushConverter().ConvertFromString(hex); }

        const string RegKey = @"Software\ClaudeCad";
        public static string Load(string name, string def)
        {
            using (var k = Registry.CurrentUser.CreateSubKey(RegKey)) return k.GetValue(name) as string ?? def;
        }
        public static void Save(string name, string value)
        {
            using (var k = Registry.CurrentUser.CreateSubKey(RegKey)) k.SetValue(name, value);
        }
        public static void Delete(string name)
        {
            using (var k = Registry.CurrentUser.CreateSubKey(RegKey)) k.DeleteValue(name, false);
        }

        // ---------- đăng nhập: tài khoản Claude / API key (spec-3.1.0 §B) ----------

        const string AccountMode = "account", ApiKeyMode = "apikey";  // mã lưu registry (F2c)
        static readonly string[] AuthLabels = { "Tài khoản Claude", "API key" };  // chỉ để hiển thị
        internal static int LoginGen;  // tăng khi "claude auth login" xong
        readonly ComboBox authMode;
        bool syncingAuth;

        // null = giá trị lạ -> không spawn (fail-closed). Chưa có giá trị (nâng cấp từ 3.0.x) = account
        static string AuthModeValue()
        {
            object raw;
            using (var k = Registry.CurrentUser.CreateSubKey(RegKey)) raw = k.GetValue("AuthMode");
            if (raw == null) return AccountMode;  // chỉ khi KHÔNG có giá trị; sai kiểu (không phải chuỗi) = lạ
            string v = raw as string;
            return v == AccountMode || v == ApiKeyMode ? v : null;
        }
        // N1: chế độ thô + băm chuỗi blob (KHÔNG giải mã) + lần đăng nhập -> đổi gì thì tiến trình ấm không còn khớp Key
        static string AuthStamp() { return (Load("AuthMode", "") ?? "") + ":" + Attach.Sha8(Load("ApiKeyDpapi", "")) + ":" + LoginGen; }
        static bool HasKeyBlob() { return !string.IsNullOrEmpty(Load("ApiKeyDpapi", null)); }  // N2: không giải mã
        static string OfficialClaude() { return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), @".local\bin\claude.exe"); }

        // MỌI tiến trình claude (chat, auth status, auth login) tạo ở đây. Đọc chế độ đúng 1 lần, trả ra qua api. Lỗi -> ném, không spawn.
        // M2: key chỉ cần tại Process.Start; xoá khỏi psi để nó không nằm trong Process.StartInfo suốt đời tiến trình ấm
        static void DropKey(ProcessStartInfo psi) { psi.EnvironmentVariables.Remove("ANTHROPIC_API_KEY"); }

        static ProcessStartInfo NewPsi(string exe, string args, out bool api)
        {
            string mode = AuthModeValue();
            if (mode == null) throw new Exception("Cài đặt Đăng nhập có giá trị lạ: chọn lại Tài khoản Claude hoặc API key trong Cài đặt.");
            api = mode == ApiKeyMode;
            string key = null;
            if (api)
            {
                // N6: key chỉ đưa cho claude của bộ cài chính thức, không lấy từ PATH (ExeOverride: chỉ harness test)
                if (ExeOverride == null && !string.Equals(exe, OfficialClaude(), StringComparison.OrdinalIgnoreCase))
                    throw new Exception("Chế độ API key chỉ dùng claude.exe cài ở %USERPROFILE%\\.local\\bin (bộ cài chính thức). Chạy lại Setup.");
                string blob = Load("ApiKeyDpapi", null);
                if (string.IsNullOrEmpty(blob)) throw new Exception("Chưa có API key: dán key vào Cài đặt → Đăng nhập → API key rồi bấm Lưu.");
                key = Attach.Unprotect(blob);
                if (key == null || Attach.ApiKeyProblem(key) != null)  // F2a: hỏng -> không spawn, không tự xoá
                    throw new Exception("API key đã lưu bị hỏng hoặc không giải mã được (đổi máy / đổi tài khoản Windows?): dán lại key trong Cài đặt → Đăng nhập → API key.");
            }
            var psi = new ProcessStartInfo(exe, args) { UseShellExecute = false, WorkingDirectory = Home };
            Attach.AuthEnv(psi.EnvironmentVariables, key);
            return psi;
        }

        ComboBox Combo(string name, string[] items, string def)
        {
            var c = (ComboBox)Page.FindName(name);
            c.ItemsSource = items;
            string v = Load(name, def);
            c.SelectedItem = items.Contains(v) ? v : def;
            c.SelectionChanged += (s, e) =>
            {
                Save(name, (string)c.SelectedItem);
                UpdateLabels();
                if ((name == "Model" || name == "Effort") && turn == null) KillLive();  // I1: args đổi -> không dùng lại tiến trình ấm
            };
            return c;
        }

        // Cài đặt lưu chung (registry), mỗi UI có ô riêng: đổi bản vẽ thì nạp lại để ô không hiện giá trị cũ
        internal void SyncSettings()
        {
            SyncCombo(model, "Model", Default);
            SyncCombo(effort, "Effort", Default);
            SyncCombo(withSel, "WithSelection", "Có");
            SelectAuthLabel();
            RefreshAuth();
            ClearKeyBox();  // M5: key đang gõ dở không đi theo sang bản vẽ khác
            UpdateLabels();
            ShowHotkeys(false);  // phím tắt lưu chung: vẽ lại ô của UI này
        }

        static void SyncCombo(ComboBox c, string name, string def)
        {
            string v = Load(name, def);
            c.SelectedItem = ((string[])c.ItemsSource).Contains(v) ? v : def;
        }

        void SelectAuthLabel()
        {
            string m = AuthModeValue();
            syncingAuth = true;
            try { authMode.SelectedIndex = m == ApiKeyMode ? 1 : m == AccountMode ? 0 : -1; }  // lạ -> trống, user phải chọn lại
            finally { syncingAuth = false; }
        }

        void ClearKeyBox() { ((PasswordBox)Page.FindName("ApiKeyBox")).Clear(); }  // M5

        void SaveKey()
        {
            var box = (PasswordBox)Page.FindName("ApiKeyBox");
            string k = (box.Password ?? "").Trim();
            box.Clear();
            SaveKeyText(k);
        }

        // M3a: không phụ thuộc focus bàn phím (AutoCAD hay nuốt Ctrl+V). Không hiển thị, không log nội dung clipboard
        void PasteKey()
        {
            string k;
            try { k = (Clipboard.GetText() ?? "").Trim(); }
            catch (Exception) { ShowAuthStatus("Không đọc được clipboard (đang bị chương trình khác giữ). Thử lại.", true); return; }
            if (k == "") { ShowAuthStatus("Clipboard không có chữ: sao chép API key rồi bấm Dán & lưu.", true); return; }
            SaveKeyText(k);
        }

        // đường lưu duy nhất: ApiKeyProblem -> Protect -> registry
        void SaveKeyText(string k)
        {
            string p = Attach.ApiKeyProblem(k);
            if (p != null) { ShowAuthStatus(p, true); return; }
            try { Save("ApiKeyDpapi", Attach.Protect(k)); }
            catch (Exception ex) { ShowAuthStatus("Không lưu được key: " + Attach.Redact(ex.Message), true); return; }
            ClearKeyBox();
            AuthChanged();
        }

        void AuthChanged() { Commands.KillIdleAll(); RefreshAuth(); }  // mọi bản vẽ, không chỉ bản vẽ này
        internal void KillIdle() { if (turn == null) KillLive(); }

        void ShowAuthStatus(string text, bool error)
        {
            var st = (TextBlock)Page.FindName("AuthStatus");
            st.Text = text;
            st.Foreground = error ? Red : Dim;
        }

        internal void RefreshAuth()
        {
            string m = AuthModeValue();
            var vis = m == ApiKeyMode ? Visibility.Visible : Visibility.Collapsed;
            ((UIElement)Page.FindName("ApiKeyLabel")).Visibility = vis;
            ((UIElement)Page.FindName("ApiKeyVal")).Visibility = vis;
            if (m != ApiKeyMode) ClearKeyBox();  // M5: ô bị ẩn thì không giữ chữ trong đó
            ShowAuthStatus(m == null ? "Giá trị Đăng nhập lạ: chọn lại Tài khoản Claude hoặc API key."
                : m == AccountMode ? "Dùng tài khoản Claude đã đăng nhập (gõ /login để đăng nhập hoặc đổi tài khoản). Bỏ qua ANTHROPIC_API_KEY của Windows."
                : HasKeyBlob() ? "Đã lưu key (mã hoá theo tài khoản Windows). Lỗi giải mã sẽ báo khi gửi."
                : "Chưa có key: dán key rồi bấm Lưu. Không dùng ANTHROPIC_API_KEY của Windows.", m == null);
        }

        void UpdateLabels()
        {
            string m = (string)model.SelectedItem, ef = (string)effort.SelectedItem;
            modelLabel.Text = (m == Default ? "model mặc định" : m) + (ef == Default ? "" : " · " + ef);
        }

        // ---------- phím tắt (spec-hotkeys §6-8) ----------

        public static void ApplyHotkeys()
        {
            Hotkeys.Apply(Load("HotkeyCapture", Hotkeys.DefaultCapture), Load("HotkeyToggle", Hotkeys.DefaultToggle));
        }

        internal void Error(string msg) { AddText(msg, Red); }

        TextBox HotkeyBox(string name)
        {
            var box = (TextBox)Page.FindName(name);
            // bấm vào ô = chuột đang trên palette -> giữ bàn phím được (R-A2). Nhả phím toàn hệ thống để bấm lại đúng tổ hợp đang dùng.
            box.GotKeyboardFocus += (s, e) =>
            {
                Commands.SetKeepFocus(true);
                Hotkeys.Suspend(true);
                box.Text = "Bấm tổ hợp phím…";
                box.Foreground = Dim;
            };
            box.LostKeyboardFocus += (s, e) =>
            {
                try { Hotkeys.Suspend(false); ShowHotkeys(true); }
                catch (Exception ex) { AddText("Phím tắt: " + ex.Message, Red); }
            };
            box.PreviewKeyDown += (s, e) =>
            {
                e.Handled = true;
                if (e.IsRepeat) return;  // giữ phím: không báo lỗi lặp
                try
                {
                    Key k = e.Key == Key.System ? e.SystemKey : e.Key == Key.ImeProcessed ? e.ImeProcessedKey
                        : e.Key == Key.DeadCharProcessed ? e.DeadCharProcessedKey : e.Key;
                    if (k == Key.Escape) { DropFocus(box); return; }
                    if (ApplyHotkeyInput(name, k, Keyboard.Modifiers)) DropFocus(box);  // mất focus -> đăng ký lại ngay
                }
                catch (Exception ex) { AddText("Phím tắt: " + ex.Message, Red); }
            };
            ((UIElement)Page.FindName(name + "Clear")).MouseLeftButtonUp += (s, e) => ClearHotkey(name);
            return box;
        }

        // xoá cả focus logic: nếu không, kích hoạt lại palette trả focus vào ô -> tự vào chế độ bấm phím (và nhả phím chụp vùng)
        static void DropFocus(TextBox box)
        {
            FocusManager.SetFocusedElement(FocusManager.GetFocusScope(box), null);
            Commands.ReleaseFocus();
        }

        // true = đã nhận (lưu + áp dụng); false = mới bấm phím bổ trợ, hoặc bị từ chối (đã báo đỏ)
        internal bool ApplyHotkeyInput(string name, Key key, ModifierKeys mods)
        {
            if (key == Key.LeftCtrl || key == Key.RightCtrl || key == Key.LeftAlt || key == Key.RightAlt
                || key == Key.LeftShift || key == Key.RightShift || key == Key.LWin || key == Key.RWin) return false;
            if (key == Key.Back) return false;  // UniKey dùng Backspace sửa dấu: bỏ qua. Tắt bằng Delete hoặc ×
            if (key == Key.Delete && mods == ModifierKeys.None) { ClearHotkey(name); return true; }
            int vk = KeyInterop.VirtualKeyFromKey(key);
            if (vk == 0) return false;  // Key.None / VK_PACKET (bộ gõ chèn phím): không được thành "tắt"
            var k = new HotKey((int)mods, vk);
            bool global = name == "HotkeyCapture";
            string err = HotKey.Problem(k, global);
            HotKey other = global ? Hotkeys.ToggleKey : Hotkeys.CaptureKey;
            if (err == null && other.Matches(k.Vk, k.Mods)) err = k + " đang dùng cho phím " + (global ? "ẩn/hiện palette." : "chụp vùng.");
            if (err != null) { AddText("Phím tắt: " + err, Red); return false; }
            Save(name, k.ToString());
            ApplyHotkeys();
            ShowHotkeys(true);
            return true;
        }

        internal void ClearHotkey(string name)
        {
            Save(name, "");
            ApplyHotkeys();
            ShowHotkeys(false);
        }

        internal void ShowHotkeys(bool log)
        {
            string err = Hotkeys.CaptureError;
            Paint(hkCapture, Hotkeys.CaptureKey, err);
            string terr = Hotkeys.ToggleError;
            Paint(hkToggle, Hotkeys.ToggleKey, terr);
            if (log && err != null && err != shownHotkeyError) AddText("Phím chụp vùng: " + err, Red);
            if (log && terr != null && terr != shownToggleError) AddText("Phím ẩn/hiện palette: " + terr, Red);
            shownHotkeyError = err;
            shownToggleError = terr;
        }

        void Paint(TextBox b, HotKey k, string err)
        {
            if (b.IsKeyboardFocused) return;  // đang chờ bấm phím
            b.Text = (k.IsNone ? "Tắt" : k.ToString()) + (err != null ? " (lỗi)" : "");
            b.Foreground = err != null ? Red : k.IsNone ? Dim : (Brush)Page.FindResource("Text");
            b.ToolTip = err ?? "Bấm vào rồi bấm tổ hợp phím (vd Ctrl+Shift+X). Delete hoặc × để tắt, Esc để thôi.";
        }

        // ---------- gợi ý lệnh ----------

        public void EditSuggestions()
        {
            if (!File.Exists(SuggestFile)) LoadSuggestions();
            Process.Start(Path.Combine(Environment.SystemDirectory, "notepad.exe"), "\"" + SuggestFile + "\"");
        }

        // ---------- skill (3.1.1) ----------

        // Chữ trạng thái hàng Skill: ShowSkillStatus nhận kết quả SkillMode / SkillExpected ĐÃ CÓ (StartTurn); RefreshSkillStatus tự quét một lần
        // (khi nạp palette, mở rộng Cài đặt, sau Tạo / Nạp skill thành công). Không chỗ nào khác quét thêm
        void ShowSkillStatus(bool on, IList<string> names, string why)
        {
            if (skillStatus == null) return;
            string tip;
            skillStatus.Text = Attach.SkillStatusText(on, names, why, skillsBlocked, out tip);
            skillStatus.ToolTip = tip;
        }

        void RefreshSkillStatus()
        {
            try
            {
                string why = null;
                bool on = !skillsBlocked && Attach.SkillMode(Home, out why);
                ShowSkillStatus(on, on ? Attach.SkillExpected(Home) : null, why);
            }
            catch (Exception) { }
        }

        // Tạo <Home>\.claude\skills nếu thiếu. .claude hoặc skills là liên kết (junction/symlink) -> null (không mở / không tạo gì thêm). Lỗi IO: ném cho người gọi.
        // Lỗi KHÔNG đưa ex.Message lên UI (có thể chứa đường dẫn): người gọi hiện hằng Attach.WhySkillIo / WhySkillDir
        static string EnsureSkillsDir()
        {
            foreach (string d in new[] { Path.GetDirectoryName(SkillsDir), SkillsDir })
            {
                if (!Directory.Exists(d)) Directory.CreateDirectory(d);
                else if (Attach.IsReparsePoint(d)) return null;
            }
            return SkillsDir;
        }

        void OpenSkillFolder()
        {
            try
            {
                string dir = EnsureSkillsDir();
                if (dir == null) { AddText(Attach.WhySkillDir, Red); return; }
                Process.Start(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "explorer.exe"), "\"" + dir + "\"");
            }
            catch (Exception) { AddText(Attach.WhySkillIo, Red); }
        }

        void NewSkillFile()
        {
            try
            {
                string dir = EnsureSkillsDir();
                if (dir == null) { AddText(Attach.WhySkillDir, Red); return; }
                Process.Start(Path.Combine(Environment.SystemDirectory, "notepad.exe"), "\"" + Attach.NewSkill(dir) + "\"");
                RefreshSkillStatus();
            }
            catch (Exception) { AddText(Attach.WhySkillIo, Red); }
        }

        // Nạp skill từ SKILL.md (nguồn = thư mục chứa nó) hoặc .zip (giải nén vào thư mục tạm GUID dưới %TEMP%, xoá trong finally).
        // Mọi kiểm tra / chép nằm trong Attach.UnzipSkill / ImportSkill. Bản nháp AI soạn (skill-nhap) nạp bằng đúng đường này, thành công thì xoá bản nháp
        // ponytail: giải nén / quét / chép chạy ĐỒNG BỘ trên luồng UI (≤ 200 tệp, 20 MB); trần = Home trên ổ cục bộ. Nâng cấp sang worker + tiến độ khi Home ở ổ mạng
        void ImportSkillUi()
        {
            if (modal) return;
            string tmpRoot = null;
            try
            {
                string skills = EnsureSkillsDir();
                if (skills == null) { AddText(Attach.WhySkillDir, Red); return; }
                var d = new OpenFileDialog { Title = "Nạp skill", Filter = "SKILL.md|SKILL.md|Gói skill (*.zip)|*.zip" };
                if (Attach.DraftNames(DraftsDir).Count > 0) d.InitialDirectory = DraftsDir;
                bool ok;
                modal = true;
                try { ok = d.ShowDialog() == true; }
                finally { modal = false; }
                if (!ok) return;
                string srcDir, why, name;
                if (d.FileName.EndsWith(".zip", StringComparison.OrdinalIgnoreCase))
                {
                    tmpRoot = Path.Combine(Path.GetTempPath(), "cad-skill-nap-" + Guid.NewGuid().ToString("N"));
                    string x = Path.Combine(tmpRoot, Path.GetFileNameWithoutExtension(d.FileName));  // tên thư mục = tên zip: dự phòng khi SKILL.md không có name:
                    Directory.CreateDirectory(x);
                    if (!Attach.UnzipSkill(d.FileName, x, out srcDir, out why)) { AddText(why, Red); return; }
                }
                else srcDir = Path.GetDirectoryName(d.FileName);
                bool fromDraft = string.Equals(Path.GetDirectoryName(srcDir), DraftsDir, StringComparison.OrdinalIgnoreCase);  // nguồn = <Home>\skill-nhap\<tên>
                // B5: xác nhận (tên, mô tả, số tệp, số .py/.lsp, xem trước nội dung) sau khi bản chép đã qua kiểm, trước khi Move
                Func<Attach.SkillInfo, bool> confirm = i =>
                {
                    modal = true;
                    try
                    {
                        // cảnh báo TRƯỚC phần nội dung (nội dung dài không được đẩy cảnh báo ra khỏi hộp); Preview đã cắt ≤ 25 dòng / 1000 ký tự
                        return MessageBox.Show("Nạp skill \"" + i.Name + "\"?\n\n"
                            + "Skill là chương trình chạy với quyền của bạn — chỉ nạp skill bạn đã đọc hoặc tin nguồn.\n"
                            + (fromDraft ? "Bản nháp do AI soạn — chỉ nạp nếu chính bạn vừa yêu cầu soạn.\n" : "")
                            + "\nMô tả: " + i.Description + "\nSố tệp: " + i.Files + " (trong đó .py/.lsp: " + i.Scripts + ")\n\n"
                            + "Nội dung (" + i.Preview.Length + "/" + i.Chars + " ký tự):\n" + i.Preview,
                            "Nạp skill", MessageBoxButton.YesNo, MessageBoxImage.Warning, MessageBoxResult.No) == MessageBoxResult.Yes;
                    }
                    finally { modal = false; }
                };
                bool done = Attach.ImportSkill(srcDir, skills, false, out name, out why, confirm);
                if (!done && why == Attach.WhySkillExists)
                {
                    modal = true;
                    MessageBoxResult r;
                    try { r = MessageBox.Show("Đã có skill cùng tên. Ghi đè?", "Nạp skill", MessageBoxButton.YesNo, MessageBoxImage.Question, MessageBoxResult.No); }
                    finally { modal = false; }
                    if (r != MessageBoxResult.Yes) return;
                    done = Attach.ImportSkill(srcDir, skills, true, out name, out why, confirm);
                }
                if (!done) { if (why == Attach.WhySkillCancel) Note(why, HorizontalAlignment.Left); else AddText(why, Red); return; }
                AddText("Đã nạp skill " + name, null);
                RefreshSkillStatus();
                // nạp từ bản nháp (<Home>\skill-nhap\<tên>\SKILL.md): xoá bản nháp đã duyệt
                if (fromDraft && !Attach.IsReparsePoint(srcDir))
                {
                    try { Directory.Delete(srcDir, true); } catch (Exception) { }
                }
            }
            catch (Exception) { AddText(Attach.WhySkillIo, Red); }
            finally { try { if (tmpRoot != null && Directory.Exists(tmpRoot)) Directory.Delete(tmpRoot, true); } catch (Exception) { } }
        }

        // Sau mỗi lượt: có bản nháp AI soạn thì báo MỘT dòng (chỉ khi danh sách đổi, để không lặp mỗi lượt)
        void ShowDrafts()
        {
            var names = Attach.DraftNames(DraftsDir);
            string s = string.Join(", ", names);
            if (s == lastDrafts) return;
            lastDrafts = s;
            if (names.Count > 0) Note("AI đã soạn skill chờ duyệt: " + s + " — bấm Nạp skill… để xem và bật (chỉ nạp nếu chính bạn vừa yêu cầu soạn)", HorizontalAlignment.Left);
        }

        void LoadSuggestions()
        {
            try
            {
                if (!File.Exists(SuggestFile)) File.WriteAllText(SuggestFile, DefaultSuggestions, Encoding.UTF8);
                allSuggestions = File.ReadAllLines(SuggestFile, Encoding.UTF8).Select(l => l.Trim()).Where(l => l != "" && !l.StartsWith(";")).ToList();
            }
            catch (Exception ex) { AddText("Không đọc được goi_y.txt: " + ex.Message, Red); }
            FilterSuggestions();
        }

        // Lọc theo chữ đang gõ: mọi từ gõ vào đều phải có trong gợi ý, không phân biệt dấu
        void FilterSuggestions()
        {
            string[] words = Fold(input.Text).Split(new[] { ' ', '\n', '\r', '\t' }, StringSplitOptions.RemoveEmptyEntries);
            suggestions.Children.Clear();
            int n = 0;
            foreach (string s in allSuggestions)
            {
                string f = Fold(s);
                if (words.All(w => f.Contains(w))) { n++; suggestions.Children.Add(SuggestionRow(s)); }
            }
            sugCount.Text = "(" + n + "/" + allSuggestions.Count + ")";
        }

        public static string Fold(string s)
        {
            var sb = new StringBuilder();
            foreach (char c in s.ToLowerInvariant().Replace('đ', 'd').Normalize(NormalizationForm.FormD))
                if (CharUnicodeInfo.GetUnicodeCategory(c) != UnicodeCategory.NonSpacingMark) sb.Append(c);
            return sb.ToString();
        }

        // Bấm: điền vào ô nhập để sửa thêm. Bấm đúp: gửi ngay
        UIElement SuggestionRow(string s)
        {
            var row = new Border { Style = (Style)Page.FindResource("Item"), ToolTip = "Bấm: điền vào ô nhập · Bấm đúp: gửi ngay",
                Child = new TextBlock { Text = s, TextWrapping = TextWrapping.Wrap } };
            row.MouseLeftButtonDown += (o, e) =>
            {
                if (e.ClickCount == 2) { Send(s); return; }
                filling = true;
                input.Text = s;
                filling = false;
                input.CaretIndex = s.Length;
                input.Focus();
            };
            return row;
        }

        // ---------- hội thoại ----------

        void Append(UIElement e)
        {
            log.Children.Add(e);
            logScroll.ScrollToEnd();
        }

        TextBox Msg(string text, Brush fg)
        {
            var t = new TextBox { Text = text, Style = (Style)Page.FindResource("Msg") };
            if (fg != null) t.Foreground = fg;
            return t;
        }

        void AddText(string text, Brush fg)
        {
            var t = Msg(text, fg);
            t.Margin = new Thickness(0, 4, 0, 4);
            Append(t);
        }

        void Note(string text, HorizontalAlignment align)
        {
            Append(new TextBlock { Text = text, Foreground = Dim, FontSize = 12,Margin = new Thickness(2, 1, 2, 1), TextWrapping = TextWrapping.Wrap, HorizontalAlignment = align });
        }

        void Welcome()
        {
            AddText("Chọn một gợi ý (bấm đúp để gửi ngay) hoặc nhập yêu cầu bên dưới. Claude đọc và sửa trực tiếp bản vẽ đang mở.", Dim);
        }

        // ---------- tệp đính kèm (spec §4-6) ----------

        sealed class Batch { public int Left, Shown, Hidden; }  // 1 lần thả/dán/chọn: tối đa 5 dòng đỏ

        // R5: mọi việc đọc/giải mã tệp chạy trên 1 luồng nền STA; kết quả về UI qua Dispatcher (AutoCAD không đứng)
        void Enqueue(Action job)
        {
            if (worker == null)
            {
                worker = new Thread(() =>
                {
                    foreach (var j in jobs.GetConsumingEnumerable())
                        try { j(); } catch (Exception) { }  // lưới an toàn; job tự báo lỗi
                }) { IsBackground = true, Name = "ClaudeAttach" };
                worker.SetApartmentState(ApartmentState.STA);
                worker.Start();
            }
            try { jobs.Add(job); } catch (InvalidOperationException) { }  // đã Shutdown (CompleteAdding)
        }

        // chip "đang đọc…" ngay; đọc thật trên luồng nền; × trên chip = huỷ (kill tiến trình .xls)
        void Load(string name, string sourcePath, Batch batch, Func<CancellationToken, Attachment> read)
        {
            var ph = new Attachment { Name = name, SourcePath = sourcePath, Kind = "loading", Detail = "Đang đọc tệp…" };
            var cts = new CancellationTokenSource();
            loading[ph] = cts;
            pending.Add(ph);
            batch.Left++;
            RefreshChips();
            var token = cts.Token;
            Enqueue(() =>
            {
                Attachment a = null;
                string err = null;
                var sw = Stopwatch.StartNew();  // T1: thời gian đọc/giải mã (vào nhật ký hiệu năng, không có tên tệp)
                try { if (!token.IsCancellationRequested) { a = read(token); if (a != null) a.LoadMs = sw.ElapsedMilliseconds; } }
                catch (OperationCanceledException) { }
                catch (AttachException ex) { err = ex.Message; }
                catch (Exception ex) { err = name + ": không đọc được tệp (" + ex.Message + ")."; }  // chỉ hiện trên UI, không vào prompt
                Page.Dispatcher.BeginInvoke((Action)(() =>
                {
                    try { Finish(ph, a, err, batch); }
                    catch (Exception ex) { try { AddText("Lỗi khi thêm tệp đính kèm: " + ex.Message, Red); } catch (Exception) { } }  // m1
                }));
            });
        }

        void Finish(Attachment ph, Attachment a, string err, Batch batch)
        {
            batch.Left--;
            CancellationTokenSource cts;
            bool alive = loading.TryGetValue(ph, out cts);  // false = đã bị × / Mới huỷ
            if (alive) { loading.Remove(ph); cts.Dispose(); }
            int i = pending.IndexOf(ph);
            if (alive && i >= 0)
            {
                if (err == null && a != null && pending.Sum(x => x.Bytes) + a.Bytes > Attach.MaxTotalBytes)
                    err = a.Name + ": tổng dung lượng đính kèm vượt " + Attach.HumanSize(Attach.MaxTotalBytes) + ". Gửi bớt rồi đính kèm tiếp.";
                if (err != null || a == null) pending.RemoveAt(i);
                else
                {
                    pending[i] = a;
                    if (a.Warning != null) Note(a.Name + ": " + a.Warning, HorizontalAlignment.Left);
                }
                RefreshChips();
                if (err != null) { if (batch.Shown < 5) { batch.Shown++; AddText(err, Red); } else batch.Hidden++; }
            }
            if (batch.Left == 0 && batch.Hidden > 0)
            {
                AddText("… và " + batch.Hidden + " tệp khác không đính kèm được.", Red);
                batch.Hidden = 0;
            }
        }

        // R5 bất biến: đường dẫn chỉ đến từ thao tác user (thả, dán FileDrop, hộp chọn tệp), không bao giờ từ tool / nội dung tệp
        internal void AddFiles(IEnumerable<string> paths)
        {
            var batch = new Batch();
            foreach (string p in paths)
            {
                string path = p;
                if (pending.Any(x => string.Equals(x.SourcePath, path, StringComparison.OrdinalIgnoreCase))) continue;  // đã có: bỏ qua
                if (pending.Count >= Attach.MaxAttachments) { AddText("Tối đa " + Attach.MaxAttachments + " tệp mỗi tin nhắn.", Red); break; }
                Load(Attach.SafeName(Path.GetFileName(path)), path, batch, ct => Attach.FromFile(path, ct));
            }
            // R-A2: KHÔNG input.Focus() ở đây: chuột đang ở vùng vẽ (sau hộp chọn tệp) -> GotKeyboardFocus -> SetKeepFocus(true) cướp bàn phím của AutoCAD
        }

        void AddImage(BitmapSource img, string prefix)
        {
            if (pending.Count >= Attach.MaxAttachments) { AddText("Tối đa " + Attach.MaxAttachments + " tệp mỗi tin nhắn.", Red); return; }
            string name = prefix + (++shotNo) + ".png";
            BitmapSource frozen;
            try { frozen = Attach.FrozenCopy(img); }  // ≤ 100 MP; bản sao Freeze để sang luồng nền
            catch (Exception ex) { AddText(name + ": " + ex.Message, Red); return; }
            string type = prefix == "chup-vung-" ? "capture" : "paste";
            Load(name, null, new Batch(), ct => { var a = Attach.FromBitmap(frozen, name); a.Type = type; return a; });
        }

        void RemoveAttachment(Attachment a)
        {
            CancellationTokenSource cts;
            if (loading.TryGetValue(a, out cts)) { cts.Cancel(); loading.Remove(a); }
            pending.Remove(a);
            RefreshChips();
        }

        // D1: autoConvert=false ở mọi lần hỏi/đọc IDataObject: không để WPF tự chuyển sang định dạng tuỳ biến (có thể đi qua BinaryFormatter)
        static bool HasFiles(IDataObject d) { return d != null && d.GetDataPresent(DataFormats.FileDrop, false); }

        // Ctrl+V và kéo thả cùng đi qua đây. true = đã xử lý (TextBox không dán chữ nữa)
        // R2: chỉ FileDrop / chữ / Bitmap; KHÔNG đọc định dạng tuỳ biến (như PNG đăng ký riêng) vì GetData có thể chạy BinaryFormatter
        internal bool TakeData(IDataObject d, bool fromPaste)
        {
            if (d == null) return false;
            if (HasFiles(d))
            {
                var files = d.GetData(DataFormats.FileDrop, false) as string[];
                if (files != null) AddFiles(files);
                return true;
            }
            if (!fromPaste) return false;  // kéo thả chỉ nhận tệp
            if (d.GetDataPresent(DataFormats.UnicodeText, false) || d.GetDataPresent(DataFormats.Text, false)) return false;  // chữ: dán như cũ
            if (!d.GetDataPresent(DataFormats.Bitmap, false)) return false;
            var b = d.GetData(DataFormats.Bitmap, false) as BitmapSource;
            if (b == null) return false;
            AddImage(new FormatConvertedBitmap(b, PixelFormats.Bgr32, null, 0), "anh-dan-");  // DIB hay có alpha = 0 -> bỏ alpha
            return true;
        }

        bool SafeTake(IDataObject d, bool fromPaste)
        {
            try { return TakeData(d, fromPaste); }
            catch (Exception ex) { AddText("Không đọc được dữ liệu dán/thả: " + ex.Message, Red); return true; }
        }

        void OnDragOver(DragEventArgs e)
        {
            try
            {
                bool files = HasFiles(e.Data);
                ShowDropHint(files);
                if (!files) return;  // kéo chữ: để TextBox xử lý như cũ
                e.Effects = DragDropEffects.Copy;
                e.Handled = true;
            }
            catch (Exception ex) { ShowDropHint(false); AddText("Lỗi kéo thả: " + ex.Message, Red); }
        }

        internal void ShowDropHint(bool on) { dropHint.Visibility = on ? Visibility.Visible : Visibility.Collapsed; }

        void PickFiles()
        {
            try
            {
                var d = new OpenFileDialog
                {
                    Multiselect = true, Title = "Đính kèm tệp cho Claude",
                    Filter = "Tệp hỗ trợ|*.png;*.jpg;*.jpeg;*.bmp;*.gif;*.webp;*.pdf;*.txt;*.csv;*.json;*.lsp;*.py;*.md;*.xlsx;*.xls|Tất cả tệp|*.*"
                };
                bool ok;
                modal = true;
                try { ok = d.ShowDialog() == true; }
                finally { modal = false; }
                if (ok) AddFiles(d.FileNames);
            }
            catch (Exception ex) { AddText("Không mở được hộp chọn tệp: " + ex.Message, Red); }
        }

        // Q2: ẩn palette -> chờ AutoCAD vẽ lại -> overlay (owner = cửa sổ AutoCAD, R7) -> hiện lại palette. R6: bắt mọi lỗi
        // returnTo: cửa sổ đứng trước lúc bấm phím tắt (IntPtr.Zero: nút / lệnh CLAUDECHUP). Spec-hotkeys §6.1.
        internal void CaptureRegion(IntPtr returnTo)
        {
            if (modal) return;
            modal = true;
            bool wasVisible = Commands.PaletteVisible;
            try { Commands.SetPaletteVisible(false); } catch (Exception) { }
            // palette đang ẩn (vd bấm phím khi đang xem PDF) -> không cần chờ vẽ lại, chụp ngay
            var t = new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(wasVisible ? CaptureSettleMs : 0) };
            t.Tick += (s, e) =>
            {
                t.Stop();
                BitmapSource b = null;
                try
                {
                    IntPtr acad = SafeAcadWindow();
                    if (acad != IntPtr.Zero && Native.IsIconic(acad)) acad = IntPtr.Zero;  // owner thu nhỏ -> overlay bị ẩn theo
                    b = RegionCapture.Pick(acad);
                }
                catch (Exception ex) { AddText("Không chụp được màn hình: " + ex.Message, Red); }
                finally
                {
                    modal = false;
                    try { Commands.SetPaletteVisible(wasVisible || b != null); } catch (Exception) { }
                }
                // D5: không để ngoại lệ thoát ra khỏi Tick của DispatcherTimer (sập AutoCAD)
                try
                {
                    if (b != null)
                    {
                        AddImage(b, "chup-vung-");  // Q3: không kéo AutoCAD lên (overlay thuộc AutoCAD, đóng overlay thì focus về AutoCAD). R-A2: KHÔNG input.Focus()
                    }
                    else if (returnTo != IntPtr.Zero && !Native.IsOwnProcess(returnTo))
                        Native.SetForegroundWindow(returnTo);  // huỷ -> trả lại PDF/Zalo… như chưa bấm gì
                }
                catch (Exception ex) { AddText("Không đính kèm được ảnh chụp: " + ex.Message, Red); }
            };
            t.Start();
        }

        static IntPtr SafeAcadWindow()
        {
            try { return AcadWindow(); } catch (Exception) { return IntPtr.Zero; }  // ngoài AutoCAD
        }

        [MethodImpl(MethodImplOptions.NoInlining)]
        static IntPtr AcadWindow() { return Acad.MainWindow.Handle; }

        void RefreshChips()
        {
            chips.Children.Clear();
            foreach (var a in pending) chips.Children.Add(Chip(a, true));
            chips.Visibility = pending.Count > 0 ? Visibility.Visible : Visibility.Collapsed;
        }

        // chip: tên · cỡ (hoặc "đang đọc…") · × (removable). Tooltip: ảnh xem trước + chi tiết (có số sheet ẩn) + cảnh báo
        UIElement Chip(Attachment a, bool removable)
        {
            bool busy = a.Kind == "loading";
            var dp = new DockPanel();
            if (removable)
            {
                var x = new TextBlock { Text = "×", Foreground = Dim, Margin = new Thickness(6, 0, 0, 0), Cursor = Cursors.Hand,
                    ToolTip = busy ? "Huỷ" : "Bỏ tệp này", VerticalAlignment = VerticalAlignment.Center };
                x.MouseEnter += (s, e) => x.Foreground = Red;
                x.MouseLeave += (s, e) => x.Foreground = Dim;
                x.MouseLeftButtonUp += (s, e) => RemoveAttachment(a);
                DockPanel.SetDock(x, Dock.Right);
                dp.Children.Add(x);
            }
            var size = new TextBlock
            {
                Text = busy ? "đang đọc…" : Attach.HumanSize(a.Bytes) + (a.Warning != null ? " · cắt bớt" : ""),
                Foreground = a.Warning != null ? Red : Dim, Margin = new Thickness(6, 0, 0, 0), VerticalAlignment = VerticalAlignment.Center
            };
            DockPanel.SetDock(size, Dock.Right);
            dp.Children.Add(size);
            dp.Children.Add(new TextBlock { Text = a.Name, MaxWidth = 150, TextTrimming = TextTrimming.CharacterEllipsis, VerticalAlignment = VerticalAlignment.Center });
            var tip = new StackPanel { MaxWidth = 260 };
            if (a.Thumb != null) tip.Children.Add(new Image { Source = a.Thumb, MaxWidth = 240, MaxHeight = 180, Stretch = Stretch.Uniform, HorizontalAlignment = HorizontalAlignment.Left });
            tip.Children.Add(new TextBlock { Text = a.Name + "\n" + a.Detail + (a.Warning != null ? "\n" + a.Warning : ""), TextWrapping = TextWrapping.Wrap });
            return new Border { Style = (Style)Page.FindResource("Chip"), Child = dp, ToolTip = tip };
        }

        WrapPanel ChipPanel(IEnumerable<Attachment> list, bool removable)
        {
            var w = new WrapPanel { Margin = new Thickness(0, 4, 0, 0) };
            foreach (var a in list) w.Children.Add(Chip(a, removable));
            return w;
        }

        // bong bóng tin nhắn user trong log: chữ + chip tệp (không có ×)
        UIElement UserBubble(string prompt, IList<Attachment> sent)
        {
            var body = new StackPanel();
            if (prompt != "") body.Children.Add(Msg(prompt, null));
            if (sent.Count > 0) body.Children.Add(ChipPanel(sent, false));
            return new Border { Background = Bubble, BorderBrush = BubbleEdge, BorderThickness = new Thickness(2, 0, 0, 0),
                Padding = new Thickness(8, 4, 6, 4), Margin = new Thickness(0, 10, 0, 2), Child = body };
        }

        // Đối tượng đang chọn trong CAD. Tách hàm riêng: chỉ nạp API AutoCAD khi thực sự gọi
        [MethodImpl(MethodImplOptions.NoInlining)]
        static List<string> SelectedHandles()
        {
            var r = Acad.DocumentManager.MdiActiveDocument.Editor.SelectImplied();
            var list = new List<string>();
            if (r.Status == Autodesk.AutoCAD.EditorInput.PromptStatus.OK)
                foreach (var id in r.Value.GetObjectIds()) list.Add(id.Handle.ToString());
            return list;
        }

        // Bản vẽ đang mở lúc bấm Gửi -> ghi ra file để server.py GHIM bản vẽ làm việc
        // (user chuyển tab trong lúc Claude chạy thì Claude vẫn làm đúng bản vẽ này).
        // C1: mỗi ClaudeUi (mỗi bản vẽ) một tệp riêng cad_claude_doc_<pid>_<n>.txt, đưa cho server.py qua biến môi trường CAD_CLAUDE_PIN_FILE
        // (không dùng tệp chung: server.py của bản sao khác / Claude Code CLI sẽ bị ghim nhầm). Xoá ở Shutdown / Terminate.
        static int PidOfThis()
        {
            using (var p = Process.GetCurrentProcess()) return p.Id;
        }

        // Xoá %TEMP%\cad_claude_doc_<pid>[_<n>].txt của AutoCAD đã chết. CHỈ tên khớp Attach.TempPid; pid còn là acad.exe thì giữ
        static void CleanStalePins()
        {
            try
            {
                int me = PidOfThis();
                foreach (string f in Directory.GetFiles(Path.GetTempPath(), "cad_claude_doc_*.txt"))
                {
                    int pid = Attach.TempPid(Path.GetFileName(f));
                    if (pid < 0 || pid == me) continue;
                    bool alive = false;
                    try { using (var p = Process.GetProcessById(pid)) alive = string.Equals(p.ProcessName, "acad", StringComparison.OrdinalIgnoreCase) && !p.HasExited; }
                    catch (ArgumentException) { }          // không có tiến trình đó
                    catch (Exception) { alive = true; }    // không đọc được (quyền): không xoá
                    if (!alive) { try { File.Delete(f); } catch (Exception) { } }
                }
            }
            catch (Exception) { }
        }

        internal void DeletePin()
        {
            try { File.Delete(PinFile); } catch (Exception) { }
        }

        [MethodImpl(MethodImplOptions.NoInlining)]
        static string ActiveDocName()
        {
            var d = Acad.DocumentManager.MdiActiveDocument;
            return d == null ? null : d.Name;
        }

        static string DocLabel(string dn)
        {
            try { return Attach.SafeName(Path.GetFileName(dn)); }
            catch (Exception) { return "?"; }
        }

        void Send(string text)
        {
            string prompt = text.Trim();
            if (prompt == "" && pending.Count == 0) return;
            if (prompt.Equals("/login", StringComparison.OrdinalIgnoreCase)) { Stop(true, "stopped"); KillLive(); input.Clear(); Login(); return; }  // I3: /login cũng bỏ tiến trình đang giữ
            // 3.1.1: lệnh gạch chéo (luật trong Attach.SlashSkill). Chế độ skill TẮT: chỉ "/skill ..." bị chặn, mọi "/" khác gửi nguyên văn như 3.1.0 ("/login" xử lý ở trên).
            // Chế độ skill BẬT: "/<tên>" / "/skill <tên>" gọi skill; "/" khác bị chặn (không còn --disable-slash-commands). Chặn = chỉ Note, KHÔNG gửi, giữ chữ trong ô nhập
            string slashPrompt = null;
            if (prompt.StartsWith("/", StringComparison.Ordinal))
            {
                string sw = null;
                bool skillsOn = !skillsBlocked && Attach.SkillMode(Home, out sw);
                string sp, sWhy;
                if (Attach.SlashSkill(prompt, skillsOn, skillsOn ? Attach.SkillExpected(Home) : null, skillsBlocked ? Attach.SkillBlockedWhy : sw, out sp, out sWhy))
                {
                    if (sWhy != null) { Note(sWhy, HorizontalAlignment.Left); return; }
                    slashPrompt = sp;
                }
            }
            if (loading.Count > 0) { Note("Đang đọc tệp đính kèm, đợi xong rồi gửi.", HorizontalAlignment.Left); return; }  // R5: chữ trong ô nhập được giữ
            if (turn != null && turn.Done) { Note("Đang hoàn tất lượt trước…", HorizontalAlignment.Left); return; }  // chữ trong ô nhập được giữ
            if (turn != null)
            {
                Stop(false, "superseded");
                Note("↻ Đè lệnh mới (đã dừng lệnh trước)", HorizontalAlignment.Left);
            }
            // C2: ghi tệp ghim TRƯỚC dòng stdin, ngay trên luồng UI (T2 không được dời bước này). Không có bản vẽ -> tệp rỗng
            // (server.py báo "không có bản vẽ làm việc", không bao giờ rơi về ActiveDocument). Ghi lỗi -> kill + báo đỏ + giữ nguyên ô nhập / tệp chờ
            string dn = null;
            try { dn = docName != null ? docName() : ActiveDocName(); } catch (Exception) { }  // ngoài AutoCAD / chưa có bản vẽ
            try { File.WriteAllText(PinFile, string.IsNullOrEmpty(dn) ? "" : dn, new UTF8Encoding(false)); }
            catch (Exception ex)
            {
                KillLive();
                AddText("Không ghi được tệp ghim bản vẽ làm việc nên không gửi: " + ex.Message, Red);
                return;
            }
            input.Clear();
            var sent = pending.ToList();
            pending.Clear();
            RefreshChips();
            Append(UserBubble(prompt, sent));
            string full = prompt == "" ? "Xem các tệp đính kèm." : (slashPrompt ?? prompt);  // khung chat (UserBubble) vẫn hiện đúng chữ người dùng gõ
            if (!string.IsNullOrEmpty(dn)) full = "[Bản vẽ làm việc: " + DocLabel(dn) + "]\n" + full;  // SafeName: tên bản vẽ có thể có ký tự phá nhãn
            if ((string)withSel.SelectedItem == "Có")
            {
                List<string> h = null;
                try { h = SelectedHandles(); }
                catch { }  // không có bản vẽ / chạy ngoài AutoCAD -> gửi không kèm
                if (h != null && h.Count > 0)
                {
                    Note("Kèm " + h.Count + " đối tượng đang chọn", HorizontalAlignment.Right);
                    // ponytail: tối đa 500 handle trong prompt; nhiều hơn thì Claude tự đọc PickfirstSelectionSet
                    full = "[Đang chọn " + h.Count + " đối tượng trong CAD, handle: " + string.Join(",", h.Take(500))
                        + (h.Count > 500 ? ",... (xem doc.PickfirstSelectionSet)" : "") + "]\n" + full;
                }
            }
            steps = 0;
            last = null;
            var t = NewTurn(full, sent);
            try { StartTurn(t, false); }
            catch (Exception ex)
            {
                AddText("Không chạy được Claude: " + Attach.Redact(ex.Message), Red);
                pending.InsertRange(0, sent);  // m4: không mất tệp đã chọn
                RefreshChips();
                t.Perf.End("spawn_fail");
                LogTurn(t);
                EndTurn(t, "Lỗi");
            }
        }

        // /login: mở cửa sổ console chạy "claude auth login" (palette chạy -p nên không đăng nhập được bên trong).
        // Chỉ chạy đường dẫn tuyệt đối từ Core.FindClaude, không qua cmd.exe.
        // Ghi vào %USERPROFILE%\.claude\.credentials.json - đúng tệp mà tiến trình claude của palette đọc.
        void Login()
        {
            if (login != null && !login.HasExited) { Note("Cửa sổ đăng nhập đang mở.", HorizontalAlignment.Left); return; }
            string exe = Exe;
            if (exe == null) { Note(NoClaude, HorizontalAlignment.Left); return; }
            Process p;
            try
            {
                bool api;
                var psi = NewPsi(exe, "auth login", out api);
                if (api) { DropKey(psi); Note("Đang không ở chế độ Tài khoản Claude. Chọn Cài đặt → Đăng nhập → Tài khoản Claude rồi gõ /login.", HorizontalAlignment.Left); return; }
                psi.CreateNoWindow = false;
                p = Process.Start(psi);
            }
            catch (Exception ex) { Note("Không mở được cửa sổ đăng nhập: " + Attach.Redact(ex.Message), HorizontalAlignment.Left); return; }
            login = p;
            Note("Mở cửa sổ đăng nhập Claude. Chọn đúng tài khoản trên trình duyệt; xong cửa sổ tự đóng.", HorizontalAlignment.Left);
            new Thread(() =>
            {
                try { p.WaitForExit(); } catch { }
                Interlocked.Increment(ref LoginGen);  // N1: Live.Key đổi -> tiến trình ấm cũ không còn khớp
                try { Page.Dispatcher.BeginInvoke((Action)Commands.KillIdleAll); } catch (Exception) { }  // mọi bản vẽ
                string s = AccountInfo();
                Page.Dispatcher.BeginInvoke((Action)(() => Note(s, HorizontalAlignment.Left)));
            }) { IsBackground = true }.Start();
        }

        // Tài khoản mà tiến trình claude của palette sẽ dùng: chỉ đọc email · tổ chức · gói từ "claude auth status --json".
        // Không đọc .credentials.json / token, không ghi kết quả ra tệp hay vào prompt.
        static string AccountInfo()
        {
            string v = "[palette " + typeof(ClaudeUi).Assembly.GetName().Version.ToString(3) + "] ";  // spec-3.1.0 §9 S1: ảnh lỗi phải nói rõ bản palette
            string exe = Exe;
            if (exe == null) return v + NoClaude;
            try
            {
                bool api;
                var psi = NewPsi(exe, "auth status --json", out api);
                if (api) { DropKey(psi); return v + "Chế độ API key · đã lưu key (mã hoá theo tài khoản Windows)"; }
                psi.CreateNoWindow = true; psi.RedirectStandardOutput = true; psi.StandardOutputEncoding = Encoding.UTF8;
                using (var p = Process.Start(psi))
                {
                    var t = p.StandardOutput.ReadToEndAsync();
                    if (!p.WaitForExit(15000)) { try { p.Kill(); } catch { } return v + "Không đọc được trạng thái đăng nhập (quá 15 s)."; }
                    if (!t.Wait(5000)) return v + "Không đọc được trạng thái đăng nhập (quá hạn đọc).";
                    string o = t.Result;
                    var d = Core.ParseJson(o) as Dictionary<string, object>;
                    if (d == null || !true.Equals(Get(d, "loggedIn"))) return v + "Palette chưa đăng nhập Claude. Gõ /login.";
                    return v + "Tài khoản Claude của palette: " + Get(d, "email") + " · " + Get(d, "orgName") + " · " + Get(d, "subscriptionType") + " · " + Get(d, "authMethod");
                }
            }
            catch (Exception ex) { return v + "Không đọc được trạng thái đăng nhập: " + Attach.Redact(ex.Message); }
        }

        // ---------- tiến trình claude: giữ ấm giữa các tin nhắn (T4), bộ đếm + nhật ký (T1), ghi stdin ngoài luồng UI (T2) ----------

        // 1 tiến trình claude.exe + Job Object của nó. Giữ sống sau mỗi result (tối đa IdleMinutes phút rảnh) để tin nhắn sau không phải khởi động lại
        sealed class Live
        {
            public Process P;
            public ProcJob Job;                      // null = không có Job -> không giữ ấm: một lượt rồi kill
            public string Key;                       // exe + args (không gồm --resume) + AuthStamp: khác thì không dùng lại. Không bao giờ chứa key
            public bool Api;                         // chế độ API key (do NewPsi trả): chỉ để chọn câu hướng dẫn lỗi
            public readonly StringBuilder Err = new StringBuilder();  // stderr, giữ tối đa 4096 ký tự cuối (M3)
            public volatile Turn Cur;                // lượt đang chạy trên tiến trình này, null = rảnh
            public volatile bool Dead;
            public bool Skills;                      // chạy ở chế độ skill (3.1.1): reader kiểm system/init
            public bool InitOk;                      // đã thấy system/init hợp lệ của tiến trình này (chỉ reader chạm)
            public volatile List<string> Expected;   // tên skill kỳ vọng trong init (Attach.SkillExpected); làm mới mỗi lượt
            public string Sid;                      // session id reader này đã ghi vào ClaudeUi.sessionId
            readonly object gate = new object();

            // thread-safe, lặp lại được. Có Job: TerminateJobObject giết cả cây (claude + python server.py)
            public void Kill()
            {
                Dead = true;
                lock (gate)
                {
                    try { if (Job != null) Job.Kill(); } catch (Exception) { }
                    try { if (P != null && !P.HasExited) P.Kill(); } catch (Exception) { }
                }
            }
        }

        // 1 lượt chat (1 lần Gửi). turn = lượt UI đang chờ; mọi sự kiện của lượt khác bị bỏ qua
        sealed class Turn
        {
            public string Full;                      // nội dung gửi (RAM): để gửi lại 1 lần khi tiến trình ấm chết (I2) và dựng dòng stdin
            public IList<Attachment> Sent;
            public TurnPerf Perf;
            public readonly Stopwatch Sw = new Stopwatch();
            public volatile string Abort;            // null | "stopped" | "superseded"
            public volatile bool Done;               // reader đã nhận result
            public volatile List<string> FullPdf;    // SHA của PDF gửi ĐẦY ĐỦ lượt này (writer điền; T6)
            public volatile bool Written;            // writer đã Flush xong dòng stdin
            public volatile string WriteErr;         // lỗi của writer (chỉ ghi lại; Died là nơi xử lý duy nhất)
            public bool HadDedup;                    // có PDF gửi bằng nhãn "đã gửi" lượt này
            public bool Retried;
            public int Logged;
        }

        static readonly string Ver = typeof(ClaudeUi).Assembly.GetName().Version.ToString();

        Turn NewTurn(string full, IList<Attachment> sent)
        {
            var t = new Turn { Full = full, Sent = sent };
            string m = (string)model.SelectedItem, ef = (string)effort.SelectedItem;
            t.Perf = new TurnPerf { Ver = Ver, Model = m == Default ? "default" : m, Effort = ef == Default ? "default" : ef };
            dedup.Sync(ReadSessionId());  // I5: trước khi quyết mode (chưa có session id / id khác -> quên)
            foreach (var a in sent)
            {
                string mode = a.Kind == "pdf" ? (dedup.Has(a.Sha) ? "dedup" : "doc") : null;
                if (mode == "dedup") t.HadDedup = true;
                t.Perf.AddAtt(a, mode);
            }
            turn = t;
            phase = "Claude đang suy nghĩ...";
            status.Text = phase + " · 0,0 s";
            stopBtn.IsEnabled = true;
            idleTimer.Stop();  // M3: bộ đếm rảnh chỉ chạy giữa result và lần Gửi kế
            tick.Start();
            return t;
        }

        string ReadSessionId() { lock (sessionGate) return Attach.IsSessionId(sessionId) ? sessionId : null; }  // chỉ nhận session id dạng GUID 8-4-4-4-12

        static bool Alive(Process p)
        {
            try { return p != null && !p.HasExited; }
            catch (Exception) { return false; }
        }

        // dùng lại tiến trình ấm nếu cùng exe + args (đổi Model/Effort/cfg thì không); không thì kill và spawn mới --resume
        void StartTurn(Turn t, bool forceNew)
        {
            string exe = Exe, cfg = Path.Combine(Home, "mcp.json");
            if (exe == null) throw new Exception(NoClaude);
            if (!File.Exists(cfg)) throw new Exception("Thiếu " + cfg + ". Chạy lại Setup để sửa bản cài.");
            string m = (string)model.SelectedItem, ef = (string)effort.SelectedItem;
            string mArg = m == Default ? null : m, eArg = ef == Default ? null : ef;
            string sid = ReadSessionId();
            dedup.Sync(sid);  // I5: chưa có session id / sai / id khác id đã ghi -> quên các PDF đã gửi
            string why = null;
            bool skills = !skillsBlocked && Attach.SkillMode(Home, out why);  // 3.1.1: tính lại MỖI lượt (không cache); đổi thì args đổi -> Key đổi -> tiến trình ấm bị thay
            if (why != null && !skillWhyShown) { skillWhyShown = true; AddText(why, Red); }  // một dòng đỏ mỗi phiên palette, rồi chạy như 3.1.0
            // chế độ skill: tập skill (tên thư mục + name:) nằm trong Key -> đổi tập skill thì spawn lại (vẫn --resume), init không còn mang tên cũ
            List<string> skillNames = skills ? Attach.SkillExpected(Home) : null;  // một lần: dùng cho Key và cho chữ trạng thái hàng Skill
            ShowSkillStatus(skills, skillNames, why);
            string key = exe + "\n" + Attach.ClaudeArgs(cfg, mArg, eArg, null, skills) + "\n" + AuthStamp() + (skills ? "\n" + string.Join("|", skillNames.OrderBy(x => x, StringComparer.Ordinal).ToArray()) : "");
            Live L = live;
            bool warm = !forceNew && L != null && !L.Dead && L.Job != null && L.Key == key && L.Cur == null && Alive(L.P);
            if (warm)
            {
                t.Sw.Start();
                L.Cur = t;
            }
            else
            {
                KillLive();
                L = Spawn(exe, Attach.ClaudeArgs(cfg, mArg, eArg, sid, skills), key, t, skills);
            }
            t.Perf.Warm = warm;
            t.Perf.Resumed = !warm && sid != null;
            Write(L, t);
        }

        Live Spawn(string exe, string args, string key, Turn t, bool skills)
        {
            bool api;
            var psi = NewPsi(exe, args, out api);  // ném lỗi TRƯỚC khi tạo Live/Job: chế độ lạ / key hỏng -> không spawn
            psi.CreateNoWindow = true;
            psi.RedirectStandardInput = true; psi.RedirectStandardOutput = true; psi.RedirectStandardError = true;
            psi.StandardOutputEncoding = Encoding.UTF8; psi.StandardErrorEncoding = Encoding.UTF8;
            psi.EnvironmentVariables["ENABLE_TOOL_SEARCH"] = "false";
            psi.EnvironmentVariables["CAD_CLAUDE_PIN_FILE"] = PinFile;  // C1: tệp ghim riêng của bản vẽ này
            psi.EnvironmentVariables["CAD_CLAUDE_SKILL_DRAFTS"] = DraftsDir;  // 3.1.1: save_skill của server chỉ ghi bản nháp vào đây (ngoài .claude), mọi chế độ
            if (skills) psi.EnvironmentVariables["CLAUDE_CODE_DISABLE_CLAUDE_MDS"] = "1";  // chế độ skill nạp setting source project: không nạp CLAUDE.md. AuthEnv không xoá biến này
            psi.EnvironmentVariables["CAD_CLAUDE_ACAD_PID"] = PidOfThis().ToString(CultureInfo.InvariantCulture);  // server.py kiểm: đúng AutoCAD có palette
#if NET
            psi.StandardInputEncoding = new UTF8Encoding(false);  // .NET 8: không BOM đầu stdin (NoStdinBom chỉ có tác dụng trên .NET Framework)
#endif
            Attach.NoStdinBom();  // R-A1: console 65001 có BOM -> Process.Start chèn BOM vào đầu stdin
            var L = new Live { Key = key, Api = api, Job = ProcJob.TryCreate(), Skills = skills, Expected = skills ? Attach.SkillExpected(Home) : null };
            t.Sw.Start();  // một đồng hồ cho cả lượt, bắt đầu ngay trước Process.Start (§4.2)
            Process p;
            try { p = Process.Start(psi); }
            catch (Exception) { if (L.Job != null) L.Job.Kill(); throw; }
            finally { DropKey(psi); }  // M2: cả nhánh thành công lẫn lỗi
            t.Perf.SpawnMs = t.Sw.ElapsedMilliseconds;
            L.P = p;
            if (L.Job != null && !L.Job.Assign(p)) { L.Job.Kill(); L.Job = null; }  // không vào được Job = không bảo đảm chết theo AutoCAD -> không giữ ấm
            L.Cur = t;
            p.ErrorDataReceived += (s, e) =>
            {
                if (e.Data == null) return;
                lock (L.Err) { L.Err.AppendLine(e.Data); if (L.Err.Length > 4096) L.Err.Remove(0, L.Err.Length - 4096); }
            };
            p.BeginErrorReadLine();
            lock (sessionGate) live = L;
            // đọc trước, ghi sau, và ghi trên luồng nền: payload tới ~20 MB; ghi trên luồng UI làm AutoCAD đứng,
            // ghi trước khi đọc có thể tắc khi pipe stdout đầy
            new Thread(() => Read(L)) { IsBackground = true, Name = "ClaudeRead" }.Start();
            return L;
        }

        // T2 + I4: dựng dòng stdin (có thể ~20 MB) và ghi trên luồng nền. Mỗi writer giữ Live + Turn của riêng mình; chỉ 1 writer chạy một lúc.
        // Lỗi ở bất kỳ bước nào: kill tiến trình (không bao giờ dùng lại sau ghi lỗi / ghi dở), báo lên UI. Ghi tệp ghim đã xong từ Send (C2)
        void Write(Live L, Turn t)
        {
            new Thread(() =>
            {
                byte[] b = null;
                try
                {
                    lock (writeGate)
                    {
                        if (L.Dead) throw new IOException("Claude đã dừng.");
                        var sentFull = new List<string>();
                        b = new UTF8Encoding(false).GetBytes(Attach.BuildUserLine(t.Full, t.Sent, dedup, sentFull));  // dựng đủ bộ đệm TRƯỚC khi ghi
                        t.FullPdf = sentFull;
                        t.Perf.StdinBytes = b.Length;
                        var s = L.P.StandardInput.BaseStream;
                        s.Write(b, 0, b.Length);
                        s.Flush();
                        t.Written = true;
                        if (L.Job == null) L.P.StandardInput.Close();  // không giữ ấm: claude tự thoát sau lượt này
                    }
                }
                catch (Exception ex)
                {
                    t.WriteErr = ex.Message;  // không báo gì: reader thấy stdout đóng và gọi Died (nơi xử lý duy nhất)
                    L.Kill();                 // không dùng lại tiến trình sau khi ghi lỗi / ghi dở (I4)
                }
                finally { b = null; }  // m5: không giữ ~20 MB suốt thời gian claude chạy
            }) { IsBackground = true, Name = "ClaudeWrite" }.Start();
        }

        static object Get(IDictionary<string, object> d, string key)
        {
            object v;
            return d != null && d.TryGetValue(key, out v) ? v : null;
        }

        // Luồng nền: lỗi không bắt ở đây sẽ làm sập cả AutoCAD -> bắt hết, báo lên khung chat
        void Read(Live L)
        {
            try { ReadLoop(L); }
            catch (Exception ex)
            {
                var t = L.Cur;
                string msg = Attach.Redact(ex.Message);
                if (t != null && !t.Done && t.Abort == null) PostAct(t, () => FailTurn(t, msg));
            }
            finally
            {
                L.Kill();  // Dead + giết nốt job (con cháu còn sót)
                try { L.P.Dispose(); } catch (Exception) { }
            }
        }

        // Đọc stream-json của Claude (nhiều lượt trên 1 stdout), đẩy sự kiện của lượt hiện tại về UI
        void ReadLoop(Live L)
        {
            var p = L.P;
            string line;
            while ((line = p.StandardOutput.ReadLine()) != null)
            {
                Dictionary<string, object> ev;
                try { ev = Core.ParseJson(line) as Dictionary<string, object>; }
                catch { continue; }
                if (ev == null) continue;
                // 3.1.1: chế độ skill, MỌI sự kiện: init phải hợp lệ (tools, skills, plugins), assistant/result không được tới trước init hợp lệ (Attach.SkillEventOk).
                // Sai -> kill, tắt skill cho cả phiên palette này, chạy lại lượt đó bằng args 3.1.0. Đặt TRƯỚC khi ghi session id: phiên của tiến trình bị từ chối không bị --resume
                string badInit;
                if (L.Skills && !Attach.SkillEventOk(ev, ref L.InitOk, L.Expected, () => Attach.ParentDirty(Home, null, UserProfile), out badInit)) { SkillInitBad(L, badInit); return; }
                dedup.OnEvent(ev);  // I5: system/compact_boundary -> quên các PDF đã gửi
                string sid = Get(ev, "session_id") as string;
                if (sid != null && sid != L.Sid && Attach.IsSessionId(sid))
                {
                    L.Sid = sid;
                    // M3: chỉ tiến trình đang là live mới được ghi session id (tiến trình đã bị kill / thay thế thì không)
                    lock (sessionGate) { if (ReferenceEquals(L, live)) sessionId = sid; }
                }
                Turn t = L.Cur;
                if (t == null || t.Done) continue;
                t.Perf.OnEvent(ev, t.Sw.ElapsedMilliseconds);
                string type = Get(ev, "type") as string;
                if (type == "assistant")
                {
                    var content = Get(Get(ev, "message") as Dictionary<string, object>, "content") as IEnumerable;
                    if (content == null) continue;
                    foreach (object o in content)
                    {
                        var c = o as Dictionary<string, object>;
                        string ct = Get(c, "type") as string;
                        if (ct == "tool_use")
                        {
                            string tn = ((Get(c, "name") as string) ?? "?").Replace("mcp__autocad__", "");
                            PostAct(t, () => OnEvent("tool", tn));
                        }
                        else if (ct == "text")
                        {
                            string tx = ((Get(c, "text") as string) ?? "").Trim();
                            PostAct(t, () => OnEvent("text", tx));
                        }
                    }
                }
                else if (type == "result")
                {
                    t.Done = true;
                    L.Cur = null;
                    bool isErr = true.Equals(Get(ev, "is_error"));
                    object r = Get(ev, "result") ?? (isErr ? Get(ev, "subtype") : "");
                    string msg = Attach.Redact(Convert.ToString(r));
                    if (isErr)
                    {
                        string hint = Attach.AuthHint(msg, L.Api);
                        if (hint != null) msg += "\n" + hint;
                        else if (msg.Contains("subscription access"))
                            msg += "\n" + AccountInfo() + "\nGõ /login để đăng nhập tài khoản khác.";
                    }
                    t.Perf.End(null);
                    LogTurn(t);
                    PostAct(t, () => OnDone(t, msg, isErr));
                }
            }
            if (!p.WaitForExit(3000)) L.Kill();
            else
            {  // WaitForExit(ms) không đợi handler stderr bất đồng bộ: đợi có hạn để tail đầy đủ
                var w = new Thread(() => { try { p.WaitForExit(); } catch (Exception) { } }) { IsBackground = true };
                w.Start();
                w.Join(2000);
            }
            // stdout đã đóng mà lượt còn dở: bị Dừng / đè (UI đã xử lý) hoặc tiến trình chết
            Turn u = L.Cur;
            if (u == null || u.Done) return;
            bool killedByUs = L.Dead;  // Dead đặt bởi Kill() (Stop, writer lỗi, UI) trước khi stdout đóng
            if (!killedByUs && Alive(p) == false)
            {
                try { u.Perf.ExitCode = p.ExitCode; u.Perf.Exited = true; u.Perf.ExitMs = u.Sw.ElapsedMilliseconds; } catch (Exception) { }
            }
            if (u.Abort != null) { u.Perf.End(u.Abort); LogTurn(u); return; }
            u.Perf.End(null);
            LogTurn(u);
            string tail;
            lock (L.Err) tail = L.Err.ToString().Trim();
            tail = Attach.Redact(tail);
            string em = tail.Length > 0 ? tail.Substring(Math.Max(0, tail.Length - 500)) : null;
            string tailHint = Attach.AuthHint(tail, L.Api);
            if (tailHint != null) em += "\n" + tailHint;
            string ex = "Claude thoát với mã " + (u.Perf.Exited ? u.Perf.ExitCode.ToString(CultureInfo.InvariantCulture) : "?");
            PostAct(u, () => Died(u, em, ex));
        }

        // Luồng đọc: system/init của chế độ skill có tool lạ. skillsBlocked: mọi lượt sau chạy 3.1.0 (không kiểm lại, không lặp kill)
        void SkillInitBad(Live L, string bad)
        {
            skillsBlocked = true;
            L.Kill();
            lock (sessionGate) { if (ReferenceEquals(L, live) || (L.Sid != null && L.Sid == sessionId)) sessionId = null; }  // lượt chạy lại KHÔNG --resume phiên của tiến trình skill này
            // KHÔNG hiện gì lên khung chat: chỉ một dòng vào nhật ký sẵn có (logs\perf.jsonl, cùng đường PerfAppend với LogTurn); nhãn lạ đã lọc KeyLabel
            try
            {
                string logLine = Attach.SkillInitLog(Ver, bad), logDir = Path.Combine(Home, "logs");
                ThreadPool.QueueUserWorkItem(_ => Attach.PerfAppend(logDir, logLine));  // PerfAppend không bao giờ ném lỗi
            }
            catch (Exception) { }
            var t = L.Cur;
            if (t == null || t.Done) return;
            t.Perf.End("skill_init");
            LogTurn(t);
            PostAct(t, () => RetryTurn(t));  // chạy lại lượt đó bằng args 3.1.0 (skillsBlocked)
        }

        // chạy trên UI khi lượt t còn là lượt hiện tại
        void PostAct(Turn t, Action a)
        {
            try
            {
                Page.Dispatcher.BeginInvoke((Action)(() =>
                {
                    if (t != turn) return;
                    try { a(); }
                    catch (Exception ex) { try { AddText("Lỗi: " + Attach.Redact(ex.Message), Red); } catch (Exception) { } }  // không để ngoại lệ thoát ra khỏi dispatcher (sập AutoCAD)
                }));
            }
            catch (Exception) { }
        }

        static void LogTurn(Turn t)
        {
            if (Interlocked.Exchange(ref t.Logged, 1) != 0) return;
            try
            {
                string line = t.Perf.ToJson(), dir = Path.Combine(Home, "logs");
                ThreadPool.QueueUserWorkItem(_ => Attach.PerfAppend(dir, line));  // PerfAppend không bao giờ ném lỗi
            }
            catch (Exception) { }
        }

        static string Sec(long ms) { return (ms / 1000.0).ToString("0.0", CultureInfo.InvariantCulture).Replace('.', ',') + " s"; }
        static string Sec1(long ms) { return (Math.Max(0, ms) / 1000.0).ToString("0.0", CultureInfo.InvariantCulture).Replace('.', ','); }

        // I2: tiến trình ấm chết trước khi có sự kiện nào của lượt này -> gửi lại đúng 1 lần (spawn mới --resume); còn lại báo đỏ
        // Nơi xử lý DUY NHẤT khi tiến trình chết giữa lượt (kể cả do writer ghi lỗi rồi kill). tail = stderr cuối (null nếu trống)
        void Died(Turn t, string tail, string exitText)
        {
            if (t.Perf.Warm && t.Perf.Events == 0 && !t.Retried) { RetryTurn(t); return; }
            string we = t.WriteErr;
            if (!t.Written) { pending.InsertRange(0, t.Sent); RefreshChips(); }  // I4: chưa gửi được -> trả tệp về hàng chờ
            FailTurn(t, tail ?? (!t.Written ? "Không gửi được tin nhắn tới Claude" + (we != null ? ": " + we : "") : exitText));
        }

        void RetryTurn(Turn old)
        {
            var t2 = NewTurn(old.Full, old.Sent);
            t2.Retried = true;
            try { StartTurn(t2, true); }
            catch (Exception ex)
            {
                pending.InsertRange(0, t2.Sent);
                RefreshChips();
                t2.Perf.End("spawn_fail");
                LogTurn(t2);
                FailTurn(t2, "Không chạy được Claude: " + Attach.Redact(ex.Message));
            }
        }

        void FailTurn(Turn t, string msg)
        {
            AddText("Lỗi: " + Attach.Redact(msg), Red);
            EndTurn(t, "Lỗi");
            KillLive();  // tiến trình hỏng / không rõ trạng thái: không giữ
        }

        // kết thúc lượt trên UI: rảnh, dừng bộ đếm
        void EndTurn(Turn t, string label)
        {
            if (t == turn) turn = null;
            tick.Stop();
            stopBtn.IsEnabled = false;
            status.Text = label + " · " + Sec(t.Sw.ElapsedMilliseconds);
            UpdateLabels();
        }

        void OnEvent(string kind, string text)
        {
            if (kind == "tool")
            {
                steps++;
                phase = "Bước " + steps + ": đang chạy " + text + "...";
                var t = turn;
                status.Text = phase + (t != null ? " · " + Sec(t.Sw.ElapsedMilliseconds) : "");
                Note("⚙ " + text, HorizontalAlignment.Left);
            }
            else if (kind == "text")
            {
                if (text != "" && text != last) { last = text; AddText(text, null); }
            }
            else if (kind == "err") AddText("Lỗi: " + text, Red);
        }

        // result: "Xong" ngay (không chờ tiến trình thoát, vì tiến trình được giữ ấm). Spec §4.1
        void OnDone(Turn t, string msg, bool isErr)
        {
            if (isErr) AddText("Lỗi: " + msg, Red);
            else OnEvent("text", msg);
            dedup.Commit(t.Perf, t.FullPdf ?? new List<string>());  // I5: chỉ khi result && !is_error (Commit tự kiểm)
            var pf = t.Perf;
            if (pf.Compact && t.HadDedup) AddText("Lịch sử vừa được nén; đính kèm lại PDF nếu Claude cần đọc.", Dim);
            string label = isErr ? "Lỗi" : steps > 0 ? "Xong (" + steps + " thao tác)" : "Xong";
            string st = label + " · " + Sec(pf.ResultMs >= 0 ? pf.ResultMs : t.Sw.ElapsedMilliseconds);
            if (!isErr && pf.FirstEventMs >= 0 && pf.ResultMs >= 0)
                st += " — khởi động " + Sec1(pf.FirstEventMs) + " · Claude " + Sec1(pf.ResultMs - pf.FirstEventMs - pf.ToolMs) + " · công cụ " + Sec1(pf.ToolMs);
            turn = null;
            tick.Stop();
            stopBtn.IsEnabled = false;
            status.Text = st;
            UpdateLabels();
            try { ShowDrafts(); } catch (Exception) { }  // 3.1.1: bản nháp skill AI soạn trong lượt này
            var L = live;
            if (L != null && L.Job == null) { lock (sessionGate) { if (ReferenceEquals(live, L)) live = null; } }  // không có Job: không giữ ấm; claude tự thoát (stdin đã đóng), không kill
            else if (L != null) { idleTimer.Stop(); idleTimer.Start(); }  // M3: chỉ bắt đầu đếm rảnh từ lúc có result
        }

        // abort: "stopped" (nút Dừng, Esc, Mới, /login) | "superseded" (Gửi đè). Ghi nhật ký do reader làm khi stdout đóng
        void Stop(bool announce, string abort)
        {
            var t = turn;
            if (t == null) return;
            turn = null;  // mọi sự kiện còn lại của lượt này bị bỏ qua
            t.Abort = abort;
            tick.Stop();
            KillLive();
            stopBtn.IsEnabled = false;
            if (announce)
            {
                AddText("(đã dừng)", Red);
                status.Text = "Đã dừng · " + Sec(t.Sw.ElapsedMilliseconds);
            }
        }

        // TerminateJobObject cho tiến trình đang giữ (đổi Model/Effort, Mới, /login, rảnh 10 phút, lỗi, thoát AutoCAD)
        void KillLive()
        {
            Live L;
            lock (sessionGate) { L = live; live = null; }
            idleTimer.Stop();
            if (L != null) L.Kill();
        }

        // QuitWillStart / Terminate (không phải BeginQuit: người dùng còn có thể huỷ thoát). Job KILL_ON_JOB_CLOSE lo phần còn lại nếu AutoCAD chết đột ngột
        public void Shutdown()
        {
            try { tick.Stop(); } catch (Exception) { }
            try { var t = turn; if (t != null) t.Abort = "stopped"; turn = null; KillLive(); } catch (Exception) { }
            try { jobs.CompleteAdding(); } catch (Exception) { }  // luồng đọc tệp thoát khi hàng đợi cạn
            DeletePin();
        }

        public void NewChat()
        {
            Stop(false, "stopped");
            KillLive();  // trước khi xoá session id: reader của tiến trình cũ không ghi lại được nữa
            dedup.Clear();
            foreach (var c in loading.Values) c.Cancel();  // dừng việc đọc (kill tiến trình .xls)
            loading.Clear();
            pending.Clear();
            RefreshChips();
            lock (sessionGate) sessionId = null;
            log.Children.Clear();
            Welcome();
            status.Text = "Sẵn sàng";
            UpdateLabels();
        }
    }
}
