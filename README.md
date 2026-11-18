# Claude AI cho AutoCAD (ClaudeCAD 3.1.1)

Claude AI cho AutoCAD (tên nội bộ cad-mcp / ClaudeCAD) là một palette trong AutoCAD để trò chuyện với Claude. Claude đọc và sửa bản vẽ đang mở thông qua một MCP server Python (`server.py`, công cụ `mcp__autocad`), điều khiển AutoCAD bằng COM.

## Tính năng

- Palette trong AutoCAD để trò chuyện với Claude; palette chạy Claude Code CLI (`claude.exe`).
- Claude đọc và sửa bản vẽ đang mở qua MCP server `server.py`.
- Lệnh AutoCAD: `CLAUDE`, `CLAUDENEW`, `CLAUDECHUP`. Chi tiết xem `installer\docs\HDSD.html`.
- Đăng nhập bằng tài khoản Claude hoặc API key (xem mục Đăng nhập).
- Làm việc với nhiều bản vẽ, mỗi bản vẽ có phiên hội thoại riêng (xem mục Làm việc với nhiều bản vẽ).
- Phím tắt chụp vùng và ẩn/hiện palette (xem `docs\THU-PHIM-TAT-TRONG-AUTOCAD.md`).
- 9 công cụ MCP mới: `mirror`, `offset`, `array_rect`, `array_polar`, `ellipse`, `create_block`, `list_blocks`, `undo`, `redo`; mỗi lượt công cụ sửa bản vẽ được gom thành một bước undo; `save` chỉ ghi ra ổ cục bộ và không ghi đè tệp khác khi thiếu cờ `overwrite`.
- Skill: tự tạo, nạp gói skill, hoặc nhờ Claude soạn skill (bản nháp chờ bạn duyệt).

## Yêu cầu

- AutoCAD 2018–2024 (R22.0–R24.3), .NET Framework 4.6.
- AutoCAD 2025–2026 (R25.0–R25.1), .NET 8.
- AutoCAD 2027 trở lên (R26.0+), .NET 10.
- Windows 10 1809 trở lên, 64-bit.
- Không hỗ trợ AutoCAD LT.
- Không cần Claude Desktop.

## Cài đặt

Cài đặt từ bộ cài (`Setup.exe`) được đóng gói bằng `installer\pack_setup.py`.

- Cài per-user, không cần quyền admin.
- Ứng dụng được cài vào `%LOCALAPPDATA%\cad-mcp`.
- Bundle được đặt tại `%APPDATA%\Autodesk\ApplicationPlugins\ClaudeCAD.bundle`.
- Các bước cài chi tiết: xem `installer\docs\HDSD.html`. Trước khi cài: xem `installer\docs\truoc-khi-cai.txt`.
- Python nhúng 3.13 và Claude Code CLI được tải lúc đóng gói/cài, không kèm trong repo.

## Đăng nhập

Chọn phương thức trong mục Cài đặt của palette.

**Tài khoản Claude (mặc định)**

- Đăng nhập qua `claude auth login`.
- Palette bỏ qua mọi biến môi trường xác thực như `ANTHROPIC_API_KEY` còn sót trong Windows.
- Đây là nguyên nhân lỗi "Invalid API key · Fix external API key" ở bản 3.0.x.

**API key**

- Nhập key trong palette bằng nút "Dán & lưu".
- Key được mã hoá bằng Windows DPAPI theo tài khoản Windows.
- Key chỉ được truyền cho tiến trình `claude` qua biến môi trường.
- Nếu key hỏng hoặc chế độ không hợp lệ, palette dừng; palette không tự chuyển sang key khác.

## Làm việc với nhiều bản vẽ

- Mỗi bản vẽ có phiên hội thoại riêng. Palette đổi nội dung theo bản vẽ đang mở.
- Bản vẽ không ở phía trước vẫn đọc/ghi đối tượng được, nhưng không chạy được lệnh AutoCAD (`SendCommand`) và không zoom.
- Hội thoại không tự nối lại sau khi khởi động lại AutoCAD.
- Phím chụp vùng (mặc định Ctrl+Shift+X) và phím ẩn/hiện palette (Ctrl+5) chỉ hoạt động khi AutoCAD ở phía trước.
- Không còn phím tắt toàn hệ thống. Chụp xong không kéo AutoCAD lên trước.

## Skill: tự tạo và nạp

### Skill là gì

- Skill là tệp `SKILL.md` hướng dẫn Claude cách làm một việc lặp lại trong AutoCAD (ví dụ quy trình ghi kích thước).
- Skill nằm ở `%LOCALAPPDATA%\cad-mcp\.claude\skills\<tên>\SKILL.md`. Mở lại palette (lượt hỏi mới) là có hiệu lực.
- Gỡ cài đặt không xoá skill và bản nháp của bạn; muốn xoá hẳn thì xoá thư mục `%LOCALAPPDATA%\cad-mcp` sau khi gỡ.

### Ba nút trong phần Cài đặt

Ba nút nằm trong phần Cài đặt của palette, dưới trạng thái đăng nhập:

- **Thư mục skill**: mở thư mục skill trong Explorer.
- **Tạo skill mới**: tạo `skill-moi-<n>\SKILL.md` mẫu và mở bằng Notepad để bạn sửa. Mẫu khung:

```
---
name: skill-moi-1
description: Mô tả ngắn skill này làm gì và khi nào dùng (Claude đọc dòng này để quyết định có dùng skill hay không).
---

# Tên skill

## Khi nào dùng
- Ví dụ: khi người dùng yêu cầu ...

## Các bước
1. Dùng công cụ mcp__autocad__... để ...
2. ...

## Lưu ý
- Đơn vị bản vẽ: ...
```

- **Nạp skill…**: chọn một tệp SKILL.md (nạp cả thư mục chứa nó) hoặc một gói .zip. Palette kiểm tra rồi hiện hộp xác nhận (tên, mô tả, số tệp, số tệp script .py/.lsp) trước khi chép; nếu đã có skill cùng tên thì hỏi ghi đè.

### Gọi skill bằng lệnh

Gõ `/tên-skill nội dung cần làm` (hoặc `/skill tên-skill nội dung`) trong ô chat để bảo Claude dùng đúng skill đó. Khi đang có skill: tên không có trong danh sách thì palette không gửi và hiện danh sách skill đang có (muốn gửi như câu thường thì bỏ dấu / ở đầu). Khi chưa có skill hoặc skill đang tắt: dòng bắt đầu bằng / được gửi bình thường như bản 3.1.0, riêng `/skill …` sẽ hiện ghi chú lý do. Không gõ lệnh thì Claude vẫn tự chọn skill phù hợp theo mô tả.

### Nhờ Claude soạn skill

- Trong khung chat, yêu cầu rõ, ví dụ "Lưu quy trình vừa làm thành skill tên ghi-kich-thuoc".
- Claude chỉ lưu BẢN NHÁP vào `%LOCALAPPDATA%\cad-mcp\skill-nhap\`; palette hiện dòng "AI đã soạn skill chờ duyệt".
- Bấm "Nạp skill…" (hộp chọn tệp mở sẵn ở thư mục nháp). Hộp xác nhận hiện tên, mô tả, số tệp, số tệp script .py/.lsp và phần đầu nội dung (1000 ký tự). Với bản nháp do AI soạn có thêm dòng nhắc "chỉ nạp nếu chính bạn vừa yêu cầu soạn". Nếu SKILL.md dài, mở thư mục nháp đọc toàn bộ trước khi xác nhận. Tối đa 20 bản nháp.

### Quy tắc tên và nội dung

- Tên chỉ gồm chữ thường a–z, số, dấu gạch ngang đơn giữa các cụm (không bắt đầu/kết thúc bằng gạch ngang, không hai gạch liền nhau), bắt đầu bằng chữ, tối đa 64 ký tự.
- Phần đầu SKILL.md chỉ nhận các khoá: name, description, when_to_use, argument-hint, arguments, disable-model-invocation, user-invocable, license, metadata, compatibility.
- Gói bị từ chối nếu có:
  - lệnh shell nhúng (`` !` `` hoặc khối ```` ```! ````), khoá hooks/allowed-tools/model…;
  - ký tự vô hình hoặc ký tự đảo chiều văn bản;
  - tệp ngoài các đuôi .md .txt .json .csv .yaml .yml .py .lsp .dxf .dwg .dwt .png .jpg;
  - tệp/thư mục tên bắt đầu bằng dấu chấm;
  - hơn một SKILL.md;
  - quá 200 tệp hoặc 20 MB;
  - SKILL.md quá 256 KB.

### Khi nào chế độ skill tự tắt

Palette chạy như 3.1.0 và hiện một dòng đỏ khi:

- trong `%LOCALAPPDATA%\cad-mcp\.claude\` có bất cứ thứ gì ngoài thư mục `skills` (ví dụ settings.json);
- có CLAUDE.md, CLAUDE.local.md hoặc AGENTS.md trong thư mục cài;
- có thư mục `.claude` hoặc tệp CLAUDE.md/AGENTS.md ở bất kỳ thư mục cha nào của thư mục cài (trừ `.claude` trong thư mục người dùng);
- một skill đã chép tay vi phạm quy tắc ở mục trên.

Trên máy nhiều người dùng, người khác tạo `C:\.claude` hoặc `C:\CLAUDE.md` cũng làm skill tự tắt; xoá mục đó để bật lại.

Trong thư mục `skills` có tệp lẻ, thư mục thiếu SKILL.md hoặc tên thư mục không đúng quy tắc tên thì MỌI skill đều bị tắt (dòng đỏ) cho tới khi bạn xoá hoặc sửa mục đó.

Khi khởi động, nếu Claude báo có công cụ, skill hoặc plugin ngoài danh sách cho phép (chỉ chấp nhận plugin dựng sẵn của Claude Code), palette lặng lẽ chạy lại phiên đó không có skill; lý do ghi trong log của palette.

Trường hợp này palette không báo trên khung chat; muốn biết lý do thì xem dòng có `"event":"skill_init_bad"` trong `%LOCALAPPDATA%\cad-mcp\logs\perf.jsonl`. Khi đó gõ `/tên-skill` sẽ thấy ghi chú skill đang tắt.

Không có skill nào thì Claude được chạy với tham số y hệt 3.1.0; máy chủ AutoCAD có thêm công cụ `save_skill`, chỉ ghi bản nháp vào thư mục `skill-nhap`. Skill cá nhân trong `%USERPROFILE%\.claude\skills` KHÔNG được palette nạp. Plugin bạn cài cho Claude Code cũng không được nạp.

### Đọc trước khi dùng skill

- Skill là chương trình: nó điều khiển được công cụ run_python, tức chạy mã bất kỳ dưới tài khoản Windows của bạn (đọc/xoá tệp, gửi dữ liệu, đọc API key đã lưu). Không có sandbox. Chỉ nạp skill bạn đã tự đọc hết hoặc từ nguồn tin cậy; script kèm trong thư mục skill cũng chạy được qua run_python.
- Chữ trong bản vẽ hoặc tệp đính kèm có thể chứa chỉ dẫn giả xúi Claude soạn skill độc. Luôn đọc kỹ bản nháp trước khi duyệt.
- Rủi ro không chặn được hoàn toàn: Claude bị chỉ dẫn giả điều khiển vẫn có thể dùng run_python ghi thẳng vào thư mục skill, bỏ qua bước duyệt. Thấy skill lạ trong "Thư mục skill" thì xoá.
- Skill cần các công cụ Read/Bash/Write của Claude Code sẽ không chạy đúng: palette chỉ cho Claude dùng công cụ Skill và công cụ AutoCAD.
- Nội dung skill được gửi tới Anthropic ở mỗi phiên (tốn token).
- Máy nhiều người dùng: giữ thư mục cài trong %LOCALAPPDATA% (mặc định).
- Bộ kiểm của palette chỉ chặn các cú pháp nguy hiểm đã biết; skill qua được bộ kiểm KHÔNG có nghĩa là an toàn.
- Cơ chế skill đã được kiểm với Claude Code 2.1.310 trên một máy. Sau khi Claude Code cập nhật, hoặc với tài khoản/bản Claude Code khác, nếu palette báo skill bị tắt (dòng đỏ) thì cứ dùng không có skill cho tới bản ClaudeCAD kế tiếp.
- Máy dùng chung với người không tin cậy: không nên bật skill.

## An toàn và giới hạn

- Claude chạy lệnh AutoCAD và mã Python với quyền của người dùng hiện tại, không hỏi lại trước khi chạy.
- Nội dung bản vẽ là dữ liệu không đáng tin. Chữ trong bản vẽ có thể chứa chỉ dẫn giả.
- Nên kiểm tra kết quả trên bản vẽ trước khi lưu.
- API key được mã hoá, nhưng mã chạy dưới cùng tài khoản Windows vẫn lấy được key. Nên dùng key riêng trong workspace có hạn mức chi.
- Sau khi bấm "Dán & lưu", key vẫn còn trong clipboard và lịch sử Win+V.
- Sau khi xoá key, phiên đang chạy có thể giữ key tới lượt sau hoặc tối đa 10 phút.
- Không hỗ trợ `ANTHROPIC_BASE_URL` (gateway/proxy riêng).

## Build từ nguồn

Yêu cầu:

- .NET SDK 10.
- Python 3.13 có `pywin32`. Phải đúng 3.13 vì thư viện Python nhúng ghim hash cho cp313.
- Inno Setup 6.7.4 (script kiểm hash).
- AutoCAD.NET (NuGet của Autodesk) được tải lúc build; không kèm trong repo.

Build DLL:

```
dotnet build installer\ClaudePalette.csproj -c Release -p:RestoreLockedMode=true
```

Đóng gói bộ cài:

```
python -I -B installer\pack_setup.py [--force]
```

`pack_setup.py` mong đợi:

- công cụ tại `C:\Autodesk\tools\dotnet` và `C:\Autodesk\tools\InnoSetup6`;
- output tại `C:\Autodesk\dist`;
- repo tại `C:\Autodesk\cad-mcp`.

Chi tiết các đường dẫn: xem phần đầu `installer\pack_setup.py`.

## Kiểm thử

```
tests\run-tests.cmd
python -I -B tests\test_multi.py
python -I -B tests\test_pin.py
python -I -B tests\test_focus.py
python -I -B tests\test_tools32.py
python -I -B tests\test_spawn_sites.py
python -I -B installer\pack_setup.py --selftest
```

| Bộ test | Số ca |
|---|---|
| Bộ test C# | 375 |
| `test_multi.py` | 15 |
| `test_pin.py` | 25 |
| `test_focus.py` | 31 |
| `test_spawn_sites.py` | 24 |
| `test_tools32.py` | 169 |
| Selftest đóng gói (`pack_setup.py --selftest`) | 108 |

Script `tests\auth_env_smoke.ps1` (không nằm trong `run-tests.cmd`; chạy tay sau khi build bằng `powershell -NoProfile -ExecutionPolicy Bypass -File tests\auth_env_smoke.ps1`, cần `claude.exe`) chạy `claude` thật trong môi trường giống AutoCAD để kiểm bộ lọc biến xác thực; chỉ in tên biến và trạng thái đăng nhập, không in giá trị key.

## Trạng thái

- Test tự động đã qua: bộ test C# 375 ca; test Python `test_multi` 15, `test_pin` 25, `test_focus` 31, `test_spawn_sites` 24, `test_tools32` 169; selftest đóng gói 108 ca; build 3 target 0 lỗi.
- Chưa thử trong AutoCAD thật tại thời điểm viết tài liệu này.
- Nhánh .NET 8 (AutoCAD 2025/2026) chưa được thử trên máy nào.
- Các công cụ mới (`mirror`, `offset`, `array_rect`, `array_polar`, `ellipse`, `create_block`, `list_blocks`, `undo`, `redo`) và việc gom undo mới được thử với AutoCAD giả lập (`tests\test_tools32.py`), chưa thử trong AutoCAD thật.

## Cấu trúc thư mục

- `docs\` — tài liệu người dùng: `THU-DINH-KEM-TRONG-AUTOCAD.md`, `THU-PHIM-TAT-TRONG-AUTOCAD.md`.
- `installer\` — dự án build (`ClaudePalette.csproj`), script đóng gói (`pack_setup.py`), script Inno Setup (`cad-mcp.iss`), script build Python nhúng và kiểm tra.
- `installer\docs\` — `HDSD.html` (hướng dẫn sử dụng), `truoc-khi-cai.txt`.
- `installer\docs\img\` — ảnh minh hoạ các bước cài đặt.
- `tests\` — test Python và C#, script chạy test (`run-tests.cmd`, `build-test.cmd`).
- `tests\fixtures-xls\` — tệp `.xls` dùng cho test.
- `vendor\` — thư viện bên thứ ba: `xlrd` 2.0.2.
- Thư mục gốc — mã nguồn chính: `server.py` (MCP server), `ClaudePalette.cs`, `ClaudePalette.xaml`, `Attachments.cs`, `Hotkeys.cs`; cùng `build.cmd`, `mcp.json`, `PackageContents.xml`, `LICENSE`.

## Phụ thuộc bên thứ ba

- `vendor\xlrd` 2.0.2, giấy phép BSD. Giấy phép kèm trong `vendor\xlrd-2.0.2.dist-info\LICENSE`.
- AutoCAD.NET (NuGet của Autodesk): tải lúc build, không kèm trong repo.
- Python nhúng 3.13 và Claude Code CLI: tải lúc đóng gói/cài, không kèm trong repo.
- Claude và AutoCAD là thương hiệu của chủ sở hữu tương ứng. Dự án không thuộc Anthropic hay Autodesk.

## Tham khảo

Ý tưởng bộ công cụ 3.2 (undo/redo, mirror, offset, array, ellipse, block) lấy từ `daobataotie/CAD-MCP` (MIT, Copyright (c) 2025 曹瑞) và `2ez4gcx/autocad-mcp` (MIT, Copyright (c) 2026 Khuong Doan). Mã được viết lại cho ClaudeCAD, không sao chép.

## Giấy phép

Phần mềm được phát hành theo giấy phép MIT. Copyright (c) 2026 Gaus. Xem tệp `LICENSE`.
