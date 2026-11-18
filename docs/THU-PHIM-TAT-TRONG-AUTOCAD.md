# Thử phím tắt của palette Claude trong AutoCAD

Ngày: 2026-11-14. Nguồn: `docs\spec-hotkeys.md` (rev 2) và `docs\plan-hotkeys.md` (Task 6). Spec rev 2 thắng khi hai tài liệu khác nhau.
Đã kiểm ngoài AutoCAD: 243/243 test tự động PASS (2026-11-14), review mã ACCEPT.

## 0. Cài đặt bản mới

- Đóng AutoCAD. Nhờ Claude cài bản mới (Claude sẽ tự chạy bước cài đặt).
- Mở lại AutoCAD. Nếu AutoCAD hỏi về ClaudeCAD, chọn **Always load**.
- Đóng palette, thoát AutoCAD, mở lại. Kỳ vọng: palette không tự mở.

## 1. Chụp vùng (mặc định Ctrl+Shift+X, toàn hệ thống)

Phím này chạy trên toàn Windows khi AutoCAD đang mở.

- [ ] **[M1]** Mở PDF trong Acrobat, Acrobat đứng trước. Bấm Ctrl+Shift+X. Kỳ vọng: màn chụp vùng hiện ngay (không trễ), bấm **Esc** là huỷ ngay mà không cần bấm chuột; Acrobat đứng trước lại; palette vẫn ẩn.
- [ ] **[M2]** Như M1, nhưng kéo chọn một khung. Kỳ vọng: AutoCAD lên trước, palette hiện, có chip `chup-vung-1`, ảnh là nội dung PDF. Gõ `LINE` rồi Enter. Kỳ vọng: vào dòng lệnh, không vào ô chat.
- [ ] **[M3]** AutoCAD đứng trước, palette đang hiện, chuột trên vùng vẽ. Bấm Ctrl+Shift+X. Kỳ vọng: palette ẩn, màn chụp hiện, chọn khung thì có chip và palette hiện lại; gõ phím thì vào dòng lệnh. Nút Chụp vùng trên palette vẫn chạy như cũ.
- [ ] **[M4]** Lặp M1 và M2 với Zalo và Excel đứng trước. Kỳ vọng: vẫn chụp được. Với Zalo, Ctrl+Alt+S vẫn chụp được.
- [ ] Bấm Ctrl+Shift+X rồi bấm **chuột phải**. Kỳ vọng: huỷ như Esc (xem M1).

## 2. UniKey (kiểm Ctrl+Shift+X)

- [ ] **[M5]** Đang gõ UniKey ở chế độ V. Bấm Ctrl+Shift+X. Kỳ vọng: chụp vùng chạy và UniKey **không** đổi V/E (xem biểu tượng UniKey). Bấm-nhả riêng Ctrl+Shift vẫn đổi V/E.
  - Nếu UniKey đổi chế độ: mở Cài đặt, đặt "Phím chụp vùng" thành Ctrl+Shift+Q, rồi báo Claude.

## 3. Ẩn/hiện palette (mặc định Ctrl+5, chỉ trong AutoCAD)

- [ ] **[M6]** Đưa chuột vào vùng vẽ, bấm Ctrl+5. Kỳ vọng: palette bật/tắt. Giữ Ctrl+5 không nhấp nháy.
- [ ] **[M6]** Đặt con trỏ trong ô chat của palette, bấm Ctrl+5. Kỳ vọng: palette ẩn và gõ tiếp vào dòng lệnh được.
- [ ] **[M6]** Gõ `CLAUDE` ở dòng lệnh. Kỳ vọng: palette bật/tắt.
- [ ] **[M6]** (Tuỳ chọn) Gán thử Ctrl+5 cho một lệnh khác trong CUI. Kỳ vọng: palette vẫn được ưu tiên; ghi lại kết quả cho Claude.

## 4. Cài đặt: đổi phím

Mở palette, vào **Cài đặt**. Có 2 dòng: **Phím chụp vùng** và **Phím ẩn/hiện palette**. Bấm vào ô, rồi bấm tổ hợp mới.

- [ ] **[M7]** Ô "Phím chụp vùng": bấm Ctrl+Shift+Q. Kỳ vọng: ô hiện `Ctrl+Shift+Q`; Ctrl+Shift+X hết tác dụng; Ctrl+Shift+Q chạy được từ Acrobat. Có thể kiểm bằng `reg query HKCU\Software\ClaudeCad /v HotkeyCapture`, kết quả `Ctrl+Shift+Q`.
- [ ] **[M7]** Bấm lại ô, rồi bấm đúng Ctrl+Shift+Q. Kỳ vọng: nhận phím, không bị bắt thành chụp vùng.
- [ ] **[M7]** Ô "Phím ẩn/hiện palette": bấm Ctrl+C. Kỳ vọng: dòng đỏ, không lưu.
- [ ] **[M7]** Bấm **×** hoặc **Delete** trong ô. Kỳ vọng: ô hiện `Tắt`, phím hết tác dụng.
- [ ] **[M7]** Bấm **Backspace** trong ô. Kỳ vọng: không làm gì.
- [ ] **[M7]** Bấm **Esc** trong ô. Kỳ vọng: giữ nguyên giá trị cũ.

## 5. Các ca đặc biệt

- [ ] **[M8]** Mở AutoCAD thứ hai. Kỳ vọng: palette phiên 2 có dòng đỏ "Phím chụp vùng: Ctrl+Shift+X đang được chương trình khác dùng…", ô đỏ `(lỗi)` có tooltip. Ctrl+5 vẫn chạy ở cả hai; phiên 1 vẫn chụp được; không sập.
- [ ] **[M9]** Thu nhỏ AutoCAD, Acrobat đứng trước, bấm Ctrl+Shift+X, chọn một khung. Kỳ vọng: AutoCAD được khôi phục và có chip.
- [ ] **[M10]** Mở hộp thoại **OPTIONS**, bấm Ctrl+Shift+X. Kỳ vọng: chỉ có tiếng bíp, không hiện màn chụp, OPTIONS vẫn dùng được.
- [ ] **[M11]** Gõ `CLAUDECHUP`. Kỳ vọng: chạy chụp vùng. Gán alias hoặc phím cho `CLAUDECHUP` trong CUI, kỳ vọng: lệnh đó chạy.
- [ ] **[M12]** Đặt "Phím chụp vùng" = Win+E. Kỳ vọng: ô đỏ `(lỗi)` (Windows giữ phím). Sau đó đặt lại Ctrl+Shift+X.
- [ ] **[M13]** Thoát AutoCAD. Kỳ vọng: không lỗi. Sau khi thoát, bấm Ctrl+Shift+X trong Acrobat: không có gì xảy ra.

## 6. Thoát AutoCAD và Ctrl+G

- [ ] **[M14]** Mở một bản vẽ và sửa (chưa lưu). Thoát AutoCAD, ở hộp "Save changes?" bấm **Cancel**. Kỳ vọng: AutoCAD chạy tiếp; Ctrl+Shift+X từ Acrobat vẫn mở màn chụp vùng; Ctrl+5 vẫn bật/tắt palette; ô Cài đặt không đỏ.
- [ ] **[M15]** Mở **OPTIONS**, bấm Ctrl+5. Kỳ vọng: palette không đổi; OPTIONS vẫn dùng được; không có tiếng bíp. Đóng OPTIONS, bấm Ctrl+5 lần nữa. Kỳ vọng: Ctrl+5 chạy lại bình thường.
- [ ] **[M16]** Trong ô "Phím ẩn/hiện palette", bấm **Ctrl+G**. Kỳ vọng: dòng đỏ "Ctrl+G đã là phím của AutoCAD/Windows/UniKey.", không lưu. Bấm Ctrl+G ngoài ô (ở vùng vẽ). Kỳ vọng: Grid của AutoCAD bật/tắt.

## 7. Báo lại cho Claude

- Nói rõ số M (ví dụ "M14") và điều đã xảy ra so với kỳ vọng.
- Nếu được, gửi kèm ảnh chụp màn hình.
