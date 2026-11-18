# Thử tính năng đính kèm trong AutoCAD

Tài liệu này dành cho người dùng AutoCAD. Làm theo thứ tự từ trên xuống.

## Lưu ý an toàn

> **Lưu ý an toàn**
>
> Khi có tệp đính kèm, Claude vẫn dùng được mọi công cụ, kể cả chạy mã Python và lệnh AutoCAD tuỳ ý trên máy với quyền của bạn (đọc/xoá tệp, gửi dữ liệu ra ngoài), không chỉ sửa bản vẽ; nội dung độc trong tệp có thể khiến Claude làm vậy cho tới khi bấm Mới. Chỉ đính kèm tệp bạn tin.
>
> Tệp đính kèm được gửi tới Anthropic và được lưu trong lịch sử phiên Claude Code trên máy này.

## Đã kiểm trước khi giao (ngoài AutoCAD)

- Chạy `tests\run-tests.cmd` ngày 2026-11-14: 126/126 kiểm tự động PASS.
- Chạy khung thử giao diện: đã dựng hình mọi trạng thái của palette.
- Rà kỹ thuật: ACCEPT. Rà bảo mật: ACCEPT sau khi sửa.

Loại tệp hỗ trợ và giới hạn:

- Ảnh: png, jpg, bmp, gif, webp.
- PDF: tối đa 8 MB.
- Văn bản: txt, csv, json, lsp, py, md; tối đa 100 000 ký tự, dài hơn sẽ bị cắt và có thông báo.
- Excel: xlsx và xls (xls đọc trong tiến trình riêng, tối đa 10 MB).
- Mỗi tin nhắn tối đa 10 tệp và 15 MB.

Cách đính kèm: Ctrl+V; kéo thả vào bất kỳ chỗ nào trên palette; nút 📎 (Đính kèm); nút Chụp vùng (palette ẩn, kéo chọn một hình chữ nhật, nhấn Esc hoặc bấm chuột phải để huỷ).

## Phần 1: Cài bản mới

1. Lưu tất cả bản vẽ đang mở.
2. Đóng hẳn AutoCAD (kiểm tra không còn cửa sổ AutoCAD nào, vì tệp DLL đang bị khoá khi AutoCAD chạy).
3. Bấm đúp tệp `C:\Autodesk\cad-mcp\build.cmd`. Cuối cửa sổ phải có dòng `Installed: …ClaudeCAD.bundle`.
4. Mở lại AutoCAD, gõ `CLAUDE` rồi Enter để mở palette Claude.

## Phần 2: Kiểm tra từng mục

Mỗi mục có một ô `[ ]`. Làm xong mục nào thì đánh dấu `[x]`. Nếu kết quả khác mong đợi, ghi lại mã mục (ví dụ M7) để báo lại.

### 2.1 Đính kèm bằng nút 📎

- [ ] **[M1]** Bấm 📎 (Đính kèm), chọn một ảnh png, một PDF và một tệp xlsx. Mong đợi: chip hiện "đang đọc…" rồi đủ 3 chip. Di chuột lên chip ảnh thấy ảnh xem trước. AutoCAD không bị đứng.

### 2.2 Dán bằng Ctrl+V

- [ ] **[M2]** Bấm PrtScn (hoặc dùng Snipping Tool rồi chọn Copy), đưa con trỏ vào ô nhập của palette, bấm Ctrl+V. Mong đợi: có chip tên `anh-dan-1.png`.
- [ ] **[M3]** Copy vài ô trong Excel, bấm Ctrl+V trong palette. Mong đợi: nội dung dán vào là chữ (không phải ảnh).
- [ ] **[M6]** Trong Windows Explorer, chọn (Ctrl+C) một tệp, rồi bấm Ctrl+V trong palette. Mong đợi: có chip cho tệp đó.

### 2.3 Kéo thả tệp

- [ ] **[M4]** Kéo 2 tệp từ Explorer vào lần lượt: khung Hội thoại, ô nhập, thanh Cài đặt. Mong đợi: khi đang kéo thì hiện lớp phủ; thả xong có chip; kéo ra ngoài thì lớp phủ ẩn.
- [ ] **[M5]** Kéo một tệp vào palette rồi bấm Esc trước khi thả. Mong đợi: lớp phủ ẩn và không có chip nào được thêm.

### 2.4 Chụp vùng màn hình

- [ ] **[M7]** Bấm nút Chụp vùng. Mong đợi: palette biến mất, lớp phủ chụp hiện ra, ảnh chụp không còn palette. Kéo chọn một vùng: có chip `chup-vung-N.png` và palette hiện lại. Lặp lại ba lần, mỗi lần kết thúc bằng một cách khác: bấm Esc, bấm chuột phải, và bấm Alt+Tab để chuyển cửa sổ. Mong đợi: lần nào cũng không có chip và palette hiện lại.
- [ ] **[M15]** Sau khi dùng hộp chọn tệp (📎) hoặc Chụp vùng, quay lại AutoCAD và gõ `COPY` rồi Enter, sau đó bấm Enter hoặc Esc. Mong đợi: lệnh chạy trong AutoCAD, không gõ vào ô chat của palette.

### 2.5 Gửi và đọc tệp

- [ ] **[M8]** Chụp một vùng bản vẽ, gửi kèm câu hỏi (ví dụ "Vùng này có gì?"). Mong đợi: tin nhắn có chip ảnh, Claude mô tả đúng nội dung vùng đã chụp.
- [ ] **[M9]** Chọn một đối tượng trong bản vẽ, đính kèm một tệp xlsx rồi gửi. Mong đợi: nhật ký có dòng "Kèm N đối tượng". Claude thấy cả đối tượng lẫn nội dung xlsx.
- [ ] **[M10]** Ở lượt sau, hỏi tiếp về tệp đã gửi ở lượt trước (không đính kèm lại). Mong đợi: Claude nhớ và trả lời đúng.
- [ ] **[M12]** Đính kèm một tệp xlsx đang mở trong Excel. Mong đợi: palette đọc được nội dung.
- [ ] **[M14]** Để ô chữ trống, đã có ít nhất một chip, rồi bấm Enter. Mong đợi: tin nhắn được gửi với câu `Xem các tệp đính kèm.`

### 2.6 Giới hạn và lỗi

- [ ] **[M11]** Đính kèm cùng lúc: một tệp .docx, một thư mục, một tệp 20 MB, và 8 tệp PDF hỏng. Mong đợi: mỗi tệp lỗi có một dòng đỏ ghi đúng lý do. Tối đa 5 dòng đỏ, các dòng còn lại gộp thành 1 dòng.
- [ ] **[M13]** Kéo cạnh palette để thu bề rộng về mức nhỏ nhất. Mong đợi: 5 nút nằm trên 1 hàng.
- [ ] **[M19]** Đính kèm một tệp xls lớn, bấm Gửi khi chip còn ghi "đang đọc…". Mong đợi: hiện ghi chú `Đang đọc tệp đính kèm, đợi xong rồi gửi.` và chữ trong ô vẫn còn nguyên.
- [ ] **[M18]** Đính kèm một tệp xls lớn, bấm × trên chip khi còn "đang đọc…" (hoặc bấm Mới). Mong đợi: chip biến mất. Mở Task Manager, không còn tiến trình python con nào.

### 2.7 Tệp .xls

- [ ] **[M16]** Đính kèm một tệp .xls thật có sheet ẩn. Mong đợi: có chip. Di chuột lên chip thấy dòng `… 1 sheet ẩn (vẫn gửi)`. Claude đọc đúng nội dung.
- [ ] **[M17]** Đính kèm một tệp .xls có mật khẩu. Mong đợi: chip báo màu đỏ `không đọc được .xls (tệp hỏng hoặc có mật khẩu).`

## Phần 3: Việc chưa kiểm được, cần bạn tự làm

Các mục sau chưa được kiểm chứng. Hãy làm và báo kết quả, kể cả khi đạt:

- Kéo thả tệp vào bên trong palette AutoCAD (xem M4).
- Ctrl+V ảnh chụp màn hình bằng PrtScn và bằng "Copy image" của trình duyệt (xem M2).
- Sau khi chụp vùng, hoặc sau khi đóng hộp chọn tệp mà chuột đang nằm trên bản vẽ: gõ `LINE` rồi Enter, lệnh phải chạy trong AutoCAD, không gõ vào ô chat (xem M7 và M15).
- Chụp vùng khi có nhiều màn hình hoặc các màn hình có tỉ lệ phóng (DPI) khác nhau (xem M7).
- Palette có biến mất hoàn toàn trong ảnh chụp không. Nếu còn thấy palette trong ảnh, hãy báo lại (xem M7).
- Gửi thật với đủ ba loại: một ảnh, một PDF và một xlsx trong cùng một tin nhắn bấm Gửi (xem M1 và M8).

## Cách báo lỗi

Nói với Claude theo mẫu: `M7 lỗi: bấm Esc thì palette không hiện lại`. Nếu có thể, đính kèm ảnh màn hình chụp lúc lỗi.
