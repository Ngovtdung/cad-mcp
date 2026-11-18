"""Fixture đo hiệu năng đính kèm (docs\\spec-attachments-perf.md). Nội dung tổng hợp, không dùng tệp user.
Dùng: python -I -B mkperf.py <thư mục ra>   (cần Pillow + openpyxl, đã có trên máy dev)
Ghi kèm gt.json = đáp án đúng để chấm câu trả lời của Claude."""
import datetime, json, os, random, sys, zlib

from PIL import Image, ImageDraw, ImageFont
import openpyxl

out = sys.argv[1]
os.makedirs(out, exist_ok=True)
rnd = random.Random(20261117)
gt = {}
PAGES = (1, 10, 30)


def font(px):
    try:
        return ImageFont.truetype("arial.ttf", px)
    except OSError:
        return ImageFont.load_default()


# ---- 1. ảnh chụp màn hình bản vẽ nét mảnh trên nền đen, 2560x1440 ----
W, H = 2560, 1440
im = Image.new("RGB", (W, H), (0, 0, 0))
d = ImageDraw.Draw(im)
d.rectangle([0, 0, W, 120], fill=(59, 68, 83))                       # ribbon
for i in range(40):
    d.text((10 + i * 63, 20), "Cmd%d" % i, fill=(245, 245, 245), font=font(11))
for x in range(0, W, 50):
    d.line([(x, 120), (x, H)], fill=(38, 38, 38), width=1)           # lưới mờ
for y in range(120, H, 50):
    d.line([(0, y), (W, y)], fill=(38, 38, 38), width=1)
cols = [(255, 0, 0), (255, 255, 0), (0, 255, 255), (0, 255, 0), (255, 255, 255)]
for _ in range(120):                                                 # nét 1 px
    x0, y0 = rnd.randrange(W), rnd.randrange(130, H)
    d.line([(x0, y0), (x0 + rnd.randrange(-400, 400), y0 + rnd.randrange(-300, 300))], fill=rnd.choice(cols), width=1)
for _ in range(25):
    cx, cy, r = rnd.randrange(100, W - 100), rnd.randrange(230, H - 100), rnd.randrange(10, 90)
    d.ellipse([cx - r, cy - r, cx + r, cy + r], outline=rnd.choice(cols), width=1)
labels = []
for k in range(12):                                                  # nhãn chữ 11 px
    t = "D-%04d" % rnd.randrange(1000, 9999)
    labels.append(t)
    d.text((150 + (k % 4) * 600, 300 + (k // 4) * 380), t, fill=(255, 255, 0), font=font(11))
gt["shot_labels"] = labels
im.save(os.path.join(out, "shot.png"), optimize=True)

# ---- 2. ảnh chụp 4000x3000 JPEG, EXIF orientation 6 ----
small = Image.new("RGB", (400, 300))
sp = small.load()
for y in range(300):
    for x in range(400):
        sp[x, y] = (int(120 + 100 * x / 400), int(90 + 120 * y / 300), int(60 + 60 * ((x + y) % 50) / 50))
ph = small.resize((4000, 3000), Image.Resampling.BICUBIC)
ph = Image.blend(ph, Image.effect_noise((4000, 3000), 40).convert("RGB"), 0.25)
ImageDraw.Draw(ph).text((200, 200), "PHOTO-58", fill=(255, 255, 255), font=font(160))
gt["photo_text"] = "PHOTO-58"
ex = Image.Exif()
ex[274] = 6
ph.save(os.path.join(out, "photo.jpg"), quality=92, exif=ex.tobytes())


# ---- PDF viết tay: mỗi trang = (content stream, MediaBox) ----
def pdf(pages, box):
    objs = ["<< /Type /Catalog /Pages 2 0 R >>", None, "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>"]
    kids = []
    for body in pages:
        z = zlib.compress(body.encode("latin-1"), 9)
        objs.append(("<< /Length %d /Filter /FlateDecode >>\nstream\n" % len(z)).encode() + z + b"\nendstream")
        objs.append("<< /Type /Page /Parent 2 0 R /MediaBox [0 0 %d %d] /Contents %d 0 R /Resources << /Font << /F1 3 0 R >> >> >>"
                    % (box[0], box[1], len(objs)))
        kids.append(len(objs))
    objs[1] = "<< /Type /Pages /Kids [%s] /Count %d >>" % (" ".join("%d 0 R" % k for k in kids), len(kids))
    b = bytearray(b"%PDF-1.4\n")
    offs = []
    for i, o in enumerate(objs, 1):
        offs.append(len(b))
        b += ("%d 0 obj\n" % i).encode() + (o if isinstance(o, bytes) else o.encode("latin-1")) + b"\nendobj\n"
    x = len(b)
    b += ("xref\n0 %d\n0000000000 65535 f \n" % (len(objs) + 1)).encode()
    for o in offs:
        b += ("%010d 00000 n \n" % o).encode()
    b += ("trailer\n<< /Size %d /Root 1 0 R >>\nstartxref\n%d\n%%%%EOF\n" % (len(objs) + 1, x)).encode()
    return bytes(b)


words = "thep tam ban ma bulong lien ket khung dam cot ton nhom go van kich thuoc chieu dai rong cao do day so luong".split()

# ---- 3. PDF chữ (lớp chữ thật, A4), mã kiểm tra ở trang cuối ----
for n in PAGES:
    pages = []
    for p in range(n):
        lines = ["Trang %d / %d - Thuyet minh ky thuat" % (p + 1, n)] + [" ".join(rnd.choice(words) for _ in range(14)) for _ in range(55)]
        if p == n - 1:
            lines[20] = "Ma kiem tra: MAGENTA-%d" % (40 + n)
        pages.append("BT /F1 10 Tf 50 800 Td 13 TL " + " ".join("(%s) '" % ln for ln in lines) + " ET")
    gt["text_%d" % n] = "trang %d: MAGENTA-%d" % (n, 40 + n)
    open(os.path.join(out, "text_%d.pdf" % n), "wb").write(pdf(pages, (595, 842)))

# ---- 4. PDF scan (ảnh JPEG mỗi trang, KHÔNG có lớp chữ), A4 150 dpi ----
for n in PAGES:
    scans = []
    for p in range(n):
        pg = Image.new("L", (1240, 1754), 245)
        g = ImageDraw.Draw(pg)
        for ln in range(45):
            g.text((90, 90 + ln * 36), " ".join(rnd.choice(words) for _ in range(9)), fill=30, font=font(22))
        if p == n - 1:
            g.text((90, 1650), "SCAN-TOKEN-%d" % (900 + n), fill=0, font=font(28))
        pg = Image.blend(pg, Image.effect_noise((1240, 1754), 25), 0.06).rotate(0.6, fillcolor=245)
        scans.append(pg.convert("RGB"))
    gt["scan_%d" % n] = "trang %d: SCAN-TOKEN-%d" % (n, 900 + n)
    scans[0].save(os.path.join(out, "scan_%d.pdf" % n), save_all=True, append_images=scans[1:], resolution=150, quality=75)

# ---- 5. PDF bản vẽ CAD vector A3 ngang: nét 0.24 pt, chữ là NÉT (giống SHX, không có lớp chữ) ----
SEG = {"0": "abcdef", "1": "bc", "2": "abged", "3": "abgcd", "4": "fgbc", "5": "afgcd", "6": "afgedc", "7": "abc", "8": "abcdefg", "9": "abcdfg"}


def strokes(txt, x, y, h):  # chữ số 7 đoạn, cao h pt
    w, s = h * 0.55, []
    pts = {"a": (0, h, w, h), "b": (w, h, w, h / 2), "c": (w, h / 2, w, 0), "d": (0, 0, w, 0),
           "e": (0, 0, 0, h / 2), "f": (0, h / 2, 0, h), "g": (0, h / 2, w, h / 2)}
    for ch in txt:
        for seg in SEG.get(ch, ""):
            a, b, c, e = pts[seg]
            s.append("%.2f %.2f m %.2f %.2f l S" % (x + a, y + b, x + c, y + e))
        x += w + h * 0.35
    return " ".join(s)


for n in PAGES:
    pages, last = [], None
    for p in range(n):
        c = ["0.24 w 0 G 20 20 1151 802 re S 0.5 w 851 20 320 90 re S 0.24 w"]
        for _ in range(1500):
            x0, y0 = rnd.uniform(30, 1160), rnd.uniform(120, 810)
            c.append("%.1f %.1f m %.1f %.1f l S" % (x0, y0, x0 + rnd.uniform(-120, 120), y0 + rnd.uniform(-90, 90)))
        for _ in range(40):
            cx, cy, r = rnd.uniform(80, 1100), rnd.uniform(150, 780), rnd.uniform(4, 40)
            k = r * 0.5523
            c.append("%.1f %.1f m %.1f %.1f %.1f %.1f %.1f %.1f c %.1f %.1f %.1f %.1f %.1f %.1f c %.1f %.1f %.1f %.1f %.1f %.1f c %.1f %.1f %.1f %.1f %.1f %.1f c S"
                     % (cx + r, cy, cx + r, cy + k, cx + k, cy + r, cx, cy + r, cx - k, cy + r, cx - r, cy + k, cx - r, cy,
                        cx - r, cy - k, cx - k, cy - r, cx, cy - r, cx + k, cy - r, cx + r, cy - k, cx + r, cy))
        nums = ["%05d" % rnd.randrange(10000, 99999) for _ in range(8)]
        for k, t in enumerate(nums):  # 8 kích thước chữ 7 pt (≈2.5 mm khi in A3)
            x, y = 80 + (k % 4) * 260, 200 + (k // 4) * 300
            c.append("%.1f %.1f m %.1f %.1f l S" % (x - 10, y - 4, x + 60, y - 4))
            c.append(strokes(t, x, y, 7))
        code = "%04d" % rnd.randrange(1000, 9999)
        c.append(strokes(code, 880, 50, 9))  # mã bản vẽ trong khung tên
        pages.append(" ".join(c))
        if p == 0:
            gt["cad_%d_page1_code" % n] = code
        last = nums
    gt["cad_%d_lastpage_dims" % n] = last
    open(os.path.join(out, "cad_%d.pdf" % n), "wb").write(pdf(pages, (1191, 842)))

# ---- 6. xlsx 5000 dòng: dòng tiêu đề gộp, hàng tên cột, ngày, tiền, %, sheet thứ 2 ----
wb = openpyxl.Workbook()
ws = wb.active
ws.title = "Vật tư"
ws["A1"] = "BẢNG KÊ VẬT TƯ"
ws.merge_cells("A1:G1")
ws.append(["Mã", "Tên vật tư", "Ngày nhập", "SL", "Đơn giá", "Thành tiền", "Hao hụt"])
d0 = datetime.date(2026, 1, 1)
for i in range(5000):
    q, u = rnd.randrange(1, 500), rnd.randrange(5000, 900000)
    ws.append(["VT%05d" % i, "Thép hộp %dx%d dày %.1f" % (rnd.choice([20, 30, 40, 50]), rnd.choice([20, 40, 60]), rnd.choice([1.2, 1.4, 2.0])),
               d0 + datetime.timedelta(days=i % 300), q, u, q * u, rnd.choice([0.02, 0.03, 0.05])])
for r in ws.iter_rows(min_row=3, max_row=5002):
    r[2].number_format = "dd/mm/yyyy"
    r[4].number_format = '#,##0 "đ"'
    r[5].number_format = '#,##0 "đ"'
    r[6].number_format = "0%"
s2 = wb.create_sheet("Tổng hợp")
s2.append(["Nhóm", "Tổng"])
s2.append(["Thép", 123456789])
wb.save(os.path.join(out, "rows5k.xlsx"))

# ---- 7. txt ~60 KB ----
txt = "\n".join("Dòng %d: ghi chú bản vẽ, kích thước %d x %d mm" % (i, rnd.randrange(100, 9000), rnd.randrange(100, 9000)) for i in range(1300))
open(os.path.join(out, "notes.txt"), "w", encoding="utf-8").write(txt)

json.dump(gt, open(os.path.join(out, "gt.json"), "w", encoding="utf-8"), ensure_ascii=False, indent=1)
print("perf fixtures ok:", out)
