"""Tạo fixture cho tests\\AttachTests.cs. Dùng: python -I mkfixtures.py <thư mục ra>"""
import datetime, os, struct, sys, zipfile, zlib

out = sys.argv[1]
os.makedirs(out, exist_ok=True)


def w(name, data):
    with open(os.path.join(out, name), "wb") as f:
        f.write(data)


# ---- văn bản ----
w("vi_utf8.txt", "Xin chào Việt Nam\nđường kính Ø20\n".encode("utf-8"))
w("vi_utf8bom.lsp", b"\xef\xbb\xbf" + '(princ "Bản vẽ")\n'.encode("utf-8"))
w("vi_utf16le.txt", b"\xff\xfe" + "Xin chào Việt Nam\n".encode("utf-16-le"))
w("vi_1258.csv", "Tên;Đơn giá\nThép;5\n".encode("cp1258"))
w("cafe_1252.txt", "café naïve\n".encode("cp1252"))
w("binary.lsp", b"(princ)\x00\x00\x01\x02")
w("big.txt", b"x" * 150000)
w("a.docx", b"PK\x03\x04 not really a docx")


# ---- ảnh ----
def chunk(tag, data):
    c = struct.pack(">I", len(data)) + tag + data
    return c + struct.pack(">I", zlib.crc32(tag + data) & 0xFFFFFFFF)


SIG = b"\x89PNG\r\n\x1a\n"


def png(wd, ht, extra=b""):
    row = b"\x00" + b"\xff\x00\x00" * (wd // 2) + b"\x00\x00\xff" * (wd - wd // 2)  # trái đỏ, phải xanh
    return (SIG + chunk(b"IHDR", struct.pack(">IIBBBBB", wd, ht, 8, 2, 0, 0, 0)) + extra
            + chunk(b"IDAT", zlib.compress(row * ht)) + chunk(b"IEND", b""))


def huge_png(wd, ht):  # 1-bit xám, nén từng dòng: tệp nhỏ nhưng header 20000×20000
    co = zlib.compressobj(9)
    row = b"\x00" + bytes((wd + 7) // 8)
    data = b"".join(co.compress(row) for _ in range(ht)) + co.flush()
    return (SIG + chunk(b"IHDR", struct.pack(">IIBBBBB", wd, ht, 1, 0, 0, 0, 0))
            + chunk(b"IDAT", data) + chunk(b"IEND", b""))


w("small.png", png(64, 32))
w("big.png", png(4000, 1000))
w("shot.png", png(2560, 1440))  # T3: ảnh chụp màn hình giữ nguyên độ phân giải
w("meta.png", png(64, 32, chunk(b"tEXt", b"Comment\x00SECRET-GPS-123")))
w("huge.png", huge_png(20000, 20000))
w("fake.webp", b"RIFF" + struct.pack("<I", 30) + b"WEBPVP8 " + bytes(22))
w("garbage.webp", b"hello, not an image at all")


# ---- PDF (1 trang, chữ ZEBRA-7) ----
def pdf(text):
    stream = b"BT /F1 24 Tf 72 700 Td (" + text + b") Tj ET"
    objs = [
        b"<< /Type /Catalog /Pages 2 0 R >>",
        b"<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
        b"<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Contents 4 0 R /Resources << /Font << /F1 5 0 R >> >> >>",
        b"<< /Length %d >>\nstream\n" % len(stream) + stream + b"\nendstream",
        b"<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>",
    ]
    d = bytearray(b"%PDF-1.4\n")
    offs = []
    for i, o in enumerate(objs, 1):
        offs.append(len(d))
        d += b"%d 0 obj\n" % i + o + b"\nendobj\n"
    xref = len(d)
    d += b"xref\n0 %d\n0000000000 65535 f \n" % (len(objs) + 1)
    for o in offs:
        d += b"%010d 00000 n \n" % o
    d += b"trailer\n<< /Size %d /Root 1 0 R >>\nstartxref\n%d\n%%%%EOF\n" % (len(objs) + 1, xref)
    return bytes(d)


w("doc.pdf", pdf(b"ZEBRA-7"))
w("notpdf.pdf", b"this is not a pdf")
w("big.pdf", b"%PDF-1.4\n" + b"0" * 9000000)
for k in range(1, 9):
    w("bad%d.pdf" % k, b"not a pdf %d" % k)

# ---- Excel .xlsx ----
HDR = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
NS = 'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"'
NSR = 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/"
PKGREL = 'xmlns="http://schemas.openxmlformats.org/package/2006/relationships"'
CT = (HDR + '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
      '<Default Extension="xml" ContentType="application/xml"/></Types>')
ROOTRELS = HDR + '<Relationships ' + PKGREL + '><Relationship Id="rId1" Type="' + REL + 'officeDocument" Target="xl/workbook.xml"/></Relationships>'
WB = (HDR + '<workbook ' + NS + ' ' + NSR + '><workbookPr/><sheets>'
      '<sheet name="Dữ liệu" sheetId="1" r:id="rId1"/><sheet name="Ẩn" sheetId="2" state="hidden" r:id="rId2"/>'
      '</sheets></workbook>')
WBRELS = (HDR + '<Relationships ' + PKGREL + '>'
          '<Relationship Id="rId1" Type="' + REL + 'worksheet" Target="worksheets/sheet1.xml"/>'
          '<Relationship Id="rId2" Type="' + REL + 'worksheet" Target="/xl/worksheets/sheet2.xml"/>'
          '<Relationship Id="rId3" Type="' + REL + 'styles" Target="styles.xml"/>'
          '<Relationship Id="rId4" Type="' + REL + 'sharedStrings" Target="sharedStrings.xml"/></Relationships>')
SST_BODY = ('<sst ' + NS + ' count="6" uniqueCount="6">'
            '<si><t>Tên</t></si><si><t>Ngày</t></si><si><t>Thép</t></si><si><t>Gộp</t></si>'
            '<si><r><rPr><b/></rPr><t>Ti</t></r><r><t xml:space="preserve">êu</t></r><rPh sb="0" eb="1"><t>x</t></rPh></si>'
            '<si><t>bí mật</t></si></sst>')
STYLES = (HDR + '<styleSheet ' + NS + '>'
          '<numFmts count="2"><numFmt numFmtId="164" formatCode="dd/mm/yyyy hh:mm"/><numFmt numFmtId="165" formatCode="[Red]0.00"/></numFmts>'
          '<cellStyleXfs count="1"><xf numFmtId="14"/></cellStyleXfs>'
          '<cellXfs count="4"><xf numFmtId="0"/><xf numFmtId="14" applyNumberFormat="1"/>'
          '<xf numFmtId="164" applyNumberFormat="1"/><xf numFmtId="165" applyNumberFormat="1"/></cellXfs></styleSheet>')
SHEET1 = (HDR + '<worksheet ' + NS + ' ' + NSR + '><dimension ref="A1:Z99"/><sheetData>'
          '<row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1" t="s"><v>1</v></c><c r="C1" t="inlineStr"><is><t>Tổng</t></is></c></row>'
          '<row r="2"><c r="A2" t="s"><v>2</v></c><c r="B2" s="1"><v>45292</v></c><c r="C2"><f>D2*2</f><v>84</v></c><c r="D2" s="3"><v>42</v></c></row>'
          '<row r="3"><c r="A3"><f>SUM(D2:D3)</f></c><c r="B3" t="b"><v>1</v></c><c r="C3" t="e"><v>#DIV/0!</v></c><c r="D3"><v>0.30000000000000004</v></c></row>'
          '<row r="4"><c r="B4" s="2"><v>45292.5</v></c></row>'
          '<row r="5"><c r="A5" t="s"><v>3</v></c><c r="D5" t="inlineStr"><is><t>x\ny</t></is></c></row>'
          '<row r="6"><c r="A6" t="s"><v>4</v></c><c r="D6" t="inlineStr"><is><t>a\tb</t></is></c></row>'
          '<row r="7"><c r="A7"/><c r="B7" t="s"/></row>'
          '</sheetData><mergeCells count="1"><mergeCell ref="A5:C5"/></mergeCells></worksheet>')
SHEET2 = HDR + '<worksheet ' + NS + '><sheetData><row r="1"><c r="A1" t="s"><v>5</v></c></row></sheetData></worksheet>'


def xlsx(name, parts):
    with zipfile.ZipFile(os.path.join(out, name), "w", zipfile.ZIP_DEFLATED) as z:
        for n, s in parts.items():
            z.writestr(n, s.encode("utf-8") if isinstance(s, str) else s)


def blank_sheet(nbytes):  # sheetData toàn khoảng trắng: nén rất nhỏ, giải nén rất lớn
    return (HDR + '<worksheet ' + NS + '><sheetData>').encode() + b" " * nbytes + b"</sheetData></worksheet>"


def one_sheet_book(sheet_bytes, nsheets=1):
    sheets = "".join('<sheet name="S%d" sheetId="%d" r:id="rId1"/>' % (k, k) for k in range(1, nsheets + 1))
    return {"xl/workbook.xml": HDR + '<workbook ' + NS + ' ' + NSR + '><sheets>' + sheets + '</sheets></workbook>',
            "xl/_rels/workbook.xml.rels": HDR + '<Relationships ' + PKGREL + '><Relationship Id="rId1" Type="' + REL
            + 'worksheet" Target="worksheets/sheet1.xml"/></Relationships>',
            "xl/worksheets/sheet1.xml": sheet_bytes}


def spoof(path, name, fake):  # ghi kích thước giải nén GIẢ vào cả local header lẫn central directory
    d = bytearray(open(path, "rb").read())
    nb = name.encode()
    i = d.find(b"PK\x01\x02")
    while i != -1:
        n = struct.unpack_from("<H", d, i + 28)[0]
        if d[i + 46:i + 46 + n] == nb:
            struct.pack_into("<I", d, i + 24, fake)
            lho = struct.unpack_from("<I", d, i + 42)[0]
            struct.pack_into("<I", d, lho + 22, fake)
        i = d.find(b"PK\x01\x02", i + 4)
    open(path, "wb").write(bytes(d))


BOOK = {"[Content_Types].xml": CT, "_rels/.rels": ROOTRELS, "xl/workbook.xml": WB, "xl/_rels/workbook.xml.rels": WBRELS,
        "xl/sharedStrings.xml": HDR + SST_BODY, "xl/styles.xml": STYLES,
        "xl/worksheets/sheet1.xml": SHEET1, "xl/worksheets/sheet2.xml": SHEET2}
xlsx("book.xlsx", BOOK)
xlsx("dtd.xlsx", dict(BOOK, **{"xl/sharedStrings.xml":
                               '<?xml version="1.0"?>\n<!DOCTYPE sst [<!ENTITY e "boom">]>\n' + SST_BODY.replace("<t>Tên</t>", "<t>&e;</t>")}))
xlsx("bomb.xlsx", one_sheet_book(blank_sheet(120000000)))
xlsx("spoof.xlsx", one_sheet_book(blank_sheet(110000000)))
spoof(os.path.join(out, "spoof.xlsx"), "xl/worksheets/sheet1.xml", 1000)
xlsx("manyrels.xlsx", one_sheet_book(blank_sheet(3000000), 150))
w("notzip.xlsx", b"not a zip file")


# ---- T7: nhiều sheet, ô gộp sau sheetData, phần trăm, ngân sách đọc ----
def rows_xml(n, ncol=7, start=1):
    out = []
    for r in range(start, start + n):
        cells = "".join('<c r="%s%d" t="inlineStr"><is><t>muc %d cot %d</t></is></c>' % (chr(65 + c), r, r, c) for c in range(ncol))
        out.append('<row r="%d">%s</row>' % (r, cells))
    return "".join(out)


def multi_book(sheets, styles=None):
    """sheets = [(tên, xml sheet)]; sheet i dùng rId(i+1)"""
    wb = HDR + '<workbook ' + NS + ' ' + NSR + '><sheets>' + "".join(
        '<sheet name="%s" sheetId="%d" r:id="rId%d"/>' % (n, i + 1, i + 1) for i, (n, _) in enumerate(sheets)) + '</sheets></workbook>'
    rels = HDR + '<Relationships ' + PKGREL + '>' + "".join(
        '<Relationship Id="rId%d" Type="%sworksheet" Target="worksheets/sheet%d.xml"/>' % (i + 1, REL, i + 1) for i in range(len(sheets)))
    if styles:
        rels += '<Relationship Id="rId99" Type="' + REL + 'styles" Target="styles.xml"/>'
    rels += '</Relationships>'
    parts = {"xl/workbook.xml": wb, "xl/_rels/workbook.xml.rels": rels}
    for i, (_, x) in enumerate(sheets):
        parts["xl/worksheets/sheet%d.xml" % (i + 1)] = x
    if styles:
        parts["xl/styles.xml"] = styles
    return parts


PCT_STYLES = (HDR + '<styleSheet ' + NS + '><numFmts count="2"><numFmt numFmtId="164" formatCode="0.0%"/>'
              '<numFmt numFmtId="165" formatCode="0.00&quot;%&quot;"/></numFmts>'
              '<cellXfs count="4"><xf numFmtId="0"/><xf numFmtId="9"/><xf numFmtId="164"/><xf numFmtId="165"/></cellXfs></styleSheet>')


def ws(dim, body, tail=""):
    return HDR + '<worksheet ' + NS + '><dimension ref="' + dim + '"/><sheetData>' + body + '</sheetData>' + tail + '</worksheet>'


summary = ('<row r="1"><c r="A1" t="inlineStr"><is><t>ty le</t></is></c><c r="B1" s="1"><v>0.02</v></c><c r="C1" s="2"><v>0.0225</v></c>'
           '<c r="D1" s="3"><v>0.5</v></c><c r="E1" s="1"><v>NaN</v></c><c r="F1" s="1"><v>1e999</v></c><c r="G1" s="1"><v>abc</v></c></row>')
xlsx("fair.xlsx", multi_book([
    ("Dữ liệu", ws("A1:G5002", rows_xml(5000), '<mergeCells count="1"><mergeCell ref="A1:G1"/></mergeCells>')),
    ("Tổng hợp", ws("A1:G1", summary)),
    ("Phụ", ws("A1:G3000", rows_xml(3000))),
], PCT_STYLES))

# 1 sheet lớn + 1 sheet nhỏ: sheet lớn phải nhận gần hết ngân sách (chia kiểu đổ nước, không phải 1/n)
xlsx("bigsmall.xlsx", multi_book([
    ("Lớn", ws("A1:G3000", rows_xml(3000))),
    ("Nhỏ", ws("A1:G5", rows_xml(5))),
]))

# 200 dòng (dưới trần ký tự) rồi 110 MB khoảng trắng: ngân sách/hạn giờ cạn khi CHƯA chạm trần ký tự -> giữ 200 dòng + cảnh báo
xlsx("mainbomb.xlsx", multi_book([
    ("Chính", (HDR + '<worksheet ' + NS + '><dimension ref="A1:G200"/><sheetData>' + rows_xml(200) + " " * 110000000
              + '</sheetData></worksheet>').encode("utf-8")),
]))

# sheet 1 vượt trần ký tự rồi còn 110 MB khoảng trắng: giai đoạn đuôi cạn ngân sách đọc -> vẫn đính kèm được (có cảnh báo)
xlsx("tailbomb.xlsx", multi_book([
    ("Lớn", (HDR + '<worksheet ' + NS + '><dimension ref="A1:G3000"/><sheetData>' + rows_xml(3000) + " " * 110000000
             + '</sheetData><mergeCells count="1"><mergeCell ref="A1:B1"/></mergeCells></worksheet>').encode("utf-8")),
    ("Sau", ws("A1:A1", rows_xml(1, 1))),
]))

# ô gộp 1 MB + ref lạ bị bỏ; <dimension> lạ không được in ra
xlsx("badmerge.xlsx", multi_book([
    ("M", ws("A1:B2 IGNORE-PREVIOUS", rows_xml(2, 2),
             '<mergeCells count="3"><mergeCell ref="A1:B1"/><mergeCell ref="' + "A" * 1000000 + '1"/><mergeCell ref="A1:B2;evil"/></mergeCells>')),
]))

# R-F1: CPU DoS qua styles.xml: formatCode 2 MB + ~900k <xf> cùng trỏ numFmtId 164 (nén còn rất nhỏ)
dos = one_sheet_book(HDR.encode() + ('<worksheet ' + NS + '><sheetData><row r="1"><c r="A1" s="5"><v>45292</v></c></row></sheetData></worksheet>').encode())
dos["xl/styles.xml"] = (HDR + '<styleSheet ' + NS + '><numFmts count="1"><numFmt numFmtId="164" formatCode="' + "a" * 2000000
                        + '"/></numFmts><cellXfs count="900000">' + '<xf numFmtId="164"/>' * 900000 + '</cellXfs></styleSheet>')
xlsx("dos.xlsx", dos)
# m7: 1.000.001 chuỗi chung rỗng (vượt trần 1.000.000 mục)
sst = one_sheet_book(HDR.encode() + ('<worksheet ' + NS + '><sheetData/></worksheet>').encode())
sst["xl/sharedStrings.xml"] = HDR + '<sst ' + NS + '>' + "<si/>" * 1000001 + '</sst>'
xlsx("sst.xlsx", sst)

# ---- Excel .xls (chỉ thử đường lỗi; tệp .xls thật do tests\mk-xls-fixtures.ps1 tạo) ----
w("old.xls", b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1" + bytes(504))  # đúng magic OLE, ruột hỏng
w("notole.xls", b"hello, not an OLE file")
w("pk.xls", b"PK\x03\x04" + bytes(100))
w("hang.py", b"import time\ntime.sleep(60)\n")
w("flood.py", b"import sys\nsys.stdout.buffer.write(b'x' * 2000000)\n")

import openpyxl  # noqa: E402  (đã có: openpyxl 3.1.5)
wb = openpyxl.Workbook()
ws = wb.active
ws.title = "Bảng"
ws.append(["Tên", "Ngày", "Số"])
ws.append(["Xin chào", datetime.datetime(2024, 1, 1), 3.5])
ws["D2"] = "=C2*2"
wb.save(os.path.join(out, "real.xlsx"))
wb = openpyxl.Workbook()
ws = wb.active
for _ in range(20000):
    ws.append(["abcdefghij"] * 5)
wb.save(os.path.join(out, "bigrows.xlsx"))
print("fixtures ok:", out)
