#!/usr/bin/env python3
"""
จัด Rank สินค้า (A/B/C/D) รายเดือน ของ TRC

อ่าน  : Data Rank เดือนนั้น (ที่รวมแล้ว) + ยอดขายชีตปี 2023-2026 (จากไฟล์ Rank เดือนก่อน)
        + ไฟล์ CSR เดือนใหม่ (ถ้ามี)
ทำ    : SUMIFS ยอดขายรายเดือน (ชิ้น/บาท) -> เฉลี่ย -> เรียง AV 6M มาก->น้อย
        -> แยกสินค้าใหม่ / รายการตัดออก -> ใส่ลำดับ + นับ Rank -> สูตร Rank
ออก   : สมุดงานมีชีต หน้าสรุป, Rank-สินค้าทั้งหมด, Rank-สินค้าใหม่, รายการที่ตัดออก, Data RANK

ตัวอย่าง:
  python3 build_rank.py --data DATA_RANK08.xlsx --master Rank_เดือนก่อน.xlsx \
      --csr CSR_09.xlsx --year 2026 --month 9 --out Rank09.xlsx
ต้องมี: pandas, openpyxl, lxml (LibreOffice ไม่บังคับ ใช้คำนวณสูตรให้เห็นค่าทันที)
"""
import argparse, collections, copy, datetime as dt, gzip, json, os, re, shutil, subprocess, sys, tempfile, zipfile
from lxml import etree
import openpyxl
from openpyxl.utils import get_column_letter as L

HERE = os.path.dirname(os.path.abspath(__file__))
TEMPLATE = os.path.join(HERE, "..", "assets", "rank_template.xlsx")
MONTHS = ["ม.ค.", "ก.พ.", "มี.ค.", "เม.ย.", "พ.ค.", "มิ.ย.", "ก.ค.", "ส.ค.", "ก.ย.", "ต.ค.", "พ.ย.", "ธ.ค."]
# ลิสต์ Category ที่ตัดออก (จับด้วย IcCategoryCode ไม่สนตัวพิมพ์เล็ก/ใหญ่)
# หมายเหตุ: Skia-BO อยู่ในลิสต์เดิม แต่ไฟล์ตัวอย่างไม่ได้ตัดตามหมวด (ตัดเฉพาะที่เป็น Barbom)
# จึงไม่ใส่ไว้ที่นี่  ส่วน เงินมัดจำ ตัวอย่างตัดออก จึงเพิ่มให้
CUT_CATS = ("อุปกรณ์คอม,อื่นๆ,สินค้า USED,ป้ายทะเบียน,ป้ายไวนิล,ชั้นโชว์,XiangGong,Stic-ST,Service05,"
            "Service04,Service03,Service02,Service01,Service,Sample,ProP,Promotion,ProGift,ProD,ProCancel,"
            "Mate-UF,Mate-ST,Material,Mate-PI,Mate-MO,Mate-LED,Mate-FO,Mate-FM,Mate-DR,Mate-DC,Mate-CW,"
            "Mate-card,Mate-BG,Mate-AT,BarCancel,เงินมัดจำ").split(",")
NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
ILLEGAL = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f]")


def clean(v):
    return ILLEGAL.sub("", v) if isinstance(v, str) else v


def num(v):
    try:
        return float(v)
    except (TypeError, ValueError):
        return 0.0


# ------------------------------------------------------------------ อ่านยอดขายจากไฟล์ Rank เดือนก่อน
def _sheet_files(z):
    wb = etree.fromstring(z.read("xl/workbook.xml"))
    rels = etree.fromstring(z.read("xl/_rels/workbook.xml.rels"))
    rid = {r.get("Id"): r.get("Target") for r in rels}
    out = {}
    for s in wb.find(NS + "sheets"):
        t = rid[s.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id")]
        out[s.get("name")] = "xl/" + t.lstrip("/").replace("xl/", "")
    return out


def read_master_sales(path):
    """คืน sales[(ปี, เดือน0-11)][รหัส] = [ชิ้น, บาท] จากชีตชื่อ 2023, 2024, ..."""
    z = zipfile.ZipFile(path)
    files = _sheet_files(z)
    ss = []
    for _, el in etree.iterparse(z.open("xl/sharedStrings.xml"), tag=NS + "si"):
        ss.append("".join(el.itertext()))
        el.clear()
    sales = collections.defaultdict(dict)
    for name, f in files.items():
        if not re.fullmatch(r"\d{4}", name.strip()):
            continue
        year = int(name)
        hdr, idx = {}, None
        n = 0
        for _, row in etree.iterparse(z.open(f), tag=NS + "row"):
            r = int(row.get("r"))
            vals = {}
            for c in row:
                col = re.match(r"[A-Z]+", c.get("r")).group()
                v = c.find(NS + "v")
                if v is None or v.text is None:
                    continue
                t = c.get("t")
                vals[col] = ss[int(v.text)] if t == "s" else (v.text if t in ("str", "e", "b") else float(v.text))
            if r == 1:
                hdr = {str(v).strip(): k for k, v in vals.items()}
                try:
                    cc = hdr["รหัสสินค้า"]; cq = hdr["จำนวนที่สั่ง"]; cu = hdr["ราคารวม"]
                except KeyError:
                    sys.exit(f"ชีต {name}: ไม่พบหัวคอลัมน์ รหัสสินค้า / จำนวนที่สั่ง / ราคารวม")
                cm = hdr.get("เดือนถอดสูตร") or hdr.get("เดือน")
                if not cm:
                    sys.exit(f"ชีต {name}: ไม่พบคอลัมน์เดือน (เดือน / เดือนถอดสูตร)")
                idx = (cc, cq, cu, cm)
            elif idx and idx[0] in vals and idx[3] in vals:
                code = str(vals[idx[0]]).lower()   # Excel SUMIFS ไม่สนตัวพิมพ์เล็ก/ใหญ่
                mon = str(vals[idx[3]]).strip()
                if mon in MONTHS:
                    d = sales[(year, MONTHS.index(mon))].setdefault(code, [0.0, 0.0])
                    d[0] += num(vals.get(idx[1])); d[1] += num(vals.get(idx[2]))
                    n += 1
            row.clear()
        print(f"  ชีต {name}: {n:,} แถวยอดขาย")
    return sales


def read_csr(path, sales):
    """ต่อท้ายยอดขายจากไฟล์ CSR ใหม่ (แทนที่เดือนที่ซ้ำ ไม่ให้นับซ้ำ)"""
    import pandas as pd
    df = pd.read_csv(path, dtype=str) if path.lower().endswith(".csv") else pd.read_excel(path)
    df.columns = [str(c).strip() for c in df.columns]
    for c in ("รหัสสินค้า", "จำนวนที่สั่ง", "ราคารวม"):
        if c not in df.columns:
            sys.exit(f"{os.path.basename(path)}: ไม่พบคอลัมน์ {c}")

    def ym(row):
        m, y = row.get("เดือน"), row.get("ปี")
        if isinstance(m, str) and m.strip() in MONTHS and y == y and y is not None:
            return int(float(y)), MONTHS.index(m.strip())
        d = row.get("วันที่เอกสาร")
        if isinstance(d, (int, float)) and d == d:
            d = dt.datetime(1899, 12, 30) + dt.timedelta(days=float(d))
        d = pd.to_datetime(d, errors="coerce")
        return (int(d.year), int(d.month) - 1) if d == d else None

    new = collections.defaultdict(dict)
    for row in df.to_dict("records"):
        k = ym(row)
        code = row.get("รหัสสินค้า")
        if not k or code != code or code is None:
            continue
        d = new[k].setdefault(str(code).lower(), [0.0, 0.0])
        d[0] += num(row.get("จำนวนที่สั่ง")); d[1] += num(row.get("ราคารวม"))
    for k, v in sorted(new.items()):
        old = len(sales.get(k, {}))
        sales[k] = v
        print(f"  CSR {MONTHS[k[1]]} {k[0]}: {len(v):,} รหัส" + (f" (แทนที่ของเดิม {old:,} รหัส)" if old else ""))
    return sales


# ------------------------------------------------------------------ อ่าน Data Rank
ALIAS = {
    "code": ["Product Code"], "cat": ["Category", "IcCategoryCode"], "desc": ["Description"],
    "brand": ["ยี่ห้อ"], "model": ["รุ่น"], "catname": ["IcCategoryName"],
    "list": ["List Price"], "member": ["Member"], "whole": ["Wholesale General"],
    "bal": ["Balance (All WH)"], "vendor": ["PropComposer"], "new": ["New Product"], "bb": ["Bar Bom", "Barbom"],
}



# ------------------------------------------------------------------ cache ยอดขาย (json.gz ไม่ใช้ pickle เพื่อความปลอดภัย)
def save_sales(sales, path):
    obj = {f"{y}-{m}": d for (y, m), d in sales.items()}
    with gzip.open(path, "wt", encoding="utf-8") as f:
        json.dump(obj, f, ensure_ascii=False)


def load_sales(path):
    with gzip.open(path, "rt", encoding="utf-8") as f:
        obj = json.load(f)
    return {tuple(int(x) for x in k.split("-")): v for k, v in obj.items()}


# ------------------------------------------------------------------ ย่อไฟล์ Rank (เทมเพลต) ให้เล็กลง เก็บแค่สไตล์/หัวตาราง
def shrink_template(src, dst):
    """ตัดข้อมูลแถวยาวๆ ออกจากทุกชีต เก็บ 13 แถวแรก (ชีต Rank) / 3 แถวแรก (ชีตอื่น) แล้วบันทึกเป็นไฟล์เล็ก"""
    zi = zipfile.ZipFile(src)
    files = _sheet_files(zi)
    rank_files = {f for n, f in files.items() if n.strip().startswith(("Rank-", "รายการที่ตัดออก", "หน้าสรุป"))}
    tmp = dst + ".tmp.xlsx"
    zo = zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED)
    for it in zi.infolist():
        b = zi.read(it.filename)
        if it.filename in files.values():
            keep = 14 if it.filename in rank_files else 4
            i = b.find(b'<row r="%d"' % keep)
            j = b.find(b"</sheetData>")
            if 0 < i < j:
                b = b[:i] + b[j:]
            b = re.sub(rb"<filterColumn.*?</filterColumn>", b"", b, flags=re.S)
            b = re.sub(rb"<sortState.*?</sortState>", b"", b, flags=re.S)
            b = b.replace(b'<sheetPr filterMode="1">', b"<sheetPr>")
        zo.writestr(it, b)
    zo.close()
    wb = openpyxl.load_workbook(tmp)
    for n in list(wb.sheetnames):
        if re.fullmatch(r"\d{4}", n.strip()):
            del wb[n]
    wb.save(dst)
    os.remove(tmp)
    return dst


def find_sheet(wb, prefix):
    for n in wb.sheetnames:
        if n.strip().startswith(prefix):
            return wb[n]
    sys.exit(f"เทมเพลตไม่มีชีตที่ชื่อขึ้นต้นด้วย '{prefix}'")


def detect_template(tws):
    """อ่านตำแหน่งคอลัมน์ในเทมเพลตจากหัวตาราง (รองรับเทมเพลตที่มีจำนวนเดือนปีล่าสุดไม่เท่ากัน)"""
    hdr = next((r for r in range(1, 15) if str(tws.cell(r, 1).value or "").strip() == "ลำดับ"), 10)
    row_h = [(c.column, str(c.value).strip()) for c in tws[hdr] if c.value is not None]
    row_m = [(c.column, str(c.value).strip()) for c in tws[hdr + 1] if c.value is not None]
    qs, bs, avy, sumy = {}, {}, {}, {}
    av6 = sum6 = fnew = fbb = None
    for col, v in row_h:
        m = re.fullmatch(r"ขาย \(ชิ้น/เดือน\) (\d{4})", v)
        if m: qs[int(m.group(1))] = col
        m = re.fullmatch(r"ขาย \(บาท/เดือน\) (\d{4})", v)
        if m: bs[int(m.group(1))] = col
        if v == "New Product": fnew = col
        if "BOM" in v.upper(): fbb = col
    for col, v in row_m:
        m = re.fullmatch(r"AV Y(\d{4})", v)
        if m: avy[int(m.group(1))] = col
        m = re.fullmatch(r"SUM Y(\d{4})", v)
        if m: sumy[int(m.group(1))] = col
        if v == "AV 6M": av6 = col
        if v == "SUM 6M": sum6 = col
    if not qs or not bs or not avy or av6 is None:
        sys.exit("อ่านโครงเทมเพลตไม่ออก (ไม่พบหัว 'ขาย (ชิ้น/เดือน) ปี' / 'AV 6M')")
    years = sorted(qs)

    def lens(starts, stop):
        ys = sorted(starts)
        return {y: ((starts[ys[i + 1]] if i + 1 < len(ys) else stop) - starts[y]) for i, y in enumerate(ys)}
    return dict(hdr=hdr, years=years, qs=qs, bs=bs, avy=avy, sumy=sumy, av6=av6, sum6=sum6,
                fnew=fnew or sum6 + 1, fbb=fbb or sum6 + 2,
                qlen=lens(qs, min(avy.values())), blen=lens(bs, min(sumy.values()) if sumy else max(bs.values()) + 12))


def read_data(path):
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    names = [s for s in wb.sheetnames if s.lower().replace(" ", "").startswith("datarank")]
    vals = [s for s in names if "(ค่า)" in s]          # ชีตค่าจากขั้นรวมข้อมูล (ไม่ต้องคำนวณสูตรก่อน)
    name = (vals or names or wb.sheetnames[:1])[0]
    ws = wb[name]
    it = ws.iter_rows(values_only=True)
    header = [str(h).strip() if h is not None else "" for h in next(it)]
    pos = {}
    for k, names in ALIAS.items():
        for n in names:
            if n in header:
                pos[k] = header.index(n); break
    for k in ("code", "cat", "desc"):
        if k not in pos:
            sys.exit(f"Data Rank ไม่พบคอลัมน์ {ALIAS[k]}")
    rows = []
    for r in it:
        code = r[pos["code"]]
        if code is None or str(code).strip() == "":
            continue
        d = {k: (r[p] if p < len(r) else None) for k, p in pos.items()}
        d["code"] = str(code).strip()
        rows.append(d)
    if rows and sum(1 for d in rows if d.get("desc") in (None, "")) > 0.5 * len(rows):
        sys.exit("Data Rank ไม่มีค่าที่คำนวณแล้ว (ไฟล์สูตรที่ยังไม่เคยเปิดใน Excel) "
                 "ให้ใช้ไฟล์จากขั้นรวมข้อมูล หรือเปิดใน Excel แล้วบันทึกก่อน")
    return rows, header, name, path


def flag(v, word):
    return isinstance(v, str) and v.strip().lower().startswith(word)


def classify(rows, cut_cats):
    cut = {c.strip().lower() for c in cut_cats}
    allr, newr, cutr = [], [], []
    for d in rows:
        if flag(d.get("bb"), "bar"):
            cutr.append(d)                      # Barbom ตัดออกก่อนเสมอ (แม้เป็น New Product)
        elif flag(d.get("new"), "new"):
            newr.append(d)                      # สินค้าใหม่ (ถ้าไม่ใช่ Barbom)
        elif str(d.get("cat") or "").strip().lower() in cut:
            cutr.append(d)
        else:
            allr.append(d)
    return allr, newr, cutr


# ------------------------------------------------------------------ คำนวณ
def window(year, month0, n):
    out = []
    y, m = year, month0
    for _ in range(n):
        out.append((y, m))
        m -= 1
        if m < 0:
            y, m = y - 1, 11
    return out[::-1]


def months_of(year, last_year, k):
    return 12 if year < last_year else k


def compute(rows, sales, years, last_year, k, n_av=6):
    """ใส่ q (ชิ้น) และ b (บาท) รายเดือนให้แต่ละแถว  คืน AV6M สำหรับเรียง"""
    win = window(last_year, k - 1, n_av)
    for d in rows:
        q, b = {}, {}
        for y in years:
            for m in range(months_of(y, last_year, k)):
                v = sales.get((y, m), {}).get(d["code"].lower())
                q[(y, m)] = v[0] if v else 0.0
                b[(y, m)] = v[1] if v else 0.0
        d["q"], d["b"] = q, b
        d["av6"] = sum(q.get(w, 0.0) for w in win) / n_av
    return win


# ------------------------------------------------------------------ เลย์เอาต์คอลัมน์
def layout(years, last_year, k):
    """คืน dict ตำแหน่งคอลัมน์ (เลขคอลัมน์ 1 เริ่ม) ให้ตรงกับเทมเพลตเดิม (เดือนล่าสุดขยายทีละคอลัมน์)"""
    c = 16
    q = {}
    for y in years:
        for m in range(months_of(y, last_year, k)):
            q[(y, m)] = c; c += 1
    avy = {}
    for y in years:
        avy[y] = c; c += 1
    av6 = c; c += 1
    b = {}
    for y in years:
        for m in range(months_of(y, last_year, k)):
            b[(y, m)] = c; c += 1
    sumy = {}
    for y in years:
        sumy[y] = c; c += 1
    sum6 = c; c += 1
    return dict(q=q, avy=avy, av6=av6, b=b, sumy=sumy, sum6=sum6, fnew=c, fbb=c + 1, last=c + 1)


# แปลงตำแหน่งคอลัมน์ของเรา -> คอลัมน์ในเทมเพลต เพื่อยืมสไตล์
def tmpl_col(tl, kind, years, last_year, y=None, m=None):
    T = tl["years"]

    def pick(yy):
        if yy == last_year:
            return T[-1]
        if yy in T and yy != T[-1]:
            return yy
        return T[-2] if len(T) > 1 else T[-1]
    if kind == "q":
        ty = pick(y); return tl["qs"][ty] + min(m, tl["qlen"][ty] - 1)
    if kind == "b":
        ty = pick(y); return tl["bs"][ty] + min(m, tl["blen"][ty] - 1)
    if kind == "avy":
        return tl["avy"][pick(y)]
    if kind == "sumy":
        return tl["sumy"].get(pick(y), tl["sum6"])
    return tl[kind]                       # av6 | sum6 | fnew | fbb


def col_width(ws, idx):
    for cd in ws.column_dimensions.values():
        if cd.min and cd.min <= idx <= (cd.max or cd.min):
            return cd.width
    return None


# ------------------------------------------------------------------ เขียนชีต Rank
def sheet_rank(wb, tws, title, rows, years, last_year, k, win, kind, first_row):
    """kind: all | new | cut   first_row = แถวข้อมูลแถวแรก (12 หรือ 10)"""
    ws = wb.create_sheet(title)
    ws.sheet_properties.tabColor = copy.copy(tws.sheet_properties.tabColor)
    ws.sheet_format.defaultRowHeight = tws.sheet_format.defaultRowHeight
    ws.sheet_format.customHeight = tws.sheet_format.customHeight
    lay = layout(years, last_year, k)
    n = len(rows)
    N = first_row + n - 1 if n else first_row
    hdr_r = first_row - 2          # แถวหัวกลุ่ม (10 หรือ 8)
    mon_r = first_row - 1          # แถวชื่อเดือน
    tl = detect_template(tws)
    thdr, tmon, tdata = tl["hdr"], tl["hdr"] + 1, tl["hdr"] + 2

    def sty(dst, tr, tc):
        dst._style = copy.copy(tws.cell(row=tr, column=tc)._style)

    # --- ความกว้างคอลัมน์ / ซ้อนหัว
    for c in range(1, lay["last"] + 1):
        tc = c if c <= 15 else None
        if tc is None:
            for kk, d in (("q", lay["q"]), ("b", lay["b"])):
                for (y, m), cc in d.items():
                    if cc == c:
                        tc = tmpl_col(tl, kk, years, last_year, y, m)
            for kk, d in (("avy", lay["avy"]), ("sumy", lay["sumy"])):
                for y, cc in d.items():
                    if cc == c:
                        tc = tmpl_col(tl, kk, years, last_year, y)
            if c == lay["av6"]: tc = tl["av6"]
            if c == lay["sum6"]: tc = tl["sum6"]
            if c == lay["fnew"]: tc = tl["fnew"]
            if c == lay["fbb"]: tc = tl["fbb"]
        w = col_width(tws, tc)
        if w:
            ws.column_dimensions[L(c)].width = w
        lay.setdefault("tc", {})[c] = tc

    tc = lay["tc"]
    # --- หัวตาราง (แถวหัวกลุ่ม + ชื่อเดือน) สไตล์จากเทมเพลต
    for c in range(1, lay["last"] + 1):
        sty(ws.cell(hdr_r, c), thdr, tc[c]); sty(ws.cell(mon_r, c), tmon, tc[c])
    for c, t in enumerate(["ลำดับ", "Category", "Product Code", "Description", "ยี่ห้อ", "รุ่น", "ประเภทสินค้า", "List Price",
                           "Member", "Wholesale General", "สต๊อกคงเหลือ", "มูลค่าสินค้า", "Vender", "สินค้าโครงการ Rank C / D",
                           "Rank สินค้าปัจุบัน"], 1):
        ws.cell(hdr_r, c).value = t
    ws.cell(mon_r, 12).value = "จำนวนคงเหลือxราคาส่ง"
    for y in years:
        first = lay["q"][(y, 0)]; lastc = lay["q"][(y, months_of(y, last_year, k) - 1)]
        ws.cell(hdr_r, first).value = f"ขาย (ชิ้น/เดือน) {y}"
        ws.merge_cells(start_row=hdr_r, start_column=first, end_row=hdr_r, end_column=lastc)
        first = lay["b"][(y, 0)]; lastc = lay["b"][(y, months_of(y, last_year, k) - 1)]
        ws.cell(hdr_r, first).value = f"ขาย (บาท/เดือน) {y}"
        ws.merge_cells(start_row=hdr_r, start_column=first, end_row=hdr_r, end_column=lastc)
        for m in range(months_of(y, last_year, k)):
            ws.cell(mon_r, lay["q"][(y, m)]).value = MONTHS[m]
            ws.cell(mon_r, lay["b"][(y, m)]).value = MONTHS[m]
        ws.cell(mon_r, lay["avy"][y]).value = f"AV Y{y}"
        ws.cell(mon_r, lay["sumy"][y]).value = f"SUM Y{y}"
    ws.cell(mon_r, lay["av6"]).value = "AV 6M"
    ws.cell(mon_r, lay["sum6"]).value = "SUM 6M"
    ws.merge_cells(start_row=hdr_r, start_column=lay["avy"][years[0]], end_row=hdr_r, end_column=lay["av6"])
    ws.cell(hdr_r, lay["fnew"]).value = "New Product"
    ws.cell(hdr_r, lay["fbb"]).value = "บาร์ BOM TYPE"

    # --- แถวข้อมูล
    ca6, cs6 = L(lay["av6"]), L(lay["sum6"])
    cut_pos = None
    if kind == "all":
        a = sum(1 for d in rows if d["av6"] >= 2.5)
        b = sum(1 for d in rows if 0.5 <= d["av6"] <= 2.49)
        c_ = sum(1 for d in rows if 0.1 <= d["av6"] <= 0.49)
        cut_pos = (a, a + b, a + b + c_)
    elif kind == "new":
        cut_pos = (int(n * 0.2), int(n * 0.5), n)
    for i, d in enumerate(rows):
        r = first_row + i
        vals = {1: i + 1, 2: d.get("cat"), 3: d["code"], 4: d.get("desc"), 5: d.get("brand"), 6: d.get("model"),
                7: d.get("catname"), 8: d.get("list"), 9: d.get("member"), 10: d.get("whole"), 11: d.get("bal"),
                12: f"=J{r}*K{r}", 13: d.get("vendor") or None}
        if kind != "cut":
            off = first_row - 1
            vals[15] = (f'=IF(ROW()-{off}<={cut_pos[0]},"A",IF(ROW()-{off}<={cut_pos[1]},"B",'
                        f'IF(ROW()-{off}<={cut_pos[2]},"C","D")))')
        for (y, m), cc in lay["q"].items():
            vals[cc] = d["q"][(y, m)]
        for (y, m), cc in lay["b"].items():
            vals[cc] = d["b"][(y, m)]
        for y in years:
            qs = [lay["q"][(y, m)] for m in range(months_of(y, last_year, k))]
            bs = [lay["b"][(y, m)] for m in range(months_of(y, last_year, k))]
            vals[lay["avy"][y]] = f"=AVERAGEA({L(qs[0])}{r}:{L(qs[-1])}{r})"
            vals[lay["sumy"][y]] = f"=SUM({L(bs[0])}{r}:{L(bs[-1])}{r})"
        wq = [lay["q"][w] for w in win]; wb_ = [lay["b"][w] for w in win]
        vals[lay["av6"]] = f"=AVERAGEA({L(wq[0])}{r}:{L(wq[-1])}{r})"
        vals[lay["sum6"]] = f"=SUM({L(wb_[0])}{r}:{L(wb_[-1])}{r})"
        if flag(d.get("new"), "new"):
            vals[lay["fnew"]] = "New Product"
        if flag(d.get("bb"), "bar"):
            vals[lay["fbb"]] = "Barbom"
        for c in range(1, lay["last"] + 1):
            cell = ws.cell(r, c)
            sty(cell, tdata, tc[c])
            v = vals.get(c)
            if v is not None and v != "":
                cell.value = clean(v)
    ws.freeze_panes = f"A{first_row}"
    # --- สรุปบนหัว
    summary(ws, tws, kind, lay, win, first_row, N, n, cut_pos, rows, k)
    return ws, lay


def cp(ws, tws, coord, value=None, formula=None):
    c = ws[coord]
    c._style = copy.copy(tws[coord]._style)
    if formula is not None:
        c.value = formula
    elif value is not None:
        c.value = value
    return c


def summary(ws, tws, kind, lay, win, fr, N, n, cut_pos, rows, k):
    ca6 = L(lay["av6"])
    # ก๊อปสไตล์ทั้งบล็อกบนหัว (A1:Z{fr-3}) จากเทมเพลต
    last_top = fr - 3
    for r in range(1, last_top + 1):
        for c in range(1, 27):
            cell = ws.cell(r, c); t = tws.cell(r, c)
            cell._style = copy.copy(t._style)
            if t.value is not None and not (isinstance(t.value, str) and t.value.startswith("=")):
                cell.value = t.value
    for rng in tws.merged_cells.ranges:
        if rng.min_row <= last_top and rng.max_col <= 26:
            ws.merge_cells(str(rng))
    if kind == "new":
        mons = win[-3:]
        qc = [L(lay["q"][w]) for w in mons]; bc = [L(lay["b"][w]) for w in mons]
        for j, w in enumerate(mons):
            ws.cell(2, 11 + j).value = MONTHS[w[1]]; ws.cell(2, 16 + j).value = MONTHS[w[1]]
        ws["D3"] = f"=COUNT(A{fr}:A1048576)"; ws["B3"] = "=SUM(B4:B6)"; ws["J3"] = "=D3"
        for j in range(3):
            q, p = L(11 + j), L(16 + j)
            ws[f"{q}3"] = f"=SUM({q}4:{q}6)"; ws[f"{p}3"] = f"=SUM({p}4:{p}6)"
        ws["N3"] = "=AVERAGE(K3:M3)"; ws["S3"] = "=SUM(P3:R3)"
        for r in (4, 5, 6):
            ws[f"D{r}"] = f"=$D$3*B{r}"; ws[f"J{r}"] = f"=D{r}"
            for j in range(3):
                ws[f"{L(11+j)}{r}"] = f"=SUMIF($O${fr}:$O${N},$I{r},{qc[j]}${fr}:{qc[j]}${N})"
                ws[f"{L(16+j)}{r}"] = f"=SUMIF($O${fr}:$O${N},$I{r},{bc[j]}${fr}:{bc[j]}${N})"
            ws[f"N{r}"] = f"=AVERAGE(K{r}:M{r})"; ws[f"O{r}"] = f"=N{r}/$N$3*$O$3"
            ws[f"S{r}"] = f"=SUM(P{r}:R{r})"; ws[f"T{r}"] = f"=$S{r}/$S$3*$T$3"
        return
    # all / cut
    for j, w in enumerate(win):
        ws.cell(2, 11 + j).value = MONTHS[w[1]]; ws.cell(2, 19 + j).value = MONTHS[w[1]]
    qc = [L(lay["q"][w]) for w in win]; bc = [L(lay["b"][w]) for w in win]
    rng = f"{ca6}{fr}:{ca6}{N}"
    ws["D3"] = f'=COUNTIF(${ca6}${fr}:${ca6}${N},">=0.01")'
    ws["J3"] = "=D3"
    a = sum(1 for d in rows if d["av6"] >= 2.5)
    b = sum(1 for d in rows if 0.5 <= d["av6"] <= 2.49)
    c_ = sum(1 for d in rows if 0.1 <= d["av6"] <= 0.49)
    ws["F4"] = f'=COUNTIF({rng},">=2.5")'
    ws["F5"] = f'=COUNTIFS({rng},">=0.5",{rng},"<=2.49")'
    ws["F6"] = f'=COUNTIFS({rng},">=0.1",{rng},"<=0.49")'
    ws["F7"] = f'=COUNTIF(${ca6}${fr}:${ca6}${N},"<=0")'
    for r, v in ((4, a), (5, b), (6, c_)):          # ขั้นตอน 8: วางเป็นค่า ไม่เอาสูตร
        ws[f"D{r}"] = v
        ws[f"J{r}"] = f"=D{r}"
        ws[f"B{r}"] = f"=$D{r}/$D$3"
    ws["D7"] = "=D8-D3"
    ws["D8"] = f"=COUNT(A{fr}:A1048576)"
    if kind == "cut":
        return
    for j in range(6):
        q, p = L(11 + j), L(19 + j)
        ws[f"{q}3"] = f"=SUM({q}4:{q}6)"; ws[f"{p}3"] = f"=SUM({p}4:{p}6)"
        for r in (4, 5, 6):
            ws[f"{q}{r}"] = f"=SUMIF($O${fr}:$O${N},$I{r},{qc[j]}${fr}:{qc[j]}${N})"
            ws[f"{p}{r}"] = f"=SUMIF($O${fr}:$O${N},$I{r},{bc[j]}${fr}:{bc[j]}${N})"
    ws["Q3"] = "=AVERAGE(K3:P3)"; ws["Y3"] = "=SUM(S3:X3)"
    for r in (4, 5, 6):
        ws[f"Q{r}"] = f"=AVERAGE(K{r}:P{r})"; ws[f"R{r}"] = f"=Q{r}/$Q$3*$R$3"
        ws[f"Y{r}"] = f"=SUM(S{r}:X{r})"; ws[f"Z{r}"] = f"=$Y{r}/$Y$3*$Z$3"


# ------------------------------------------------------------------ หน้าสรุป + ชีตข้อมูล
def sheet_summary(wb, tws, name_all, win):
    ws = wb.create_sheet("หน้าสรุป")
    ws.sheet_properties.tabColor = copy.copy(tws.sheet_properties.tabColor)
    for cd in tws.column_dimensions.values():
        ws.column_dimensions[cd.index].width = cd.width
    q = f"'{name_all}'!"
    for r in range(1, 9):
        for c in range(1, 24):
            t = tws.cell(r, c); cell = ws.cell(r, c)
            cell._style = copy.copy(t._style)
            v = t.value
            if isinstance(v, str) and v.startswith("="):
                m = re.search(r"!([A-Z]+\d+)$", v)
                if m:
                    cell.value = f"={q}{m.group(1)}"
                else:
                    cell.value = v
            elif v is not None:
                cell.value = v
    for rng in tws.merged_cells.ranges:
        ws.merge_cells(str(rng))
    for j, w in enumerate(win):
        ws.cell(2, 8 + j).value = MONTHS[w[1]]; ws.cell(2, 16 + j).value = MONTHS[w[1]]
    for r in (4, 5, 6):
        ws[f"W{r}"] = f"={q}Z{r}"
    ws["W3"] = f"={q}Z3"
    for c in range(7, 24):          # แถว 7 ในต้นฉบับลิงก์ไปไฟล์นอกที่เสีย จึงเว้นว่าง
        ws.cell(7, c).value = None
    return ws


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--data", required=True, help="Data Rank เดือนนี้ (xlsx ที่รวมแล้ว)")
    ap.add_argument("--master", help="ไฟล์ Rank เดือนก่อนที่มีชีตปี 2023, 2024, ... (ใช้ดึงยอดขาย)")
    ap.add_argument("--sales-cache", help="ไฟล์ .json.gz เก็บยอดขายที่อ่านแล้ว (ไม่ต้องอ่านไฟล์ใหญ่ซ้ำ)")
    ap.add_argument("--template", help="ไฟล์เทมเพลต Rank (ไม่ใส่ = ใช้ --master หรือเทมเพลตที่ติดมากับสกิล)")
    ap.add_argument("--csr", nargs="*", default=[], help="ไฟล์ CSR เดือนใหม่ที่จะต่อท้ายปีนั้น")
    ap.add_argument("--year", type=int, required=True)
    ap.add_argument("--month", type=int, required=True, help="เดือนที่จัด Rank 1-12")
    ap.add_argument("--cut-cats", help="ไฟล์ .txt ลิสต์ Category ตัดออก (บรรทัดละ 1 ค่า) ใช้แทนค่าเริ่มต้น")
    ap.add_argument("--out", required=True)
    ap.add_argument("--no-recalc", action="store_true")
    a = ap.parse_args(argv)
    k = a.month

    # 1) ยอดขาย
    if a.sales_cache and os.path.exists(a.sales_cache):
        sales = load_sales(a.sales_cache)
        print(f"โหลดยอดขายจาก cache {a.sales_cache}")
    else:
        if not a.master:
            sys.exit("ต้องใส่ --master หรือ --sales-cache ที่มีอยู่แล้ว")
        print("อ่านยอดขายจากไฟล์ master (ใช้เวลาสักครู่)...")
        sales = read_master_sales(a.master)
    for p in a.csr:
        print(f"ต่อท้าย CSR: {p}")
        sales = read_csr(p, sales)
    if a.sales_cache:
        save_sales(sales, a.sales_cache)
    yrs = sorted({y for y, _ in sales} | {a.year})
    years = [y for y in yrs if y <= a.year]
    if not any(key[0] == a.year and key[1] == k - 1 for key in sales):
        print(f"! ยังไม่มียอดขายเดือน {MONTHS[k-1]} {a.year} (ใส่ --csr หรือเช็กชีตปีนั้น)")

    # 2) Data Rank + แยกกลุ่ม
    rows, header, dname, _ = read_data(a.data)
    cut_cats = [x.strip() for x in open(a.cut_cats, encoding="utf-8")] if a.cut_cats else CUT_CATS
    allr, newr, cutr = classify(rows, cut_cats)
    print(f"Data Rank {len(rows):,} รายการ -> ทั้งหมด {len(allr):,} | สินค้าใหม่ {len(newr)} | ตัดออก {len(cutr):,}")

    # 3) คำนวณ + เรียง AV 6M มาก->น้อย (เรียงแบบคงลำดับเดิมเมื่อเท่ากัน)
    win = None
    for grp in (allr, newr, cutr):
        win = compute(grp, sales, years, a.year, k)
        grp.sort(key=lambda d: -d["av6"])
    print("ช่วง 6 เดือน:", ", ".join(f"{MONTHS[m]} {y}" for y, m in win))

    # 4) เขียนสมุดงาน
    tpl_src = a.template or a.master
    if tpl_src:
        print("เตรียมเทมเพลตจาก", os.path.basename(tpl_src))
        tpl_path = shrink_template(tpl_src, os.path.join(tempfile.mkdtemp(), "tpl.xlsx"))
    else:
        tpl_path = TEMPLATE
    wb = openpyxl.load_workbook(tpl_path)
    t_all, t_new, t_cut, t_sum = (find_sheet(wb, "Rank-สินค้าทั้งหมด"), find_sheet(wb, "Rank-สินค้าใหม่"),
                                  find_sheet(wb, "รายการที่ตัดออก"), find_sheet(wb, "หน้าสรุป"))
    nm_all = f"Rank-สินค้าทั้งหมด {k:02d}-{a.year}"
    ws_all, lay = sheet_rank(wb, t_all, nm_all, allr, years, a.year, k, win, "all", 12)
    ws_new, _ = sheet_rank(wb, t_new, "Rank-สินค้าใหม่", newr, years, a.year, k, win, "new", 10)
    ws_cut, _ = sheet_rank(wb, t_cut, "รายการที่ตัดออก", cutr, years, a.year, k, win, "cut", 12)
    ws_sum = sheet_summary(wb, t_sum, nm_all, win)
    ws_dat = wb.create_sheet(f"Data RANK{k:02d}")
    ws_dat.append(header)
    src = openpyxl.load_workbook(a.data, data_only=True, read_only=True)[dname]
    for i, r in enumerate(src.iter_rows(values_only=True)):
        if i:
            ws_dat.append([clean(x) for x in r])
    keep = {id(x) for x in (ws_all, ws_new, ws_cut, ws_sum, ws_dat)}
    for n in list(wb.sheetnames):               # ทิ้งชีตของเทมเพลตทั้งหมด เหลือเฉพาะชีตที่สร้างใหม่
        if id(wb[n]) not in keep:
            del wb[n]
    ws_new.title, ws_cut.title, ws_sum.title, ws_dat.title = "Rank-สินค้าใหม่", "รายการที่ตัดออก", "หน้าสรุป", f"Data RANK{k:02d}"
    ws_all.title = nm_all
    order = ["หน้าสรุป", nm_all, "Rank-สินค้าใหม่", "รายการที่ตัดออก", f"Data RANK{k:02d}"]
    wb._sheets = [wb[n] for n in order]
    wb.active = 1
    wb.save(a.out)
    print("บันทึกแล้ว:", a.out)
    if not a.no_recalc:
        soffice = shutil.which("soffice") or shutil.which("libreoffice")
        if soffice:
            tmp = tempfile.mkdtemp()
            subprocess.run([soffice, "--headless", "--convert-to", "xlsx:Calc MS Excel 2007 XML", "--outdir", tmp, a.out],
                           check=True, capture_output=True, timeout=1800)
            shutil.move(os.path.join(tmp, os.path.basename(a.out)), a.out)
            print("คำนวณสูตรด้วย LibreOffice แล้ว")
        else:
            print("! ไม่พบ LibreOffice (เปิดใน Excel สูตรจะคำนวณเอง)")
    a6 = sum(1 for d in allr if d["av6"] >= 2.5); b6 = sum(1 for d in allr if 0.5 <= d["av6"] <= 2.49)
    c6 = sum(1 for d in allr if 0.1 <= d["av6"] <= 0.49)
    print(f"\nสรุป Rank-สินค้าทั้งหมด: A={a6:,} B={b6:,} C={c6:,} (มียอดขาย {a6+b6+c6:,}) D={len(allr)-a6-b6-c6:,}")


if __name__ == "__main__":
    main()
