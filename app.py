"""Rank สินค้า TRC  (Streamlit)
ส่วนที่ 1  รวมข้อมูล -> Data Rank   (scripts/merge_data_rank.py)
ส่วนที่ 2  จัด Rank สินค้า A/B/C/D   (scripts/build_rank.py)
"""
import contextlib
import io
import os
import re
import shutil
import sys
import tempfile
import traceback

import pandas as pd
import streamlit as st

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(HERE, "scripts"))
import build_rank  # noqa: E402
import merge_data_rank  # noqa: E402

st.set_page_config(page_title="Rank สินค้า TRC", page_icon="📊", layout="wide")


# ------------------------------------------------------------------ ตัวช่วย
def password_gate():
    """ถ้าตั้ง APP_PASSWORD ใน Secrets จะต้องใส่รหัสก่อนใช้งาน"""
    try:
        pw = st.secrets.get("APP_PASSWORD")
    except Exception:
        pw = None
    if not pw or st.session_state.get("auth"):
        return
    st.title("📊 Rank สินค้า TRC")
    x = st.text_input("รหัสผ่าน", type="password")
    if x:
        if x == pw:
            st.session_state["auth"] = True
            st.rerun()
        else:
            st.error("รหัสผ่านไม่ถูกต้อง")
    st.stop()


def save_upload(up, workdir, name=None):
    path = os.path.join(workdir, name or up.name)
    with open(path, "wb") as f:
        shutil.copyfileobj(up, f)
    return path


def run(func, argv):
    """รันสคริปต์ในโปรเซสเดียวกัน จับข้อความที่พิมพ์ออกมา  คืน (สำเร็จไหม, log)"""
    buf = io.StringIO()
    ok = True
    with contextlib.redirect_stdout(buf):
        try:
            func(argv)
        except SystemExit as e:
            if e.code not in (None, 0):
                print("\n❌", e.code)
                ok = False
        except Exception:
            print("\n❌ เกิดข้อผิดพลาด\n" + traceback.format_exc())
            ok = False
    return ok, buf.getvalue()


MONTH_NAMES = ["ม.ค.", "ก.พ.", "มี.ค.", "เม.ย.", "พ.ค.", "มิ.ย.", "ก.ค.", "ส.ค.", "ก.ย.", "ต.ค.", "พ.ย.", "ธ.ค."]


# ------------------------------------------------------------------ ส่วนที่ 1
def page_merge():
    st.header("1) รวมข้อมูลเป็น Data Rank")
    st.caption("ใส่ PPS, PLR, New Product, BarBom และ Data Rank เดือนก่อน  ระบบรวมด้วย Product Code "
               "และเติมยี่ห้อ/รุ่นที่เดือนก่อนไม่มีให้จากรายละเอียดสินค้า")
    files = st.file_uploader("PPS, PLR, New Product, BarBom (เลือกพร้อมกันทั้ง 4 ไฟล์)",
                             type=["xlsx", "csv"], accept_multiple_files=True, key="src4")
    found = {}
    for f in files or []:
        n = f.name.lower()
        for k, kw in (("bb", "bar"), ("newp", "new"), ("plr", "plr"), ("pps", "pps")):
            if kw in n:
                found.setdefault(k, f)
                break
    labels = dict(pps="PPS", plr="PLR", newp="New Product", bb="BarBom")
    if files:
        st.write("  ".join(("✅ " if k in found else "❌ ") + v for k, v in labels.items()))
        if len(found) < len(files):
            st.warning("มีไฟล์ที่ชื่อไม่มีคำว่า pps / plr / new / bar จึงไม่ถูกใช้ กรุณาตั้งชื่อไฟล์ให้มีคำเหล่านี้")
    pps, plr, newp, bb = (found.get(k) for k in ("pps", "plr", "newp", "bb"))
    c1, c2 = st.columns(2)
    with c1:
        prev = st.file_uploader("Data Rank เดือนก่อน (xlsx)", type=["xlsx"], key="prev")
    with c2:
        month = st.selectbox("เดือนของ Data Rank ที่จะสร้าง", [f"{i:02d}" for i in range(1, 13)], index=8)

    if st.button("▶ รวมข้อมูล", type="primary", disabled=not all([pps, plr, newp, bb, prev])):
        with tempfile.TemporaryDirectory() as d:
            with st.status("กำลังรวมข้อมูล...", expanded=True) as status:
                paths = {k: save_upload(v, d, f"{k}_{v.name}") for k, v in
                         dict(pps=pps, plr=plr, newp=newp, bb=bb, prev=prev).items()}
                out = os.path.join(d, f"DATA_RANK{month}.xlsx")
                ok, log = run(merge_data_rank.main, [
                    "--pps", paths["pps"], "--plr", paths["plr"], "--newproduct", paths["newp"],
                    "--barbom", paths["bb"], "--prev", paths["prev"], "--month", month,
                    "--out", out, "--no-recalc"])
                if ok:
                    with open(out, "rb") as f:
                        st.session_state["dr_bytes"] = f.read()
                    st.session_state["dr_name"] = f"DATA_RANK{month}.xlsx"
                    st.session_state["dr_month"] = month
                    try:
                        st.session_state["dr_fill"] = pd.read_excel(out, sheet_name="เติมยี่ห้อรุ่น (auto)")
                    except Exception:
                        st.session_state["dr_fill"] = None
                    status.update(label="รวมข้อมูลเสร็จแล้ว", state="complete")
                else:
                    status.update(label="มีข้อผิดพลาด", state="error")
                st.session_state["dr_log"] = log
                st.session_state["dr_ok"] = ok

    if "dr_log" in st.session_state:
        st.code(st.session_state["dr_log"], language="text")
    if st.session_state.get("dr_ok"):
        st.success("ได้ไฟล์ Data Rank แล้ว  ใช้ต่อในส่วนที่ 2 ได้ทันที หรือดาวน์โหลดเก็บไว้")
        st.download_button("⬇ ดาวน์โหลด " + st.session_state["dr_name"], st.session_state["dr_bytes"],
                           file_name=st.session_state["dr_name"],
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        fill = st.session_state.get("dr_fill")
        if fill is not None and len(fill):
            with st.expander(f"รายการที่เติมยี่ห้อ/รุ่นให้อัตโนมัติ ({len(fill)}) ช่วยตรวจ"):
                st.dataframe(fill, use_container_width=True)
        st.info("ไฟล์มีชีตสูตร VLOOKUP (Data Rank) และชีต 'Data Rank.. (ค่า)' ที่เป็นตัวเลขจริง "
                "ส่วนที่ 2 ใช้ชีตค่าอัตโนมัติ  ถ้ามีช่องยี่ห้อ/รุ่นที่ยังว่าง ให้เติมในชีต (ค่า) ก่อนจัด Rank")


# ------------------------------------------------------------------ ส่วนที่ 2
def page_rank():
    st.header("2) จัด Rank สินค้า A/B/C/D")
    st.caption("ใช้ Data Rank จากส่วนที่ 1 + เทมเพลต/ไฟล์ Rank เดือนก่อน + CSR เดือนใหม่  "
               "คำนวณยอดขายรายเดือน AV 6M เรียงลำดับ แยกสินค้าใหม่/รายการตัดออก และใส่ Rank")

    has1 = "dr_bytes" in st.session_state
    src = st.radio("Data Rank", ["ใช้ผลจากส่วนที่ 1", "อัปโหลดไฟล์ Data Rank"], horizontal=True,
                   index=0 if has1 else 1)
    if src.startswith("ใช้") and not has1:
        st.warning("ยังไม่มีผลจากส่วนที่ 1 ให้ไปทำส่วนที่ 1 ก่อน หรือเลือกอัปโหลดไฟล์")
    dr_up = None
    if src.startswith("อัปโหลด"):
        dr_up = st.file_uploader("Data Rank (xlsx)", type=["xlsx"], key="dr_up")
    else:
        st.write(f"ใช้ {st.session_state['dr_name']}")

    tpl = st.file_uploader("เทมเพลต Rank (ไฟล์ Rank เดือนก่อน ที่มีชีตยอดขายรายปี)", type=["xlsx"], key="tpl",
                           help="ใช้เป็นต้นแบบหน้าตา และเป็นแหล่งยอดขายย้อนหลัง")
    cache = st.file_uploader("ยอดขายที่เก็บไว้แล้ว (.json.gz) ไม่ใส่ก็ได้", type=["gz"], key="cache",
                             help="ไฟล์ที่ระบบให้ดาวน์โหลดไว้ครั้งก่อน ช่วยข้ามการอ่านชีตปีที่ใหญ่ (ประหยัดเวลา)")
    csr = st.file_uploader("CSR เดือนใหม่ (xlsx / csv) ใส่ได้หลายไฟล์", type=["xlsx", "csv"],
                           accept_multiple_files=True, key="csr")
    c1, c2 = st.columns(2)
    year = c1.number_input("ปี (ค.ศ.)", 2023, 2100, 2026)
    month = c2.selectbox("เดือนที่จัด Rank", list(range(1, 13)), index=8,
                         format_func=lambda m: f"{m:02d} ({MONTH_NAMES[m - 1]})")
    with st.expander("ลิสต์ Category ที่ตัดออก (แก้ได้ บรรทัดละ 1 ค่า)"):
        cuts = st.text_area("Category", "\n".join(build_rank.CUT_CATS), height=220, label_visibility="collapsed")

    have_dr = dr_up is not None or (has1 and src.startswith("ใช้"))
    if st.button("▶ จัด Rank", type="primary", disabled=not (have_dr and (tpl or cache))):
        with tempfile.TemporaryDirectory() as d:
            with st.status("กำลังคำนวณ (ใช้เวลาประมาณ 3-6 นาที)...", expanded=True) as status:
                if dr_up is not None:
                    dr_path = save_upload(dr_up, d, "data_rank.xlsx")
                else:
                    dr_path = os.path.join(d, "data_rank.xlsx")
                    with open(dr_path, "wb") as f:
                        f.write(st.session_state["dr_bytes"])
                args = ["--data", dr_path, "--year", str(int(year)), "--month", str(month)]
                if tpl is not None:
                    tpl_path = save_upload(tpl, d, "template.xlsx")
                    args += ["--template", tpl_path]
                cache_path = os.path.join(d, "sales.json.gz")
                if cache is not None:
                    save_upload(cache, d, "sales.json.gz")
                else:
                    args += ["--master", tpl_path]
                args += ["--sales-cache", cache_path]
                if csr:
                    args += ["--csr"] + [save_upload(c, d, f"csr{i}_{c.name}") for i, c in enumerate(csr)]
                cuts_path = os.path.join(d, "cuts.txt")
                with open(cuts_path, "w", encoding="utf-8") as f:
                    f.write("\n".join(x.strip() for x in cuts.splitlines() if x.strip()))
                out = os.path.join(d, f"Rank{month:02d}_{int(year)}.xlsx")
                args += ["--cut-cats", cuts_path, "--out", out, "--no-recalc"]
                ok, log = run(build_rank.main, args)
                st.session_state["rk_log"] = log
                st.session_state["rk_ok"] = ok
                if ok:
                    with open(out, "rb") as f:
                        st.session_state["rk_bytes"] = f.read()
                    st.session_state["rk_name"] = os.path.basename(out)
                    with open(cache_path, "rb") as f:
                        st.session_state["rk_cache"] = f.read()
                    status.update(label="จัด Rank เสร็จแล้ว", state="complete")
                else:
                    status.update(label="มีข้อผิดพลาด", state="error")

    if "rk_log" in st.session_state:
        log = st.session_state["rk_log"]
        m = re.search(r"A=([\d,]+) B=([\d,]+) C=([\d,]+).*?D=([\d,]+)", log)
        if st.session_state.get("rk_ok") and m:
            cols = st.columns(4)
            for col, name, v in zip(cols, "ABCD", m.groups()):
                col.metric(f"Rank {name}", v)
        st.code(log, language="text")
    if st.session_state.get("rk_ok"):
        st.success("เสร็จแล้ว  เปิดใน Excel เพื่อให้สูตรคำนวณ (AV, Rank, หน้าสรุป)")
        st.download_button("⬇ ดาวน์โหลด " + st.session_state["rk_name"], st.session_state["rk_bytes"],
                           file_name=st.session_state["rk_name"],
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        st.download_button("⬇ เก็บยอดขายล่าสุด (.json.gz) ไว้ใช้เดือนหน้า", st.session_state["rk_cache"],
                           file_name="sales.json.gz", mime="application/gzip")


# ------------------------------------------------------------------ หลัก
password_gate()
st.title("📊 Rank สินค้า TRC")
st.caption("ทำต้นเดือนทุกเดือน: ส่วนที่ 1 รวมข้อมูล แล้วต่อด้วยส่วนที่ 2 จัด Rank")
page_merge()
st.divider()
page_rank()
