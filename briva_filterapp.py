import streamlit as st
import pandas as pd
import re
import os
from io import BytesIO
import time
import tempfile
import pathlib

# ==== BACA PREFIX DARI corporate_code.xlsx (harus di folder yang sama) ====
prefix_file = "corporate_code.xlsx"
df_prefix = pd.read_excel(prefix_file)
df_prefix.columns = df_prefix.columns.str.strip().str.lower()
briva_prefixes = df_prefix["corporate_code"].astype(str).tolist()

# ==== fungsi ambil BRIVA ====
def ambil_briva(remark, prefixes):
    text = str(remark)
    text = re.sub(r"[^0-9]", "", text)
    for prefix in prefixes:
        match = re.search(prefix + r"\d{10}", text)  # prefix + 10 digit
        if match:
            return match.group(0)
    return None

# ==== cari kolom otomatis ====
def cari_kolom(df, keywords):
    for col in df.columns:
        norm = str(col).strip().lower()
        for key in keywords:
            if key in norm:
                return col
    return None

# ==== fungsi bersihkan nominal ====
def bersihkan_nominal(x):
    if pd.isna(x):
        return 0
    s = str(x).strip()
    s = s.replace(",", "")
    s = re.sub(r"\.00$", "", s)
    try:
        return int(s)
    except:
        return 0

# ==== STREAMLIT APP ====
st.set_page_config(page_title="BRIVA Converter", layout="wide")

st.title("💳 Pemisah Transaksi BRIVA")

# tampilkan nama Anda di layar
st.markdown("👩‍💻 Created by **Tri**@2025")

uploaded_files = st.file_uploader(
    "Upload file Excel rekening koran [bisa banyak]",
    type=["xlsx", "xls"],   # ✅ sekarang bisa xls & xlsx
    accept_multiple_files=True
)

if uploaded_files:
    rekap_match = []
    rekap_lain = []

    total_files = len(uploaded_files)
    progress = st.progress(0)
    status_text = st.empty()

    for i, f in enumerate(uploaded_files, start=1):
        file_name = os.path.splitext(f.name)[0]
        ext = pathlib.Path(f.name).suffix.lower()
        status_text.text(f"▶️ Proses {f.name} ({i}/{total_files}) ...")

        # ==== jika file .xls, convert ke .xlsx dulu ====
        if ext == ".xls":
            try:
                df_xls = pd.read_excel(f, engine="xlrd")
                tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
                with pd.ExcelWriter(tmp.name, engine="openpyxl") as writer:
                    df_xls.to_excel(writer, index=False, sheet_name="Sheet1")
                temp_file = tmp.name
                st.info(f"ℹ️ File {f.name} otomatis dikonversi ke XLSX sebelum diproses.")
            except Exception as e:
                st.error(f"❌ Gagal konversi {f.name}: {e}")
                progress.progress(i / total_files)
                continue
        else:
            temp_file = f

        # ==== cari header otomatis (0–15) ====
        df = None
        for h in range(0, 16):
            try:
                temp = pd.read_excel(temp_file, header=h)
                cols = [c.lower() for c in temp.columns.astype(str)]
                if any("date" in c for c in cols) and any("remark" in c for c in cols):
                    df = temp
                    st.write(f"✅ Header ditemukan di baris {h+1} untuk file {f.name}")
                    break
            except Exception:
                continue

        if df is None:
            st.warning(f"⚠️ Tidak ketemu header di {f.name}, dilewati")
            progress.progress(i / total_files)
            continue

        # cari nama kolom
        col_date   = cari_kolom(df, ["date", "tanggal"])
        col_time   = cari_kolom(df, ["time", "jam"])
        col_remark = cari_kolom(df, ["remark", "uraian", "deskripsi", "keterangan"])
        col_debet  = cari_kolom(df, ["debet", "debit"])
        col_credit = cari_kolom(df, ["credit", "kredit"])

        if None in [col_date, col_time, col_remark, col_debet, col_credit]:
            st.warning(f"⚠️ Kolom tidak lengkap di {f.name}, dilewati")
            progress.progress(i / total_files)
            continue

        # ambil BRIVA
        df["BRIVA"] = df[col_remark].apply(lambda x: ambil_briva(x, briva_prefixes))

        # konversi nominal
        df[col_debet]  = df[col_debet].apply(bersihkan_nominal)
        df[col_credit] = df[col_credit].apply(bersihkan_nominal)

        # kolom TIPE
        def tentukan_tipe(row):
            if row[col_credit] > 0:
                return "MASUK"
            elif row[col_debet] > 0:
                return "KELUAR"
            return ""
        df["TIPE"] = df.apply(tentukan_tipe, axis=1)

        # kolom ASAL_FILE
        df["ASAL_FILE"] = file_name

        # BRIVA_MATCH
        df_match = df[df["BRIVA"].notna() & df["BRIVA"].str[:5].isin(briva_prefixes)].copy()
        df_match[col_remark] = df_match["BRIVA"]

        # LAIN-LAIN
        df_lain = df.drop(df_match.index).copy()

        # pilih kolom output
        kolom_output = [col_date, col_time, col_remark, col_debet, col_credit, "TIPE", "ASAL_FILE"]
        df_match = df_match[kolom_output]
        df_lain  = df_lain[kolom_output]

        if not df_match.empty:
            rekap_match.append(df_match)
        if not df_lain.empty:
            rekap_lain.append(df_lain)

        progress.progress(i / total_files)
        time.sleep(0.2)

    # gabung hasil semua file
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        if rekap_match:
            df_rekap_match = pd.concat(rekap_match, ignore_index=True)
            df_rekap_match.to_excel(writer, index=False, sheet_name="BRIVA_MATCH")
        if rekap_lain:
            df_rekap_lain = pd.concat(rekap_lain, ignore_index=True)
            df_rekap_lain.to_excel(writer, index=False, sheet_name="LAIN-LAIN")
        df_prefix.to_excel(writer, index=False, sheet_name="PREFIX_LIST")

    st.success("✅ Semua file selesai diproses.")
    st.download_button(
        label="⬇️ Download Rekap BRIVA",
        data=buffer.getvalue(),
        file_name="rekap_briva.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )



