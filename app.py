import streamlit as st
from docx import Document
import os
import calendar
import random
import pandas as pd
from docx2pdf import convert

st.set_page_config(page_title="COA Generator Pro+", layout="wide")

# ---------------------------
# FILES
# ---------------------------
HISTORY_FILE = "coa_history.csv"

# ---------------------------
# DOCX REPLACE
# ---------------------------
def replace_text(doc, replacements):
    for para in doc.paragraphs:
        for key, val in replacements.items():
            if f"{{{{{key}}}}}" in para.text:
                para.text = para.text.replace(f"{{{{{key}}}}}", val)

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for key, val in replacements.items():
                    if f"{{{{{key}}}}}" in cell.text:
                        cell.text = cell.text.replace(f"{{{{{key}}}}}", val)

# ---------------------------
# BEST BEFORE
# ---------------------------
def calculate_best_before(date_str):
    try:
        date_str = date_str.upper().strip()
        month, year = date_str.split()
        year = int(year)

        month_num = list(calendar.month_name).index(month.capitalize())
        new_month = month_num + 23
        new_year = year + (new_month - 1) // 12
        new_month = ((new_month - 1) % 12) + 1

        return f"{calendar.month_name[new_month].upper()} {new_year}"
    except:
        return ""

# ---------------------------
# CALCULATION
# ---------------------------
def generate_values(moisture):
    protein = round(random.uniform(2.45, 2.52), 2)
    ash = round(random.uniform(0.48, 0.54), 2)
    air = round(random.uniform(2.95, 3.05), 2)
    fat = round(random.uniform(0.48, 0.52), 2)
    gum = round(100 - moisture - (protein + ash + air + fat), 2)
    return gum, protein, ash, air, fat

# ---------------------------
# SAVE HISTORY
# ---------------------------
def save_history(data):
    df = pd.DataFrame([data])
    if os.path.exists(HISTORY_FILE):
        df.to_csv(HISTORY_FILE, mode='a', header=False, index=False)
    else:
        df.to_csv(HISTORY_FILE, index=False)

# ---------------------------
# LOAD HISTORY
# ---------------------------
def load_history():
    if os.path.exists(HISTORY_FILE):
        return pd.read_csv(HISTORY_FILE)
    return pd.DataFrame()

# ---------------------------
# UI HEADER
# ---------------------------
st.title("🏭 COA Generator Pro+")
st.caption("Industrial COA System with History & Preview")

tab1, tab2 = st.tabs(["🧾 Generate COA", "📊 History"])

# =========================================================
# TAB 1 - GENERATE
# =========================================================
with tab1:

    st.subheader("📦 Batch Info")
    c1, c2, c3 = st.columns(3)

    with c1:
        code = st.selectbox("Quality", [f"{i}-{i+500}" for i in range(500, 10001, 500)])

    with c2:
        date = st.text_input("Date (MARCH 2026)")

    with c3:
        batch = st.text_input("Batch No")

    st.subheader("⚗️ Lab Inputs")
    c4, c5, c6 = st.columns(3)

    with c4:
        moisture = st.number_input("Moisture", 0.0, 20.0, 10.0)

    with c5:
        ph = st.number_input("pH", 5.5, 7.0, 6.5)

    with c6:
        mesh = st.number_input("200 Mesh", 90.0, 100.0, 99.0)

    c7, c8 = st.columns(2)

    with c7:
        vis2 = st.number_input("Viscosity 2H", 5000, 6000, 5200)

    with c8:
        vis24 = st.number_input("Viscosity 24H", 5200, 6000, 5400)

    best_before = calculate_best_before(date)
    gum, protein, ash, air, fat = generate_values(moisture)

    st.subheader("📊 Live Values")
    m1, m2, m3, m4, m5 = st.columns(5)

    m1.metric("Gum", gum)
    m2.metric("Protein", protein)
    m3.metric("Ash", ash)
    m4.metric("Air", air)
    m5.metric("Fat", fat)

    st.info(f"Best Before: {best_before}")

    if st.button("🚀 Generate COA"):

        data = {
            "DATE": date,
            "BEST_BEFORE": best_before,
            "BATCH_NO": batch,
            "MOISTURE": f"{moisture}%",
            "PH": ph,
            "MESH_200": f"{mesh}%",
            "VISCOSITY_2H": vis2,
            "VISCOSITY_24H": vis24,
            "GUM_CONTENT": f"{gum}%",
            "PROTEIN": f"{protein}%",
            "ASH_CONTENT": f"{ash}%",
            "AIR": f"{air}%",
            "FAT": f"{fat}%"
        }

        template = f"COA {code}.docx"
        output = f"COA_{batch}.docx"

        if os.path.exists(template):

            doc = Document(template)
            replace_text(doc, data)
            doc.save(output)

            # Save history
            save_history({
                "batch": batch,
                "date": date,
                "quality": code,
                "gum": gum,
                "moisture": moisture
            })

            st.success("✅ COA Generated")

            with open(output, "rb") as f:
                st.download_button("📥 Download DOCX", f.read(), file_name=output)

            # PDF
            convert(output)
            pdf_file = output.replace(".docx", ".pdf")

            if os.path.exists(pdf_file):
                with open(pdf_file, "rb") as f:
                    st.download_button("📄 Download PDF", f.read(), file_name=pdf_file)

        else:
            st.error("Template not found")

# =========================================================
# TAB 2 - HISTORY
# =========================================================
with tab2:

    st.subheader("📊 COA History")

    df = load_history()

    if not df.empty:

        search = st.text_input("🔍 Search Batch")

        if search:
            df = df[df["batch"].astype(str).str.contains(search)]

        st.dataframe(df, use_container_width=True)

    else:
        st.info("No history found")
