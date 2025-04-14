
import streamlit as st
import pandas as pd
import openpyxl
from openpyxl import Workbook
from openpyxl.drawing.image import Image as ExcelImage
from PIL import Image
import io
from fpdf import FPDF

# Load CPI table
cpi_df = pd.read_excel("Refined_CPI_Lookup_Table.xlsx", index_col=0)

# Programming logic
prog_rates = {
    "Simple": (350, 50),
    "Standard": (500, 75),
    "Complex": (800, 100)
}

def calculate_quote(loi, ir, n, prog_type, dp_type):
    cpi = round(round(cpi_df.loc[loi, ir] / 0.05) * 0.05, 2)
    sample_cost = cpi * n
    base, per_min = prog_rates[prog_type]
    prog_cost = base + per_min * loi
    dp_cost = 180 if dp_type == "SPSS" else loi * 50
    total = sample_cost + prog_cost + dp_cost
    return cpi, sample_cost, prog_cost, dp_cost, total

# Streamlit UI
st.title("PureSpectrum Quote Generator")

with st.form("quote_form"):
    client = st.text_input("Client Name")
    job = st.text_input("Job Name")
    ref = st.text_input("Reference Number")
    lois = st.multiselect("Select LOI options (mins)", [5, 10, 15, 20, 25, 30, 40, 50, 60], default=[10, 15, 20])
    ir = st.slider("Incidence Rate (IR %)", 1, 99, 50)
    sample_sizes = st.multiselect("Select sample sizes", [250, 500, 1000, 2000], default=[500, 1000])
    prog = st.selectbox("Programming Type", ["Simple", "Standard", "Complex"])
    dp = st.selectbox("DP Type", ["SPSS", "Tables"])
    notes = st.text_area("Quoting Notes (optional)")
    submit = st.form_submit_button("Generate Quote Table")

if submit:
    table_data = []
    for loi in lois:
        row = {"LOI": f"{loi} min"}
        for n in sample_sizes:
            cpi, sample, prog_cost, dp_cost, total = calculate_quote(loi, ir, n, prog, dp)
            row[f"n={n}"] = f"${total:,.0f} (${cpi:.2f} CPI)"
        table_data.append(row)

    df = pd.DataFrame(table_data)
    st.subheader(f"Quote Summary (IR {ir}%)")
    st.dataframe(df)

    # Export as Excel
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name="Quote Table", index=False)
        workbook = writer.book
        worksheet = writer.sheets["Quote Table"]
        worksheet.cell(row=1, column=len(df.columns)+2, value="Client Name:").offset(0, 1).value = client
        worksheet.cell(row=2, column=len(df.columns)+2, value="Job Name:").offset(0, 1).value = job
        worksheet.cell(row=3, column=len(df.columns)+2, value="Reference No:").offset(0, 1).value = ref
        worksheet.cell(row=4, column=len(df.columns)+2, value="Notes:").offset(0, 1).value = notes
    output.seek(0)

    st.download_button("Download Excel Quote Table", data=output, file_name=f"Quote_{ref}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
