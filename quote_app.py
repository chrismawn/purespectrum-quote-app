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

def generate_excel(client, job, ref, loi, ir, n, prog, dp, cpi, sample, prog_cost, dp_cost, total):
    wb = Workbook()
    ws = wb.active
    ws.title = "Quote"

    fields = [
        ("Client Name", client),
        ("Job Name", job),
        ("Reference Number", ref),
        ("Length of Interview (LOI)", loi),
        ("Incidence Rate (IR)", ir),
        ("Sample Size", n),
        ("Programming Type", prog),
        ("DP Type", dp),
        ("CPI", cpi),
        ("Sample Cost", sample),
        ("Programming Cost", prog_cost),
        ("DP Cost", dp_cost),
        ("Total Quote", total)
    ]

    for i, (label, val) in enumerate(fields, start=1):
        ws[f"A{i}"] = label
        ws[f"B{i}"] = val

    try:
        logo = ExcelImage("purespectrum_logo.png")
        logo.width, logo.height = 200, 100
        ws.add_image(logo, "D1")
    except:
        pass

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf

def generate_pdf(client, job, ref, loi, ir, n, prog, dp, cpi, sample, prog_cost, dp_cost, total):
    pdf = FPDF()
    pdf.add_page()
    pdf.set_auto_page_break(auto=True, margin=15)
    pdf.set_font("Arial", "B", 16)
    pdf.cell(0, 10, "PureSpectrum Project Quote Summary", ln=True, align="C")
    pdf.ln(10)
    pdf.set_font("Arial", size=12)

    info = {
        "Client Name": client,
        "Job Name": job,
        "Reference Number": ref,
        "Length of Interview (LOI)": f"{loi} minutes",
        "Incidence Rate (IR)": f"{ir}%",
        "Sample Size": n,
        "Programming Type": prog,
        "DP Type": dp,
        "CPI": f"${cpi:.2f}",
        "Sample Cost": f"${sample:,.2f}",
        "Programming Cost": f"${prog_cost:,.2f}",
        "DP Cost": f"${dp_cost:,.2f}",
        "Total Quote": f"${total:,.2f}"
    }

    for label, value in info.items():
        pdf.cell(70, 10, f"{label}:", ln=0)
        pdf.cell(0, 10, str(value), ln=1)

    pdf.ln(10)
    pdf.set_font("Arial", "I", 10)
    pdf.cell(0, 10, "Prepared by PureSpectrum - All quotes valid for 30 days", ln=True, align="C")

    pdf_bytes = pdf.output(dest='S').encode('latin1')
    return io.BytesIO(pdf_bytes)

st.title("PureSpectrum Quote Generator")

with st.form("quote_form"):
    client = st.text_input("Client Name")
    job = st.text_input("Job Name")
    ref = st.text_input("Reference Number")
    loi = st.slider("Length of Interview (LOI)", 1, 60, 10)
    ir = st.slider("Incidence Rate (IR %)", 1, 99, 50)
    n = st.number_input("Sample Size (n=)", min_value=1, value=1000)
    prog = st.selectbox("Programming Type", ["Simple", "Standard", "Complex"])
    dp = st.selectbox("DP Type", ["SPSS", "Tables"])
    submit = st.form_submit_button("Generate Quote")

if submit:
    cpi, sample, prog_cost, dp_cost, total = calculate_quote(loi, ir, n, prog, dp)
    excel_file = generate_excel(client, job, ref, loi, ir, n, prog, dp, cpi, sample, prog_cost, dp_cost, total)
    pdf_file = generate_pdf(client, job, ref, loi, ir, n, prog, dp, cpi, sample, prog_cost, dp_cost, total)

    st.success(f"Quote generated successfully! Total: ${total:,.2f}")
    st.download_button("Download Excel Quote", data=excel_file, file_name=f"Quote_{ref}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    st.download_button("Download PDF Quote", data=pdf_file, file_name=f"Quote_{ref}.pdf", mime="application/pdf")
