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

def generate_pdf(quote_df, audience_df, client, contact, company, ref, notes, project_name):
    pdf = FPDF()
    pdf.add_page()
    pdf.set_auto_page_break(auto=True, margin=15)
    pdf.set_font("Arial", "B", 16)
    pdf.cell(0, 10, "PureSpectrum Multi-Option Quote", ln=True, align="C")
    pdf.ln(5)

    pdf.set_font("Arial", size=12)
    for label, val in [
        ("Client", client),
        ("Contact", contact),
        ("Company", company),
        ("Reference", ref),
        ("Project Name", project_name)
    ]:
        pdf.cell(50, 10, f"{label}:", ln=0)
        pdf.cell(0, 10, str(val), ln=1)

    if notes:
        pdf.ln(5)
        pdf.set_font("Arial", "B", 12)
        pdf.cell(0, 10, "Notes:", ln=True)
        pdf.set_font("Arial", size=12)
        pdf.multi_cell(0, 8, notes)

    if not audience_df.empty:
        pdf.ln(5)
        pdf.set_font("Arial", "B", 12)
        pdf.cell(0, 10, "Audience:", ln=True)
        pdf.set_font("Arial", size=12)
        for _, row in audience_df.iterrows():
            pdf.multi_cell(0, 8, f"- {row_
