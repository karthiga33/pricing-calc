"""Convert Pipeline-V3-Iceberg-Architecture.md to PDF - simple text approach"""
import re
from fpdf import FPDF

md_path = r"c:\Users\KarthigaS\Downloads\aws -account- invoices\aws -account- invoices\Pipeline-V3-Iceberg-Architecture.md"
pdf_path = r"c:\Users\KarthigaS\Downloads\aws -account- invoices\aws -account- invoices\Pipeline-V3-Iceberg-Architecture.pdf"

with open(md_path, "r", encoding="utf-8") as f:
    content = f.read()

def clean(text):
    replacements = {
        '\u2014': '-', '\u2013': '-', '\u2018': "'", '\u2019': "'",
        '\u201c': '"', '\u201d': '"', '\u2022': '*', '\u2192': '->',
        '\u2190': '<-', '\u2026': '...', '\u2265': '>=', '\u2264': '<=',
        '\u2260': '!=', '\u00a0': ' ', '\u200b': '', '\u2502': '|',
        '\u251c': '|', '\u2500': '-', '\u2514': '|', '\u2518': '+',
        '\u250c': '+', '\u2510': '+', '\u253c': '+',
    }
    for k, v in replacements.items():
        text = text.replace(k, v)
    return text.encode('latin-1', errors='ignore').decode('latin-1')

pdf = FPDF()
pdf.set_auto_page_break(auto=True, margin=15)
pdf.add_page()

lines = content.split("\n")
in_code_block = False

for line in lines:
    stripped = line.rstrip()
    cleaned = clean(stripped)

    if cleaned.strip().startswith("```"):
        in_code_block = not in_code_block
        continue

    if in_code_block:
        pdf.set_font("Courier", "", 8)
        pdf.set_text_color(50, 50, 50)
        try:
            pdf.multi_cell(0, 4, "  " + cleaned)
        except Exception:
            pass
        continue

    s = cleaned.strip()

    if s.startswith("# "):
        pdf.ln(6)
        pdf.set_font("Helvetica", "B", 16)
        pdf.set_text_color(26, 82, 118)
        try:
            pdf.multi_cell(0, 8, s[2:])
        except Exception:
            pass
        pdf.ln(2)
    elif s.startswith("## "):
        pdf.ln(4)
        pdf.set_font("Helvetica", "B", 13)
        pdf.set_text_color(44, 62, 80)
        try:
            pdf.multi_cell(0, 7, s[3:])
        except Exception:
            pass
        pdf.ln(2)
    elif s.startswith("### "):
        pdf.ln(3)
        pdf.set_font("Helvetica", "B", 11)
        pdf.set_text_color(52, 73, 94)
        try:
            pdf.multi_cell(0, 6, s[4:])
        except Exception:
            pass
        pdf.ln(1)
    elif s.startswith("#### "):
        pdf.ln(2)
        pdf.set_font("Helvetica", "BI", 10)
        pdf.set_text_color(52, 73, 94)
        try:
            pdf.multi_cell(0, 5, s[5:])
        except Exception:
            pass
    elif s == "---":
        pdf.ln(2)
        pdf.line(10, pdf.get_y(), 200, pdf.get_y())
        pdf.ln(2)
    elif s == "":
        pdf.ln(2)
    elif s.startswith("- ") or s.startswith("* "):
        pdf.set_font("Helvetica", "", 9)
        pdf.set_text_color(0, 0, 0)
        text = s[2:]
        text = re.sub(r"\*\*(.*?)\*\*", r"\1", text)
        text = re.sub(r"`(.*?)`", r"\1", text)
        try:
            pdf.multi_cell(0, 4.5, "   * " + text)
        except Exception:
            pass
    elif s.startswith("|") and "|" in s[1:]:
        # Table row
        pdf.set_font("Courier", "", 7)
        pdf.set_text_color(0, 0, 0)
        try:
            pdf.multi_cell(0, 3.5, s)
        except Exception:
            pass
    else:
        pdf.set_font("Helvetica", "", 9)
        pdf.set_text_color(0, 0, 0)
        text = re.sub(r"\*\*(.*?)\*\*", r"\1", s)
        text = re.sub(r"`(.*?)`", r"\1", text)
        text = re.sub(r"\[(.*?)\]\(.*?\)", r"\1", text)
        try:
            pdf.multi_cell(0, 4.5, text)
        except Exception:
            pass

pdf.output(pdf_path)
import os
print(f"PDF created: {pdf_path}")
print(f"Size: {os.path.getsize(pdf_path):,} bytes")
