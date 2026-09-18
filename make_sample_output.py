"""Generate a sample output Excel showing the new A1 (Attribute1) column"""
import pandas as pd
from io import BytesIO

# Sample rows based on the "Bills Adjustment Details" image
# Existing extracted columns + NEW "A1" column holding the Cr/Dr indicator
rows = [
    {"DOCUMENT_NUMBER": "3005010025770", "DOCUMENT_DATE": "25/06/2026", "INVOICE_AMOUNT": 31242.04, "AMOUNT": 31242.04, "A1": "Cr"},
    {"DOCUMENT_NUMBER": "JV-9.260626.20646", "DOCUMENT_DATE": "", "INVOICE_AMOUNT": 21005.09, "AMOUNT": 21005.09, "A1": "Dr"},
    {"DOCUMENT_NUMBER": "3005010025769", "DOCUMENT_DATE": "25/06/2026", "INVOICE_AMOUNT": 353996.18, "AMOUNT": 353996.18, "A1": "Cr"},
    {"DOCUMENT_NUMBER": "3005010025711", "DOCUMENT_DATE": "16/06/2026", "INVOICE_AMOUNT": 111393.98, "AMOUNT": 111393.98, "A1": "Cr"},
    {"DOCUMENT_NUMBER": "3005010025652", "DOCUMENT_DATE": "08/06/2026", "INVOICE_AMOUNT": 337967.20, "AMOUNT": 337967.20, "A1": "Cr"},
    {"DOCUMENT_NUMBER": "3005010025712", "DOCUMENT_DATE": "16/06/2026", "INVOICE_AMOUNT": 396081.11, "AMOUNT": 396081.11, "A1": "Cr"},
    {"DOCUMENT_NUMBER": "3005010025709", "DOCUMENT_DATE": "16/06/2026", "INVOICE_AMOUNT": 200984.22, "AMOUNT": 200984.22, "A1": "Cr"},
    {"DOCUMENT_NUMBER": "JV-9.310326.20907", "DOCUMENT_DATE": "", "INVOICE_AMOUNT": 697247.85, "AMOUNT": 697247.85, "A1": "Dr"},
]

df = pd.DataFrame(rows)

out_path = r"c:\Users\KarthigaS\Downloads\Pricing-calc-updated-code\SAMPLE-OUTPUT-with-A1.xlsx"
with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
    df.to_excel(writer, sheet_name="Remittance", index=False)

print(f"Sample output written to: {out_path}")
print()
print(df.to_string(index=False))
