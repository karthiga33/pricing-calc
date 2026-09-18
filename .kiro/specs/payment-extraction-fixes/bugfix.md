# Bugfix Requirements Document

## Introduction

The payment advice PDF extraction system (`lambda_full.py`) has data quality bugs in two areas: (1) incorrect invoice row capture — summary rows, balance-forward lines, and ERS reference rows are being captured as valid invoice transactions, and (2) customer name prediction — the system sometimes picks address fragments or the receiver name instead of the actual payer company name. These bugs result in incorrect data being written to DynamoDB and output Excel files.

## Bug Analysis

### Current Behavior (Defect)

1.1 WHEN a PDF contains "Balance carried forward" or "Balance brought forward" summary lines with amounts THEN the system captures these as valid invoice transaction rows in the output

1.2 WHEN a PDF contains page subtotals, document totals, or "ERS Forward" reference lines (e.g., "ERS Forward 301001034722602...2026 1,23,321.94 12,67,000.00") THEN the system captures these summary rows as if they were actual invoice line items

1.3 WHEN a document contains invoice references with "ERS-" prefix (e.g., "ERS-302822000210", "ERS-3028220002101", "ERS-3028220002106") THEN the system captures the full ERS-prefixed string as the DOCUMENT_NUMBER without stripping the prefix

1.4 WHEN an ERS reference row appears adjacent to a real invoice row (e.g., invoice 3010010347216) THEN the system captures the ERS reference row's data instead of or in addition to the actual invoice data, producing incorrect INVOICE_AMOUNT values for that invoice

1.5 WHEN the document's payer company name is not clearly in a letterhead or labeled field THEN the system picks address fragments (GAT, PLOT, road names, city names, PIN codes) or the receiver name (Tube Investments of India Ltd / TIDC India) as the CUSTOMER_NAME

### Expected Behavior (Correct)

2.1 WHEN a PDF contains "Balance carried forward", "Balance brought forward", or similar balance summary lines THEN the system SHALL exclude these rows from the extracted invoice data and not write them to DynamoDB or the output Excel

2.2 WHEN a PDF contains page subtotals, document totals, "ERS Forward" reference lines, or any row that is clearly a summary/aggregation rather than an individual invoice THEN the system SHALL filter out these rows during post-processing before writing to output

2.3 WHEN a document contains invoice references with "ERS-" prefix THEN the system SHALL strip the "ERS-" prefix to extract the underlying numeric document number, ensuring the DOCUMENT_NUMBER field contains only the clean numeric invoice number

2.4 WHEN an ERS reference row appears adjacent to a real invoice row THEN the system SHALL correctly associate the amounts with the actual invoice row and discard the ERS reference row as a non-invoice entry

2.5 WHEN extracting CUSTOMER_NAME THEN the system SHALL identify only the payer company name (not the receiver), SHALL never use address fragments (street names, plot numbers, city names, PIN codes, building names) as company names, and SHALL return "Unknown" if no confident match is found

2.6 WHEN Claude returns the extracted JSON output THEN the system SHALL apply a rule-based post-processing validation and correction step on Claude's JSON BEFORE pushing results to DynamoDB and Excel. The flow SHALL be: Claude extracts → validation corrects errors → only corrected data is written to output. This validation SHALL:

  **Row filtering rules (remove the row entirely from Claude's output):**
  - Remove rows where DOCUMENT_NUMBER contains summary keywords: "Balance", "Forward", "Total", "Subtotal", "Grand Total", "Net Amount", "carried forward", "brought forward"
  - Remove rows where DOCUMENT_NUMBER contains "ERS Forward" or is clearly a page/section summary reference (not a real invoice)
  - Remove rows where DOCUMENT_NUMBER is blank/null AND INVOICE_AMOUNT is blank/null (empty rows)

  **Document number correction rules (fix the value in Claude's output):**
  - IF DOCUMENT_NUMBER starts with "ERS-" prefix THEN strip the prefix and keep only the numeric portion (e.g., "ERS-3028220002101" becomes "3028220002101")
  - IF DOCUMENT_NUMBER contains pipe-separated segments (e.g., "S|3002010005572|261300037731") THEN extract the 13-digit segment starting with 3, 4, or 7
  - IF DOCUMENT_NUMBER has non-numeric characters other than "ERS-" prefix THEN log a warning but keep the value as-is

  **Customer name correction rules (fix the value in Claude's output):**
  - IF CUSTOMER_NAME matches any known receiver name pattern (case-insensitive): "Tube Investments", "TIDC India", "TII", "M/S TUBE INVESTMENTS" THEN replace with "Unknown"
  - IF CUSTOMER_NAME contains address keywords (case-insensitive): "GAT", "PLOT", "ROAD", "HIGHWAY", "LANE", "STREET", "SECTOR", "BLOCK", "FLOOR", "BUILDING", "DIST", "TALUKA" THEN replace with "Unknown"
  - IF CUSTOMER_NAME is a PIN code pattern (6 digits) or phone number pattern THEN replace with "Unknown"
  - IF CUSTOMER_NAME is empty or null THEN set to "Unknown"

### Unchanged Behavior (Regression Prevention)

3.1 WHEN a PDF contains valid invoice rows with 13-digit document numbers starting with 3, 4, or 7 THEN the system SHALL CONTINUE TO extract these correctly with proper DOCUMENT_NUMBER, DOCUMENT_DATE, INVOICE_AMOUNT, TDS_AMOUNT, and AMOUNT values

3.2 WHEN a document has line-wrapped invoice numbers (e.g., 12 digits on one line, 1-3 digits on the next) THEN the system SHALL CONTINUE TO concatenate them into complete 13-digit invoice numbers

3.3 WHEN a document contains multiple vouchers/payments THEN the system SHALL CONTINUE TO extract each voucher separately with its own UTR_REFERENCE_NUMBER, PAYMENT_DATE, and PAYMENT_AMOUNT

3.4 WHEN a document has TDS entries (either as separate rows with "TDS" suffix or inline) THEN the system SHALL CONTINUE TO correctly merge TDS amounts with their corresponding invoice rows

3.5 WHEN a PDF has a clearly labeled company letterhead or "For M/s." pattern THEN the system SHALL CONTINUE TO correctly extract the payer company name from these standard locations

3.6 WHEN processing Excel, CSV, HTML, Word, or image files THEN the system SHALL CONTINUE TO extract payment data correctly from these formats using the existing processing pipelines
