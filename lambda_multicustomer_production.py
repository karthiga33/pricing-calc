import json
import boto3
import os
import time
import re
import urllib.parse
import pandas as pd
import logging
import subprocess
import tempfile
from io import BytesIO
from collections import OrderedDict
from decimal import Decimal
from datetime import datetime

OUTPUT_BUCKET = "production-ti"
OUTPUT_PREFIX = "multi-output/"
MODEL_ID = "global.anthropic.claude-sonnet-4-6"

s3_client       = boto3.client("s3")
textract        = boto3.client("textract")
bedrock_runtime = boto3.client("bedrock-runtime", region_name="us-east-1")
dynamodb        = boto3.resource("dynamodb")

logger = logging.getLogger()
logger.setLevel(logging.INFO)

REJECTION_TABLE_NAME = "prod_rejected_files"
DEDUP_TABLE_NAME = "prod_processed_file_etags"

def _log_rejection(input_key, etag, reason):
    """Log rejected file to DynamoDB for display in application."""
    try:
        table = dynamodb.Table(REJECTION_TABLE_NAME)
        file_name = os.path.basename(input_key)
        table.put_item(Item={
            "file_name": file_name,
            "rejected_at": datetime.utcnow().isoformat(),
            "input_key": input_key,
            "etag": etag or "",
            "reason": reason,
            "source": "multi_customer",
        })
        logger.info("Logged rejection: %s — %s", file_name, reason)
    except Exception as e:
        logger.warning("Failed to log rejection for %s: %s", input_key, str(e))


def _move_to_rejected_folder(input_bucket, input_key):
    """Move a rejected file from multi-input/ to rejected/ folder in S3."""
    try:
        file_name = os.path.basename(input_key)
        rejected_key = f"rejected/{file_name}"
        # Copy to rejected/
        s3_client.copy_object(
            Bucket=input_bucket,
            CopySource={"Bucket": input_bucket, "Key": input_key},
            Key=rejected_key,
        )
        # Delete from original location
        s3_client.delete_object(Bucket=input_bucket, Key=input_key)
        logger.info("Moved rejected file: s3://%s/%s → s3://%s/%s", input_bucket, input_key, input_bucket, rejected_key)
    except Exception as e:
        logger.warning("Failed to move rejected file %s to rejected/: %s", input_key, str(e))


def _is_etag_processed(etag):
    """Check if this ETag was already processed."""
    if not etag:
        return False
    try:
        table = dynamodb.Table(DEDUP_TABLE_NAME)
        resp = table.get_item(Key={"etag": etag})
        return "Item" in resp
    except Exception as e:
        logger.warning("ETag dedup check failed: %s — proceeding with processing", str(e))
        return False


def _store_etag(etag, input_key):
    """Store ETag after successful processing."""
    if not etag:
        return
    try:
        table = dynamodb.Table(DEDUP_TABLE_NAME)
        table.put_item(Item={
            "etag": etag,
            "input_key": input_key,
            "processed_at": datetime.utcnow().isoformat(),
        })
        logger.info("Stored ETag %s in dedup table", etag)
    except Exception as e:
        logger.warning("Failed to store ETag %s: %s", etag, str(e))


# ── DATE CONVERSION ───────────────────────────────────────────────────────────
def convert_date(raw_date):
    if not raw_date or str(raw_date).strip() == "":
        return ""
    raw_date = str(raw_date).strip()
    formats = [
        "%d-%b-%y", "%d-%b-%Y", "%d-%m-%Y", "%Y-%m-%d",
        "%d/%m/%Y", "%m/%d/%Y", "%d%m%Y", "%Y%m%d",
        "%Y-%m-%d %H:%M:%S", "%Y-%m-%dT%H:%M:%S", "%Y-%m-%dT%H:%M:%SZ",
    ]
    for fmt in formats:
        try:
            return datetime.strptime(raw_date, fmt).strftime("%d/%m/%Y")
        except ValueError:
            continue
    logger.warning("Could not parse date: '%s' — storing as-is", raw_date)
    return raw_date


# ── READ EMAIL METADATA ───────────────────────────────────────────────────────
def get_email_metadata(bucket_name, input_key):
    att_name = os.path.basename(input_key)
    meta_key = f"metadata/{att_name}.json"
    try:
        response = s3_client.get_object(Bucket=bucket_name, Key=meta_key)
        metadata = json.loads(response["Body"].read().decode("utf-8"))
        logger.info("✅ Metadata found for %s — MAIL_ID=%s, MAIL_RECEIVED_DATE=%s",
                    att_name, metadata.get("MAIL_ID", ""), metadata.get("MAIL_RECEIVED_DATE", ""))
        return metadata
    except Exception:
        logger.warning("⚠️ No metadata found for %s", att_name)
        return {}


# ── MAIN HANDLER ─────────────────────────────────────────────────────────────
def lambda_handler(event, context):
    logger.info("Received event: %s", json.dumps(event))
    records = event.get("Records", [])

    if records:
        processed_files = []
        for record in records:
            try:
                if record.get("eventSource") == "aws:sqs":
                    body = json.loads(record["body"])
                    s3_records = body.get("Records", [])
                else:
                    s3_records = [record] if record.get("eventSource") == "aws:s3" else []

                for s3_record in s3_records:
                    if s3_record.get("eventSource") != "aws:s3":
                        continue
                    input_bucket = s3_record["s3"]["bucket"]["name"]
                    input_key    = urllib.parse.unquote_plus(s3_record["s3"]["object"]["key"])
                    etag         = s3_record["s3"]["object"].get("eTag", "")
                    logger.info("S3 trigger via SQS: s3://%s/%s (ETag: %s)", input_bucket, input_key, etag)

                    # ETag dedup check
                    if _is_etag_processed(etag):
                        logger.info("⏭️ ETag %s already processed — skipping s3://%s/%s", etag, input_bucket, input_key)
                        _log_rejection(input_key, etag, "Duplicate file - already processed (same ETag)")
                        continue

                    result = _process_single_file(input_bucket, input_key)
                    if result:
                        _store_etag(etag, input_key)
                        processed_files.append(result)
            except Exception as e:
                logger.exception("Error processing record: %s", str(e))
                raise

        return {
            "statusCode": 200,
            "body": json.dumps({"message": "Processing completed", "files": processed_files}),
        }

    input_bucket = event.get("input_bucket")
    input_key    = event.get("input_key")
    if input_bucket and input_key:
        logger.info("Manual trigger: s3://%s/%s", input_bucket, input_key)
        try:
            result = _process_single_file(input_bucket, input_key)
            return {
                "statusCode": 200,
                "body": json.dumps({"message": "Processing completed", "files": [result] if result else []}),
            }
        except Exception as e:
            logger.exception("Error processing file: %s", str(e))
            return {"statusCode": 500, "body": json.dumps(str(e))}

    return {
        "statusCode": 400,
        "body": json.dumps("Invalid event: no S3 Records and no input_bucket/input_key found"),
    }


def _process_single_file(input_bucket, input_key):
    file_ext  = input_key.lower()
    supported = (".pdf", ".jpg", ".jpeg", ".png", ".xlsx", ".xls", ".csv", ".txt", ".html", ".doc", ".docx")
    if not any(file_ext.endswith(ext) for ext in supported):
        logger.info("Skipping unsupported file type: %s", input_key)
        return None

    email_metadata     = get_email_metadata(OUTPUT_BUCKET, input_key)
    mail_id            = email_metadata.get("from", "")
    mail_received_date = convert_date(email_metadata.get("MAIL_RECEIVED_DATE", ""))
    logger.info("from: %s  |  MAIL_RECEIVED_DATE: %s", mail_id, mail_received_date)

    base_name = os.path.splitext(os.path.basename(input_key))[0]
    parts     = input_key.split("/")
    if len(parts) >= 3:
        subject_folder = parts[-2]
        output_key  = f"{OUTPUT_PREFIX}{subject_folder}/{base_name}.xlsx"
        table_name  = re.sub(r"[^a-zA-Z0-9_]", "_", f"{subject_folder}_{base_name}")[:255]
    else:
        output_key  = f"{OUTPUT_PREFIX}{base_name}.xlsx"
        table_name  = re.sub(r"[^a-zA-Z0-9_]", "_", base_name)[:255]

    if table_name and table_name[0].isdigit():
        table_name = "t_" + table_name

    logger.info("Processing  : s3://%s/%s", input_bucket, input_key)
    logger.info("Output Excel: s3://%s/%s", OUTPUT_BUCKET, output_key)
    logger.info("DynamoDB    : %s", table_name)

    excel_bytes, merged_rows = process_file_from_s3(input_bucket, input_key, mail_id, mail_received_date)

    # ── Check if extraction produced meaningful data ──
    has_meaningful_data = any(
        row.get("DOCUMENT_NUMBER") or row.get("INVOICE_AMOUNT") or row.get("AMOUNT")
        or row.get("TRX_NUMBER") or row.get("OUTSTANDING_AMT") or row.get("APPLIED_AMT")
        for row in merged_rows
    )
    if not has_meaningful_data:
        logger.info("⏭️ No meaningful extraction from %s — skipping output", input_key)
        _log_rejection(input_key, "", "No payment/invoice data extracted — file does not contain relevant content")
        _move_to_rejected_folder(input_bucket, input_key)
        return None

    for row in merged_rows:
        row["MAIL_ID"]            = mail_id
        row["MAIL_RECEIVED_DATE"] = mail_received_date

    s3_client.put_object(
        Bucket=OUTPUT_BUCKET, Key=output_key, Body=excel_bytes,
        ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
    logger.info("Saved Excel to S3: s3://%s/%s", OUTPUT_BUCKET, output_key)

    write_to_dynamodb(table_name, merged_rows)
    logger.info("Written %d rows to DynamoDB table: %s", len(merged_rows), table_name)

    return {
        "input":               f"s3://{input_bucket}/{input_key}",
        "output":              f"s3://{OUTPUT_BUCKET}/{output_key}",
        "table":               table_name,
        "row_count":           len(merged_rows),
        "mail_id":             mail_id,
        "mail_received_date":  mail_received_date,
    }


def process_file_from_s3(input_bucket, input_key, mail_id="", mail_received_date=""):
    file_ext = input_key.lower()
    if file_ext.endswith(".xlsx") or file_ext.endswith(".xls"):
        logger.info("Processing as Excel ...")
        return process_excel_from_s3(input_bucket, input_key, mail_id, mail_received_date)
    if file_ext.endswith(".csv") or file_ext.endswith(".txt"):
        logger.info("Processing as CSV/TXT ...")
        return process_text_from_s3(input_bucket, input_key, mail_id, mail_received_date)
    # PDF / image / Word — use Textract
    return process_image_or_pdf_from_s3(input_bucket, input_key, mail_id, mail_received_date)


def process_excel_from_s3(input_bucket, input_key, mail_id="", mail_received_date=""):
    import io
    response    = s3_client.get_object(Bucket=input_bucket, Key=input_key)
    excel_bytes = response["Body"].read()
    try:
        df_dict = pd.read_excel(io.BytesIO(excel_bytes), sheet_name=None)
    except Exception as e:
        logger.error("Failed to read Excel: %s", str(e))
        raise

    # ── Only process first sheet; skip files with multiple sheets ──
    sheet_names = list(df_dict.keys())
    if len(sheet_names) > 1:
        logger.info("⏭️ Excel file has %d sheets — processing only first sheet: '%s'", len(sheet_names), sheet_names[0])
    first_sheet_name = sheet_names[0]
    first_sheet_df = df_dict[first_sheet_name]

    text_content = "=== EXCEL FILE CONTENT ===\n\n"
    text_content += f"\n=== SHEET: {first_sheet_name} ===\n"
    text_content += f"Columns: {', '.join(first_sheet_df.columns.astype(str))}\n\n"
    text_content += "Table Data:\n"
    for idx, row in first_sheet_df.iterrows():
        row_text = " | ".join([f"{col}: {val}" for col, val in zip(first_sheet_df.columns, row)])
        text_content += f"Row {idx + 1}: {row_text}\n"
    text_content += "\n"

    logger.info("Excel -> text (first 1000 chars): %s", text_content[:1000])
    extracted_data, claude_rows = _call_bedrock(text_content, source_type="EXCEL")
    return _build_output(claude_rows, extracted_data, default_source_type="EXCEL",
                         mail_id=mail_id, mail_received_date=mail_received_date)


def process_text_from_s3(input_bucket, input_key, mail_id="", mail_received_date=""):
    import io
    file_ext = input_key.lower()
    is_csv   = file_ext.endswith(".csv")
    source_type = "CSV" if is_csv else "TXT"

    response  = s3_client.get_object(Bucket=input_bucket, Key=input_key)
    raw_bytes = response["Body"].read()
    try:
        raw_text = raw_bytes.decode("utf-8")
    except UnicodeDecodeError:
        raw_text = raw_bytes.decode("latin-1")

    if is_csv:
        try:
            df = pd.read_csv(io.StringIO(raw_text))
            text_content = "=== CSV FILE CONTENT ===\n\n"
            text_content += f"Columns: {', '.join(df.columns.astype(str))}\n\n"
            text_content += "Table Data:\n"
            for idx, row in df.iterrows():
                row_text = " | ".join([f"{col}: {val}" for col, val in zip(df.columns, row)])
                text_content += f"Row {idx + 1}: {row_text}\n"
        except Exception as e:
            logger.warning("pandas CSV parse failed (%s), sending raw text", str(e))
            text_content = f"=== CSV FILE CONTENT ===\n\n{raw_text}"
    else:
        text_content = f"=== TEXT FILE CONTENT ===\n\n{raw_text}"

    logger.info("%s -> text (first 1000 chars): %s", source_type, text_content[:1000])
    extracted_data, claude_rows = _call_bedrock(text_content, source_type=source_type)
    return _build_output(claude_rows, extracted_data, default_source_type=source_type,
                         mail_id=mail_id, mail_received_date=mail_received_date)


def process_image_or_pdf_from_s3(input_bucket, input_key, mail_id="", mail_received_date=""):
    file_ext = input_key.lower()
    is_image = file_ext.endswith(".jpg") or file_ext.endswith(".jpeg") or file_ext.endswith(".png")

    if is_image:
        logger.info("Processing image with Textract ...")
        response = textract.analyze_document(
            Document={"S3Object": {"Bucket": input_bucket, "Name": input_key}},
            FeatureTypes=["FORMS", "TABLES", "SIGNATURES"],
        )
        blocks = response.get("Blocks", [])
        logger.info("Textract completed. Blocks: %d", len(blocks))
    else:
        logger.info("Starting Textract async job for PDF ...")
        job    = textract.start_document_analysis(
            DocumentLocation={"S3Object": {"Bucket": input_bucket, "Name": input_key}},
            FeatureTypes=["FORMS", "TABLES", "SIGNATURES"],
        )
        job_id = job["JobId"]
        logger.info("Textract job started: %s", job_id)

        blocks     = []
        next_token = None
        while True:
            kwargs = {"JobId": job_id}
            if next_token:
                kwargs["NextToken"] = next_token
            status     = textract.get_document_analysis(**kwargs)
            job_status = status["JobStatus"]
            if job_status == "SUCCEEDED":
                blocks.extend(status.get("Blocks", []))
                next_token = status.get("NextToken")
                while next_token:
                    page = textract.get_document_analysis(JobId=job_id, NextToken=next_token)
                    blocks.extend(page.get("Blocks", []))
                    next_token = page.get("NextToken")
                break
            if job_status == "FAILED":
                raise RuntimeError(f"Textract failed: {status.get('StatusMessage', 'Unknown')}")
            logger.info("Textract status: %s ... waiting", job_status)
            time.sleep(5)
        logger.info("Textract SUCCEEDED. Total blocks: %d", len(blocks))

    text_content = textract_blocks_to_text(blocks)
    logger.info("Textract extracted text_content:\n%s", text_content)
    extracted_data, claude_rows = _call_bedrock(text_content, source_type="PDF")
    return _build_output(claude_rows, extracted_data, default_source_type="PDF",
                         mail_id=mail_id, mail_received_date=mail_received_date)


def textract_blocks_to_text(blocks):
    lines, kv_pairs, tables = [], [], []
    block_map = {b["Id"]: b for b in blocks}
    for block in blocks:
        if block["BlockType"] == "LINE" and "Text" in block:
            lines.append(block["Text"])
        if block["BlockType"] == "KEY_VALUE_SET" and "KEY" in block.get("EntityTypes", []):
            key_text = " ".join([
                block_map.get(cid, {}).get("Text", "")
                for rel in block.get("Relationships", []) if rel["Type"] == "CHILD"
                for cid in rel["Ids"]
            ]).strip()
            value_text = ""
            for rel in block.get("Relationships", []):
                if rel["Type"] == "VALUE":
                    for vid in rel["Ids"]:
                        val = block_map.get(vid)
                        if val:
                            value_text = " ".join([
                                block_map.get(w, {}).get("Text", "")
                                for r in val.get("Relationships", []) if r["Type"] == "CHILD"
                                for w in r["Ids"]
                            ]).strip()
            if key_text and value_text:
                kv_pairs.append(f"{key_text}: {value_text}")
        if block["BlockType"] == "TABLE":
            table_str = "Table:\n"
            for rel in block.get("Relationships", []):
                if rel["Type"] == "CHILD":
                    for cell_id in rel["Ids"]:
                        cell = block_map.get(cell_id)
                        if cell and cell["BlockType"] == "CELL":
                            cell_text = " ".join([
                                block_map.get(w, {}).get("Text", "")
                                for r in cell.get("Relationships", []) if r["Type"] == "CHILD"
                                for w in r["Ids"]
                            ]).strip()
                            table_str += f"[{cell.get('RowIndex','?')},{cell.get('ColumnIndex','?')}] {cell_text}\n"
            if len(table_str.splitlines()) > 2:
                tables.append(table_str)
    parts = []
    # If tables exist, send ONLY tables to Claude (not LINES which has same data repeated)
    parts = []
    if tables:
        parts.append("=== TABLES ===\n" + "\n\n".join(tables))
        if kv_pairs:
            parts.append("=== KEY-VALUES ===\n" + "\n".join(kv_pairs))
    else:
        if lines:
            parts.append("=== LINES ===\n" + "\n".join(lines))
        if kv_pairs:
            parts.append("=== KEY-VALUES ===\n" + "\n".join(kv_pairs))
    return "\n\n".join(parts) or "[No content extracted]"


# ── BEDROCK CLAUDE — MULTI-CUSTOMER PROMPT ────────────────────────────────────
def _call_bedrock(text_content, source_type):
    prompt = f"""You are an intelligent document extraction AI. Extract ALL rows from a MULTI-CUSTOMER remittance document exactly as they appear.

CRITICAL RULES:
1. Extract EVERY row exactly as it appears — do NOT skip, merge, or duplicate any row
2. Preserve the EXACT ORDER of rows as they appear in the document — do NOT rearrange or group them
3. Preserve the CLASS value exactly as in the document (INV, Invoice, PMT, Payment, PAY, RECEIPT etc.)
4. Each row in the input = exactly one row in the output — NEVER output more or fewer rows than exist in the input
5. Do NOT nest invoices under payments — every row is independent
6. Payment rows (PMT/Payment) and Invoice rows (INV/Invoice) are ALL output as separate rows
7. Preserve the original sign of amounts (negative for payments, positive for invoices)
8. Empty/blank rows between customers should be ignored (do not output blank rows)
9. If the SAME customer appears in MULTIPLE SEPARATE BLOCKS in the document, output them as SEPARATE BLOCKS in the SAME ORDER — do NOT merge them into one block
10. If the SAME row (same TRX_NUMBER, same date, same amount) appears TWICE in the document, output it TWICE — do NOT deduplicate

CUSTOMER ASSIGNMENT RULES (VERY IMPORTANT):
- Each row has its own CUST_NO and CUSTOMER_NAME already shown in that row
- Use ONLY the CUST_NO and CUSTOMER_NAME that appear on THAT SPECIFIC ROW
- Do NOT carry forward or inherit CUST_NO/CUSTOMER_NAME from a previous row
- NEVER assign a row's data to a different customer than what the document shows on that row

For each row extract:
- CUST_NO: the customer number shown ON THIS ROW
- CUSTOMER_NAME: the customer name shown ON THIS ROW
- TRX_NUMBER: transaction/reference number on THIS ROW
- TXN_DATE: date in DD/MM/YYYY format
- CLASS: exactly as in document (INV, Invoice, PMT, Payment, etc.)
- OUTSTANDING_AMT: outstanding amount (preserve sign)
- REJECTION_SHORT: rejection/short paid amount
- PAID: paid amount
- TDS: TDS amount
- APPLIED_AMT: applied/net amount (preserve sign)

Return ONLY valid JSON (no markdown, no extra text):
{{
  "rows": [
    {{
      "CUST_NO": "<customer number>",
      "CUSTOMER_NAME": "<customer name>",
      "TRX_NUMBER": "<transaction number>",
      "TXN_DATE": "<DD/MM/YYYY>",
      "CLASS": "<exactly as in document>",
      "OUTSTANDING_AMT": <number or null>,
      "REJECTION_SHORT": <number or null>,
      "PAID": <number or null>,
      "TDS": <number or null>,
      "APPLIED_AMT": <number or null>
    }}
  ]
}}

Document Content:
{text_content}"""

    logger.info("Calling Bedrock Claude (source_type=%s) ...", source_type)
    response = bedrock_runtime.converse(
        modelId=MODEL_ID,
        system=[{"text": ("You output ONLY raw JSON. No markdown, no explanation, no preamble, "
                          "no ```json fences. Your entire response must be valid JSON starting with {{ and ending with }}.")}],
        messages=[{"role": "user", "content": [{"text": prompt}]}],
        inferenceConfig={"maxTokens": 16384, "temperature": 0.0},
    )

    output_text = "".join(
        block["text"] for block in response["output"]["message"]["content"] if "text" in block
    )
    logger.info("Bedrock raw output (first 800 chars): %s", output_text[:800])

    cleaned = output_text.strip()
    cleaned = re.sub(r"^```json\s*", "", cleaned, flags=re.IGNORECASE)
    cleaned = re.sub(r"^```\s*",     "", cleaned)
    cleaned = re.sub(r"\s*```$",     "", cleaned)
    cleaned = re.sub(r",\s*([}\]])", r"\1", cleaned)
    cleaned = cleaned.strip()

    json_start = cleaned.find("{")
    if json_start > 0:
        logger.warning("Stripping preamble (%d chars) before JSON", json_start)
        cleaned = cleaned[json_start:]

    json_end = cleaned.rfind("}")
    if json_end != -1 and json_end < len(cleaned) - 1:
        trailing = cleaned[json_end + 1:].strip()
        if trailing:
            logger.warning("Stripping trailing text after JSON: %s", trailing[:120])
            cleaned = cleaned[:json_end + 1]

    try:
        extracted_data = json.loads(cleaned)
        logger.info("JSON parse SUCCESS")
    except json.JSONDecodeError as e:
        logger.error("JSON parse FAILED: %s", str(e))
        return {"error": "JSON parsing failed", "parse_error": str(e), "raw_output": output_text}, []

    # Flat rows — each input row = one output row
    rows = extracted_data.get("rows", [])
    if not rows:
        logger.warning("No rows found in Claude output")
        return extracted_data, []

    logger.info("Total rows from Claude: %d", len(rows))

    # ── NO DEDUP — extract exactly as-is from the document ──
    # If the Excel has the same row twice, we output it twice.
    # Do NOT remove "duplicate" rows — the document is the source of truth.

    # Attach _header placeholder so _build_output works
    claude_rows = []
    for row in rows:
        row["_header"] = {
            "CUST_NO":      row.get("CUST_NO", ""),
            "CUSTOMER_NAME": row.get("CUSTOMER_NAME", ""),
        }
        claude_rows.append(row)

    return extracted_data, claude_rows


# ── BUILD OUTPUT ──────────────────────────────────────────────────────────────
def _build_output(table_rows, extracted_data, default_source_type="EXCEL",
                  mail_id="", mail_received_date=""):
    desired_order = [
        "CUST_NO", "CUSTOMER_NAME", "TRX_NUMBER", "TXN_DATE", "CLASS",
        "OUTSTANDING_AMT", "REJECTION_SHORT", "PAID", "TDS", "APPLIED_AMT",
        "MAIL_ID", "MAIL_RECEIVED_DATE",
    ]

    merged_rows = []
    for row in table_rows:
        merged_rows.append({
            "CUST_NO":           row.get("CUST_NO", ""),
            "CUSTOMER_NAME":     row.get("CUSTOMER_NAME", ""),
            "TRX_NUMBER":        row.get("TRX_NUMBER"),
            "TXN_DATE":          convert_date(row.get("TXN_DATE")),
            "CLASS":             row.get("CLASS", ""),
            "OUTSTANDING_AMT":   normalize_amount(row.get("OUTSTANDING_AMT")),
            "REJECTION_SHORT":   normalize_amount(row.get("REJECTION_SHORT")),
            "PAID":              normalize_amount(row.get("PAID")),
            "TDS":               normalize_amount(row.get("TDS")),
            "APPLIED_AMT":       normalize_amount(row.get("APPLIED_AMT")),
            "MAIL_ID":           mail_id,
            "MAIL_RECEIVED_DATE": mail_received_date,
        })

    if not merged_rows:
        logger.warning("No rows to output")

    # ── Output rows EXACTLY as they appear in the document ──
    # Do NOT group by customer — preserve original block order from the Excel
    # If same customer appears in 2 separate blocks, keep them separate
    # Add blank rows between blocks where the customer changes
    final_rows = []
    prev_customer = None
    for row in merged_rows:
        current_customer = (row.get("CUST_NO", ""), row.get("CUSTOMER_NAME", ""))
        if prev_customer is not None and current_customer != prev_customer:
            # Customer changed — add blank separator row
            blank = {col: None for col in desired_order}
            final_rows.append(blank)
        final_rows.append(row)
        prev_customer = current_customer

    df_flat = pd.DataFrame(final_rows)
    df_flat = df_flat[[c for c in desired_order if c in df_flat.columns]]

    excel_buffer = BytesIO()
    with pd.ExcelWriter(excel_buffer, engine="openpyxl") as writer:
        df_flat.to_excel(writer, sheet_name="MultiCustomer", index=False)
        pd.DataFrame([{"full_json": json.dumps(extracted_data, indent=2)}]).to_excel(
            writer, sheet_name="Raw_JSON", index=False
        )
    excel_buffer.seek(0)
    return excel_buffer.getvalue(), merged_rows


# ── HELPERS ───────────────────────────────────────────────────────────────────
def normalize_amount(value):
    if value is None:
        return None
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        num = float(value)
        return int(num) if num == int(num) else num
    text = str(value).strip().replace(",", "").replace(" ", "").replace("\u2212", "-")
    is_neg = text.startswith("-")
    if text.endswith("-"):
        text = text[:-1]
    if is_neg and not text.startswith("-"):
        text = "-" + text
    m = re.search(r"^-?\d+(?:\.\d+)?", text)
    if not m:
        return None
    result = float(m.group(0))
    return int(result) if result == int(result) else result


# ── WRITE TO DYNAMODB ─────────────────────────────────────────────────────────
def write_to_dynamodb(table_name, rows):
    dynamodb_client = boto3.client("dynamodb")
    try:
        dynamodb_client.describe_table(TableName=table_name)
        logger.info("DynamoDB table %s already exists", table_name)
    except dynamodb_client.exceptions.ResourceNotFoundException:
        logger.info("Creating DynamoDB table: %s", table_name)
        dynamodb_client.create_table(
            TableName=table_name,
            KeySchema=[
                {"AttributeName": "document_name", "KeyType": "HASH"},
                {"AttributeName": "row_number",    "KeyType": "RANGE"},
            ],
            AttributeDefinitions=[
                {"AttributeName": "document_name", "AttributeType": "S"},
                {"AttributeName": "row_number",    "AttributeType": "N"},
            ],
            BillingMode="PAY_PER_REQUEST",
        )
        waiter = dynamodb_client.get_waiter("table_exists")
        waiter.wait(TableName=table_name, WaiterConfig={"Delay": 2, "MaxAttempts": 30})
        logger.info("DynamoDB table %s created", table_name)

    column_order = [
        "CUST_NO", "CUSTOMER_NAME", "TRX_NUMBER", "TXN_DATE", "CLASS",
        "OUTSTANDING_AMT", "REJECTION_SHORT", "PAID", "TDS", "APPLIED_AMT",
        "MAIL_ID", "MAIL_RECEIVED_DATE",
    ]

    table = dynamodb.Table(table_name)
    with table.batch_writer() as batch:
        for idx, row in enumerate(rows, start=1):
            # Skip blank separator rows
            if not any(row.get(k) for k in ["TRX_NUMBER", "CUSTOMER_NAME", "CUST_NO"]):
                continue
            item = {"document_name": table_name, "row_number": idx}
            for key in column_order:
                value = row.get(key)
                if value is None or value == "":
                    item[key] = ""
                elif isinstance(value, (int, float)):
                    item[key] = Decimal(str(value))
                else:
                    item[key] = str(value)
            batch.put_item(Item=item)
            if idx % 10 == 0:
                logger.info("Written %d rows to DynamoDB", idx)
    logger.info("Total %d rows written to table %s", len(rows), table_name)
