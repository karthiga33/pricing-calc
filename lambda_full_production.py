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
OUTPUT_PREFIX = "output/"
MODEL_ID = "global.anthropic.claude-sonnet-4-6"

from botocore.config import Config

s3_client = boto3.client("s3")
textract = boto3.client("textract")
bedrock_runtime = boto3.client("bedrock-runtime", region_name="us-east-1",
                               config=Config(read_timeout=300, connect_timeout=10, retries={"max_attempts": 2}))
dynamodb = boto3.resource("dynamodb")

logger = logging.getLogger()
logger.setLevel(logging.INFO)


# ── DATE CONVERSION ───────────────────────────────────────────────────────────
def convert_date(raw_date):
    """Normalize any date format to DD/MM/YYYY before storing in DynamoDB."""
    if not raw_date or str(raw_date).strip() == "":
        return ""
    raw_date = str(raw_date).strip()
    formats = [
        "%d-%b-%y", "%d-%b-%Y", "%d-%m-%Y", "%Y-%m-%d",
        "%d/%m/%Y", "%m/%d/%Y", "%d%m%Y", "%Y%m%d",
        "%d.%m.%Y", "%d.%m.%y",
        "%Y-%m-%d %H:%M:%S", "%Y-%m-%dT%H:%M:%S", "%Y-%m-%dT%H:%M:%SZ",
    ]
    for fmt in formats:
        try:
            return datetime.strptime(raw_date, fmt).strftime("%d/%m/%Y")
        except ValueError:
            continue
    logger.warning("Could not parse date: '%s' — storing as-is", raw_date)
    return raw_date


# ── READ EMAIL METADATA SAVED BY LAMBDA 1 ────────────────────────────────────
def get_email_metadata(bucket_name, input_key):
    att_name = os.path.basename(input_key)
    meta_key = f"metadata/{att_name}.json"
    try:
        response = s3_client.get_object(Bucket=bucket_name, Key=meta_key)
        metadata = json.loads(response["Body"].read().decode("utf-8"))
        logger.info(
            "✅ Metadata found for %s — MAIL_ID=%s, MAIL_RECEIVED_DATE=%s",
            att_name, metadata.get("MAIL_ID", ""), metadata.get("MAIL_RECEIVED_DATE", ""),
        )
        return metadata
    except Exception:
        logger.warning(
            "⚠️ No metadata found for %s — MAIL_ID and MAIL_RECEIVED_DATE will be empty", att_name,
        )
        return {}


# ── ETAG DEDUP TABLE ──────────────────────────────────────────────────────────
DEDUP_TABLE_NAME = "prod_processed_file_etags"


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


REJECTION_TABLE_NAME = "prod_rejected_files"
CORRECTIONS_TABLE_NAME = "prod_extraction_corrections"


def _log_rejection(input_key, etag, reason, mail_id="", mail_received_date=""):
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
            "source": "single_customer",
            "mail_id": mail_id or "",
            "mail_received_date": mail_received_date or "",
        })
        logger.info("Logged rejection: %s — %s", file_name, reason)
    except Exception as e:
        logger.warning("Failed to log rejection for %s: %s", input_key, str(e))


def _move_to_rejected_folder(input_bucket, input_key):
    """Move a rejected file from emails/ to rejected/ folder in S3."""
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


# ── FEEDBACK/CORRECTION LOOP — Self-learning from past mistakes ───────────────
def _get_recent_corrections(customer_name=None, limit=5):
    """
    Fetch recent corrections from DynamoDB to include as few-shot examples
    in the prompt. This allows the model to learn from past extraction mistakes.
    """
    try:
        table = dynamodb.Table(CORRECTIONS_TABLE_NAME)
        # Scan for recent corrections (optionally filtered by customer)
        scan_kwargs = {"Limit": limit * 3}  # Over-fetch then filter
        if customer_name and customer_name.lower() != "unknown":
            scan_kwargs["FilterExpression"] = boto3.dynamodb.conditions.Attr("customer_name").contains(
                customer_name.split()[0]  # Match first word of company name
            )
        try:
            from boto3.dynamodb.conditions import Attr
            response = table.scan(**scan_kwargs)
        except Exception:
            response = table.scan(Limit=limit * 3)
        items = response.get("Items", [])
        # Sort by timestamp descending
        items.sort(key=lambda x: x.get("corrected_at", ""), reverse=True)
        return items[:limit]
    except Exception as e:
        logger.info("No corrections table or no corrections found: %s", str(e))
        return []


def _build_correction_examples(corrections):
    """Build a prompt section from past corrections to help Claude avoid repeating mistakes."""
    if not corrections:
        return ""
    examples = "\n\nLEARN FROM THESE PAST CORRECTIONS (avoid repeating these mistakes):\n"
    for i, corr in enumerate(corrections, 1):
        mistake = corr.get("mistake_description", "")
        correct = corr.get("correct_value", "")
        field = corr.get("field_name", "")
        if mistake and correct:
            examples += f"  {i}. Field '{field}': WRONG='{mistake}' → CORRECT='{correct}'\n"
    return examples


def store_correction(customer_name, field_name, wrong_value, correct_value, file_name=""):
    """Store a user correction so the system can learn from it. Called by the UI/API."""
    try:
        table = dynamodb.Table(CORRECTIONS_TABLE_NAME)
        table.put_item(Item={
            "correction_id": f"{field_name}_{datetime.utcnow().strftime('%Y%m%d%H%M%S')}",
            "customer_name": customer_name or "Unknown",
            "field_name": field_name,
            "mistake_description": str(wrong_value),
            "correct_value": str(correct_value),
            "file_name": file_name,
            "corrected_at": datetime.utcnow().isoformat(),
        })
        logger.info("Stored correction: %s %s='%s' → '%s'", customer_name, field_name, wrong_value, correct_value)
    except Exception as e:
        logger.warning("Failed to store correction: %s", str(e))


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

    # DynamoDB table names must be at least 3 characters
    if len(table_name) < 3:
        table_name = "tbl_" + table_name

    logger.info("Processing  : s3://%s/%s", input_bucket, input_key)
    logger.info("Output Excel: s3://%s/%s", OUTPUT_BUCKET, output_key)
    logger.info("DynamoDB    : %s", table_name)

    excel_bytes, merged_rows = process_file_from_s3(input_bucket, input_key, mail_id, mail_received_date)

    # ── Check if extraction produced meaningful data ──
    # If all rows have no DOCUMENT_NUMBER, no INVOICE_AMOUNT, and no AMOUNT → skip
    has_meaningful_data = any(
        row.get("DOCUMENT_NUMBER") or row.get("INVOICE_AMOUNT") or row.get("AMOUNT")
        for row in merged_rows
    )
    if not has_meaningful_data:
        logger.info("⏭️ No meaningful extraction from %s — skipping output", input_key)
        _log_rejection(input_key, "", "No payment/invoice data extracted — file does not contain relevant content", mail_id, mail_received_date)
        _move_to_rejected_folder(input_bucket, input_key)
        return None

    # ── Validate this is actually a PAYMENT document (not planning/PO/inventory) ──
    # A valid payment document MUST have a UTR/payment reference in the header
    # If no payment reference exists, it's likely a non-payment document that slipped through
    first_header = merged_rows[0].get("_header", {}) if merged_rows and isinstance(merged_rows[0], dict) else {}
    utr_ref = None
    for row in merged_rows:
        h = row.get("_header", {}) if isinstance(row, dict) else {}
        utr_ref = h.get("UTR_REFERENCE_NUMBER")
        if utr_ref and str(utr_ref).strip() and str(utr_ref).strip().lower() not in ("", "none", "null", "unknown", "-"):
            break
    if not utr_ref or str(utr_ref).strip().lower() in ("", "none", "null", "unknown", "-"):
        # No valid payment reference found — check if document numbers look like real invoices
        has_valid_invoice = any(
            row.get("DOCUMENT_NUMBER") and re.fullmatch(r"\d{7,13}", str(row.get("DOCUMENT_NUMBER")).strip())
            for row in merged_rows
        )
        if not has_valid_invoice:
            logger.info("⏭️ No payment reference (UTR) and no valid invoice numbers in %s — rejecting as non-payment document", input_key)
            _log_rejection(input_key, "", "Not a payment/remittance document — no UTR reference and no valid numeric invoice numbers found", mail_id, mail_received_date)
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


def _compute_payment_amount(header_amount, invoice_rows):
    if header_amount is not None:
        return header_amount
    total = None
    for row in invoice_rows:
        inv = normalize_amount(row.get("INVOICE_AMOUNT"))
        if inv is not None:
            try:
                total = (total or 0) + float(inv)
            except (ValueError, TypeError):
                pass
    if total is None:
        return None
    return int(total) if float(total).is_integer() else total


def process_file_from_s3(input_bucket, input_key, mail_id="", mail_received_date=""):
    file_ext = input_key.lower()
    if file_ext.endswith(".xlsx") or file_ext.endswith(".xls"):
        logger.info("Processing as Excel ...")
        return process_excel_from_s3(input_bucket, input_key, mail_id, mail_received_date)
    if file_ext.endswith(".docx") or file_ext.endswith(".doc"):
        logger.info("Processing as Word document ...")
        return process_word_from_s3(input_bucket, input_key, mail_id, mail_received_date)
    if file_ext.endswith(".html"):
        logger.info("Processing as HTML ...")
        return process_text_from_s3(input_bucket, input_key, mail_id, mail_received_date)
    if file_ext.endswith(".csv") or file_ext.endswith(".txt"):
        logger.info("Processing as text/CSV ...")
        return process_text_from_s3(input_bucket, input_key, mail_id, mail_received_date)
    return process_image_or_pdf_from_s3(input_bucket, input_key, mail_id, mail_received_date)


# ── Word document processor (.doc / .docx) ────────────────────────────────────
def process_word_from_s3(input_bucket, input_key, mail_id="", mail_received_date=""):
    import io
    import zipfile
    import xml.etree.ElementTree as ET

    file_ext    = input_key.lower()
    source_type = "DOCX" if file_ext.endswith(".docx") else "DOC"
    response   = s3_client.get_object(Bucket=input_bucket, Key=input_key)
    word_bytes = response["Body"].read()

    def _extract_via_zip(raw_bytes):
        text_parts  = []
        table_parts = []
        tbl_idx     = 0
        with zipfile.ZipFile(io.BytesIO(raw_bytes)) as zf:
            with zf.open("word/document.xml") as doc_xml:
                tree = ET.parse(doc_xml)
                root = tree.getroot()
                ns   = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
                body = root.find(".//w:body", ns)
                if body is None:
                    body = root
                for child in body:
                    tag = child.tag.split("}")[-1]
                    if tag == "p":
                        run_texts = [t.text or "" for t in child.findall(".//w:t", ns)]
                        line = "".join(run_texts).strip()
                        if line:
                            text_parts.append(line)
                    elif tag == "tbl":
                        tbl_idx += 1
                        table_parts.append(f"\nTable {tbl_idx}:")
                        for row_el in child.findall(".//w:tr", ns):
                            cells = []
                            for cell_el in row_el.findall(".//w:tc", ns):
                                cell_text = "".join(
                                    t.text or "" for t in cell_el.findall(".//w:t", ns)
                                ).strip()
                                cells.append(cell_text)
                            row_line = " | ".join(cells)
                            if row_line.strip():
                                table_parts.append(row_line)
        content = "=== WORD DOCUMENT CONTENT ===\n\n"
        if text_parts:
            content += "Paragraphs:\n" + "\n".join(text_parts) + "\n"
        if table_parts:
            content += "\nTables:\n" + "\n".join(table_parts) + "\n"
        logger.info("zipfile/XML extraction succeeded — %d paragraph lines, %d table blocks",
                    len(text_parts), tbl_idx)
        return content

    if file_ext.endswith(".docx"):
        try:
            text_content = _extract_via_zip(word_bytes)
        except Exception as zip_err:
            logger.error("zipfile/XML extraction failed for .docx (%s)", str(zip_err))
            raise RuntimeError(f"Could not extract text from {input_key} (docx/zip): {zip_err}")
    else:
        logger.info(".doc file detected — trying ZIP/OOXML extraction first ...")
        zip_ok = False
        try:
            text_content = _extract_via_zip(word_bytes)
            zip_ok = True
            logger.info(".doc ZIP/OOXML extraction succeeded (%d chars)", len(text_content))
        except Exception as zip_err:
            logger.warning(
                ".doc ZIP/OOXML extraction failed (%s) — falling back to antiword/LibreOffice",
                str(zip_err),
            )
        if not zip_ok:
            text_content = _extract_text_from_ole_doc(word_bytes, input_key)

    logger.info("%s -> text (first 1000 chars): %s", source_type, text_content[:1000])
    extracted_data, claude_rows = _call_bedrock(text_content, source_type=source_type)
    return _build_output(claude_rows, extracted_data, default_source_type=source_type,
                         mail_id=mail_id, mail_received_date=mail_received_date)


# ── OLE .doc extractor: antiword → LibreOffice → binary scrape ───────────────
def _extract_text_from_ole_doc(data, input_key="file.doc"):
    tmp_path = None
    try:
        with tempfile.NamedTemporaryFile(suffix=".doc", delete=False, dir="/tmp") as tmp:
            tmp.write(data)
            tmp_path = tmp.name

        antiword_bin = _find_executable(["antiword", "/opt/bin/antiword"])
        if antiword_bin:
            try:
                env = os.environ.copy()
                if not env.get("ANTIWORDHOME") and os.path.isdir("/opt/share/antiword"):
                    env["ANTIWORDHOME"] = "/opt/share/antiword"
                result = subprocess.run(
                    [antiword_bin, tmp_path],
                    capture_output=True, text=True, timeout=30, env=env
                )
                if result.returncode == 0 and result.stdout.strip():
                    logger.info("antiword extraction succeeded (%d chars)", len(result.stdout))
                    return "=== WORD DOCUMENT CONTENT ===\n\n" + result.stdout
                else:
                    logger.warning("antiword returned code %d: %s", result.returncode, result.stderr[:200])
            except Exception as aw_err:
                logger.warning("antiword subprocess error: %s", str(aw_err))
        else:
            logger.info("antiword not found — skipping to LibreOffice")

        soffice_bin = _find_executable(["soffice", "/opt/libreoffice/program/soffice", "/usr/bin/soffice"])
        if soffice_bin:
            try:
                out_dir = tempfile.mkdtemp(dir="/tmp")
                env = os.environ.copy()
                env["HOME"] = "/tmp"
                result2 = subprocess.run(
                    [soffice_bin, "--headless", "--convert-to", "txt:Text", "--outdir", out_dir, tmp_path],
                    capture_output=True, text=True, timeout=60, env=env
                )
                txt_file = os.path.join(out_dir, os.path.basename(tmp_path).replace(".doc", ".txt"))
                if os.path.exists(txt_file):
                    with open(txt_file, "r", errors="ignore") as f:
                        text = f.read()
                    if text.strip():
                        logger.info("LibreOffice extraction succeeded (%d chars)", len(text))
                        return "=== WORD DOCUMENT CONTENT ===\n\n" + text
                    else:
                        logger.warning("LibreOffice produced empty output")
                else:
                    logger.warning("LibreOffice conversion did not produce %s — stderr: %s",
                                   txt_file, result2.stderr[:200])
            except Exception as lo_err:
                logger.warning("LibreOffice subprocess error: %s", str(lo_err))
        else:
            logger.info("soffice not found — skipping to binary scrape")

        logger.warning("Both antiword and LibreOffice unavailable or failed — "
                       "falling back to raw binary scrape (quality may be poor)")
        return _binary_scrape_doc(data)
    finally:
        if tmp_path and os.path.exists(tmp_path):
            try:
                os.unlink(tmp_path)
            except Exception:
                pass


def _find_executable(candidates):
    import shutil
    for candidate in candidates:
        found = shutil.which(candidate) or (candidate if os.path.isfile(candidate) else None)
        if found:
            return found
    return None


def _binary_scrape_doc(data):
    try:
        import olefile
        ole = olefile.OleFileIO(data)
        if ole.exists("WordDocument"):
            word_stream = ole.openstream("WordDocument").read()
            text_parts = []
            current    = []
            for byte in word_stream:
                if 32 <= byte <= 126 or byte in (9, 10, 13):
                    current.append(chr(byte))
                elif 128 <= byte <= 255:
                    try:
                        ch = bytes([byte]).decode("cp1252")
                        current.append(ch)
                    except Exception:
                        if len(current) >= 4:
                            text_parts.append("".join(current))
                        current = []
                else:
                    if len(current) >= 4:
                        text_parts.append("".join(current))
                    current = []
            if len(current) >= 4:
                text_parts.append("".join(current))
            if text_parts:
                best = max(text_parts, key=len)
                if len(best) > 100:
                    clean = best.replace("\r", "\n")
                    clean = re.sub(r"\n{3,}", "\n\n", clean).strip()
                    logger.info("olefile CP1252 extraction succeeded (%d chars)", len(clean))
                    return "=== WORD DOCUMENT CONTENT ===\n\n" + clean
        logger.warning("olefile: WordDocument stream empty or too short — falling back")
    except ImportError:
        logger.warning("olefile not installed — falling back to raw binary scrape")
    except Exception as ole_err:
        logger.warning("olefile extraction failed (%s) — falling back", str(ole_err))

    logger.info("Using raw binary string scrape for .doc")

    def _scrape_strings(raw, min_len=4):
        ascii_pat = re.compile(rb"[ -~\t\r\n]{" + str(min_len).encode() + rb",}")
        for m in ascii_pat.finditer(raw):
            yield m.group(0).decode("ascii", errors="ignore")
        try:
            utf16 = raw.decode("utf-16-le", errors="ignore")
            utf16_pat = re.compile(r"[ -~\t\r\n]{" + str(min_len) + r",}")
            for m in utf16_pat.finditer(utf16):
                yield m.group(0)
        except Exception:
            pass

    seen       = set()
    text_lines = []
    for s in _scrape_strings(data, min_len=4):
        s = s.strip()
        if s and s not in seen:
            seen.add(s)
            text_lines.append(s)
    clean_lines = [line.strip() for line in text_lines
                   if line.strip() and any(c.isalnum() for c in line)]
    return "=== WORD DOCUMENT CONTENT ===\n\n" + "\n".join(clean_lines)


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

    tables       = extract_tables_from_blocks(blocks)
    text_content = textract_blocks_to_text(blocks)
    logger.info("Textract extracted text_content:\n%s", text_content)
    extracted_data, claude_rows = _call_bedrock(text_content, source_type="PDF")

    # ── POST-PROCESSING: Fix invoice numbers if Claude picked wrong column ──
    # If Claude returned rows but the DOCUMENT_NUMBER is not 13 digits,
    # check if the Textract tables have a 13-digit column and override.
    if claude_rows:
        claude_rows = _fix_invoice_numbers_from_tables(claude_rows, tables)

    if not claude_rows:
        logger.info("Claude returned no rows, falling back to table parsing")
        table_rows = []
        for table in tables:
            rows = build_rows_from_table(table)
            if rows:
                table_rows.extend(rows)
        table_rows = merge_tds_rows(table_rows)
        table_rows = calculate_net_amounts(table_rows)
        fallback_header = extracted_data.get("header", {}) if isinstance(extracted_data, dict) else {}
        fallback_header["_computed_payment_amount"] = _compute_payment_amount(
            fallback_header.get("AMOUNT"), table_rows
        )
        for row in table_rows:
            row.setdefault("_header", fallback_header)
        claude_rows = table_rows

    return _build_output(claude_rows, extracted_data, default_source_type="PDF",
                         mail_id=mail_id, mail_received_date=mail_received_date)


def process_excel_from_s3(input_bucket, input_key, mail_id="", mail_received_date=""):
    import io
    response    = s3_client.get_object(Bucket=input_bucket, Key=input_key)
    excel_bytes = response["Body"].read()
    try:
        df_dict = pd.read_excel(io.BytesIO(excel_bytes), sheet_name=None)
    except Exception as e:
        logger.error("Failed to read Excel: %s", str(e))
        raise

    text_content = "=== EXCEL FILE CONTENT ===\n\n"
    for sheet_name, sheet_df in df_dict.items():
        text_content += f"\n=== SHEET: {sheet_name} ===\n"
        text_content += f"Columns: {', '.join(sheet_df.columns.astype(str))}\n\n"
        text_content += "Table Data:\n"
        for idx, row in sheet_df.iterrows():
            row_text = " | ".join([f"{col}: {val}" for col, val in zip(sheet_df.columns, row)])
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
    is_html  = file_ext.endswith(".html")

    if is_html:
        source_type = "HTML"
    elif is_csv:
        source_type = "CSV"
    else:
        source_type = "TXT"

    response  = s3_client.get_object(Bucket=input_bucket, Key=input_key)
    raw_bytes = response["Body"].read()
    try:
        raw_text = raw_bytes.decode("utf-8")
    except UnicodeDecodeError:
        raw_text = raw_bytes.decode("latin-1")

    if is_html:
        clean = re.sub(r"<(script|style)[^>]*>.*?</(script|style)>", " ", raw_text,
                       flags=re.IGNORECASE | re.DOTALL)
        clean = re.sub(r"<(br|tr|p|div|h[1-6]|li|td|th)[^>]*>", "\n", clean,
                       flags=re.IGNORECASE)
        clean = re.sub(r"<[^>]+>", " ", clean)
        clean = clean.replace("&nbsp;", " ").replace("&amp;", "&") \
                     .replace("&lt;", "<").replace("&gt;", ">") \
                     .replace("&quot;", '"').replace("&#39;", "'")
        clean = re.sub(r"[ \t]+", " ", clean)
        clean = re.sub(r"\n{3,}", "\n\n", clean).strip()
        text_content = f"=== HTML FILE CONTENT ===\n\n{clean}"
        logger.info("HTML -> text (first 1000 chars): %s", text_content[:1000])
        extracted_data, claude_rows = _call_bedrock(text_content, source_type=source_type)
        return _build_output(claude_rows, extracted_data, default_source_type=source_type,
                             mail_id=mail_id, mail_received_date=mail_received_date)

    if is_csv:
        try:
            df           = pd.read_csv(io.StringIO(raw_text))
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


def _repair_truncated_json(text):
    """
    Attempt to repair JSON that was truncated mid-output (e.g., Claude hit token limit).
    Strategy: find the last complete row, close the arrays/objects properly.
    """
    if not text:
        return None

    # Find the last complete invoice_row object (ends with "}")
    # Then close the remaining structure: ]}]}
    last_complete_obj = text.rfind("}")
    if last_complete_obj == -1:
        return None

    # Work backwards from the last "}" to find where the last complete row ends
    # We need to find a "}" that is followed by incomplete content (or end of string)
    attempt = text[:last_complete_obj + 1]

    # Count open/close brackets to determine what's missing
    open_braces = attempt.count("{") - attempt.count("}")
    open_brackets = attempt.count("[") - attempt.count("]")

    # Close them
    closing = "]" * open_brackets + "}" * open_braces

    repaired = attempt + closing
    # Validate it's parseable
    try:
        json.loads(repaired)
        return repaired
    except json.JSONDecodeError:
        pass

    # Second attempt: be more aggressive — find last complete row entry
    # Look for pattern: }, { or }, ] which indicates row boundary
    patterns_to_try = [
        # Find last "}" followed by any whitespace/comma
        (r'\}\s*,\s*\{[^}]*$', '}'),  # Remove last incomplete row
    ]

    for i in range(len(text) - 1, max(0, len(text) - 2000), -1):
        if text[i] == '}':
            chunk = text[:i + 1]
            ob = chunk.count("{") - chunk.count("}")
            obrk = chunk.count("[") - chunk.count("]")
            if ob >= 0 and obrk >= 0:
                closing2 = "]" * obrk + "}" * ob
                candidate = chunk + closing2
                try:
                    json.loads(candidate)
                    logger.info("JSON repair succeeded at position %d (removed last %d chars)",
                                i, len(text) - i - 1)
                    return candidate
                except json.JSONDecodeError:
                    continue

    return None


def _call_bedrock(text_content, source_type):
    # ── Fetch past corrections for self-learning ──
    correction_examples = ""
    try:
        corrections = _get_recent_corrections(limit=5)
        correction_examples = _build_correction_examples(corrections)
    except Exception as e:
        logger.info("Could not fetch corrections (non-critical): %s", str(e))

    prompt = f"""You are an intelligent document extraction AI that learns from context. Extract payment remittance data from ANY format (printed, handwritten, scanned, digital).

CRITICAL: If the document contains MULTIPLE VOUCHERS/PAYMENTS, extract each voucher separately with its own payment details.

YOU MUST:
1. Detect multiple vouchers/payments in the same document (look for repeated headers, multiple check numbers, multiple payment amounts)
2. For each voucher, extract the UTR_REFERENCE_NUMBER using these priority rules:

RULE A — BANK DOCUMENT (UTR format):
- If the document is from a bank (bank letterhead, NEFT/RTGS/IMPS advice, bank payment confirmation), look for a UTR number.
- A genuine bank UTR is EXACTLY 22 characters long, alphanumeric (mix of letters and digits), e.g. "HDFC0012345678901234AB".
- If a 22-character alphanumeric reference exists in the document, use it as UTR_REFERENCE_NUMBER.
- Do NOT use a number that is not exactly 22 characters as the UTR even if labelled "UTR".

RULE B — VENDOR / CUSTOMER DOCUMENT (non-bank):
- If the document is from a vendor or customer (remittance advice, payment voucher, debit note), a 22-character UTR may not be present.
- In this case use the best available reference in this priority order:
  1. Voucher Number / Pay Voucher No / PV No
  2. Check Number / Cheque Number
  3. Transaction ID / Reference Number / Ref No
  4. Any other unique payment identifier in the document
- Take whatever is available; do NOT leave UTR_REFERENCE_NUMBER empty if any reference exists.

RULE C — FALLBACK:
- If a field is explicitly labelled "UTR" but is not 22 characters, still use it (vendor may have typed a partial UTR).
- Always prefer the most specific and unique identifier available.

3. Extract Payment Date (may be same or different per voucher)
4. Extract Payment Amount/Check Amount (total for that voucher)
5. Identify the PAYER company (customer making payment) - NOT the receiver (Tube Investments/TIDC India)
6. Find payment reference per Rule A/B/C above
7. Extract ALL invoice rows with their amounts
8. Recognize TDS/Tax entries and merge them with corresponding invoices
9. Handle ANY document layout - tables, lists, paragraphs, handwritten notes
10. Parse complex patterns like "S|3002010005572|261300037731|ERS-2613000" -> extract first segment "3002010005572"
11. Ignore totals, summaries, and "Unsettled Transactions" sections
12. NEVER include "Balance carried forward", "Balance b/f", "Balance c/f", or page totals as invoice rows

CRITICAL RULE — ERS PREFIX HANDLING:
- "ERS" stands for Evaluated Receipt Settlement — these are AUTO-GENERATED settlement references, NOT actual invoice numbers.
- If a Document Number starts with "ERS-" or "ERS" followed by digits (e.g., "ERS-302822000210", "ERS-3028220002101"):
  * Strip the "ERS-" or "ERS" prefix completely
  * Use ONLY the numeric portion as the DOCUMENT_NUMBER
  * Example: "ERS-302822000210" → DOCUMENT_NUMBER = "302822000210"
  * Example: "ERS-3028220002101" → DOCUMENT_NUMBER = "3028220002101"
- If the table has BOTH an ERS-prefixed column AND a separate 13-digit invoice number column, ALWAYS prefer the 13-digit column.
- ERS references in a "Surcharge/Inv No" or "Description" column are NOT invoice numbers — ignore them.

CRITICAL RULE — EXCLUDE SUMMARY/TOTAL/FORWARD ROWS:
- NEVER include these as invoice rows:
  * "Balance carried forward" / "Balance b/f" / "Balance c/f" / "Carried forward"
  * Any row labelled "Total", "Subtotal", "Grand Total", "Net Amount", "Sum"
  * Page summary rows (usually last row on a page with running totals)
  * Rows where the Document Number field contains text like "Forward", "Total", "Subtotal"
  * Rows where the amount equals the PAYMENT_AMOUNT (header total) — these are summary rows
- If a row's DOCUMENT_NUMBER is the same as the page/document reference number (e.g., "1500016021" which is the Document Number in the header), it is a SUMMARY ROW — exclude it.
- Look at the CONTEXT: if a row appears after "Balance carried forward" text or at the bottom of a page with a running total, it is NOT an invoice row.

CRITICAL RULE — CUSTOMER_NAME EXTRACTION:
- CUSTOMER_NAME is the company/person who is MAKING the payment (the sender/payer), NOT the receiver.
- The receiver is ALWAYS Tube Investments of India Ltd / TIDC India / TII / M/S TUBE INVESTMENTS — NEVER return this as CUSTOMER_NAME.
- IMPORTANT: The "Company:" field in many payment advices refers to the RECEIVER (Tube Investments), NOT the payer. IGNORE the "Company:" label — it is misleading.
- Look for the PAYER name in these locations (priority order):
  1. THE VERY FIRST LINE(S) of the document — the company letterhead/logo name printed at the TOP (e.g., "Jay Bharat Maruti Limited", "TENNECO", "Gabriel India Ltd"). This is ALWAYS the payer.
  2. "For M/s." or "For Messrs." pattern (company name after this prefix)
  3. "From:" or "Remitter:" or "Customer Name:" or "Buyer:" labeled field
  4. Signature block ("Authorised Signatory" section — the company name near it)
  5. Email sender domain (if metadata shows from: user@company.com, company name may be derived)
- KEY INSIGHT: In Indian payment advices, the document is SENT BY the payer company. The payer's name appears as the LETTERHEAD (first prominent text at top). The "Company:" field usually shows the RECEIVER (Tube Investments). Do NOT use the "Company:" field as CUSTOMER_NAME.
- NEVER use address fragments as company names. These are NOT company names:
  - GAT, PLOT, ROAD, HIGHWAY, LANE, STREET, SECTOR, BLOCK, FLOOR, BUILDING
  - City names, PIN codes, district names, state names
  - Phone numbers, email addresses, registration numbers
- The CUSTOMER_NAME must be a proper company/business name (e.g., "TENNECO", "Jay Bharat Maruti Limited", "TI CYCLES OF INDIA", "DHANSHREE ENTERPRISES", "GABRIEL INDIA LTD", "VST TILLERS TRACTORS LIMITED")
- If you cannot confidently identify the payer name, use "Unknown" — NEVER guess from address lines.
- Company name is usually present in the logo, letterhead, or the very top of the document
- IMPORTANT: The company name on the letterhead IS the payer. For example, if "TENNECO" appears as the letterhead/logo and "TUBE INVESTMENTS OF INDIA LTD" appears as the payee/beneficiary, then CUSTOMER_NAME = "TENNECO"
- Another example: If "Jay Bharat Maruti Limited" is printed at the top and "Company: TUBE INVESTMENTS OF INDIA LIMITED" is in the body, then CUSTOMER_NAME = "Jay Bharat Maruti Limited" (NOT Tube Investments)
- NEVER return null, empty string, or "Unknown" for CUSTOMER_NAME when a clear company letterhead/logo exists at the top of the document.
- If the document says "Payment advice" and has a company logo/name at the top, that company IS the CUSTOMER_NAME (the one making the payment).
- Common payer companies you may encounter: TENNECO, Jay Bharat Maruti Limited, TI CYCLES OF INDIA, GABRIEL INDIA, DHANSHREE ENTERPRISES, VST TILLERS TRACTORS — always use the exact name as it appears on the document.
- NEVER leave CUSTOMER_NAME as null or empty. If absolutely no company name is identifiable, use "Unknown" — but ALWAYS try to extract it first.

CRITICAL RULE — DOCUMENT_NUMBER MUST BE 13-DIGIT INVOICE NUMBER:
- Look for the column that contains EXACTLY 13-digit numeric values — these always START WITH 3, 4, or 7 (e.g., 3004010162475, 4001010284952, 7001010266491).
- This 13-digit number may appear under ANY column name: "Your Document", "Your Doc", "Reference", "Ref No", "Doc No", "Invoice No", "Bill No", "Document Number", etc.
- The table may have MULTIPLE number columns. ALWAYS pick the one with 13-digit values starting with 3, 4, or 7.
- NEVER use shorter numbers (10-digit like 2637035520, or 6-8 digit numbers) as DOCUMENT_NUMBER when a 13-digit number exists.
- If a row does NOT have a 13-digit number starting with 3/4/7, check the row immediately below — if it has the 13-digit number for the same transaction, use that.
- If no 13-digit number starting with 3/4/7 is found anywhere, check for 10-digit pure numeric columns starting with 3/4/7 (e.g., 3004010421, 7001010266) — these are shorter-format invoice numbers used by some vendors.
- The DOCUMENT_NUMBER should be a PURE NUMERIC value (10-13 digits, starting with 3, 4, or 7). NEVER use alphanumeric references like "1RD526-00932" or "INV-12345" as DOCUMENT_NUMBER when a pure numeric column exists in the same table.
- CRITICAL: If the table has BOTH a pure-numeric column (e.g., "3004010421") AND an alphanumeric reference column (e.g., "1RD526-00932"), ALWAYS pick the pure-numeric column as DOCUMENT_NUMBER.
- The Invoice Number must be a numeric value. If a numeric value cannot be found, take whatever value is present in that column.

CRITICAL RULE — INV.NO vs BILL NO COLUMN PRIORITY:
- Many invoices have BOTH an "INV.NO" (or "Invoice No") column AND a "Bill No" column in the same table.
- PRIORITY RULE:
  1. FIRST check the "Bill No" column — if it contains a 13-digit number (starting with 3, 4, or 7), either as a single value OR split across multiple rows (e.g., "30080" / "10091" / "403" → concatenate to "3008010091403"), use that as DOCUMENT_NUMBER.
  2. ONLY IF the "Bill No" column does NOT contain a 13-digit number starting with 3/4/7, then fall back to the "INV.NO" / "Invoice No" column.
- The "Bill No" column value may be split across rows because the cell is narrow — always concatenate the fragments to form the complete 13-digit number.
- Example: Table with "INV.NO | Bill No | Bill Date | Bill Amount | TDS | Net Amount":
  * Bill No rows: 30080 → 10091 → 403 → concatenated = "3008010091403" (13 digits, starts with 3) → USE THIS
  * INV.NO = 60025053 → IGNORE (only 8 digits, not 13)
- Another example where Bill No has no 13-digit value: Bill No = "PO-1234" → fall back to INV.NO column.

CRITICAL RULE — LINE-WRAPPED INVOICE/DOCUMENT NUMBERS:
- OCR and document scanners often split a single invoice/document number across two lines when the cell is narrow.
- Example: a cell containing "4001010284" on line 1 and "952" on line 2 means the FULL invoice number is "4001010284952" — you MUST concatenate them WITHOUT any separator.
- Another example: "4001010284" followed immediately by "953" in the same cell/row context → full number is "4001010284953".
- NEVER output only the first fragment (e.g. "4001010284") when a second numeric fragment immediately follows it in the same cell or in the very next line of the same table cell.
- Apply this concatenation rule to every invoice number, document number, or reference number field in every row — without exception.
- Each row that has a different suffix (952, 953, etc.) is a SEPARATE invoice row with its own complete document number.

KEY INTELLIGENCE:
- Look for voucher separators: "VOUCHER 1", "VOUCHER 2", multiple "Pay Voucher No", multiple "Check Number"
- Each voucher has its own Check Number and Check Amount
- If you see invoice number variations (with pipe separators, suffixes like "TDS", "E-TDS-CM"), extract the core invoice number (first segment before pipe)
- Match TDS entries to their invoices by invoice number similarity

CRITICAL RULE — AMOUNT FIELDS (STRICT DEFINITIONS):
- INVOICE_AMOUNT = the amount shown on the invoice line in the document. This is the number you SEE in the document. Do NOT calculate or invent a number that is not visible in the document.
- TDS_AMOUNT = the tax deducted at source. This is always a POSITIVE number (do NOT put negative symbol).
- AMOUNT = NET PAID amount. Formula: AMOUNT = INVOICE_AMOUNT - TDS_AMOUNT.
- If there is NO TDS for a row, then AMOUNT = INVOICE_AMOUNT (they are the same).

STRICT RULE — DO NOT INVENT NUMBERS:
- Every INVOICE_AMOUNT value MUST be a number that actually appears in the document.
- If you add TDS to a net amount and get a gross — that gross number MUST also appear in the document. If it does NOT appear, then DO NOT use it.
- The INVOICE_AMOUNT is simply the main amount value shown on that invoice line in the table.

VALIDATION:
- Sum of all AMOUNT (net) values should approximately equal the total PAYMENT_AMOUNT in the header.
- If sum of INVOICE_AMOUNT values equals PAYMENT_AMOUNT, there is NO TDS — set TDS_AMOUNT to null and AMOUNT = INVOICE_AMOUNT.
- If INVOICE_AMOUNT - TDS_AMOUNT gives the NET, then sum of NET should equal PAYMENT_AMOUNT.

BANK VOUCHER FORMAT (Cr/Dr pattern):
- In bank payment vouchers with Cr/Dr entries:
  * The MAIN amount on each invoice line (typically the Dr amount) = INVOICE_AMOUNT
  * The Cr amount against the same invoice = TDS_AMOUNT
  * AMOUNT (net) = INVOICE_AMOUNT - TDS_AMOUNT
- Learn from document structure: identify sections, headers, table patterns

Return ONLY valid JSON (no markdown, no extra text):
{{
  "vouchers": [
    {{
      "header": {{
        "UTR_REFERENCE_NUMBER": "<per Rule A/B/C above>",
        "PAYMENT_DATE": "<date in DD/MM/YYYY format>",
        "CUSTOMER_NAME": "<company making payment>",
        "AMOUNT": <total payment amount for this voucher as number>
      }},
      "invoice_rows": [
        {{
          "DOCUMENT_NUMBER": "<complete invoice number — concatenate all line-wrapped fragments, never truncate>",
          "DOCUMENT_DATE": "<DD/MM/YYYY>",
          "INVOICE_AMOUNT": <positive number>,
          "TDS_AMOUNT": <positive number or null>,
          "AMOUNT": <positive net paid amount>
        }}
      ]
    }}
  ]
}}

Document Content:
{text_content}
{correction_examples}"""

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
            logger.warning("Stripping trailing text after JSON (%d chars): %s",
                           len(trailing), trailing[:120])
            cleaned = cleaned[:json_end + 1]

    try:
        extracted_data = json.loads(cleaned)
        logger.info("JSON parse SUCCESS")
    except json.JSONDecodeError as e:
        logger.warning("JSON parse FAILED (attempting repair): %s", str(e))
        # ── Attempt to repair truncated JSON ──
        # Claude may have hit token limit mid-response for documents with many rows.
        # Try to close open brackets/braces to make it valid.
        repaired = _repair_truncated_json(cleaned)
        if repaired:
            try:
                extracted_data = json.loads(repaired)
                logger.info("JSON repair SUCCESS — extracted partial data")
            except json.JSONDecodeError as e2:
                logger.error("JSON repair also FAILED: %s", str(e2))
                return {"error": "JSON parsing failed", "parse_error": str(e), "raw_output": output_text}, []
        else:
            logger.error("JSON repair returned None")
            return {"error": "JSON parsing failed", "parse_error": str(e), "raw_output": output_text}, []

    claude_rows = []
    vouchers    = extracted_data.get("vouchers", [])
    if vouchers:
        logger.info("Detected %d voucher(s)", len(vouchers))
        for voucher in vouchers:
            header = voucher.get("header", {})
            header.setdefault("IMPORT_REFERENCE",  None)
            header.setdefault("SOURCE_TYPE",        source_type)
            header.setdefault("KOD_CUST_CODE",      None)
            header.setdefault("MAIL_ID",             None)
            header.setdefault("MAIL_RECEIVED_DATE", None)
            voucher_rows = voucher.get("invoice_rows", [])
            header["_computed_payment_amount"] = _compute_payment_amount(header.get("AMOUNT"), voucher_rows)
            logger.info("Voucher %s: %d rows", header.get("UTR_REFERENCE_NUMBER"), len(voucher_rows))
            for row in voucher_rows:
                row["_header"] = header
            claude_rows.extend(voucher_rows)
    else:
        header = extracted_data.get("header", {})
        header.setdefault("IMPORT_REFERENCE",  None)
        header.setdefault("SOURCE_TYPE",        source_type)
        header.setdefault("KOD_CUST_CODE",      None)
        header.setdefault("MAIL_ID",             None)
        header.setdefault("MAIL_RECEIVED_DATE", None)
        claude_rows = extracted_data.get("invoice_rows", [])
        header["_computed_payment_amount"] = _compute_payment_amount(header.get("AMOUNT"), claude_rows)
        for row in claude_rows:
            row["_header"] = header

    logger.info("Total rows from Claude (before dedup): %d", len(claude_rows))

    # ── FALLBACK: If CUSTOMER_NAME is empty/null, try to extract from text ──
    _ensure_customer_name(claude_rows, text_content)

    # Deduplicate rows — Claude may extract the same invoice row twice when
    # Textract provides overlapping data sections.
    seen = set()
    unique_rows = []
    for row in claude_rows:
        dedup_key = (
            str(row.get("DOCUMENT_NUMBER", "")).strip(),
            str(row.get("DOCUMENT_DATE", "")).strip(),
            str(row.get("INVOICE_AMOUNT", "")).strip(),
            str(row.get("AMOUNT", "")).strip(),
        )
        if dedup_key not in seen:
            seen.add(dedup_key)
            unique_rows.append(row)
        else:
            logger.info("Dedup: skipping duplicate row %s", dedup_key)

    if len(unique_rows) < len(claude_rows):
        logger.info("Deduplication removed %d duplicate rows", len(claude_rows) - len(unique_rows))
    claude_rows = unique_rows

    # ── POST-PROCESSING: Validate and clean extracted rows ──────────────────
    claude_rows = _post_process_claude_rows(claude_rows)

    logger.info("Total rows from Claude (after dedup + validation): %d", len(claude_rows))
    return extracted_data, claude_rows


def _ensure_customer_name(claude_rows, text_content):
    """
    If Claude returned empty/null CUSTOMER_NAME, try to extract it from text content.
    Looks for common patterns: company names in the first few lines (letterhead),
    known company patterns, etc.
    """
    if not claude_rows:
        return

    # Check if any row already has a customer name
    first_header = claude_rows[0].get("_header", {}) if claude_rows else {}
    customer_name = first_header.get("CUSTOMER_NAME")

    if customer_name and customer_name.strip() and customer_name.strip().lower() not in ("", "unknown", "null", "none", "-", "—"):
        return  # Already has a valid customer name

    # ── Try to extract from text content ──
    # Known receiver names to exclude (these are NEVER the customer/payer)
    receiver_names = [
        "tube investments", "tidc india", "tii", "m/s tube investments",
        "tube investments of india", "ti cycles"
    ]

    # Get first 1000 chars of content (letterhead + header area)
    header_text = text_content[:1000] if text_content else ""

    # Address/location words to skip
    address_words = {"gat", "plot", "road", "highway", "lane", "street", "sector",
                     "block", "floor", "building", "gurgaon", "haryana", "india",
                     "mumbai", "delhi", "chennai", "pune", "bangalore", "kolkata",
                     "hyderabad", "noida", "manesar", "imt"}

    lines = header_text.replace("\r", "\n").split("\n")
    candidate = None

    for line in lines[:15]:  # Check first 15 lines
        line_clean = line.strip()
        if not line_clean or len(line_clean) < 3:
            continue
        # Skip lines that are clearly addresses or metadata
        line_lower = line_clean.lower()
        if any(addr in line_lower for addr in ["plot no", "sector", "gurgaon", "haryana",
                                                "pin", "email:", "phone", "fax", "tel:",
                                                "===", "table", "[", "key-value"]):
            continue
        # Skip if it matches receiver names
        if any(recv in line_lower for recv in receiver_names):
            continue
        # Skip if it looks like a label with value (contains ":" with short prefix)
        if ":" in line_clean and len(line_clean.split(":")[0]) < 20:
            continue
        # Skip dates, numbers-only lines
        if re.fullmatch(r"[\d/.\-\s]+", line_clean):
            continue
        # Skip very long lines (likely paragraphs, not company names)
        if len(line_clean) > 60:
            continue

        # ── Match 1: ALL CAPS word(s) — very likely a company logo/name ──
        if line_clean.isupper() and len(line_clean) >= 3:
            words = line_clean.lower().split()
            if not any(w in address_words for w in words):
                candidate = line_clean
                break

        # ── Match 2: Title case or mixed case company name pattern ──
        if re.match(r"^[A-Z][A-Za-z\s&.,]+(?:LTD|LIMITED|PVT|PRIVATE|INC|CORP|ENTERPRISES?|CO|COMPANY)?\.?$",
                    line_clean, re.IGNORECASE):
            words = line_clean.lower().split()
            if not any(w in address_words for w in words):
                if not any(recv in line_clean.lower() for recv in receiver_names):
                    candidate = line_clean
                    break

        # ── Match 3: Single prominent word (like "TENNECO", "BOSCH", "MAHINDRA") ──
        if re.fullmatch(r"[A-Z]{3,30}", line_clean):
            if line_clean.lower() not in address_words:
                candidate = line_clean
                break

    if candidate:
        logger.info("Fallback CUSTOMER_NAME extraction: '%s'", candidate)
        for row in claude_rows:
            header = row.get("_header", {})
            if not header.get("CUSTOMER_NAME") or header["CUSTOMER_NAME"].strip().lower() in ("", "unknown", "null", "none", "-", "—"):
                header["CUSTOMER_NAME"] = candidate
    else:
        logger.warning("Could not extract CUSTOMER_NAME from text content fallback")


def _post_process_claude_rows(claude_rows):
    """
    Post-processing validation layer that catches errors Claude may still make:
    1. Remove summary/total/balance-forward rows that slipped through
    2. Strip ERS- prefixes from document numbers
    3. Remove rows where document number matches header/page reference
    4. Remove rows with no meaningful data
    """
    if not claude_rows:
        return claude_rows

    validated_rows = []
    for row in claude_rows:
        doc_num = str(row.get("DOCUMENT_NUMBER") or "").strip()
        doc_num_lower = doc_num.lower()

        # ── Rule 1: Skip summary/total/balance-forward rows ──
        skip_keywords = [
            "balance carried forward", "balance b/f", "balance c/f",
            "carried forward", "total", "subtotal", "grand total",
            "net amount", "sum total", "page total", "closing balance",
            "opening balance", "forward"
        ]
        if any(keyword in doc_num_lower for keyword in skip_keywords):
            logger.info("Post-process: skipping summary row with DOCUMENT_NUMBER='%s'", doc_num)
            continue

        # Also check if the amount text contains summary keywords
        amount_str = str(row.get("AMOUNT") or "").strip().lower()
        invoice_str = str(row.get("INVOICE_AMOUNT") or "").strip().lower()
        if any(keyword in amount_str for keyword in skip_keywords):
            logger.info("Post-process: skipping row with summary keyword in AMOUNT='%s'", amount_str)
            continue

        # ── Rule 2: Strip ERS- prefix from document numbers ──
        if re.match(r"^ERS[-\s]?", doc_num, re.IGNORECASE):
            original_doc = doc_num
            # Strip "ERS-", "ERS ", or "ERS" prefix
            doc_num = re.sub(r"^ERS[-\s]?", "", doc_num, flags=re.IGNORECASE).strip()
            row["DOCUMENT_NUMBER"] = doc_num
            logger.info("Post-process: stripped ERS prefix: '%s' → '%s'", original_doc, doc_num)

        # ── Rule 3: DOCUMENT_NUMBER must be NUMERIC ONLY ──
        # Invoice/document numbers are always pure digits (no letters, no dashes, no slashes)
        # If it contains alphabetic characters or special chars like "-", it's a reference code, not an invoice number
        if doc_num:
            clean_doc_num = re.sub(r"[\s\n]+", "", doc_num)
            if not re.fullmatch(r"\d+", clean_doc_num):
                # Contains non-numeric characters — try to extract just the numeric portion
                numeric_only = re.sub(r"[^\d]", "", clean_doc_num)
                if numeric_only and len(numeric_only) >= 7:
                    # Has a substantial numeric portion — use just the digits
                    row["DOCUMENT_NUMBER"] = numeric_only
                    doc_num = numeric_only
                    logger.info("Post-process: stripped non-numeric chars from DOCUMENT_NUMBER → '%s'", numeric_only)
                elif numeric_only and len(numeric_only) < 7:
                    # Too short after stripping — likely not a real invoice number, skip row
                    logger.info("Post-process: skipping row with alphanumeric DOCUMENT_NUMBER='%s' (not a valid invoice)", doc_num)
                    continue
                else:
                    # No digits at all — skip
                    logger.info("Post-process: skipping row with non-numeric DOCUMENT_NUMBER='%s'", doc_num)
                    continue

        # ── Rule 4: Skip rows where the amount equals zero or is negative ──
        # (these are typically adjustment/reversal rows that shouldn't be in output)
        raw_amount = row.get("AMOUNT")
        raw_invoice = row.get("INVOICE_AMOUNT")
        if raw_amount is not None and raw_invoice is not None:
            try:
                amt_val = float(str(raw_amount).replace(",", ""))
                inv_val = float(str(raw_invoice).replace(",", ""))
                # If both are exactly 0, skip
                if amt_val == 0 and inv_val == 0:
                    logger.info("Post-process: skipping zero-amount row doc='%s'", doc_num)
                    continue
            except (ValueError, TypeError):
                pass

        # ── Rule 4b: Validate INVOICE_AMOUNT >= AMOUNT ──
        # INVOICE_AMOUNT is what's in the document (gross). AMOUNT = INVOICE_AMOUNT - TDS.
        # If Claude put the net in INVOICE_AMOUNT and gross in AMOUNT, swap them.
        if raw_amount is not None and raw_invoice is not None:
            try:
                amt_val = float(str(raw_amount).replace(",", ""))
                inv_val = float(str(raw_invoice).replace(",", ""))
                if amt_val > 0 and inv_val > 0 and inv_val < amt_val:
                    # INVOICE_AMOUNT < AMOUNT means they are swapped — fix
                    row["INVOICE_AMOUNT"] = raw_amount
                    row["AMOUNT"] = raw_invoice
                    logger.info("Post-process: swapped INVOICE_AMOUNT/AMOUNT for doc='%s' "
                                "(gross=%s, net=%s)", doc_num, raw_amount, raw_invoice)
            except (ValueError, TypeError):
                pass

        # ── Rule 4c: If TDS exists, ensure AMOUNT = INVOICE_AMOUNT - TDS ──
        raw_tds = row.get("TDS_AMOUNT")
        if row.get("INVOICE_AMOUNT") is not None and raw_tds is not None:
            try:
                inv_val = float(str(row.get("INVOICE_AMOUNT")).replace(",", ""))
                tds_val = float(str(raw_tds).replace(",", ""))
                expected_net = inv_val - abs(tds_val)
                if expected_net > 0:
                    row["AMOUNT"] = int(expected_net) if expected_net == int(expected_net) else round(expected_net, 2)
            except (ValueError, TypeError):
                pass

        # ── Rule 5: Skip rows that have only a document number but no amounts at all ──
        has_any_amount = (
            row.get("INVOICE_AMOUNT") is not None or
            row.get("AMOUNT") is not None or
            row.get("TDS_AMOUNT") is not None
        )
        if doc_num and not has_any_amount:
            logger.info("Post-process: skipping row with doc='%s' but no amounts", doc_num)
            continue

        validated_rows.append(row)

    if len(validated_rows) < len(claude_rows):
        logger.info("Post-processing removed %d invalid rows", len(claude_rows) - len(validated_rows))

    return validated_rows


def _fix_invoice_numbers_from_tables(claude_rows, tables):
    """
    Post-processing: If Claude returned DOCUMENT_NUMBERs that are NOT 13 digits,
    but the Textract tables have a column with 13-digit values (or 12+space+digit),
    override the DOCUMENT_NUMBERs with the correct values from that column.

    Also ensures that a "Bill No" column is never used to override when an
    "INV.NO"/"Invoice No" column exists in the same table.
    """
    if not claude_rows or not tables:
        return claude_rows

    # Check if Claude's document numbers are already 13 digits
    all_13 = True
    for row in claude_rows:
        doc = str(row.get("DOCUMENT_NUMBER") or "").strip()
        clean = re.sub(r"[\s\n]+", "", doc)
        if not re.fullmatch(r"\d{13}", clean):
            all_13 = False
            break

    if all_13:
        return claude_rows  # Already correct, no fix needed

    # Find 13-digit invoice numbers from Textract tables
    for table in tables:
        if not table or len(table) < 2:
            continue
        header_row = [str(x).strip() for x in table[0]]
        data_rows = table[1:]

        # ── Detect Bill No and INV.NO columns in this table ──
        bill_no_col = None
        inv_no_col = None
        for col_i, hdr in enumerate(header_row):
            h = hdr.lower().strip()
            if inv_no_col is None and ("invoice no" in h or "inv no" in h or "inv. no" in h
                    or "inv.no" in h or h == "inv"):
                inv_no_col = col_i
            if bill_no_col is None and ("bill no" in h or "bill number" in h):
                bill_no_col = col_i

        # Priority: Bill No if it has 13-digit value, else INV.NO, else auto-detect
        if bill_no_col is not None:
            bill_no_has_13 = _column_has_13_digit_value(data_rows, bill_no_col)
            if bill_no_has_13:
                # Force 13-digit search to use Bill No column
                exclude_col = inv_no_col  # exclude INV.NO so auto-detect picks Bill No
                logger.info("_fix_invoice_numbers: Bill No col=%d has 13-digit — prioritising it", bill_no_col)
            else:
                # Bill No has no 13-digit — exclude it, let INV.NO or other col be used
                exclude_col = bill_no_col
                logger.info("_fix_invoice_numbers: Bill No col=%d has no 13-digit — excluding, falling back to INV.NO", bill_no_col)
        else:
            exclude_col = None

        thirteen_col = _find_13_digit_invoice_column(data_rows, len(header_row), exclude_col=exclude_col)

        # If Bill No has 13-digit but auto-detect still missed it, force it
        if thirteen_col is None and bill_no_col is not None and _column_has_13_digit_value(data_rows, bill_no_col):
            thirteen_col = bill_no_col
            logger.info("_fix_invoice_numbers: forcing Bill No column idx=%d", bill_no_col)
        elif thirteen_col is None:
            continue

        # Extract 13-digit values from this column (skip summary rows)
        invoice_numbers = []
        skip_indices = set()

        # Pre-compute continuation rows
        for row_idx, row in enumerate(data_rows):
            val = str(get_value(row, thirteen_col) or "").strip()
            clean_val = re.sub(r"[\s\n]+", "", val)
            if re.fullmatch(r"\d{12}", clean_val):
                next_idx = row_idx + 1
                if next_idx < len(data_rows):
                    next_val = str(get_value(data_rows[next_idx], thirteen_col) or "").strip()
                    next_clean = re.sub(r"[\s\n]+", "", next_val)
                    if re.fullmatch(r"\d{1,3}", next_clean):
                        if len(clean_val + next_clean) == 13:
                            skip_indices.add(next_idx)

        for row_idx, row in enumerate(data_rows):
            if row_idx in skip_indices:
                continue
            if not any(str(x).strip() for x in row):
                continue
            # Skip summary rows
            skip = False
            for cell in row:
                if str(cell).lower().strip().rstrip(":").strip() in ("total", "grand total", "subtotal", "net amount", "sum"):
                    skip = True
                    break
            if skip:
                continue

            val = str(get_value(row, thirteen_col) or "").strip()
            completed = _complete_13_digit_invoice(val, row_idx, data_rows, thirteen_col)
            if completed:
                clean_completed = re.sub(r"[\s\n]+", "", str(completed))
                if re.fullmatch(r"\d{12,13}", clean_completed):
                    invoice_numbers.append(clean_completed)

        # If we found 13-digit numbers, map them to Claude rows
        if invoice_numbers and len(invoice_numbers) == len(claude_rows):
            logger.info("Fixing DOCUMENT_NUMBERs from Claude: replacing with 13-digit values from Textract table")
            for i, row in enumerate(claude_rows):
                row["DOCUMENT_NUMBER"] = invoice_numbers[i]
            return claude_rows
        elif invoice_numbers and len(invoice_numbers) >= 1 and len(claude_rows) == 1:
            # Single row case
            logger.info("Fixing single DOCUMENT_NUMBER from Claude: %s -> %s",
                        claude_rows[0].get("DOCUMENT_NUMBER"), invoice_numbers[0])
            claude_rows[0]["DOCUMENT_NUMBER"] = invoice_numbers[0]
            return claude_rows
        elif invoice_numbers and len(invoice_numbers) != len(claude_rows):
            # Row counts don't match exactly — try to match by partial number overlap
            logger.info("Row count mismatch: %d invoice numbers vs %d Claude rows — attempting partial match",
                        len(invoice_numbers), len(claude_rows))
            for row in claude_rows:
                doc = str(row.get("DOCUMENT_NUMBER") or "").strip()
                clean_doc = re.sub(r"[\s\n]+", "", doc)
                # If Claude's value is already 13 digits starting with 3/4/7, keep it
                if re.fullmatch(r"[347]\d{12}", clean_doc):
                    continue
                # Try to find a matching 13-digit number that contains Claude's value as substring
                for inv_num in invoice_numbers:
                    if clean_doc and clean_doc in inv_num:
                        row["DOCUMENT_NUMBER"] = inv_num
                        logger.info("Partial match fix: %s -> %s", clean_doc, inv_num)
                        break
                    elif clean_doc and inv_num.startswith(clean_doc):
                        row["DOCUMENT_NUMBER"] = inv_num
                        logger.info("Prefix match fix: %s -> %s", clean_doc, inv_num)
                        break
            return claude_rows

    return claude_rows


def _build_output(table_rows, extracted_data, default_source_type="PDF",
                  mail_id="", mail_received_date=""):
    desired_order = [
        "CUST_PAYMENT_ID", "SOURCE_TYPE", "IMPORT_REFERENCE", "UTR_REFERENCE_NUMBER",
        "CUSTOMER_NAME", "PAYMENT_DATE", "PAYMENT_AMOUNT", "DOCUMENT_NUMBER",
        "DOCUMENT_DATE", "INVOICE_AMOUNT", "TDS_AMOUNT", "DEDUCTION_AMOUNT",
        "CASH_DISCOUNT", "AMOUNT", "KOD_CUST_CODE", "MAIL_ID", "MAIL_RECEIVED_DATE",
    ]

    merged_rows = []
    for row in table_rows:
        row_header = row.get("_header", {})
        merged_rows.append({
            "CUST_PAYMENT_ID":      row.get("CUST_PAYMENT_ID"),
            "SOURCE_TYPE":          row_header.get("SOURCE_TYPE", default_source_type),
            "IMPORT_REFERENCE":     row_header.get("IMPORT_REFERENCE"),
            "UTR_REFERENCE_NUMBER": row_header.get("UTR_REFERENCE_NUMBER"),
            "CUSTOMER_NAME":        row_header.get("CUSTOMER_NAME"),
            "PAYMENT_DATE":         convert_date(row_header.get("PAYMENT_DATE")),
            "PAYMENT_AMOUNT":       row_header.get("_computed_payment_amount"),
            "DOCUMENT_NUMBER":      normalize_document_number(row.get("DOCUMENT_NUMBER")),
            "DOCUMENT_DATE":        convert_date(row.get("DOCUMENT_DATE")),
            "INVOICE_AMOUNT":       normalize_amount(row.get("INVOICE_AMOUNT")),
            "TDS_AMOUNT":           normalize_amount(row.get("TDS_AMOUNT")),
            "DEDUCTION_AMOUNT":     row.get("DEDUCTION_AMOUNT"),
            "CASH_DISCOUNT":        row.get("CASH_DISCOUNT"),
            "AMOUNT":               normalize_amount(row.get("AMOUNT")),
            "KOD_CUST_CODE":        row_header.get("KOD_CUST_CODE"),
            "MAIL_ID":              mail_id,
            "MAIL_RECEIVED_DATE":   mail_received_date,
        })

    if not merged_rows:
        default_header = extracted_data.get("header", {}) if isinstance(extracted_data, dict) else {}
        merged_rows = [{
            "CUST_PAYMENT_ID":      None,
            "SOURCE_TYPE":          default_header.get("SOURCE_TYPE", default_source_type),
            "IMPORT_REFERENCE":     default_header.get("IMPORT_REFERENCE"),
            "UTR_REFERENCE_NUMBER": default_header.get("UTR_REFERENCE_NUMBER"),
            "CUSTOMER_NAME":        default_header.get("CUSTOMER_NAME"),
            "PAYMENT_DATE":         convert_date(default_header.get("PAYMENT_DATE")),
            "PAYMENT_AMOUNT":       default_header.get("AMOUNT"),
            "DOCUMENT_NUMBER":      None,
            "DOCUMENT_DATE":        None,
            "INVOICE_AMOUNT":       None,
            "TDS_AMOUNT":           None,
            "DEDUCTION_AMOUNT":     None,
            "CASH_DISCOUNT":        None,
            "AMOUNT":               None,
            "KOD_CUST_CODE":        default_header.get("KOD_CUST_CODE"),
            "MAIL_ID":              mail_id,
            "MAIL_RECEIVED_DATE":   mail_received_date,
        }]

    df_flat = pd.DataFrame(merged_rows)
    df_flat = df_flat[[c for c in desired_order if c in df_flat.columns]]
    excel_buffer = BytesIO()
    with pd.ExcelWriter(excel_buffer, engine="openpyxl") as writer:
        df_flat.to_excel(writer, sheet_name="Remittance", index=False)
        pd.DataFrame([{"full_json": json.dumps(extracted_data, indent=2)}]).to_excel(
            writer, sheet_name="Raw_JSON", index=False
        )
    excel_buffer.seek(0)
    return excel_buffer.getvalue(), merged_rows


# ── Textract helpers ──────────────────────────────────────────────────────────
def extract_tables_from_blocks(blocks):
    block_map = {b["Id"]: b for b in blocks}
    tables    = []
    for block in blocks:
        if block.get("BlockType") != "TABLE":
            continue
        cells              = {}
        max_row = max_col  = 0
        for rel in block.get("Relationships", []):
            if rel["Type"] != "CHILD":
                continue
            for cell_id in rel["Ids"]:
                cell = block_map.get(cell_id)
                if not cell or cell.get("BlockType") != "CELL":
                    continue
                r, c           = cell.get("RowIndex", 0), cell.get("ColumnIndex", 0)
                max_row, max_col = max(max_row, r), max(max_col, c)
                cells[(r, c)]  = get_block_text(cell, block_map)
        table_data = [[cells.get((r, c), "").strip() for c in range(1, max_col + 1)]
                      for r in range(1, max_row + 1)]
        tables.append(table_data)
    return tables


def get_block_text(block, block_map):
    words = []
    for rel in block.get("Relationships", []):
        if rel["Type"] == "CHILD":
            for child_id in rel["Ids"]:
                child = block_map.get(child_id)
                if not child:
                    continue
                if child["BlockType"] == "WORD":
                    words.append(child.get("Text", ""))
                elif child["BlockType"] == "SELECTION_ELEMENT" and child.get("SelectionStatus") == "SELECTED":
                    words.append("X")
    return " ".join(words)


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

    # If tables were found, use ONLY tables (+ key-values for metadata).
    # Sending both LINES and TABLES causes Claude to extract duplicate rows
    # because the same data appears in both sections.
    # HOWEVER: Always include first 30 lines for letterhead/company name extraction.
    parts = []
    if tables:
        # Include first 30 lines for header context (company name, dates, references)
        # This ensures Claude can see the letterhead/logo text even when tables are present
        if lines:
            header_lines = lines[:30]
            parts.append("=== DOCUMENT HEADER (first lines) ===\n" + "\n".join(header_lines))
        parts.append("=== TABLES ===\n" + "\n\n".join(tables))
        if kv_pairs:
            parts.append("=== KEY-VALUES ===\n" + "\n".join(kv_pairs))
    else:
        # Fallback: no tables detected, use lines
        if lines:
            parts.append("=== LINES ===\n" + "\n".join(lines))
        if kv_pairs:
            parts.append("=== KEY-VALUES ===\n" + "\n".join(kv_pairs))
    return "\n\n".join(parts) or "[No content extracted]"


# ── Table parsing helpers ─────────────────────────────────────────────────────
def build_rows_from_table(table):
    if not table or len(table) < 2:
        return []

    header_row = [str(x).strip() for x in table[0]]
    data_rows  = table[1:]

    doc_num_idx = ref_idx = date_idx = inv_date_idx = None
    tds_idx = gross_amt_idx = net_amt_idx = None
    bill_no_idx = inv_no_idx = None  # Track these separately to resolve priority

    for idx, name in enumerate(header_row):
        n = name.lower().strip()
        # ── Track INV.NO and Bill No columns separately for priority resolution ──
        if inv_no_idx is None and ("invoice no" in n or "invoice number" in n
                or "inv no" in n or "inv. no" in n or n == "inv" or "inv.no" in n):
            inv_no_idx = idx
        if bill_no_idx is None and ("bill no" in n or "bill number" in n or n == "bill no" or n == "bill number"):
            bill_no_idx = idx

        if doc_num_idx is None and ("your document" in n or "invoice no" in n or "invoice number" in n
                or "inv no" in n or "inv. no" in n or "inv.no" in n or "bill no" in n or "bill number" in n
                or n == "description" or n == "inv" or "ref. no" in n or "ref no" in n or n == "ref."):
            doc_num_idx = idx
        if ref_idx is None and (n == "reference" or n == "ref" or "pv no" in n):
            ref_idx = idx
        if inv_date_idx is None and ("inv date" in n or "inv. date" in n or "invoice date" in n):
            inv_date_idx = idx
        if date_idx is None and ("bill date" in n or "invoice date" in n or "post date" in n or "date" in n):
            date_idx = idx
        if tds_idx is None and ("tds" in n or "t.d.s" in n):
            tds_idx = idx
        if gross_amt_idx is None and ("gross" in n or "invoice amount" in n or "invoice amt" in n or "bill amount" in n or "bill amt" in n):
            gross_amt_idx = idx
        if net_amt_idx is None and ("total amt" in n
                or ("amount" in n and "gross" not in n and "invoice" not in n and "tds" not in n and "bill" not in n)
                or "net" in n or "total amount" in n or "paid" in n):
            net_amt_idx = idx

    # ── PRIORITY RULE: Bill No first, then INV.NO as fallback ──
    # 1. If Bill No column exists AND contains a 13-digit number (single or split across
    #    rows, starting with 3/4/7) → use Bill No as DOCUMENT_NUMBER.
    # 2. If Bill No column has no valid 13-digit number → fall back to INV.NO column.
    # 3. If neither exists, doc_num_idx stays as whatever was detected above.
    if bill_no_idx is not None:
        # Check if Bill No column has any 13-digit value (including split rows)
        bill_no_has_13 = _column_has_13_digit_value(data_rows, bill_no_idx)
        if bill_no_has_13:
            doc_num_idx = bill_no_idx
            logger.info("build_rows: Bill No column (idx=%d) has 13-digit value — using as DOCUMENT_NUMBER", bill_no_idx)
        elif inv_no_idx is not None:
            doc_num_idx = inv_no_idx
            logger.info("build_rows: Bill No has no 13-digit value — falling back to INV.NO (idx=%d)", inv_no_idx)
        # else: bill_no_idx but no inv_no_idx and no 13-digit — keep whatever doc_num_idx was

    if inv_date_idx is not None:
        date_idx = inv_date_idx

    # ── 13-DIGIT INVOICE NUMBER PRIORITY LOGIC ────────────────────────────────
    # Scan all columns in data rows to find one that has 13-digit numeric values.
    # EXCEPTION: If we already pinned to Bill No (because it has 13-digit values),
    # exclude INV.NO from auto-detection so it doesn't override our decision.
    _exclude_for_13 = None
    if bill_no_idx is not None and doc_num_idx == bill_no_idx and inv_no_idx is not None:
        _exclude_for_13 = inv_no_idx  # Bill No already selected — don't let INV.NO override
    elif bill_no_idx is not None and doc_num_idx == inv_no_idx and inv_no_idx is not None:
        _exclude_for_13 = bill_no_idx  # Fell back to INV.NO — don't let Bill No sneak back in

    thirteen_digit_col = _find_13_digit_invoice_column(
        data_rows, len(header_row), exclude_col=_exclude_for_13
    )
    if thirteen_digit_col is not None:
        # If the 13-digit column was previously detected as ref_idx, clear ref_idx
        if ref_idx == thirteen_digit_col:
            ref_idx = None
        # Only override if we haven't already pinned to Bill No or INV.NO via priority rule
        if bill_no_idx is None and inv_no_idx is None:
            doc_num_idx = thirteen_digit_col
        elif _exclude_for_13 is None or thirteen_digit_col != _exclude_for_13:
            doc_num_idx = thirteen_digit_col
    # ── END 13-DIGIT PRIORITY LOGIC ───────────────────────────────────────────

    if gross_amt_idx is not None and tds_idx is not None and gross_amt_idx == tds_idx:
        next_idx = gross_amt_idx + 1
        if next_idx < len(header_row):
            for row in data_rows:
                val = str(get_value(row, next_idx) or "").strip()
                if val and re.match(r"-?[\d,.]+$", val):
                    tds_idx = next_idx
                    break
            else:
                tds_idx = None

    if doc_num_idx is None:
        for idx, name in enumerate(header_row):
            n = name.lower().strip()
            if n in ("sno", "s.no", "s no", "sr", "sr.", "sr no", "sr. no", "#", "sl", "sl.", "sl no"):
                continue
            if "document" in n or "doc" in n:
                doc_num_idx = idx
                break

    if doc_num_idx is None:
        for idx, name in enumerate(header_row):
            n = name.lower().strip()
            if n in ("sno", "s.no", "s no", "sr", "sr.", "sr no", "sr. no", "#", "sl", "sl.", "sl no"):
                continue
            doc_num_idx = idx
            break

    if doc_num_idx is None:
        doc_num_idx = 0

    if tds_idx is not None and gross_amt_idx is not None and tds_idx == gross_amt_idx + 1:
        net_candidate = tds_idx + 1
        if net_candidate < len(header_row) and net_amt_idx is None:
            for row in data_rows:
                val = str(get_value(row, net_candidate) or "").strip()
                if val and re.match(r"-?[\d,.]+$", val):
                    net_amt_idx = net_candidate
                    break

    has_combined_amounts = False
    if gross_amt_idx is not None and tds_idx is None:
        for row in data_rows:
            val = str(get_value(row, gross_amt_idx) or "").strip()
            if re.match(r"[\d,.]+\s*-[\d,.]+\s+[\d,.]+", val) or re.match(r"[\d,.]+\s+[\d,.]+\s*-\s*[\d,.]+\s+[\d,.]+", val):
                has_combined_amounts = True
                break

    tds_in_rows = False
    if tds_idx is None and not has_combined_amounts:
        for row in data_rows:
            if not any(str(x).strip() for x in row):
                continue
            if net_amt_idx is not None:
                amt_val = get_value(row, net_amt_idx)
                if amt_val and str(amt_val).strip().upper().startswith("TDS"):
                    tds_in_rows = True
                    break
            if doc_num_idx is not None:
                doc_val = get_value(row, doc_num_idx)
                if doc_val and re.search(r"\d+\s*TDS$", str(doc_val).strip(), re.IGNORECASE):
                    tds_in_rows = True
                    break

    if tds_in_rows:
        return _build_rows_with_tds_in_rows(data_rows, ref_idx, doc_num_idx, date_idx, gross_amt_idx, net_amt_idx)

    if has_combined_amounts:
        return _build_rows_with_combined_amounts(data_rows, doc_num_idx, date_idx, gross_amt_idx)

    amt_check_idx = net_amt_idx if net_amt_idx is not None else (gross_amt_idx if gross_amt_idx is not None else (4 if len(header_row) > 4 else None))
    if _detect_paired_rows(data_rows, doc_num_idx, date_idx, amt_check_idx):
        return _build_rows_with_paired_amounts(data_rows, doc_num_idx, date_idx, amt_check_idx)

    if date_idx is None and len(header_row) > 2:
        date_idx = 2
    if tds_idx is None and len(header_row) > 3:
        tds_idx = 3
    if net_amt_idx is None and len(header_row) > 4:
        net_amt_idx = 4

    # ── Pre-compute which rows are "continuation rows" (just 1-3 digits used to
    #    complete a 12-digit invoice in the previous row). These must be skipped. ──
    skip_row_indices = set()
    if thirteen_digit_col is not None:
        for row_idx, row in enumerate(data_rows):
            doc_val = str(get_value(row, doc_num_idx) or "").strip()
            clean_doc = re.sub(r"[\s\n]+", "", doc_val)
            if re.fullmatch(r"\d{12}", clean_doc):
                next_row_idx = row_idx + 1
                if next_row_idx < len(data_rows):
                    next_val = str(get_value(data_rows[next_row_idx], doc_num_idx) or "").strip()
                    next_clean = re.sub(r"[\s\n]+", "", next_val)
                    if re.fullmatch(r"\d{1,3}", next_clean):
                        combined = clean_doc + next_clean
                        if len(combined) == 13:
                            skip_row_indices.add(next_row_idx)

    rows = []
    for row_idx, row in enumerate(data_rows):
        if not any(str(x).strip() for x in row):
            continue

        # Skip continuation rows (rows that are just digit fragments for the previous row)
        if row_idx in skip_row_indices:
            continue

        is_last_row = row_idx == len(data_rows) - 1
        if should_skip_summary_row_full(row, doc_num_idx, date_idx, is_last_row):
            continue

        document_number  = get_value(row, doc_num_idx)
        document_date    = get_value(row, date_idx)
        tds_amount       = get_value(row, tds_idx)
        invoice_amount   = normalize_amount(get_value(row, gross_amt_idx)) if gross_amt_idx is not None else None
        amount_raw       = normalize_amount(get_value(row, net_amt_idx))

        if invoice_amount is None and amount_raw is not None:
            invoice_amount = amount_raw
            amount         = None
        else:
            amount = amount_raw

        # ── 13-DIGIT CONCATENATION: if doc number is 12 digits, check next row ──
        if thirteen_digit_col is not None and document_number is not None:
            document_number = _complete_13_digit_invoice(
                document_number, row_idx, data_rows, doc_num_idx
            )

        if should_skip_summary_row(document_number, None, document_date, amount or invoice_amount):
            continue

        if not any([document_number, document_date, tds_amount, invoice_amount, amount]):
            continue

        rows.append({
            "DOCUMENT_NUMBER":  document_number,
            "DOCUMENT_DATE":    document_date,
            "TDS_AMOUNT":       tds_amount,
            "INVOICE_AMOUNT":   invoice_amount,
            "AMOUNT":           amount,
            "CUST_PAYMENT_ID":  None,
            "DEDUCTION_AMOUNT": None,
            "CASH_DISCOUNT":    None,
        })
    return rows


def _column_has_13_digit_value(data_rows, col_idx):
    """
    Check if a given column contains a 13-digit numeric value starting with 3, 4, or 7.
    Handles both single-cell 13-digit values and split-row values (e.g., 12 digits
    in one row + 1-3 digits in the next row that together form 13 digits).
    Also handles same-cell space-separated fragments.
    """
    for row_idx, row in enumerate(data_rows):
        val = str(get_value(row, col_idx) or "").strip()
        clean_val = re.sub(r"[\s\n]+", "", val)

        # Direct 13-digit match starting with 3/4/7
        if re.fullmatch(r"[347]\d{12}", clean_val):
            return True

        # Same-cell space-separated (e.g., "300801009140 3")
        parts = val.split()
        if len(parts) >= 2:
            first_part = parts[0].strip()
            remaining = "".join(parts[1:]).strip()
            combined = re.sub(r"[\s\n]+", "", first_part + remaining)
            if re.fullmatch(r"[347]\d{12}", combined):
                return True

        # Split across rows: current row has some digits, next row has remainder
        if re.fullmatch(r"\d{5,12}", clean_val) and clean_val[0] in "347":
            next_row_idx = row_idx + 1
            while next_row_idx < len(data_rows):
                next_val = str(get_value(data_rows[next_row_idx], col_idx) or "").strip()
                next_clean = re.sub(r"[\s\n]+", "", next_val)
                if not next_clean:
                    next_row_idx += 1
                    continue
                if re.fullmatch(r"\d+", next_clean):
                    combined = clean_val + next_clean
                    if re.fullmatch(r"[347]\d{12}", combined):
                        return True
                break

    return False


def _find_13_digit_invoice_column(data_rows, num_cols, exclude_col=None):
    """
    Scan all columns in data rows. If any column has values that are exactly
    13 digits (or 12 digits that could be part of a 13-digit number with the
    next row or same-cell space-separated), return that column index.
    Priority: a column where we find at least one 13-digit numeric value wins.
    Also detects 12-digit values with next-row or same-cell continuation.
    Fallback: detects 10-digit numeric columns starting with 3/4/7 (shorter invoice formats).

    exclude_col: column index to skip (used to prevent Bill No from being picked
                 when an INV.NO column already exists).
    """
    best_col = None
    best_count = 0

    for col_idx in range(num_cols):
        if exclude_col is not None and col_idx == exclude_col:
            continue  # Skip explicitly excluded column (e.g. Bill No when INV.NO exists)
        count_13 = 0
        count_12 = 0
        for row_idx, row in enumerate(data_rows):
            val = str(get_value(row, col_idx) or "").strip()
            # Remove spaces/newlines that OCR may insert
            clean_val = re.sub(r"[\s\n]+", "", val)
            if re.fullmatch(r"\d{13}", clean_val):
                count_13 += 1
            elif re.fullmatch(r"\d{12}", clean_val):
                count_12 += 1
                # Check if next row has 1-3 digit continuation -> treat as 13
                if row_idx + 1 < len(data_rows):
                    next_val = str(get_value(data_rows[row_idx + 1], col_idx) or "").strip()
                    next_clean = re.sub(r"[\s\n]+", "", next_val)
                    if re.fullmatch(r"\d{1,3}", next_clean):
                        if len(clean_val + next_clean) == 13:
                            count_13 += 1
                            count_12 -= 1  # It's actually 13 via continuation
            else:
                # Check same-cell space pattern: "304501004223 0" or "300114000333 9"
                parts = val.split()
                if len(parts) >= 2:
                    first_part = parts[0].strip()
                    remaining = "".join(parts[1:]).strip()
                    if re.fullmatch(r"\d{12}", first_part) and re.fullmatch(r"\d{1,3}", remaining):
                        combined = first_part + remaining
                        if len(combined) == 13:
                            count_13 += 1

        # A column qualifies if it has at least one 13-digit value (including via continuation)
        total_qualifying = count_13 + count_12
        if count_13 > 0 and total_qualifying > best_count:
            best_count = total_qualifying
            best_col = col_idx

    # If no column has a 13-digit value even with continuation,
    # check for columns with 12-digit values (likely invoice numbers)
    if best_col is None:
        for col_idx in range(num_cols):
            if exclude_col is not None and col_idx == exclude_col:
                continue
            count_12 = 0
            for row_idx, row in enumerate(data_rows):
                val = str(get_value(row, col_idx) or "").strip()
                clean_val = re.sub(r"[\s\n]+", "", val)
                if re.fullmatch(r"\d{12}", clean_val):
                    count_12 += 1
                else:
                    parts = val.split()
                    if len(parts) >= 2:
                        first_part = parts[0].strip()
                        if re.fullmatch(r"\d{12}", first_part):
                            count_12 += 1
            if count_12 > 0:
                best_col = col_idx
                break

    # ── FALLBACK: Detect 10-digit numeric invoice columns starting with 3/4/7 ──
    # Some vendors use 10-digit invoice numbers (e.g., 3004010421, 7001010266)
    # Prefer pure-numeric columns over alphanumeric reference columns
    if best_col is None:
        best_10_col = None
        best_10_count = 0
        for col_idx in range(num_cols):
            if exclude_col is not None and col_idx == exclude_col:
                continue
            count_10 = 0
            for row_idx, row in enumerate(data_rows):
                val = str(get_value(row, col_idx) or "").strip()
                clean_val = re.sub(r"[\s\n]+", "", val)
                # Match 10-digit numbers starting with 3, 4, or 7
                if re.fullmatch(r"[347]\d{9}", clean_val):
                    count_10 += 1
                # Also match 10-13 digit pure numeric (broader catch)
                elif re.fullmatch(r"[347]\d{9,12}", clean_val):
                    count_10 += 1
            if count_10 >= 2 and count_10 > best_10_count:
                best_10_count = count_10
                best_10_col = col_idx
        if best_10_col is not None:
            best_col = best_10_col

    return best_col


def _complete_13_digit_invoice(document_number, current_row_idx, data_rows, doc_num_idx):
    """
    If the document_number is exactly 12 digits, look at the next row's same column.
    If the next row has a short numeric value (1-3 digits), concatenate it to form
    a 13-digit invoice number.
    Even if the next row has '0', it gets appended.

    Also handles the case where OCR/Textract puts the continuation digit in the
    SAME CELL separated by a space (e.g., "304501004223 0" -> "3045010042230").
    """
    if document_number is None:
        return document_number

    doc_str = str(document_number).strip()

    # CASE 1: Same-cell space-separated continuation
    # Textract often gives "304501004223 0" as a single cell value
    parts = doc_str.split()
    if len(parts) >= 2:
        first_part = parts[0].strip()
        remaining_parts = "".join(parts[1:]).strip()
        if re.fullmatch(r"\d{12}", first_part) and re.fullmatch(r"\d{1,3}", remaining_parts):
            combined = first_part + remaining_parts
            if len(combined) == 13:
                return combined
            elif len(combined) > 13:
                return first_part + remaining_parts[:13 - 12]
        # If all parts together form 13 digits, just concatenate
        all_digits = re.sub(r"[\s\n]+", "", doc_str)
        if re.fullmatch(r"\d{13}", all_digits):
            return all_digits

    # Remove any internal spaces/newlines (OCR artifact within same cell)
    clean_doc = re.sub(r"[\s\n]+", "", doc_str)

    # Already 13 digits, return as-is
    if re.fullmatch(r"\d{13}", clean_doc):
        return clean_doc

    # CASE 2: Next-row continuation
    # If it's 12 digits, check the next row for continuation
    if re.fullmatch(r"\d{12}", clean_doc):
        next_row_idx = current_row_idx + 1
        if next_row_idx < len(data_rows):
            next_row = data_rows[next_row_idx]
            next_val = str(get_value(next_row, doc_num_idx) or "").strip()
            next_clean = re.sub(r"[\s\n]+", "", next_val)
            # If next row value is 1-3 digits (including "0"), concatenate
            if re.fullmatch(r"\d{1,3}", next_clean):
                combined = clean_doc + next_clean
                # Only accept if the result is exactly 13 digits
                if len(combined) == 13:
                    return combined
                # If concatenation gives more than 13, just take enough to make 13
                elif len(combined) > 13:
                    return clean_doc + next_clean[:13 - 12]
        # Could not find continuation, return the 12-digit value as-is
        return clean_doc

    return doc_str


def _detect_paired_rows(data_rows, doc_num_idx, date_idx, amt_idx):
    if doc_num_idx is None or amt_idx is None:
        return False
    non_empty   = [r for r in data_rows if any(str(x).strip() for x in r)]
    if len(non_empty) < 2:
        return False
    pair_count = 0
    i = 0
    while i < len(non_empty) - 1:
        doc1  = str(get_value(non_empty[i],     doc_num_idx) or "").strip()
        doc2  = str(get_value(non_empty[i + 1], doc_num_idx) or "").strip()
        date1 = str(get_value(non_empty[i],     date_idx)    or "").strip() if date_idx is not None else ""
        date2 = str(get_value(non_empty[i + 1], date_idx)    or "").strip() if date_idx is not None else ""
        if doc1 and doc1 == doc2 and date1 == date2:
            amt1 = normalize_amount(get_value(non_empty[i],     amt_idx))
            amt2 = normalize_amount(get_value(non_empty[i + 1], amt_idx))
            if amt1 is not None and amt2 is not None and amt1 != amt2:
                pair_count += 1
                i += 2
                continue
        i += 1
    return pair_count >= 2


def _build_rows_with_paired_amounts(data_rows, doc_num_idx, date_idx, amt_idx):
    rows      = []
    non_empty = [r for r in data_rows if any(str(x).strip() for x in r)]
    i = 0
    while i < len(non_empty):
        doc1  = str(get_value(non_empty[i], doc_num_idx) or "").strip()
        date1 = str(get_value(non_empty[i], date_idx)    or "").strip() if date_idx is not None else ""
        amt1  = normalize_amount(get_value(non_empty[i], amt_idx))

        if i + 1 < len(non_empty):
            doc2  = str(get_value(non_empty[i + 1], doc_num_idx) or "").strip()
            date2 = str(get_value(non_empty[i + 1], date_idx)    or "").strip() if date_idx is not None else ""
            amt2  = normalize_amount(get_value(non_empty[i + 1], amt_idx))
            if doc1 and doc1 == doc2 and date1 == date2 and amt1 is not None and amt2 is not None and amt1 != amt2:
                tds_amount, invoice_amount = (amt1, amt2) if abs(float(amt1)) < abs(float(amt2)) else (amt2, amt1)
                if not should_skip_summary_row(doc1, None, date1, invoice_amount):
                    rows.append({
                        "DOCUMENT_NUMBER":  doc1,
                        "DOCUMENT_DATE":    date1,
                        "TDS_AMOUNT":       tds_amount,
                        "INVOICE_AMOUNT":   invoice_amount,
                        "AMOUNT":           None,
                        "CUST_PAYMENT_ID":  None,
                        "DEDUCTION_AMOUNT": None,
                        "CASH_DISCOUNT":    None,
                    })
                i += 2
                continue

        if doc1 and amt1 is not None and not should_skip_summary_row(doc1, None, date1, amt1):
            rows.append({
                "DOCUMENT_NUMBER":  doc1,
                "DOCUMENT_DATE":    date1,
                "TDS_AMOUNT":       None,
                "INVOICE_AMOUNT":   amt1,
                "AMOUNT":           None,
                "CUST_PAYMENT_ID":  None,
                "DEDUCTION_AMOUNT": None,
                "CASH_DISCOUNT":    None,
            })
        i += 1
    return rows


def _build_rows_with_combined_amounts(data_rows, doc_num_idx, date_idx, amt_idx):
    rows = []
    for row_idx, row in enumerate(data_rows):
        if not any(str(x).strip() for x in row):
            continue
        if should_skip_summary_row_full(row, doc_num_idx, date_idx, row_idx == len(data_rows) - 1):
            continue
        document_number = get_value(row, doc_num_idx)
        document_date   = get_value(row, date_idx)
        combined        = str(get_value(row, amt_idx) or "").strip()
        invoice_amount  = tds_amount = amount = None

        m1 = re.match(r"([\d,.]+)\s+([\d,.]+)\s*-\s*([\d,.]+)\s+([\d,.]+)", combined)
        if m1:
            invoice_amount = normalize_amount(m1.group(1))
            tds_amount     = normalize_amount(m1.group(2))
            amount         = normalize_amount(m1.group(3))
        else:
            m2 = re.match(r"([\d,.]+)\s*(-\s*[\d,.]+)\s+([\d,.]+)", combined)
            if m2:
                invoice_amount = normalize_amount(m2.group(1))
                tds_amount     = normalize_amount(m2.group(2).replace(",", "").replace(" ", ""))
                amount         = normalize_amount(m2.group(3))
            else:
                parts = combined.split()
                if len(parts) >= 3:
                    invoice_amount = normalize_amount(parts[0])
                    for p in parts[1:]:
                        if p.startswith("-") and tds_amount is None:
                            tds_amount = normalize_amount(p)
                        elif tds_amount is not None and amount is None:
                            amount = normalize_amount(p)
                    if amount is None:
                        amount = normalize_amount(parts[-1])
                else:
                    invoice_amount = normalize_amount(combined)
                    amount         = invoice_amount

        if should_skip_summary_row(document_number, None, document_date, amount):
            continue
        if not any([document_number, document_date, invoice_amount, amount]):
            continue

        rows.append({
            "DOCUMENT_NUMBER":  document_number,
            "DOCUMENT_DATE":    document_date,
            "TDS_AMOUNT":       tds_amount,
            "INVOICE_AMOUNT":   invoice_amount,
            "AMOUNT":           amount,
            "CUST_PAYMENT_ID":  None,
            "DEDUCTION_AMOUNT": None,
            "CASH_DISCOUNT":    None,
        })
    return rows


def _build_rows_with_tds_in_rows(data_rows, ref_idx, doc_num_idx, date_idx, gross_amt_idx, net_amt_idx):
    grouped = OrderedDict()
    for row in data_rows:
        if not any(str(x).strip() for x in row):
            continue
        ref_val     = get_value(row, ref_idx) if ref_idx is not None else get_value(row, 0)
        amt_val     = str(get_value(row, net_amt_idx) or "").strip()
        doc_num     = get_value(row, doc_num_idx)
        doc_date    = get_value(row, date_idx)
        doc_num_str = str(doc_num or "").strip()

        is_tds_row  = False
        tds_val     = None
        if amt_val.upper().startswith("TDS"):
            is_tds_row = True
            m          = re.search(r"[Tt][Dd][Ss][\s\-:]*([0-9,.]+)", amt_val)
            tds_val    = m.group(1).replace(",", "") if m else None
        elif re.search(r"\d+\s*TDS$", doc_num_str, re.IGNORECASE):
            is_tds_row = True
            tds_val    = amt_val.replace(",", "") if amt_val else None

        if tds_val and tds_val.startswith("-"):
            tds_val = tds_val[1:]

        if is_tds_row:
            clean_doc = re.sub(r"\s*TDS$", "", doc_num_str, flags=re.IGNORECASE).strip() if doc_num_str else None
            key = clean_doc or ref_val
            if key and key in grouped:
                grouped[key]["TDS_AMOUNT"] = tds_val
            elif key:
                grouped[key] = {
                    "DOCUMENT_NUMBER":  clean_doc,
                    "DOCUMENT_DATE":    None,
                    "TDS_AMOUNT":       tds_val,
                    "INVOICE_AMOUNT":   None,
                    "AMOUNT":           None,
                    "CUST_PAYMENT_ID":  None,
                    "DEDUCTION_AMOUNT": None,
                    "CASH_DISCOUNT":    None,
                }
        else:
            amount         = normalize_amount(amt_val)
            invoice_amount = amount if gross_amt_idx is None else normalize_amount(get_value(row, gross_amt_idx))
            key = doc_num_str or ref_val
            if should_skip_summary_row(doc_num, None, doc_date, amount):
                continue
            if key and key in grouped:
                existing = grouped[key]
                if invoice_amount and not existing.get("INVOICE_AMOUNT"):
                    existing["INVOICE_AMOUNT"] = invoice_amount
                if doc_date and not existing.get("DOCUMENT_DATE"):
                    existing["DOCUMENT_DATE"] = doc_date
            elif key:
                grouped[key] = {
                    "DOCUMENT_NUMBER":  doc_num,
                    "DOCUMENT_DATE":    doc_date,
                    "TDS_AMOUNT":       None,
                    "INVOICE_AMOUNT":   invoice_amount,
                    "AMOUNT":           None,
                    "CUST_PAYMENT_ID":  None,
                    "DEDUCTION_AMOUNT": None,
                    "CASH_DISCOUNT":    None,
                }
    return list(grouped.values())


def calculate_net_amounts(rows):
    for row in rows:
        if row.get("AMOUNT") is None and row.get("INVOICE_AMOUNT") is not None and row.get("TDS_AMOUNT") is not None:
            try:
                inv_val = float(row["INVOICE_AMOUNT"])
                tds_val = float(row["TDS_AMOUNT"])
                if tds_val > 0:
                    tds_val = -tds_val
                net_val    = inv_val + tds_val
                row["AMOUNT"] = int(net_val) if net_val == int(net_val) else net_val
            except (ValueError, TypeError):
                pass
    return rows


def merge_tds_rows(rows):
    has_tds_suffix = any(
        r.get("DOCUMENT_NUMBER") and re.search(r"\d+\s*TDS$", str(r["DOCUMENT_NUMBER"]).strip(), re.IGNORECASE)
        for r in rows
    )
    doc_counts = {}
    for r in rows:
        dn = str(r.get("DOCUMENT_NUMBER") or "").strip()
        if dn:
            doc_counts[dn] = doc_counts.get(dn, 0) + 1

    if not has_tds_suffix and not any(v > 1 for v in doc_counts.values()):
        return rows

    grouped = OrderedDict()
    for row in rows:
        doc_num = str(row.get("DOCUMENT_NUMBER") or "").strip()
        if re.search(r"\d+\s*TDS$", doc_num, re.IGNORECASE):
            clean_doc = re.sub(r"\s*TDS$", "", doc_num, flags=re.IGNORECASE).strip()
            tds_val   = row.get("AMOUNT")
            if tds_val is not None and isinstance(tds_val, (int, float)) and tds_val < 0:
                tds_val = abs(tds_val)
            if clean_doc in grouped:
                grouped[clean_doc]["TDS_AMOUNT"] = tds_val
            else:
                grouped[clean_doc] = {
                    "DOCUMENT_NUMBER":  clean_doc,
                    "DOCUMENT_DATE":    None,
                    "TDS_AMOUNT":       tds_val,
                    "INVOICE_AMOUNT":   row.get("INVOICE_AMOUNT"),
                    "AMOUNT":           None,
                    "CUST_PAYMENT_ID":  row.get("CUST_PAYMENT_ID"),
                    "DEDUCTION_AMOUNT": row.get("DEDUCTION_AMOUNT"),
                    "CASH_DISCOUNT":    row.get("CASH_DISCOUNT"),
                }
        else:
            key = doc_num
            if key in grouped:
                existing = grouped[key]
                if row.get("AMOUNT") is not None and not existing.get("AMOUNT"):
                    existing["AMOUNT"] = row["AMOUNT"]
                if row.get("TDS_AMOUNT") is not None and not existing.get("TDS_AMOUNT"):
                    existing["TDS_AMOUNT"] = row["TDS_AMOUNT"]
                if row.get("AMOUNT") is not None and row.get("DOCUMENT_DATE"):
                    existing["DOCUMENT_DATE"] = row["DOCUMENT_DATE"]
                if row.get("INVOICE_AMOUNT") and not existing.get("INVOICE_AMOUNT"):
                    existing["INVOICE_AMOUNT"] = row["INVOICE_AMOUNT"]
            else:
                grouped[key] = dict(row)
    return list(grouped.values())


def should_skip_summary_row(document_number, invoice_number, invoice_date, invoice_amount):
    doc_text = to_text(document_number).lower()
    combined = " ".join([
        doc_text, to_text(invoice_number).lower(),
        to_text(invoice_date).lower(), to_text(invoice_amount).lower()
    ]).strip()
    for word in ["net amount", "total", "grand total", "balance", "closing balance",
                 "opening balance", "subtotal", "total :", "total:",
                 "balance carried forward", "balance b/f", "balance c/f",
                 "carried forward", "page total", "sum total", "forward"]:
        if word in combined:
            return True
    if doc_text.strip().rstrip(":").strip() in ("total", "grand total", "subtotal", "net amount", "forward"):
        return True
    if document_number is not None and not looks_like_document_number(document_number):
        if any(ch.isalpha() for ch in doc_text):
            return True
    return False


def should_skip_summary_row_full(row, doc_num_idx, date_idx, is_last_row=False):
    doc_val = get_value(row, doc_num_idx) if doc_num_idx is not None else None
    if is_last_row and doc_val is None:
        return True
    for cell in row:
        cell_lower = str(cell).lower().strip().rstrip(":").strip()
        if cell_lower in (
            "total", "grand total", "subtotal", "net amount", "sum",
            "balance carried forward", "balance b/f", "balance c/f",
            "carried forward", "page total"
        ):
            return True
        # Check if cell contains these keywords as part of longer text
        if any(kw in cell_lower for kw in ["balance carried forward", "carried forward", "balance b/f"]):
            return True
    return False


def looks_like_document_number(value):
    text = to_text(value)
    if not text:
        return False
    v = text.replace(" ", "")
    if re.fullmatch(r"[0-9]{6,}", v):
        return True
    if re.fullmatch(r"[A-Za-z0-9/\-]{6,}", v) and any(ch.isdigit() for ch in v):
        return True
    return False


def normalize_document_number(value):
    if value is None:
        return None
    text = str(value).strip()
    if not text:
        return None
    # Strip ERS- prefix if present
    if re.match(r"^ERS[-\s]?", text, re.IGNORECASE):
        text = re.sub(r"^ERS[-\s]?", "", text, flags=re.IGNORECASE).strip()
    no_spaces = re.sub(r"[\s\n]+", "", text)
    return no_spaces if re.fullmatch(r"\d+", no_spaces) else text


def normalize_amount(value):
    if value is None:
        return None
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        num = float(value)
        return int(num) if num.is_integer() else num
    text = str(value).strip().replace(",", "").replace(" ", "").replace("\u2212", "-")
    is_neg = text.startswith("-")
    if text.endswith("-"):
        text = text[:-1]
    if is_neg and not text.startswith("-"):
        text = "-" + text
    m = re.search(r"^-?\d+(?:\.\d+)?", text)
    if not m:
        return None
    num    = m.group(0)
    result = float(num)
    return int(result) if result.is_integer() else result


def to_text(value):
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


def get_value(row, idx):
    if idx is None or idx < 0 or idx >= len(row):
        return None
    v = row[idx]
    return None if (v is None or v == "") else v


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
        "CUST_PAYMENT_ID", "SOURCE_TYPE", "IMPORT_REFERENCE", "UTR_REFERENCE_NUMBER",
        "CUSTOMER_NAME", "PAYMENT_DATE", "PAYMENT_AMOUNT", "DOCUMENT_NUMBER",
        "DOCUMENT_DATE", "INVOICE_AMOUNT", "TDS_AMOUNT", "DEDUCTION_AMOUNT",
        "CASH_DISCOUNT", "AMOUNT", "KOD_CUST_CODE", "MAIL_ID", "MAIL_RECEIVED_DATE",
    ]

    table = dynamodb.Table(table_name)
    with table.batch_writer() as batch:
        for idx, row in enumerate(rows, start=1):
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
