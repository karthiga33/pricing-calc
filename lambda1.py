import json
import os
import re
import html
import io
import csv
import base64
import urllib.request
import urllib.parse
import boto3
from botocore.exceptions import ClientError
from datetime import datetime, timezone, timedelta

try:
    import openpyxl
except ImportError:
    openpyxl = None  # multi-customer xlsx split disabled if layer not present

s3 = boto3.client("s3")
sqs = boto3.client("sqs")
dynamodb = boto3.resource("dynamodb")
secrets_client = boto3.client("secretsmanager", region_name=os.environ.get("AWS_REGION"))
bedrock_runtime = boto3.client("bedrock-runtime", region_name=os.environ.get("BEDROCK_REGION", "us-east-1"))

# ── Rejection logging table (same table used by lambda_full) ──
REJECTION_TABLE_NAME = "rejected_files"


def _log_rejection_lambda1(file_name, reason, mail_id="", mail_received_date="", subject="", from_address=""):
    """Log rejected/skipped file to DynamoDB for visibility in the application."""
    try:
        table = dynamodb.Table(REJECTION_TABLE_NAME)
        table.put_item(Item={
            "file_name": file_name,
            "rejected_at": datetime.now(timezone.utc).isoformat(),
            "input_key": f"emails/{file_name}",
            "etag": "",
            "reason": reason,
            "source": "lambda1_email_processing",
            "mail_id": mail_id or "",
            "mail_received_date": mail_received_date or "",
            "subject": subject or "",
            "from_address": from_address or "",
        })
        print(f"  📝 Logged rejection: {file_name} — {reason}")
    except Exception as e:
        print(f"  ⚠️ Failed to log rejection for {file_name}: {e}")


def _save_to_rejected_folder(bucket_name, file_name, file_bytes, content_type="application/octet-stream"):
    """Save a rejected/skipped attachment file to rejected/ folder in S3."""
    try:
        rejected_key = f"rejected/{file_name}"
        s3.put_object(
            Bucket=bucket_name,
            Key=rejected_key,
            Body=file_bytes,
            ContentType=content_type,
        )
        print(f"  📁 Saved rejected file: s3://{bucket_name}/{rejected_key}")
    except Exception as e:
        print(f"  ⚠️ Failed to save rejected file {file_name}: {e}")

# Same model used in your other Lambda pipeline (Textract/Bedrock remittance extraction)
MODEL_ID = "global.anthropic.claude-sonnet-4-6"

STATE_KEY = "state/last_processed.txt"

# ── Payment advice keywords for body detection ──
PAYMENT_ADVICE_KEYWORDS = [
    "payment advice", "remittance advice", "remittance notification",
    "payment notification", "bank transfer", "fund transfer",
    "amount paid", "invoice payment", "payment confirmation",
    "paid amount", "transaction reference", "UTR", "NEFT", "RTGS", "IMPS",
    # Additional common variants
    "remittance", "payment voucher", "pay voucher", "cheque", "check payment",
    "debit note", "credit note", "settlement", "wire transfer", "electronic transfer",
    "eft", "tds", "tax deducted", "net payment", "gross amount", "invoice no",
    "invoice number", "document number", "voucher no", "ref no", "reference no",
]

# ── Allowed attachment file extensions ──
ALLOWED_ATTACHMENT_EXTENSIONS = {
    ".pdf", ".xlsx", ".xls", ".csv", ".txt", ".html", ".htm",
    ".jpg", ".jpeg", ".png",
    ".doc", ".docx",
}

# ── Stage 2 content validation ONLY applies to these text-based formats ──
# PDF, Excel, Images, Word → skip Stage 2, always pass through if extension is valid
# TXT, CSV, HTML → apply Stage 2 keyword scan (these are easy to read as text)
STAGE2_CHECK_EXTENSIONS = {".txt", ".csv", ".html", ".htm"}

# ── Extensions for which we attempt multi-customer (Custno/Custname) table splitting ──
MULTI_CUSTOMER_SPLIT_EXTENSIONS = {".xlsx", ".csv", ".txt", ".html", ".htm"}

# ── Keywords for Stage 2 text/CSV/HTML body scan ──
PAYMENT_KEYWORDS_BROAD = [
    # Standard payment terms
    "payment", "remittance", "invoice", "amount", "voucher",
    "cheque", "check", "transfer", "utr", "neft", "rtgs", "imps",
    "tds", "tax deducted", "net amount", "gross amount",
    "document no", "document number", "invoice no", "invoice number",
    "ref no", "reference", "bank", "debit", "credit",
    "settlement", "advice", "pay", "paid",
    # Indian business terms
    "gst", "pan", "cgst", "sgst", "igst", "hsn",
    "vendor", "customer", "supplier", "buyer",
    # Amount indicators
    "rs.", "inr", "rupees", "₹",
    # Multi-customer ledger keywords
    "custno", "cust no", "cust name", "custname", "customer name",
    "trx number", "txn date", "outstanding", "applied amt",
    "party name", "party code", "account name",
]

# Minimum keyword matches for Stage 2 text scan
MIN_KEYWORD_MATCHES = 2

# ── Password map for password-protected PDFs ──
# Passwords are managed via frontend and stored in DynamoDB table "pdf_password_map"
# Key: email subject (or part of it), Value: password
PDF_PASSWORD_MAP = {}  # No static fallback — all passwords come from DynamoDB

# DynamoDB table for password map (managed via frontend)
PDF_PASSWORD_TABLE = "pdf_password_map"


def is_pdf_encrypted(pdf_bytes):
    """Check if a PDF requires a USER password to open.
    Returns False for PDFs with only owner-password (permissions restrictions)
    since those can be opened and read without a password."""
    try:
        from PyPDF2 import PdfReader
        reader = PdfReader(io.BytesIO(pdf_bytes))
        if not reader.is_encrypted:
            return False
        # Try to decrypt with empty password — works for owner-only encrypted PDFs
        try:
            result = reader.decrypt("")
            # result > 0 means decryption succeeded with empty password
            if result > 0:
                print("  ℹ️ PDF has owner-password only (no user password needed) — treating as unencrypted")
                return False
        except Exception:
            pass
        return True
    except ImportError:
        pass
    try:
        import pikepdf
        try:
            pikepdf.open(io.BytesIO(pdf_bytes))
            return False
        except pikepdf._core.PasswordError:
            return True
    except ImportError:
        pass
    # If neither library is available, check PDF header bytes for /Encrypt
    return b"/Encrypt" in pdf_bytes[:10000]


def decrypt_pdf_bytes(pdf_bytes, password):
    """Decrypt a password-protected PDF and return decrypted bytes."""
    try:
        import pikepdf
    except ImportError:
        try:
            from PyPDF2 import PdfReader, PdfWriter
            reader = PdfReader(io.BytesIO(pdf_bytes))
            if reader.is_encrypted:
                reader.decrypt(password)
            writer = PdfWriter()
            for page in reader.pages:
                writer.add_page(page)
            output = io.BytesIO()
            writer.write(output)
            output.seek(0)
            print(f"  🔓 PDF decrypted successfully using PyPDF2 (password: {password})")
            return output.read()
        except ImportError:
            print("  ❌ Neither pikepdf nor PyPDF2 available — cannot decrypt PDF")
            return None
        except Exception as e:
            print(f"  ❌ PyPDF2 decrypt failed: {e}")
            return None

    # pikepdf path (preferred)
    try:
        pdf_in = pikepdf.open(io.BytesIO(pdf_bytes), password=password)
        output = io.BytesIO()
        pdf_in.save(output)
        pdf_in.close()
        output.seek(0)
        print(f"  🔓 PDF decrypted successfully using pikepdf (password: {password})")
        return output.read()
    except Exception as e:
        print(f"  ❌ pikepdf decrypt failed: {e}")
        return None


def find_password_by_subject(email_subject):
    """Look up password from DynamoDB by email subject.
    The table stores subject (or keyword from subject) as partition key.
    We do a scan and check if any stored key is contained in the email subject."""
    if not email_subject:
        return None

    subject_lower = email_subject.lower().strip()

    try:
        table = dynamodb.Table(PDF_PASSWORD_TABLE)
        scan_resp = table.scan()
        for item in scan_resp.get("Items", []):
            # Try both 'subject' and 'file_name' as key column (supports both table schemas)
            key = item.get("subject", "") or item.get("file_name", "")
            key = str(key).strip()
            pwd = item.get("password", "")
            if not key or not pwd:
                continue
            # Check if the stored key matches the subject (case-insensitive)
            key_lower = key.lower()
            if key_lower in subject_lower or subject_lower in key_lower:
                print(f"  🔑 Password found in DynamoDB (subject match: '{key[:50]}' ↔ '{email_subject[:50]}')")
                return pwd
    except Exception as e:
        print(f"  ⚠️ DynamoDB password lookup by subject failed: {e}")

    return None


def get_secret(secret_name):
    try:
        response = secrets_client.get_secret_value(SecretId=secret_name)
        return json.loads(response["SecretString"])
    except ClientError as e:
        raise Exception(f"Failed to retrieve secret '{secret_name}': {e}")


def get_last_processed_time(bucket_name):
    try:
        response = s3.get_object(Bucket=bucket_name, Key=STATE_KEY)
        timestamp = response["Body"].read().decode("utf-8").strip()
        print(f"  Last processed time: {timestamp}")
        return timestamp
    except Exception:
        print("  No state file found — this is the first run.")
        return None


def save_last_processed_time(bucket_name, timestamp):
    s3.put_object(
        Bucket=bucket_name,
        Key=STATE_KEY,
        Body=timestamp.encode("utf-8"),
        ContentType="text/plain",
    )
    print(f"  ✅ State saved: {timestamp}")


def increment_timestamp(ts):
    """Add 1 second to avoid re-fetching the last processed email on next run."""
    if "." in ts:
        ts = ts.split(".")[0] + "Z"
    dt = datetime.strptime(ts, "%Y-%m-%dT%H:%M:%SZ").replace(tzinfo=timezone.utc)
    dt = dt + timedelta(seconds=1)
    return dt.strftime("%Y-%m-%dT%H:%M:%SZ")


def utc_to_ist_display(raw_datetime):
    """Convert UTC ISO timestamp (from Graph API) to IST readable string for metadata display."""
    try:
        if raw_datetime:
            dt_utc = datetime.strptime(raw_datetime.replace("Z", ""), "%Y-%m-%dT%H:%M:%S").replace(tzinfo=timezone.utc)
            dt_ist = dt_utc + timedelta(hours=5, minutes=30)
            return dt_ist.strftime("%Y-%m-%d %H:%M:%S")
        return ""
    except Exception:
        return raw_datetime


def html_to_text_preserve_lines(html_body):
    """
    Convert HTML to plain text while preserving table structure.

    For HTML tables: each <tr> becomes one line with cells separated by " | "
    so a table with 35 rows produces exactly 35 lines (not 35×N lines).
    This prevents token overflow when Claude processes large payment advice tables.

    For non-table HTML: falls back to simple line-break preservation.
    """
    if not html_body:
        return ""

    # ── Try structured table extraction first ──
    # If the HTML contains <table> tags, extract each row as tab-separated cells
    if re.search(r"(?i)<table", html_body):
        try:
            output_lines = []

            # ── Robust row extraction: split on <tr> openers ──
            # This handles missing </tr> closing tags and deeply nested tables.
            # Strategy: split the HTML on every <tr...> opener, then extract
            # <td>/<th> cells from each chunk until the next <tr> or </table>.
            tr_split = re.split(r"(?i)<tr[^>]*>", html_body)

            for chunk in tr_split[1:]:  # skip everything before first <tr>
                # Stop at </table> or next table boundary
                chunk_end = re.search(r"(?i)</table\s*>", chunk)
                if chunk_end:
                    chunk = chunk[:chunk_end.start()]

                cell_pattern = re.compile(r"(?is)<t[dh][^>]*>(.*?)</t[dh]\s*>")
                cells = []
                for cell_match in cell_pattern.finditer(chunk):
                    cell_html = cell_match.group(1)
                    cell_text = re.sub(r"<[^>]+>", " ", cell_html)
                    cell_text = html.unescape(cell_text)
                    cell_text = re.sub(r"\s+", " ", cell_text).strip()
                    cells.append(cell_text)

                if cells:
                    row_line = "\t".join(cells)
                    if row_line.strip():
                        output_lines.append(row_line)

            if output_lines:
                # ── Second pass: extract S| pattern rows from raw HTML ──
                # Some emails have the remittance data inside a single <td> cell
                # with line breaks, or in <p> tags outside proper <tr> rows.
                # Extract ALL S| rows directly from the stripped HTML text.
                full_stripped = re.sub(r"<[^>]+>", "\n", html_body)
                full_stripped = html.unescape(full_stripped)
                # Split by S| pattern to find all invoice rows
                s_pattern_rows = re.findall(
                    r"(S\|[0-9]{10,13}\|[0-9]+\|[^\n]*?(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\s+\d{1,2},?\s+\d{4}[^\n]*?\d+[\.,]\d{2})",
                    full_stripped
                )

                # Check if we have S| rows already in output_lines
                existing_s_rows = set()
                for line in output_lines:
                    m = re.search(r"(S\|\d{10,13}\|\d+\|)", line)
                    if m:
                        existing_s_rows.add(m.group(1))

                # Count how many S| rows the raw text has vs what we extracted
                raw_s_count = full_stripped.count("S|")
                extracted_s_count = len(existing_s_rows)

                if raw_s_count > extracted_s_count * 1.5 and raw_s_count > extracted_s_count + 10:
                    # The HTML table only had a fraction of the rows — the rest
                    # are in a different structure. Fall through to the fallback
                    # which handles this format better by splitting on S| pattern.
                    print(f"  ℹ️ html_to_text: table had only {extracted_s_count} S| rows but raw has {raw_s_count} — using raw S| extraction")
                    # Build structured output from raw text by splitting on S|
                    raw_lines = full_stripped.split("\n")
                    raw_clean = [l.strip() for l in raw_lines if l.strip()]
                    # Rebuild: find header metadata + all S| rows in sequence
                    header_lines = []
                    data_lines = []
                    in_data = False
                    combined_text = "\n".join(raw_clean)
                    # Re-split the entire text by finding S| boundaries
                    parts = re.split(r"(?=S\|\d{10,})", combined_text)
                    for part in parts:
                        part = part.strip()
                        if part.startswith("S|"):
                            data_lines.append(part)
                        elif not in_data and part:
                            header_lines.append(part)
                        if part.startswith("S|"):
                            in_data = True

                    output_lines = header_lines + data_lines

                # Also extract any text outside tables (headers, metadata, etc.)
                non_table = re.sub(r"(?is)<table[^>]*>.*?</table\s*>", "\n", html_body)
                non_table = re.sub(r"(?i)<br\s*/?>", "\n", non_table)
                non_table = re.sub(r"(?i)</p\s*>", "\n", non_table)
                non_table = re.sub(r"<[^>]+>", " ", non_table)
                non_table = html.unescape(non_table)
                non_table = re.sub(r"[ \t]{2,}", " ", non_table)
                non_table_lines = [l.strip() for l in non_table.splitlines() if l.strip()]

                # ── Dedup: if the table rows already contain invoice/remittance data
                # (S|... pattern or pipe-delimited rows), drop any non-table lines
                # that duplicate that data to avoid sending both broken and clean
                # versions to Claude. Keep only header/metadata lines.
                table_has_invoice_data = any(
                    re.match(r"S\|", line.strip()) or re.search(r"\d{10,13}", line)
                    for line in output_lines
                )
                if table_has_invoice_data:
                    # Filter out non-table lines that look like invoice/amount data
                    # Keep only short metadata lines (company names, dates, labels)
                    filtered_non_table = []
                    for line in non_table_lines:
                        # Skip lines that look like invoice rows or raw data duplicates
                        if re.match(r"S\|", line.strip()):
                            continue
                        if re.search(r"\d{10,13}", line):
                            continue
                        # Skip lines that are just amounts/dates (pure noise from plain-text section)
                        if re.fullmatch(r"[-\d,\.]+", line.strip()):
                            continue
                        if re.fullmatch(r"(INR|USD|EUR|GBP)", line.strip()):
                            continue
                        filtered_non_table.append(line)
                    non_table_lines = filtered_non_table

                # Combine: non-table header lines first, then structured table rows
                combined = non_table_lines + [""] + output_lines if non_table_lines else output_lines
                result = "\n".join(combined)
                result = re.sub(r"\n{3,}", "\n\n", result)
                result = result.strip()

                # ── Trim everything before "Payment Remittance Advice" if present ──
                # The email may have a noisy alert/disclaimer section before the
                # actual payment content. Find the real start and cut above it.
                trim_markers = [
                    "Payment Remittance Advice",
                    "payment remittance advice",
                    "PAYMENT REMITTANCE ADVICE",
                    "Payment Advice",
                    "PAYMENT ADVICE",
                    "Remittance Advice",
                    "REMITTANCE ADVICE",
                ]
                for marker in trim_markers:
                    idx = result.find(marker)
                    if idx > 0:
                        result = result[idx:]
                        break

                return result
        except Exception:
            pass  # Fall through to simple conversion if anything goes wrong

    # ── Fallback: simple line-break preservation for non-table HTML ──
    text = html_body
    text = re.sub(r"(?i)<br\s*/?>", "\n", text)
    text = re.sub(r"(?i)</p\s*>", "\n", text)
    text = re.sub(r"(?i)</div\s*>", "\n", text)
    text = re.sub(r"(?i)</tr\s*>", "\n", text)
    text = re.sub(r"(?i)</li\s*>", "\n", text)
    text = re.sub(r"(?i)<tr[^>]*>", "\n", text)
    text = re.sub(r"(?i)</td\s*>", " | ", text)
    text = re.sub(r"<[^>]+>", " ", text)
    text = html.unescape(text)
    text = re.sub(r" \| \s*\n", "\n", text)  # clean trailing pipe before newline
    text = re.sub(r"[ \t]{2,}", " ", text)
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text.strip()


def strip_quoted_reply_trail(body_text, body_type, source_name="email"):
    """
    Trims a reply/forward email body down to only the newest message,
    dropping any older quoted mail trail beneath it.
    """
    if not body_text:
        return body_text

    if body_type == "html":
        marker_match = re.search(r'(?is)<div[^>]+id=["\']divRplyFwdMsg["\'][^>]*>', body_text)
        if marker_match:
            trimmed = body_text[:marker_match.start()]
            print(f"  ✂️ '{source_name}': trimmed quoted reply trail at divRplyFwdMsg marker")
            return trimmed

        hr_match = re.search(r'(?is)<hr[^>]*>', body_text)
        if hr_match:
            # Only trim at <hr> if it looks like a reply separator —
            # i.e., the text after the <hr> contains From/Sent/To/Subject
            # pattern typical of a forwarded/replied email.
            # If the <hr> is inside a payment table, do NOT trim.
            after_hr = body_text[hr_match.end():hr_match.end() + 2000]
            after_hr_plain = re.sub(r"<[^>]+>", " ", after_hr)
            is_reply_separator = (
                re.search(r'(?i)\bFrom\s*:', after_hr_plain)
                and re.search(r'(?i)\bSent\s*:', after_hr_plain)
                and re.search(r'(?i)\bTo\s*:', after_hr_plain)
            )
            if is_reply_separator:
                trimmed = body_text[:hr_match.start()]
                print(f"  ✂️ '{source_name}': trimmed quoted reply trail at <hr> marker (reply pattern confirmed)")
                return trimmed
            else:
                print(f"  ℹ️ '{source_name}': <hr> found but NOT a reply separator — keeping full body")

    earliest_idx = None
    for m in re.finditer(r'(?i)\bFrom\s*:', body_text):
        window = body_text[m.start(): m.start() + 2000]
        if (re.search(r'(?i)\bSent\s*:', window)
                and re.search(r'(?i)\bTo\s*:', window)
                and re.search(r'(?i)\bSubject\s*:', window)):
            earliest_idx = m.start()
            break

    other_patterns = [
        r'(?i)-{3,}\s*Original Message\s*-{3,}',
        r'(?i)-{3,}\s*Forwarded Message\s*-{3,}',
        r'(?im)^On\s.+?\swrote:\s*$',
    ]
    for pat in other_patterns:
        m = re.search(pat, body_text)
        if m and (earliest_idx is None or m.start() < earliest_idx):
            earliest_idx = m.start()

    if earliest_idx is not None:
        trimmed = body_text[:earliest_idx]
        print(f"  ✂️ '{source_name}': trimmed quoted reply trail at offset {earliest_idx}")
        return trimmed

    return body_text


def sanitize_s3_key(name):
    """Make a filename safe for S3 keys."""
    if "." in name:
        base, ext = name.rsplit(".", 1)
    else:
        base, ext = name, ""
    base = re.sub(r"[^\w\-]", "-", base)
    base = re.sub(r"-{2,}", "-", base)
    base = base.strip("-")
    return f"{base}.{ext}" if ext else base


def post_form(url, data):
    encoded_data = urllib.parse.urlencode(data).encode("utf-8")
    req = urllib.request.Request(url, data=encoded_data, method="POST")
    req.add_header("Content-Type", "application/x-www-form-urlencoded")
    try:
        with urllib.request.urlopen(req, timeout=10) as response:
            return json.loads(response.read().decode("utf-8"))
    except urllib.error.HTTPError as e:
        error_body = e.read().decode("utf-8")
        raise Exception(f"Token request failed: {error_body}")


def get_json(url, headers):
    req = urllib.request.Request(url, method="GET")
    for k, v in headers.items():
        req.add_header(k, v)
    try:
        with urllib.request.urlopen(req, timeout=10) as response:
            return json.loads(response.read().decode("utf-8"))
    except urllib.error.HTTPError as e:
        error_body = e.read().decode("utf-8")
        raise Exception(f"Graph API error {e.code}: {error_body}")


def check_already_processed(bucket_name, s3_safe_name):
    """Check if attachment already exists in emails/ folder."""
    try:
        s3.head_object(Bucket=bucket_name, Key=f"emails/{s3_safe_name}")
        return True
    except Exception:
        return False


def _make_unique_s3_name(bucket_name, s3_safe_name):
    """
    If a file with this name already exists in emails/, append an incrementing
    number to make it unique. E.g.:
      Payment-Advice.pdf → Payment-Advice.pdf (if not exists)
      Payment-Advice.pdf → Payment-Advice-1.pdf (if original exists)
      Payment-Advice.pdf → Payment-Advice-2.pdf (if -1 also exists)
    This prevents overwriting when same-named attachments arrive from different emails.
    """
    try:
        s3.head_object(Bucket=bucket_name, Key=f"emails/{s3_safe_name}")
    except Exception:
        # File doesn't exist — use original name
        return s3_safe_name

    # File exists — find a unique name
    if "." in s3_safe_name:
        base, ext = s3_safe_name.rsplit(".", 1)
        ext = "." + ext
    else:
        base, ext = s3_safe_name, ""

    counter = 1
    while counter < 1000:  # Safety limit
        candidate = f"{base}-{counter}{ext}"
        try:
            s3.head_object(Bucket=bucket_name, Key=f"emails/{candidate}")
            counter += 1
        except Exception:
            print(f"  🔄 File '{s3_safe_name}' already exists — renamed to '{candidate}'")
            return candidate

    # Fallback: use timestamp
    ts = datetime.now().strftime("%Y%m%d%H%M%S")
    fallback = f"{base}-{ts}{ext}"
    print(f"  🔄 File '{s3_safe_name}' counter exhausted — using timestamp: '{fallback}'")
    return fallback


def is_payment_advice_email(body_text):
    """Return True if the email body looks like a payment advice.
    Uses word-boundary matching to avoid false positives from substrings
    (e.g. 'eft' inside 'left', 'pay' inside 'repay').
    """
    if not body_text:
        return False
    lower = body_text.lower()
    matched = []
    for kw in PAYMENT_ADVICE_KEYWORDS:
        # Use word-boundary regex so 'eft' matches only standalone "eft"
        # and not 'left', 'theft', etc.
        pattern = r'\b' + re.escape(kw.lower()) + r'\b'
        if re.search(pattern, lower):
            matched.append(kw)
    if matched:
        print(f"  🔍 Payment advice keywords found in body: {matched[:5]}")
        return True
    return False


def has_valid_extension(att_name):
    """Stage 1 gate: reject hard binary formats immediately."""
    lower_name = att_name.lower()
    ext        = "." + lower_name.rsplit(".", 1)[-1] if "." in lower_name else ""
    if ext in ALLOWED_ATTACHMENT_EXTENSIONS:
        print(f"  ✅ Extension allowed: '{att_name}' (ext='{ext}')")
        return True
    print(f"  🚫 Extension not allowed: '{att_name}' (ext='{ext}') — rejected immediately")
    return False


def content_has_payment_fields(raw_bytes, att_name):
    """
    Stage 2 gate — ONLY applied to TXT, CSV, HTML files.
    PDF, Excel, Images, DOC/DOCX → always return True (pass through).
    """
    lower_name = att_name.lower()
    ext        = "." + lower_name.rsplit(".", 1)[-1] if "." in lower_name else ""

    if ext not in STAGE2_CHECK_EXTENSIONS:
        print(f"  ✅ Stage 2 skipped for '{att_name}' (ext='{ext}') — passed to Lambda 2")
        return True

    text = ""
    try:
        if ext in (".html", ".htm"):
            try:
                text = raw_bytes.decode("utf-8")
            except UnicodeDecodeError:
                text = raw_bytes.decode("latin-1", errors="replace")
            text = re.sub(r"<[^>]+>", " ", text)
            text = re.sub(r"[ \t]{2,}", " ", text).strip()
            print(f"  🌐 HTML tags stripped: {len(text)} chars from '{att_name}'")
        else:
            try:
                text = raw_bytes.decode("utf-8")
            except UnicodeDecodeError:
                text = raw_bytes.decode("latin-1", errors="replace")
    except Exception as e:
        print(f"  ⚠️ Could not decode '{att_name}': {e} — passing through anyway")
        return True

    text_lower = text.lower()
    matched = []
    for kw in PAYMENT_KEYWORDS_BROAD:
        pattern = r'\b' + re.escape(kw.lower()) + r'\b'
        if re.search(pattern, text_lower):
            matched.append(kw)
    print(f"  🔎 Keyword scan '{att_name}': {len(matched)} keywords matched → {matched[:5]}")

    if len(matched) >= MIN_KEYWORD_MATCHES:
        print(f"  ✅ Content validated: {len(matched)} payment keywords found")
        return True

    print(f"  🚫 Content rejected: only {len(matched)} keyword(s) found, need at least {MIN_KEYWORD_MATCHES}")
    return False


def is_valid_payment_advice_content(body_text):
    """For email body TXT: broad keyword scan with word-boundary matching."""
    text_lower = body_text.lower()
    matched = []
    for kw in PAYMENT_KEYWORDS_BROAD:
        pattern = r'\b' + re.escape(kw.lower()) + r'\b'
        if re.search(pattern, text_lower):
            matched.append(kw)
    print(f"  🔎 Body keyword scan: {len(matched)} keywords matched → {matched[:5]}")
    return len(matched) >= MIN_KEYWORD_MATCHES


# ════════════════════════════════════════════════════════════════════
#  MULTI-CUSTOMER (Custno/Custname) TABLE DETECTION — format-agnostic
#  NOTE: this ONLY detects whether a file/body is a multi-customer
#  ledger and counts how many distinct customers it has. It does NOT
#  split the file — the whole original content is saved as ONE file
#  under multi-input/ (see save_multi_customer_file below).
# ════════════════════════════════════════════════════════════════════

def detect_and_split_multi_customer_rows(rows, source_name):
    """
    Generic multi-customer detection/split — takes parsed list of rows.
    Returns dict {file_key: text_block} if multi-customer, or None.
    """
    if not rows:
        return None

    header_row_idx = None
    custno_col     = None
    custname_col   = None

    for i, row in enumerate(rows[:5]):
        for j, cell in enumerate(row):
            if cell is None:
                continue
            cell_str = str(cell).strip().lower()
            cell_clean = re.sub(r"[^a-z0-9]", "", cell_str)

            if custno_col is None:
                if ("cust" in cell_str and ("no" in cell_str or "number" in cell_str or "code" in cell_str)):
                    custno_col = j
                elif cell_clean in ("custno", "cno", "custnumber", "customercode", "customerno", "customernumber"):
                    custno_col = j
                elif re.match(r"^c\.?\s*no\.?$", cell_str):
                    custno_col = j
                elif "party" in cell_str and ("code" in cell_str or "no" in cell_str):
                    custno_col = j
                elif "account" in cell_str and ("no" in cell_str or "code" in cell_str):
                    custno_col = j

            if custname_col is None:
                if ("cust" in cell_str and "name" in cell_str):
                    custname_col = j
                elif cell_clean in ("custname", "cname", "customername", "customer", "partyname", "accountname"):
                    custname_col = j
                elif re.match(r"^c\.?\s*name$", cell_str):
                    custname_col = j
                elif "party" in cell_str and "name" in cell_str:
                    custname_col = j
                elif "account" in cell_str and "name" in cell_str:
                    custname_col = j
                elif cell_str == "customer":
                    custname_col = j

        if custno_col is not None and custname_col is not None:
            header_row_idx = i
            break

    if header_row_idx is None or custno_col is None or custname_col is None:
        print(f"  ℹ️ '{source_name}' doesn't look like a multi-customer ledger")
        return None

    headers   = rows[header_row_idx]
    data_rows = rows[header_row_idx + 1:]

    groups = {}
    for row in data_rows:
        if row is None or all(c is None or str(c).strip() == "" for c in row):
            continue
        custno   = row[custno_col]   if custno_col   < len(row) else None
        custname = row[custname_col] if custname_col < len(row) else None
        if (custno is None or str(custno).strip() == "") and (custname is None or str(custname).strip() == ""):
            continue
        key = (
            str(custno).strip()   if custno   is not None else "",
            str(custname).strip() if custname is not None else "",
        )
        groups.setdefault(key, []).append(row)

    if len(groups) < 1:
        print(f"  ℹ️ '{source_name}' has no customer rows — skipping multi-customer path")
        return None

    print(f"  ✅ '{source_name}' detected as multi-customer ledger: {len(groups)} distinct customers")

    customer_blocks = {}
    header_line = "\t".join(str(h) if h is not None else "" for h in headers)

    for (custno, custname), group_rows in groups.items():
        lines = [header_line]
        for row in group_rows:
            lines.append("\t".join(str(c) if c is not None else "" for c in row))
        text_block = "\n".join(lines)

        safe_custno = re.sub(r"[^\w\-]", "", custno) if custno else ""
        name_part   = custname if custname else f"customer-{safe_custno or 'unknown'}"
        file_key    = f"{safe_custno}-{name_part}" if safe_custno else name_part
        customer_blocks[file_key] = text_block

    return customer_blocks


def parse_rows_from_xlsx_bytes(raw_bytes, source_name):
    """Parse .xlsx bytes into a list-of-rows using openpyxl."""
    if openpyxl is None:
        print("  ⚠️ openpyxl not available — skipping xlsx multi-customer split detection")
        return None
    try:
        wb = openpyxl.load_workbook(io.BytesIO(raw_bytes), data_only=True)
        ws = wb.active
        return list(ws.iter_rows(values_only=True))
    except Exception as e:
        print(f"  ⚠️ Could not parse '{source_name}' as Excel: {e}")
        return None


def parse_rows_from_csv_bytes(raw_bytes, source_name):
    """Parse .csv bytes into a list-of-rows."""
    try:
        try:
            text = raw_bytes.decode("utf-8-sig")
        except UnicodeDecodeError:
            text = raw_bytes.decode("latin-1", errors="replace")
        sample = text[:2048]
        try:
            dialect = csv.Sniffer().sniff(sample, delimiters=",\t;")
        except csv.Error:
            dialect = csv.excel
        reader = csv.reader(io.StringIO(text), dialect)
        return [row for row in reader]
    except Exception as e:
        print(f"  ⚠️ Could not parse '{source_name}' as CSV: {e}")
        return None


def parse_rows_from_text_bytes(raw_bytes, source_name):
    """Parse .txt bytes into a list-of-rows (tab or multi-space delimited)."""
    try:
        try:
            text = raw_bytes.decode("utf-8") if isinstance(raw_bytes, (bytes, bytearray)) else raw_bytes
        except UnicodeDecodeError:
            text = raw_bytes.decode("latin-1", errors="replace")

        lines = [l for l in text.splitlines()]
        rows = []
        for line in lines:
            if "\t" in line:
                rows.append(line.split("\t"))
            else:
                parts = re.split(r" {2,}", line.strip())
                rows.append(parts)
        return rows
    except Exception as e:
        print(f"  ⚠️ Could not parse '{source_name}' as text table: {e}")
        return None


def parse_rows_from_html_table(html_text, source_name):
    """Parse the first <table> found in HTML content into a list-of-rows."""
    try:
        table_match = re.search(r"(?is)<table[^>]*>(.*?)</table>", html_text)
        if not table_match:
            return None
        table_html = table_match.group(1)

        row_matches = re.findall(r"(?is)<tr[^>]*>(.*?)</tr>", table_html)
        if not row_matches:
            return None

        rows = []
        for row_html in row_matches:
            cell_matches = re.findall(r"(?is)<t[dh][^>]*>(.*?)</t[dh]>", row_html)
            cells = []
            for cell_html in cell_matches:
                cell_text = re.sub(r"<[^>]+>", " ", cell_html)
                cell_text = html.unescape(cell_text)
                cell_text = re.sub(r"\s+", " ", cell_text).strip()
                cells.append(cell_text)
            if cells:
                rows.append(cells)
        return rows if rows else None
    except Exception as e:
        print(f"  ⚠️ Could not parse '{source_name}' as HTML table: {e}")
        return None


def try_multi_customer_split(raw_content, source_name, ext_or_kind):
    """Dispatcher: picks the right row-parser based on ext_or_kind.
    Used ONLY for detection/counting — the returned dict's length tells
    us how many distinct customers are present. The actual content that
    gets saved is always the ORIGINAL raw_content, untouched.
    """
    ext_or_kind = ext_or_kind.lower()

    if ext_or_kind == ".xlsx":
        rows = parse_rows_from_xlsx_bytes(raw_content, source_name)
    elif ext_or_kind == ".csv":
        rows = parse_rows_from_csv_bytes(raw_content, source_name)
    elif ext_or_kind == ".txt":
        rows = parse_rows_from_text_bytes(raw_content, source_name)
    elif ext_or_kind in (".html", ".htm"):
        try:
            text = raw_content.decode("utf-8") if isinstance(raw_content, (bytes, bytearray)) else raw_content
        except UnicodeDecodeError:
            text = raw_content.decode("latin-1", errors="replace")
        rows = parse_rows_from_html_table(text, source_name)
    elif ext_or_kind == "body_html":
        rows = parse_rows_from_html_table(raw_content, source_name)
    elif ext_or_kind == "body_text":
        rows = parse_rows_from_text_bytes(raw_content, source_name)
    else:
        return None

    if not rows:
        return None

    return detect_and_split_multi_customer_rows(rows, source_name)


def build_multi_customer_filename(customer_count, received_datetime_ist):
    """DEPRECATED — kept for backward compatibility but no longer used.
    Multi-customer files now use original attachment filenames."""
    time_part = "00.00.00"
    if received_datetime_ist:
        m = re.search(r"(\d{2}):(\d{2}):(\d{2})", received_datetime_ist)
        if m:
            time_part = f"{m.group(1)}.{m.group(2)}.{m.group(3)}"
    return f"multy.{customer_count}c.{time_part}"


def _make_unique_multi_input_name(bucket_name, s3_safe_name):
    """Same logic as _make_unique_s3_name but for multi-input/ prefix."""
    try:
        s3.head_object(Bucket=bucket_name, Key=f"multi-input/{s3_safe_name}")
    except Exception:
        return s3_safe_name

    if "." in s3_safe_name:
        base, ext = s3_safe_name.rsplit(".", 1)
        ext = "." + ext
    else:
        base, ext = s3_safe_name, ""

    counter = 1
    while counter < 1000:
        candidate = f"{base}-{counter}{ext}"
        try:
            s3.head_object(Bucket=bucket_name, Key=f"multi-input/{candidate}")
            counter += 1
        except Exception:
            print(f"  🔄 Multi-input file '{s3_safe_name}' already exists — renamed to '{candidate}'")
            return candidate

    ts = datetime.now().strftime("%Y%m%d%H%M%S")
    fallback = f"{base}-{ts}{ext}"
    return fallback


def save_multi_customer_file(raw_content, is_binary, file_ext, content_type, customer_count,
                              source_name, bucket_name, msg_id, received_datetime_ist,
                              subject, from_address, source_label):
    """
    Saves a detected multi-customer email/attachment as ONE file, UNCHANGED
    (no per-customer splitting), into multi-input/ instead of emails/.

    Uses the ORIGINAL attachment filename (sanitized for S3).
    If same filename already exists in multi-input/, appends -1, -2, etc.
    """
    # Use original source filename instead of multy.Nc.HH.MM.SS pattern
    s3_safe_name = sanitize_s3_key(source_name) if source_name and source_name != "email_body" else None

    if not s3_safe_name:
        # Fallback for email body content (no attachment filename)
        s3_safe_name = f"email-body-{msg_id[:8]}{file_ext}"

    # Ensure the extension matches
    if not s3_safe_name.lower().endswith(file_ext.lower()):
        # Strip existing extension and add the correct one
        if "." in s3_safe_name:
            base_no_ext = s3_safe_name.rsplit(".", 1)[0]
        else:
            base_no_ext = s3_safe_name
        s3_safe_name = f"{base_no_ext}{file_ext}"

    # Content-hash dedup for multi-customer files too
    if is_binary:
        body_bytes = raw_content
    else:
        body_bytes = raw_content.encode("utf-8")

    # Make unique name if same filename already exists
    s3_safe_name = _make_unique_multi_input_name(bucket_name, s3_safe_name)

    multi_key = f"multi-input/{s3_safe_name}"
    meta_key  = f"metadata/{s3_safe_name}.json"

    s3.put_object(
        Bucket=bucket_name,
        Key=multi_key,
        Body=body_bytes,
        ContentType=content_type,
    )
    print(f"  ✅ Saved multi-customer file: {multi_key} ({len(body_bytes)} bytes)")

    metadata = {
        "MAIL_ID":            msg_id,
        "MAIL_RECEIVED_DATE": received_datetime_ist,
        "attachment_name":    source_name,
        "s3_key":             s3_safe_name,
        "subject":            subject,
        "from":               from_address,
        "source_attachment":  source_name,
        "source":             source_label,
        "customer_count":     customer_count,
    }
    s3.put_object(
        Bucket=bucket_name,
        Key=meta_key,
        Body=json.dumps(metadata, indent=2).encode("utf-8"),
        ContentType="application/json",
    )
    print(f"  ✅ Saved multi-customer metadata: {meta_key}")

    return s3_safe_name


def is_item_attachment(att):
    """True if this Graph attachment is an embedded/forwarded email."""
    return att.get("@odata.type", "") == "#microsoft.graph.itemAttachment"


def get_nested_item_attachment(mailbox, msg_id, attachment_id, headers):
    """Fetch the nested email (item attachment)'s content."""
    url = (
        f"https://graph.microsoft.com/v1.0/users/{mailbox}"
        f"/messages/{msg_id}/attachments/{attachment_id}"
        f"?$expand=microsoft.graph.itemattachment/item"
    )
    return get_json(url, headers)


def process_nested_email_attachment(item_att, bucket_name, msg_id, mailbox, headers,
                                     outer_subject, outer_from, received_datetime_ist):
    """Handles an embedded/forwarded email attachment."""
    saved   = []
    skipped = []

    attachment_id = item_att.get("id")
    nested_name   = item_att.get("name", "nested-email")

    try:
        detail = get_nested_item_attachment(mailbox, msg_id, attachment_id, headers)
    except Exception as e:
        print(f"  ⚠️ Could not fetch nested item attachment '{nested_name}': {e}")
        skipped.append({"file": nested_name, "reason": f"Failed to fetch nested item: {e}"})
        return saved, skipped

    nested_item = detail.get("item", {})
    if not nested_item:
        print(f"  ⚠️ Nested item attachment '{nested_name}' had no expandable item content")
        skipped.append({"file": nested_name, "reason": "No expandable item content found"})
        return saved, skipped

    nested_subject = nested_item.get("subject", nested_name)
    nested_from    = nested_item.get("from", {}).get("emailAddress", {}).get("address", outer_from)

    nested_body_content = nested_item.get("body", {})
    nested_body_text    = nested_body_content.get("content", "")
    nested_body_type    = nested_body_content.get("contentType", "text")

    print(f"  🧪 DEBUG nested body_type = '{nested_body_type}'")
    print(f"  🧪 DEBUG nested raw body_text (first 500 chars) = {nested_body_text[:500]!r}")

    nested_body_text_raw     = nested_body_text
    nested_body_text_trimmed = strip_quoted_reply_trail(nested_body_text, nested_body_type, source_name=nested_name)

    # Check trimmed and untrimmed SEPARATELY, then decide ONCE which version
    # is the real content — same reasoning and same pattern as the outer
    # email body check. That single decision (nested_body_text_effective)
    # is then used consistently for BOTH the multi-customer table detection
    # and the single-customer company-name extraction below.
    trimmed_has_keywords = (
        is_payment_advice_email(nested_body_text_trimmed)
        or is_valid_payment_advice_content(nested_body_text_trimmed)
    )
    raw_has_keywords = (
        is_payment_advice_email(nested_body_text_raw)
        or is_valid_payment_advice_content(nested_body_text_raw)
    )

    if not (trimmed_has_keywords or raw_has_keywords):
        print(f"  ⏭️ Nested email '{nested_subject}' body has no payment advice content — skipping")
        skipped.append({
            "file":   nested_name,
            "reason": "Nested email body has no payment advice content",
        })
        _log_rejection_lambda1(
            nested_name, "Nested email body has no payment advice content",
            mail_id=nested_from, mail_received_date=received_datetime_ist,
            subject=nested_subject, from_address=nested_from,
        )
        return saved, skipped

    if trimmed_has_keywords:
        nested_body_text_effective = nested_body_text_trimmed
    else:
        print(f"  ℹ️ Trimmed nested body '{nested_name}' had no payment keywords — using UNTRIMMED body downstream")
        nested_body_text_effective = nested_body_text_raw

    kind = "body_html" if nested_body_type == "html" else "body_text"
    customer_blocks = try_multi_customer_split(nested_body_text_effective, nested_name, kind)

    if customer_blocks:
        # Multi-customer body → save the WHOLE body (converted to plain text)
        # as ONE file in multi-input/, not split per customer.
        if nested_body_type == "html":
            whole_plain_text = html_to_text_preserve_lines(nested_body_text_effective)
        else:
            whole_plain_text = nested_body_text_effective.strip()

        saved_name = save_multi_customer_file(
            whole_plain_text, is_binary=False, file_ext=".txt", content_type="text/plain",
            customer_count=len(customer_blocks),
            source_name=nested_name, bucket_name=bucket_name, msg_id=msg_id,
            received_datetime_ist=received_datetime_ist, subject=outer_subject, from_address=outer_from,
            source_label="multi_customer_nested_body_whole",
        )
        if saved_name:
            saved.append(saved_name)
        return saved, skipped

    company_name  = extract_company_name(nested_body_text_effective, nested_body_type, msg_id)
    body_txt_key  = f"emails/{company_name}.txt"
    body_meta_key = f"metadata/{company_name}.txt.json"

    if nested_body_type == "html":
        plain_text = html_to_text_preserve_lines(nested_body_text_effective)
    elif re.search(r"<[a-zA-Z][^>]*>", nested_body_text_effective):
        plain_text = html_to_text_preserve_lines(nested_body_text_effective)
    else:
        plain_text = nested_body_text_effective.strip()

    s3.put_object(
        Bucket=bucket_name,
        Key=body_txt_key,
        Body=plain_text.encode("utf-8"),
        ContentType="text/plain",
    )
    print(f"  ✅ Saved nested body as TXT: {body_txt_key}")

    body_metadata = {
        "MAIL_ID":            msg_id,
        "MAIL_RECEIVED_DATE": received_datetime_ist,
        "attachment_name":    f"{company_name}.txt",
        "s3_key":             f"{company_name}.txt",
        "subject":            outer_subject,
        "from":               outer_from,
        "nested_subject":     nested_subject,
        "nested_from":        nested_from,
        "source":             "nested_email_body",
    }
    s3.put_object(
        Bucket=bucket_name,
        Key=body_meta_key,
        Body=json.dumps(body_metadata, indent=2).encode("utf-8"),
        ContentType="application/json",
    )
    print(f"  ✅ Saved nested body metadata: {body_meta_key}")
    saved.append(f"{company_name}.txt")

    return saved, skipped


def extract_company_name_with_claude(plain_text, msg_id):
    """Fallback company-name extraction using Bedrock Claude."""
    try:
        prompt = (
            "Below is the body of a payment/remittance advice email. Identify ONLY the "
            "name of the company or person who SENT this payment (the vendor/customer who "
            "is paying), not the recipient. The recipient is always Tube Investments of "
            "India Ltd / TIDC India / TII — never return that name.\n\n"
            "Respond with ONLY the company or sender name, nothing else — no explanation, "
            "If no company name is found, or if the company name cannot be identified, or if the identified company name is anything other than Tube Investments,then output only Tube Investments."
            "no punctuation around it, no quotes. If you genuinely cannot identify a sender "
            "name, respond with exactly: UNKNOWN\n\n"
            f"Email body:\n{plain_text[:4000]}"
        )

        response = bedrock_runtime.converse(
            modelId=MODEL_ID,
            messages=[{"role": "user", "content": [{"text": prompt}]}],
            inferenceConfig={"maxTokens": 60, "temperature": 0},
        )

        result_text = response["output"]["message"]["content"][0]["text"].strip()
        result_text = result_text.strip('"\'')

        print(f"  🤖 Claude company name extraction result: '{result_text}'")

        if result_text and result_text.upper() != "UNKNOWN" and 3 <= len(result_text) <= 80:
            return sanitize_s3_key(result_text)

    except Exception as e:
        print(f"  ⚠️ Claude company name extraction failed: {e} — falling back to unknown-{msg_id[:8]}")

    return None


def extract_company_name(body_text, body_type, msg_id):
    """
    Extract company/vendor name from email body to use as the TXT filename.
    Strategies (tried in order):
      1. "Customer Name" table row
      2. Remitter's name field
      3. ALL-CAPS company name with Indian suffix
      3b. Title-Case company name with suffix
      4. M/s. or Messrs. prefix
      5. Labeled inline patterns
      6. ALL-CAPS line in first 30 lines
      7. Title-Case multi-word line near top
      8. Bedrock Claude fallback
      9. Fallback → unknown-<msg_id[:8]>
    """
    if body_type == "html":
        plain = html_to_text_preserve_lines(body_text)
    else:
        plain = body_text.strip()

    lines = [l.strip() for l in plain.splitlines() if l.strip()]

    # ── Strategy 1: "Customer Name" table row → value on next line ──
    for i, line in enumerate(lines):
        if re.match(r"^customer\s*name", line, re.IGNORECASE):
            if i + 1 < len(lines):
                next_line = lines[i + 1].strip()
                name = re.sub(r"\s+\d+\s*$", "", next_line).strip()
                if 3 <= len(name) <= 80:
                    print(f"  🏢 Company name (Customer Name row): {name}")
                    return sanitize_s3_key(name)

    # ── Strategy 2: Remitter's name field ──
    remitter_match = re.search(
        r"remitter.{0,2}s?\s*name\s*[:\-]\s*(.+?)(?:Remitting|\n|$)",
        plain, re.IGNORECASE
    )
    if remitter_match:
        name = remitter_match.group(1).strip().rstrip(".,")
        if 3 <= len(name) <= 80:
            print(f"  🏢 Company name (remitter): {name}")
            return sanitize_s3_key(name)

    # ── Strategy 3: ALL-CAPS company name with Indian suffix ──
    TII_NAMES = {
        "TUBE INVESTMENT OF INDIA LTD",
        "TUBE INVESTMENTS OF INDIA LTD",
        "TUBE INVESTMENT OF INDIA LIMITED",
        "TUBE INVESTMENTS OF INDIA LIMITED",
        "TUBE INVESTMENTS INDIA LTD",
    }
    for allcaps_match in re.finditer(
        r"([A-Z][A-Z\s&()\-\.]{5,}(?:PVT\.?\s*LTD\.?|LTD\.?|LIMITED|WORKS|INDUSTRIES|ENGINEERS?|ENTERPRISES?|CORPORATION|INFRA(?:STRUCTURE)?|TRADING|BROTHERS?|BROS\.?))",
        plain
    ):
        name = allcaps_match.group(1).strip().rstrip(".")
        name = re.sub(r"\s{2,}", " ", name)
        if name.upper() in TII_NAMES:
            continue
        if 3 <= len(name) <= 80:
            print(f"  🏢 Company name (all-caps regex): {name}")
            return sanitize_s3_key(name)

    # ── Strategy 3b: Title-Case company name with suffix ──
    for mixedcase_match in re.finditer(
        r"([A-Z][\w&,\.\-]*(?:\s+[A-Z][\w&,\.\-]*){0,5}\s+"
        r"(?:Pvt\.?\s*Ltd\.?|Private\s+Limited|Ltd\.?|Limited|LLP|Inc\.?|"
        r"Industries|Enterprises|Corporation|Trading\s+Co\.?|&\s*Co\.?|Brothers|Bros\.?))",
        plain
    ):
        name = mixedcase_match.group(1).strip().rstrip(".,")
        name = re.sub(r"\s{2,}", " ", name)
        if name.upper() in TII_NAMES:
            continue
        if 3 <= len(name) <= 80:
            print(f"  🏢 Company name (mixed-case regex): {name}")
            return sanitize_s3_key(name)


    # ── Strategy 4: M/s. or Messrs. prefix ──
    for line in lines[:40]:
        m = re.search(r"(?:m/s\.?\s+|messrs\.?\s+)(.+)", line, re.IGNORECASE)
        if m:
            name = m.group(1).strip().rstrip(".,")
            if 3 <= len(name) <= 80:
                print(f"  🏢 Company name (M/s prefix): {name}")
                return sanitize_s3_key(name)

    # ── Strategy 5: labeled inline patterns ──
    label_patterns = [
        r"(?:customer\s*name|vendor|company|from|sender|party)\s*[:\-]\s*(.+)",
    ]
    for line in lines[:40]:
        for pat in label_patterns:
            m = re.search(pat, line, re.IGNORECASE)
            if m:
                name = re.sub(r"\s+\d+\s*$", "", m.group(1)).strip().rstrip(".,")
                if 3 <= len(name) <= 80:
                    print(f"  🏢 Company name (labeled pattern): {name}")
                    return sanitize_s3_key(name)

    # ── Strategy 6: ALL-CAPS line in first 30 lines ──
    for line in lines[:30]:
        if re.match(r"^[A-Z][A-Z\s&()\-\.]{5,}$", line) and len(line.split()) >= 2:
            print(f"  🏢 Company name (all-caps): {line}")
            return sanitize_s3_key(line)

    # ── Strategy 7: Title-Case multi-word line near top ──
    skip_words = {"payment", "advice", "dear", "sir", "madam", "date", "ref",
                  "invoice", "amount", "total", "utr", "reference", "number",
                  "tube", "investments", "india", "limited", "kindly", "please"}
    for line in lines[:20]:
        words = line.split()
        if (
            2 <= len(words) <= 6
            and all(w[0].isupper() for w in words if w.isalpha())
            and not any(w.lower() in skip_words for w in words)
        ):
            print(f"  🏢 Company name (title-case): {line}")
            return sanitize_s3_key(line)

    # ── Strategy 8: Bedrock Claude fallback ──
    claude_name = extract_company_name_with_claude(plain, msg_id)
    if claude_name:
        print(f"  🏢 Company name (Claude fallback): {claude_name}")
        return claude_name

    # ── Fallback ──
    fallback = f"unknown-{msg_id[:8]}"
    print(f"  ⚠️ Could not extract company name — using fallback: {fallback}")
    return fallback


def store_email_to_s3(msg, headers, bucket_name, mailbox):
    """
    ╔══════════════════════════════════════════════════════════════════╗
    ║  V2 CHANGE: MANDATORY BODY KEYWORD CHECK BEFORE PROCESSING     ║
    ║                                                                  ║
    ║  The email body MUST contain payment-related keywords before     ║
    ║  any attachments are processed. If the body has no payment       ║
    ║  keywords → skip the entire email (attachments included).        ║
    ║                                                                  ║
    ║  This prevents non-payment emails with random PDF/Excel          ║
    ║  attachments from being processed through the pipeline.          ║
    ╚══════════════════════════════════════════════════════════════════╝

    After the body check passes:
    - If real (non-inline) attachments exist → process ONLY attachments.
    - If NO real attachments → fall back to body content.
    """
    msg_id       = msg["id"]
    raw_datetime = msg.get("receivedDateTime", "")
    subject      = msg.get("subject", "")
    from_address = msg.get("from", {}).get("emailAddress", {}).get("address", "")

    received_datetime_ist = utc_to_ist_display(raw_datetime)

    # ══════════════════════════════════════════════════════════════
    # V2 NEW: MANDATORY BODY KEYWORD GATE
    # Email body must have payment keywords — otherwise skip entirely
    # ══════════════════════════════════════════════════════════════
    body_content = msg.get("body", {})
    body_text    = body_content.get("content", "")
    body_type    = body_content.get("contentType", "text")

    # Trim quoted reply trail first (same as existing logic)
    body_text_trimmed = strip_quoted_reply_trail(body_text, body_type, source_name=subject or "email_body")

    # Check keywords on the trimmed body and the raw untrimmed body SEPARATELY.
    # The quoted-reply trimmer can sometimes cut off the actual payment/
    # multi-customer table content along with the old reply chain (e.g. if
    # the real table sits below an "Original Message" / <hr> marker). Rather
    # than just gating on "either has keywords", we decide ONCE which body
    # version is the real content, and use THAT SAME version everywhere
    # downstream — for single-customer body saving AND multi-customer table
    # detection/saving alike. This keeps both paths consistent instead of
    # each re-deciding independently.
    trimmed_has_keywords = is_payment_advice_email(body_text_trimmed) or is_valid_payment_advice_content(body_text_trimmed)
    raw_has_keywords      = is_payment_advice_email(body_text) or is_valid_payment_advice_content(body_text)

    body_has_keywords = trimmed_has_keywords or raw_has_keywords

    # ── Decide the ONE effective body to use downstream ──
    if trimmed_has_keywords:
        # Extra check: if trimming cut off a large portion of invoice rows,
        # prefer the raw body. This handles cases where <hr> splits a
        # payment table mid-way, causing strip_quoted_reply_trail to
        # discard hundreds of rows while leaving just enough for keyword match.
        trimmed_s_rows = len(re.findall(r'(?m)^S\|', body_text_trimmed))
        raw_s_rows     = len(re.findall(r'(?m)^S\|', body_text))
        # Also count S| patterns that aren't at line start (inline HTML)
        trimmed_s_count = body_text_trimmed.count('S|')
        raw_s_count     = body_text.count('S|')

        if raw_s_count > trimmed_s_count * 1.5 and raw_s_count > trimmed_s_count + 5:
            # Raw body has significantly more invoice rows — trimming cut off data
            print(f"  ℹ️ Trim cut off invoice rows (trimmed={trimmed_s_count} vs raw={raw_s_count}) — using UNTRIMMED body")
            body_text_effective = body_text
        else:
            body_text_effective = body_text_trimmed
    elif raw_has_keywords:
        print("  ℹ️ Trimmed body had no payment keywords — using UNTRIMMED body downstream (trim had cut off real content)")
        body_text_effective = body_text
    else:
        body_text_effective = body_text_trimmed if body_text_trimmed else body_text

    # ── Fetch attachments (with retry on 404 — Graph API may need time to index) ──
    attachments_url = (
        f"https://graph.microsoft.com/v1.0/users/{mailbox}"
        f"/messages/{msg_id}/attachments"
    )
    import time as _time
    attachments_response = None
    for _attempt in range(3):
        try:
            attachments_response = get_json(attachments_url, headers)
            break
        except Exception as _e:
            if "404" in str(_e) and _attempt < 2:
                print(f"  ⏳ Attachments fetch got 404 — retrying in 5s (attempt {_attempt + 1}/3)")
                _time.sleep(5)
            else:
                raise
    all_attachments = attachments_response.get("value", [])
    print(f"  📎 Total attachments (including inline): {len(all_attachments)}")

    # ── Debug: log all attachment names and types for troubleshooting ──
    for _idx, _att in enumerate(all_attachments):
        print(f"  📎 Attachment[{_idx}]: name='{_att.get('name', '')}' "
              f"isInline={_att.get('isInline', False)} "
              f"hasContentBytes={bool(_att.get('contentBytes'))} "
              f"contentType='{_att.get('contentType', '')}' "
              f"size={_att.get('size', 0)}")

    # ── Separate real attachments from inline ones ──
    real_attachments = [
        att for att in all_attachments
        if not att.get("isInline", False) and att.get("contentBytes")
    ]
    print(f"  📎 Real (non-inline) attachments: {len(real_attachments)}")

    # ── Also include inline PDFs that may lack contentBytes (needs separate download) ──
    has_pdf_already = any(att.get("name", "").lower().endswith(".pdf") for att in real_attachments)
    if not has_pdf_already:
        # First try inline PDFs with contentBytes
        inline_pdfs = [
            att for att in all_attachments
            if att.get("isInline", False)
            and att.get("contentBytes")
            and att.get("name", "").lower().endswith(".pdf")
        ]
        if inline_pdfs:
            print(f"  � Found {len(inline_pdfs)} inline PDF(s) with contentBytes — adding to real attachments")
            real_attachments.extend(inline_pdfs)
        else:
            # Try ALL PDFs without contentBytes (inline or not) — download them
            pdf_attachments_no_bytes = [
                att for att in all_attachments
                if att.get("name", "").lower().endswith(".pdf")
                and not att.get("contentBytes")
                and att.get("id")
            ]
            if not pdf_attachments_no_bytes:
                # Maybe the PDF name isn't set — try by contentType
                pdf_attachments_no_bytes = [
                    att for att in all_attachments
                    if "pdf" in att.get("contentType", "").lower()
                    and not att.get("contentBytes")
                    and att.get("id")
                    and att.get("name", "").lower() != "removed attachments.txt"
                ]
            for att in pdf_attachments_no_bytes:
                att_id = att["id"]
                try:
                    att_url = (f"https://graph.microsoft.com/v1.0/users/{mailbox}"
                              f"/messages/{msg_id}/attachments/{att_id}")
                    att_detail = get_json(att_url, headers)
                    if att_detail.get("contentBytes"):
                        att["contentBytes"] = att_detail["contentBytes"]
                        if not att.get("name"):
                            att["name"] = att_detail.get("name", "attachment.pdf")
                        real_attachments.append(att)
                        print(f"  � Downloaded PDF attachment: '{att.get('name')}' — added to real attachments")
                    else:
                        print(f"  ⚠️ PDF '{att.get('name')}' has no contentBytes even after direct fetch")
                except Exception as e:
                    print(f"  ⚠️ Failed to download PDF attachment '{att.get('name')}': {e}")

    # ── Separate item attachments (nested/embedded emails) ──
    item_attachments = [att for att in all_attachments if is_item_attachment(att)]
    print(f"  📨 Item (nested email) attachments: {len(item_attachments)}")

    # ── If NO attachments at all AND body has no keywords → skip email ──
    if not real_attachments and not item_attachments and not body_has_keywords:
        print(f"  🚫 No attachments and body has no payment keywords — skipping: '{subject}'")
        _log_rejection_lambda1(
            f"email_{msg_id[:16]}",
            "No attachments and body has no payment keywords — email skipped entirely",
            mail_id=from_address, mail_received_date=received_datetime_ist,
            subject=subject, from_address=from_address,
        )
        return {
            "attachments": [],
            "skipped":     [{"file": "entire_email", "reason": "No attachments and body has no payment keywords — skipped"}],
        }

    # ── If attachments exist, process them regardless of body keywords ──
    if real_attachments or item_attachments:
        print(f"  ✅ Attachments found — proceeding with processing")
    else:
        print(f"  ✅ No attachments but body has keywords — using body as content")

    saved_attachments   = []
    skipped_attachments = []


    # ══════════════════════════════════════════════════════════════
    # BRANCH A — real attachments present → process attachments only
    # ══════════════════════════════════════════════════════════════
    if real_attachments:
        print(f"  📂 Attachment-first path: processing {len(real_attachments)} attachment(s)")

        for att in real_attachments:
            att_name      = att.get("name", "unnamed_file")
            s3_safe_name  = sanitize_s3_key(att_name)
            att_bytes_b64 = att.get("contentBytes")
            content_type  = att.get("contentType", "application/octet-stream")

            # ── Unique naming: if file already exists, append incrementing number ──
            s3_safe_name = _make_unique_s3_name(bucket_name, s3_safe_name)

            att_key  = f"emails/{s3_safe_name}"
            meta_key = f"metadata/{s3_safe_name}.json"

            lower_name = att_name.lower()
            ext        = "." + lower_name.rsplit(".", 1)[-1] if "." in lower_name else ""

            # ── Stage 1: Reject hard binary formats immediately ──
            if not has_valid_extension(att_name):
                skipped_attachments.append({
                    "file":   att_name,
                    "reason": f"Extension not allowed",
                })
                _log_rejection_lambda1(
                    att_name, f"Extension not allowed — unsupported file format",
                    mail_id=from_address, mail_received_date=received_datetime_ist,
                    subject=subject, from_address=from_address,
                )
                # Save to rejected folder if bytes available
                if att_bytes_b64:
                    _save_to_rejected_folder(bucket_name, s3_safe_name, base64.b64decode(att_bytes_b64), content_type)
                continue

            # ── Decode bytes ──
            att_bytes = base64.b64decode(att_bytes_b64)

            # ── Check if PDF is password-protected and decrypt if needed ──
            if ext == ".pdf":
                if is_pdf_encrypted(att_bytes):
                    print(f"  🔐 PDF is password-protected: '{att_name}' — looking up password by subject")
                    password = find_password_by_subject(subject)
                    if password:
                        decrypted = decrypt_pdf_bytes(att_bytes, password)
                        if decrypted:
                            att_bytes = decrypted
                        else:
                            print(f"  ❌ Failed to decrypt '{att_name}' — skipping")
                            skipped_attachments.append({
                                "file":   att_name,
                                "reason": f"Failed to decrypt password-protected PDF",
                            })
                            _log_rejection_lambda1(
                                att_name, "Failed to decrypt password-protected PDF — decryption error",
                                mail_id=from_address, mail_received_date=received_datetime_ist,
                                subject=subject, from_address=from_address,
                            )
                            _save_to_rejected_folder(bucket_name, s3_safe_name, att_bytes, "application/pdf")
                            continue
                    else:
                        print(f"  ⚠️ No password found for subject '{subject[:50]}' — skipping encrypted PDF")
                        skipped_attachments.append({
                            "file":   att_name,
                            "reason": f"Password-protected PDF but no password found in DynamoDB for this subject",
                        })
                        _log_rejection_lambda1(
                            att_name, f"Password-protected PDF — no password found for subject '{subject[:80]}'",
                            mail_id=from_address, mail_received_date=received_datetime_ist,
                            subject=subject, from_address=from_address,
                        )
                        _save_to_rejected_folder(bucket_name, s3_safe_name, att_bytes, "application/pdf")
                        continue

            # ── Content-hash dedup: DISABLED ──
            # Dedup is handled by Lambda Full's ETag check instead.
            # Lambda1 always saves the file to S3 — if Lambda Full fails,
            # the file can be retried without getting stuck.

            # ── Multi-customer table check — if detected, save the WHOLE
            #    original attachment (unsplit) to multi-input/ ──
            if ext in MULTI_CUSTOMER_SPLIT_EXTENSIONS:
                customer_blocks = try_multi_customer_split(att_bytes, att_name, ext)
                if customer_blocks:
                    saved_name = save_multi_customer_file(
                        att_bytes, is_binary=True, file_ext=ext, content_type=content_type,
                        customer_count=len(customer_blocks),
                        source_name=att_name, bucket_name=bucket_name, msg_id=msg_id,
                        received_datetime_ist=received_datetime_ist, subject=subject, from_address=from_address,
                        source_label=f"multi_customer_{ext.lstrip('.')}_whole",
                    )
                    if saved_name:
                        saved_attachments.append(saved_name)
                    continue

            # ── Stage 2: Only for TXT/CSV/HTML — PDF/Excel/Images always pass ──
            if not content_has_payment_fields(att_bytes, att_name):
                skipped_attachments.append({
                    "file":   att_name,
                    "reason": f"Content validation failed: no payment keywords in '{att_name}'",
                })
                _log_rejection_lambda1(
                    att_name, f"Content validation failed — no payment keywords found in file content",
                    mail_id=from_address, mail_received_date=received_datetime_ist,
                    subject=subject, from_address=from_address,
                )
                _save_to_rejected_folder(bucket_name, s3_safe_name, att_bytes, content_type)
                continue


            # ── Save attachment → triggers Lambda 2 ──
            s3.put_object(
                Bucket=bucket_name,
                Key=att_key,
                Body=att_bytes,
                ContentType=content_type,
            )
            print(f"  ✅ Saved attachment: {att_key} ({len(att_bytes)} bytes)")

            # ── Save sidecar metadata ──
            metadata = {
                "MAIL_ID":            msg_id,
                "MAIL_RECEIVED_DATE": received_datetime_ist,
                "attachment_name":    att_name,
                "s3_key":             s3_safe_name,
                "subject":            subject,
                "from":               from_address,
            }
            s3.put_object(
                Bucket=bucket_name,
                Key=meta_key,
                Body=json.dumps(metadata, indent=2).encode("utf-8"),
                ContentType="application/json",
            )
            print(f"  ✅ Saved metadata: {meta_key}")
            saved_attachments.append(att_name)


    # ══════════════════════════════════════════════════════════════
    # BRANCH A2 — item attachments (nested emails)
    # ══════════════════════════════════════════════════════════════
    if item_attachments:
        print(f"  📨 Processing {len(item_attachments)} nested email attachment(s)")
        for item_att in item_attachments:
            nested_saved, nested_skipped = process_nested_email_attachment(
                item_att, bucket_name, msg_id, mailbox, headers,
                subject, from_address, received_datetime_ist
            )
            saved_attachments.extend(nested_saved)
            skipped_attachments.extend(nested_skipped)

    # ══════════════════════════════════════════════════════════════
    # BRANCH B — no real attachments AND no item attachments →
    # fall back to the outer email's own body content
    # (Body already passed the keyword check above, so proceed)
    # ══════════════════════════════════════════════════════════════
    if not real_attachments and not item_attachments:
        print("  📭 No real attachments — using email body as payment advice content")

        # ── Multi-customer table check on the email body ──
        # Use the SAME effective body (decided once above) for both the
        # multi-customer table detection and, if it turns out to be a
        # single customer, the company-name extraction below. This keeps
        # both paths consistent instead of each re-deciding independently.
        kind = "body_html" if body_type == "html" else "body_text"
        customer_blocks = try_multi_customer_split(body_text_effective, "email_body", kind)

        if customer_blocks:
            # Multi-customer body → save the WHOLE body (converted to plain
            # text) as ONE file in multi-input/, not split per customer.
            if body_type == "html":
                whole_plain_text = html_to_text_preserve_lines(body_text_effective)
            elif re.search(r"<[a-zA-Z][^>]*>", body_text_effective):
                whole_plain_text = html_to_text_preserve_lines(body_text_effective)
            else:
                whole_plain_text = body_text_effective.strip()

            saved_name = save_multi_customer_file(
                whole_plain_text, is_binary=False, file_ext=".txt", content_type="text/plain",
                customer_count=len(customer_blocks),
                source_name="email_body", bucket_name=bucket_name, msg_id=msg_id,
                received_datetime_ist=received_datetime_ist, subject=subject, from_address=from_address,
                source_label="multi_customer_body_whole",
            )
            if saved_name:
                saved_attachments.append(saved_name)
        else:
            company_name    = extract_company_name(body_text_effective, body_type, msg_id)
            body_txt_key    = f"emails/{company_name}.txt"
            body_meta_key   = f"metadata/{company_name}.txt.json"
            unique_filename = company_name

            if body_type == "html":
                plain_text = html_to_text_preserve_lines(body_text_effective)
            elif re.search(r"<[a-zA-Z][^>]*>", body_text_effective):
                # Content has HTML tags even if body_type says "text" — parse as HTML
                plain_text = html_to_text_preserve_lines(body_text_effective)
            else:
                plain_text = body_text_effective.strip()

            # ── POST-CHECK: If the raw body has significantly more S| rows than ──
            # ── what ended up in plain_text, re-extract by simply stripping tags ──
            raw_body_stripped = re.sub(r"<[^>]+>", "\n", body_text_effective)
            raw_body_stripped = html.unescape(raw_body_stripped)
            raw_s_count = raw_body_stripped.count("S|")
            plain_s_count = plain_text.count("S|")
            if raw_s_count > plain_s_count + 5:
                print(f"  ℹ️ POST-CHECK: plain_text has {plain_s_count} S| rows but raw body has {raw_s_count} — rebuilding from raw")
                # Simple approach: strip HTML tags, split on S| boundaries, rebuild
                lines_raw = [l.strip() for l in raw_body_stripped.splitlines() if l.strip()]
                combined_raw = "\n".join(lines_raw)
                # Split at each S| pattern start
                parts = re.split(r"(?=S\|\d{10,})", combined_raw)
                rebuilt_lines = []
                for part in parts:
                    part = part.strip()
                    if part:
                        rebuilt_lines.append(part)
                plain_text = "\n".join(rebuilt_lines)
                # Trim to start from Payment Remittance Advice
                for marker in ["Payment Remittance Advice", "Payment Advice", "Remittance Advice"]:
                    idx = plain_text.find(marker)
                    if idx > 0:
                        plain_text = plain_text[idx:]
                        break

            s3.put_object(
                Bucket=bucket_name,
                Key=body_txt_key,
                Body=plain_text.encode("utf-8"),
                ContentType="text/plain",
            )
            print(f"  ✅ Saved body as TXT: {body_txt_key}")

            body_metadata = {
                "MAIL_ID":            msg_id,
                "MAIL_RECEIVED_DATE": received_datetime_ist,
                "attachment_name":    f"{unique_filename}.txt",
                "s3_key":             f"{unique_filename}.txt",
                "subject":            subject,
                "from":               from_address,
                "source":             "email_body",
            }
            s3.put_object(
                Bucket=bucket_name,
                Key=body_meta_key,
                Body=json.dumps(body_metadata, indent=2).encode("utf-8"),
                ContentType="application/json",
            )
            print(f"  ✅ Saved body metadata: {body_meta_key}")
            saved_attachments.append(f"{unique_filename}.txt")

    return {
        "attachments": saved_attachments,
        "skipped":     skipped_attachments,
    }


# ════════════════════════════════════════════════════════════════════
#  lambda_handler — splits into two modes:
#  MODE 1 → triggered by SQS: process ONE email (calls store_email_to_s3)
#  MODE 2 → triggered by EventBridge scheduler: fetch emails → queue to SQS
# ════════════════════════════════════════════════════════════════════

def lambda_handler(event, context):

    SQS_QUEUE_URL = os.environ.get("SQS_QUEUE_URL")

    # ──────────────────────────────────────────────────────────────
    # MODE 1: SQS trigger — process a single queued email
    # ──────────────────────────────────────────────────────────────
    if event.get("Records") and event["Records"][0].get("eventSource") == "aws:sqs":
        print("=== MODE 1: SQS — processing single email ===")

        for record in event["Records"]:
            body         = json.loads(record["body"])
            msg          = body["msg"]
            mailbox      = body["mailbox"]
            bucket_name  = body["bucket_name"]
            access_token = body["access_token"]

            headers = {
                "Authorization": f"Bearer {access_token}",
                "Accept":        "application/json",
            }

            subject = msg.get("subject", "no-subject")
            print(f"\n📧 Processing from queue: {subject}")

            try:
                result = store_email_to_s3(msg, headers, bucket_name, mailbox)
                if result["skipped"]:
                    for skip in result["skipped"]:
                        print(f"  ⚠️ Skipped: {skip['reason']}")
                print(f"  ✅ Done: saved={result['attachments']}")
            except Exception as e:
                print(f"  ❌ Failed: {e}")
                raise

        return {"statusCode": 200, "body": "SQS record processed"}


    # ──────────────────────────────────────────────────────────────
    # MODE 2: EventBridge / manual trigger — fetch emails → queue them
    # ──────────────────────────────────────────────────────────────
    print("=== MODE 2: SCHEDULER — fetching emails and queuing ===")

    if not SQS_QUEUE_URL:
        raise EnvironmentError("Missing required env var: SQS_QUEUE_URL")

    secret_name = os.environ.get("SECRET_NAME")
    bucket_name = os.environ.get("s3_bucket")

    if not secret_name:
        raise EnvironmentError("Missing required env var: SECRET_NAME")
    if not bucket_name:
        raise EnvironmentError("Missing required env var: s3_bucket")

    # ── Step 1: Load credentials from Secrets Manager ──
    print("=== STEP 1: Loading credentials from Secrets Manager ===")
    secret = get_secret(secret_name)
    required_keys = ["tenant_id", "client_id", "client_secret", "mailbox"]
    missing = [k for k in required_keys if not secret.get(k)]
    if missing:
        raise Exception(f"Missing keys in secret: {missing}")

    tenant_id     = secret["tenant_id"]
    client_id     = secret["client_id"]
    client_secret = secret["client_secret"]
    mailbox       = secret["mailbox"]
    print("✅ Credentials loaded")

    # ── Step 2: Get access token ──
    print("\n=== STEP 2: Requesting access token ===")
    token_url = f"https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token"
    token_data = {
        "client_id":     client_id,
        "client_secret": client_secret,
        "scope":         "https://graph.microsoft.com/.default",
        "grant_type":    "client_credentials",
    }
    token_response = post_form(token_url, token_data)
    access_token   = token_response.get("access_token")
    if not access_token:
        raise Exception(f"No access token received: {token_response}")
    print(f"✅ Token acquired (length: {len(access_token)})")

    headers = {
        "Authorization": f"Bearer {access_token}",
        "Accept":        "application/json",
    }


    # ── Step 3: Read last processed timestamp from S3 ──
    print("\n=== STEP 3: Checking last processed time ===")
    last_processed = get_last_processed_time(bucket_name)

    # ── Step 4: Fetch emails newer than last processed time ──
    print("\n=== STEP 4: Fetching new emails ===")
    if last_processed:
        filter_query = f"receivedDateTime gt {last_processed}"
    else:
        since        = datetime.now(timezone.utc).strftime("%Y-%m-%dT00:00:00Z")
        filter_query = f"receivedDateTime gt {since}"
        print(f"  First run — fetching emails since {since}")

    query_params = urllib.parse.urlencode({
        "$filter":  filter_query,
        "$top":     "50",
        "$select":  "id,subject,from,receivedDateTime,isRead,hasAttachments,body",
        "$orderby": "receivedDateTime asc",
    })

    graph_url         = f"https://graph.microsoft.com/v1.0/users/{mailbox}/messages?{query_params}"
    messages_response = get_json(graph_url, headers)
    messages          = messages_response.get("value", [])
    print(f"New emails found: {len(messages)}")

    if not messages:
        print("No new emails since last run.")
        return {
            "statusCode": 200,
            "body": json.dumps({"message": "No new emails found", "results": []}),
        }


    # ── Step 5: Queue each email to SQS ──
    # V2 CHANGE: Body MUST have payment keywords to be queued,
    # regardless of whether it has attachments or not.
    print("\n=== STEP 5: Queuing emails to SQS ===")
    results          = []
    latest_timestamp = last_processed

    for msg in messages:
        subject         = msg.get("subject", "no-subject")
        received        = msg.get("receivedDateTime", "")
        has_attachments = msg.get("hasAttachments", False)

        if not latest_timestamp or received > latest_timestamp:
            latest_timestamp = received

        # ── If email has attachments, queue it (attachment keywords checked in store_email_to_s3)
        # ── If no attachments, check body keywords before queuing ──
        if not has_attachments:
            body_text        = msg.get("body", {}).get("content", "")
            has_payment_body = is_payment_advice_email(body_text) or is_valid_payment_advice_content(body_text)
            if not has_payment_body:
                print(f"\n⏭️ Skipping (no attachments and body has no payment keywords): {subject}")
                _log_rejection_lambda1(
                    f"email_{msg.get('id', '')[:16]}",
                    "No attachments and body has no payment keywords — not queued for processing",
                    mail_id=msg.get("from", {}).get("emailAddress", {}).get("address", ""),
                    mail_received_date=utc_to_ist_display(received),
                    subject=subject,
                    from_address=msg.get("from", {}).get("emailAddress", {}).get("address", ""),
                )
                continue

        print(f"\n📨 Queuing: {subject} ({received})")
        try:
            # Strip large body content to keep SQS message under 256KB
            # Mode 1 only needs: id, subject, from, receivedDateTime, hasAttachments
            # The body was already checked for keywords above — no need to send it again
            msg_for_queue = dict(msg)
            body_content = msg_for_queue.get("body", {}).get("content", "")
            if len(body_content) > 50000:
                # Truncate body to 50KB for SQS — keep enough for keyword check in Mode 1
                msg_for_queue["body"] = {
                    "contentType": msg_for_queue.get("body", {}).get("contentType", "text"),
                    "content": body_content[:50000],
                }
                print(f"  ℹ️ Email body truncated for SQS (was {len(body_content)} bytes)")

            sqs_payload = json.dumps({
                "msg":          msg_for_queue,
                "mailbox":      mailbox,
                "bucket_name":  bucket_name,
                "access_token": access_token,
            })

            # Final size check — SQS limit is 256KB (262144 bytes)
            if len(sqs_payload.encode("utf-8")) > 250000:
                # Still too large — strip body entirely
                msg_for_queue["body"] = {
                    "contentType": msg_for_queue.get("body", {}).get("contentType", "text"),
                    "content": "",
                }
                sqs_payload = json.dumps({
                    "msg":          msg_for_queue,
                    "mailbox":      mailbox,
                    "bucket_name":  bucket_name,
                    "access_token": access_token,
                })
                print(f"  ℹ️ Email body stripped entirely for SQS (message too large)")

            sqs.send_message(
                QueueUrl=SQS_QUEUE_URL,
                MessageBody=sqs_payload,
            )
            print(f"  ✅ Queued successfully")
            results.append({"subject": subject, "status": "queued"})
        except Exception as e:
            print(f"  ❌ Failed to queue: {e}")
            _log_rejection_lambda1(
                f"email_{msg.get('id', '')[:16]}",
                f"Failed to queue to SQS: {str(e)[:200]}",
                mail_id=msg.get("from", {}).get("emailAddress", {}).get("address", ""),
                mail_received_date=utc_to_ist_display(received),
                subject=subject,
                from_address=msg.get("from", {}).get("emailAddress", {}).get("address", ""),
            )
            results.append({"subject": subject, "status": "failed", "error": str(e)})

    # ── Step 6: Save latest timestamp + 1 second ──
    if latest_timestamp and latest_timestamp != last_processed:
        print("\n=== STEP 6: Saving last processed time ===")
        try:
            next_timestamp = increment_timestamp(latest_timestamp)
            save_last_processed_time(bucket_name, next_timestamp)
            print(f"  Next run will fetch emails after: {next_timestamp}")
        except Exception as e:
            print(f"  ❌ Failed to save state: {e}")

    return {
        "statusCode": 200,
        "body": json.dumps(
            {
                "message": f"Queued {len(results)} emails to SQS",
                "results": results,
            },
            indent=2,
        ),
    }