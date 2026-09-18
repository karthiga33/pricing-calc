# Bugfix Requirements Document

## Introduction

The Lambda1 email-fetching pipeline saves payment advice attachments to S3 under the `emails/` prefix. When the same attachment content arrives again (due to scheduler re-runs, SQS retries, or the same file being forwarded in multiple emails), the `_make_unique_s3_name` function creates new files with incrementing suffixes (`-1`, `-2`, etc.) instead of detecting the duplicate content and skipping it. This causes the downstream extraction pipeline (Lambda2) to process the same document multiple times, producing duplicate vendor records in the UI.

The fix should use a content-hash/ETag-based deduplication approach (similar to the existing `_is_etag_processed` pattern in `lambda_full.py`) — compute a hash of the attachment bytes before saving, check a DynamoDB dedup table, and skip the file if the same content was already processed by Lambda1.

## Bug Analysis

### Current Behavior (Defect)

1.1 WHEN an attachment with the same filename as an existing S3 object is processed THEN the system appends an incrementing suffix (e.g., `Payment-Advice-1.pdf`, `Payment-Advice-2.pdf`) and saves a new copy to S3

1.2 WHEN the same email is re-processed due to a scheduler re-run or SQS retry THEN the system creates duplicate files in S3 with different names for identical content

1.3 WHEN the same attachment file content is forwarded in a different email THEN the system saves it again under a new suffixed name, resulting in duplicate extraction downstream

### Expected Behavior (Correct)

2.1 WHEN an attachment with the same content (identical bytes) as an existing S3 object is processed THEN the system SHALL compute a content hash (MD5/ETag) of the attachment bytes, check it against a DynamoDB dedup table, and skip saving the file if the hash already exists

2.2 WHEN the same email is re-processed due to a scheduler re-run or SQS retry THEN the system SHALL recognize the attachment content hash has already been recorded and skip it without creating a new file

2.3 WHEN the same attachment file content is forwarded in a different email (possibly under a different filename) THEN the system SHALL detect the content hash already exists in the dedup table and skip saving a duplicate

### Unchanged Behavior (Regression Prevention)

3.1 WHEN two attachments have the same filename but different content (e.g., `Payment-Advice.pdf` from two different vendors) THEN the system SHALL CONTINUE TO save both files with unique names so neither is lost

3.2 WHEN an attachment with a unique content that has never been seen before arrives THEN the system SHALL CONTINUE TO save it to the `emails/` prefix in S3 with appropriate metadata

3.3 WHEN attachments pass the extension filter and content validation stages THEN the system SHALL CONTINUE TO save them and trigger the downstream Lambda2 extraction pipeline

3.4 WHEN multi-customer files are detected THEN the system SHALL CONTINUE TO route them to the `multi-input/` prefix unchanged
