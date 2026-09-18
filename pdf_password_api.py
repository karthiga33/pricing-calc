"""
API handler for PDF Password Management.
Add these routes to your AR Automation frontend API (API Gateway + Lambda).

DynamoDB Table: pdf_password_map
  - Partition Key: file_name (String)
  - Attributes: password (String), created_at (String)
"""

import json
import boto3
from datetime import datetime

dynamodb = boto3.resource("dynamodb")
PDF_PASSWORD_TABLE = "pdf_password_map"


def lambda_handler(event, context):
    """Handle API Gateway requests for PDF password management."""
    http_method = event.get("httpMethod", event.get("requestContext", {}).get("http", {}).get("method", ""))
    path = event.get("path", event.get("rawPath", ""))

    # CORS headers
    headers = {
        "Access-Control-Allow-Origin": "*",
        "Access-Control-Allow-Headers": "Content-Type",
        "Access-Control-Allow-Methods": "GET, POST, DELETE, OPTIONS",
        "Content-Type": "application/json",
    }

    if http_method == "OPTIONS":
        return {"statusCode": 200, "headers": headers, "body": ""}

    try:
        if http_method == "GET" and "pdf-passwords" in path:
            return get_all_passwords(headers)

        elif http_method == "POST" and "pdf-passwords" in path:
            body = json.loads(event.get("body", "{}"))
            return save_password(body, headers)

        elif http_method == "DELETE" and "pdf-passwords" in path:
            # file_name from path param or query string
            params = event.get("queryStringParameters") or {}
            file_name = params.get("file_name", "")
            if not file_name:
                # Try path parameter
                path_params = event.get("pathParameters") or {}
                file_name = path_params.get("file_name", "")
            return delete_password(file_name, headers)

        return {
            "statusCode": 404,
            "headers": headers,
            "body": json.dumps({"error": "Not found"}),
        }

    except Exception as e:
        return {
            "statusCode": 500,
            "headers": headers,
            "body": json.dumps({"error": str(e)}),
        }


def get_all_passwords(headers):
    """GET /pdf-passwords — return all entries."""
    table = dynamodb.Table(PDF_PASSWORD_TABLE)
    resp = table.scan()
    items = resp.get("Items", [])

    # Sort by created_at descending (newest first)
    items.sort(key=lambda x: x.get("created_at", ""), reverse=True)

    return {
        "statusCode": 200,
        "headers": headers,
        "body": json.dumps({"passwords": items}),
    }


def save_password(body, headers):
    """POST /pdf-passwords — add or update an entry."""
    file_name = body.get("file_name", "").strip()
    password = body.get("password", "").strip()

    if not file_name or not password:
        return {
            "statusCode": 400,
            "headers": headers,
            "body": json.dumps({"error": "file_name and password are required"}),
        }

    table = dynamodb.Table(PDF_PASSWORD_TABLE)
    table.put_item(Item={
        "file_name": file_name,
        "password": password,
        "created_at": datetime.utcnow().isoformat(),
    })

    return {
        "statusCode": 200,
        "headers": headers,
        "body": json.dumps({"message": f"Password saved for '{file_name}'"}),
    }


def delete_password(file_name, headers):
    """DELETE /pdf-passwords?file_name=xxx — remove an entry."""
    if not file_name:
        return {
            "statusCode": 400,
            "headers": headers,
            "body": json.dumps({"error": "file_name is required"}),
        }

    table = dynamodb.Table(PDF_PASSWORD_TABLE)
    table.delete_item(Key={"file_name": file_name})

    return {
        "statusCode": 200,
        "headers": headers,
        "body": json.dumps({"message": f"Password deleted for '{file_name}'"}),
    }
