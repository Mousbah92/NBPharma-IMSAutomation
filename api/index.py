"""
NBPharma IMS — Vercel Serverless Function
==========================================
Accepts distributor Excel files via HTTP POST,
processes them, and pushes data to Dataverse.

Auth: Refresh token (no App Registration needed).
"""

import os
import io
import json
import time
import logging
from datetime import datetime
from collections import defaultdict
from http.server import BaseHTTPRequestHandler

import msal
import requests as http_requests
import openpyxl

# =============================================================================
# CONFIGURATION
# =============================================================================

DATAVERSE_URL = os.environ.get("DATAVERSE_URL", "https://ppcustomprojects.crm4.dynamics.com")
API_URL = f"{DATAVERSE_URL}/api/data/v9.2"

REFRESH_TOKEN = os.environ.get("REFRESH_TOKEN", "")
CLIENT_ID = "51f81489-12ee-4a9e-aaae-a2591f45987d"
AUTHORITY = "https://login.microsoftonline.com/organizations"

API_SECRET_KEY = os.environ.get("API_SECRET_KEY", "nbpharma-ims-poc-2025")

SCOPES = [f"{DATAVERSE_URL}/.default"]

TABLE_SKU_BATCHES = "ma_imsskubatches"
TABLE_DISTRIBUTORS = "ma_distributors"
TABLE_PRODUCTS = "ma_imsproducts"

UNICARE_PRODUCT_MAP = {
    "U190001": {"nbp_desc": "CIMZIA 200mg/ml 2 Prefilled Syringes ME", "brand": "CIMZIA"},
    "U190003": {"nbp_desc": "Orladeyo 150mg Caps 28's US1", "brand": "ORLADEYO"},
    "U190004": {"nbp_desc": "Livmarli oral solution 30ml bottle, 1's", "brand": "LIVMARLI"},
    "U190005": {"nbp_desc": "Bimzelx Inj 160mg/ml 2 Prefilled Injectors ME1", "brand": "BIMZELX"},
}

logging.basicConfig(level=logging.INFO)
log = logging.getLogger(__name__)

# =============================================================================
# AUTH
# =============================================================================
_token_cache = {"token": None, "expires_at": 0}


def get_token():
    now = time.time()
    if _token_cache["token"] and _token_cache["expires_at"] > now + 60:
        return _token_cache["token"]

    if not REFRESH_TOKEN:
        raise Exception("REFRESH_TOKEN env var not set.")

    app_client = msal.PublicClientApplication(client_id=CLIENT_ID, authority=AUTHORITY)
    result = app_client.acquire_token_by_refresh_token(refresh_token=REFRESH_TOKEN, scopes=SCOPES)

    if "access_token" in result:
        _token_cache["token"] = result["access_token"]
        _token_cache["expires_at"] = now + result.get("expires_in", 3600)
        return result["access_token"]

    raise Exception(f"Auth failed: {result.get('error_description', 'Unknown')}")


def get_headers():
    return {
        "Authorization": f"Bearer {get_token()}",
        "Content-Type": "application/json",
        "OData-MaxVersion": "4.0",
        "OData-Version": "4.0",
        "Accept": "application/json",
    }


# =============================================================================
# DATAVERSE HELPERS
# =============================================================================
_odata_cache = {}


def resolve_odata_set(logical_name):
    if logical_name in _odata_cache:
        return _odata_cache[logical_name]
    for suffix in ["s", "es", "set", ""]:
        candidate = f"{logical_name}{suffix}"
        resp = http_requests.get(f"{API_URL}/{candidate}?$top=1", headers=get_headers(), timeout=10)
        if resp.status_code == 200:
            _odata_cache[logical_name] = candidate
            return candidate
    return f"{logical_name}s"


def discover_valid_columns(table_name):
    url = (f"{API_URL}/EntityDefinitions(LogicalName='{table_name}')"
           f"/Attributes?$select=LogicalName,AttributeType"
           f"&$filter=AttributeType ne Microsoft.Dynamics.CRM.AttributeTypeCode'Virtual'")
    resp = http_requests.get(url, headers=get_headers(), timeout=15)
    columns = set()
    primary_name = None
    if resp.status_code == 200:
        for a in resp.json().get("value", []):
            ln = a["LogicalName"]
            if ln.startswith("ma_"):
                columns.add(ln)
                if a["AttributeType"] == "String" and not primary_name:
                    if "batchid" in ln or "name" in ln:
                        primary_name = ln
    return columns, primary_name


def safe_payload(payload, valid_columns):
    clean = {}
    for k, v in payload.items():
        if "@odata.bind" in k:
            clean[k] = v
        elif k.startswith("ma_") and k in valid_columns:
            clean[k] = v
    return clean


# =============================================================================
# PARSERS
# =============================================================================

def parse_closing_stock(file_bytes):
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes), read_only=True)
    ws = wb.active
    batches = {}
    for row in ws.iter_rows(min_row=2, values_only=True):
        material = str(row[0]).strip() if row[0] else None
        if not material:
            continue
        batch_nb = str(row[4]).strip() if row[4] else ""
        if not batch_nb:
            continue
        from_date = row[9]
        if isinstance(from_date, str):
            parts = from_date.split(".")
            month, year = int(parts[1]), int(parts[2])
        elif isinstance(from_date, datetime):
            month, year = from_date.month, from_date.year
        else:
            month, year = 9, 2025
        expiry = row[6]
        expiry_str = expiry.strftime("%Y-%m-%d") if isinstance(expiry, datetime) else (str(expiry)[:10] if expiry else None)
        batches[batch_nb] = {
            "material": material, "description": str(row[1] or ""),
            "opening": int(row[7]) if row[7] is not None else 0,
            "closing": int(row[8]) if row[8] is not None else 0,
            "expiry": expiry_str, "month": month, "year": year,
            "nbp_info": UNICARE_PRODUCT_MAP.get(material, {}),
        }
    wb.close()
    return batches


def parse_mtd_sales(file_bytes):
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes), read_only=True)
    ws = wb.active
    batches = defaultdict(lambda: {"qty": 0, "value": 0, "material": "", "name": "", "brand": "", "expiry": None})
    for row in ws.iter_rows(min_row=2, values_only=True):
        batch = str(row[18] or "").strip()
        if not batch:
            continue
        batches[batch]["qty"] += (row[10] or 0)
        batches[batch]["value"] += (row[11] or 0)
        batches[batch]["material"] = str(row[7] or "")
        batches[batch]["name"] = str(row[8] or "")
        batches[batch]["brand"] = str(row[21] or "")
        if row[19] and isinstance(row[19], datetime):
            batches[batch]["expiry"] = row[19].strftime("%Y-%m-%d")
    wb.close()
    return dict(batches)


def detect_file_type(filename):
    fn = filename.lower()
    if "closing" in fn or "stock" in fn:
        return "closing_stock"
    elif "mtd" in fn or "sales" in fn:
        return "mtd_sales"
    return "unknown"


# =============================================================================
# CORE PROCESSING
# =============================================================================

def process_files(files_dict):
    results = {"status": "processing", "steps": [], "patched": 0, "created": 0, "failed": 0, "errors": []}

    try:
        stock_data = parse_closing_stock(files_dict["closing_stock"]) if "closing_stock" in files_dict else {}
        mtd_data = parse_mtd_sales(files_dict["mtd_sales"]) if "mtd_sales" in files_dict else {}

        if stock_data:
            results["steps"].append(f"Parsed {len(stock_data)} batches from Closing Stock")
        if mtd_data:
            results["steps"].append(f"Parsed {len(mtd_data)} batches from MTD Sales")

        all_batch_numbers = set(stock_data.keys()) | set(mtd_data.keys())
        if not all_batch_numbers:
            results["status"] = "error"
            results["errors"].append("No batch data found")
            return results

        batch_set = resolve_odata_set(TABLE_SKU_BATCHES)
        dist_set = resolve_odata_set(TABLE_DISTRIBUTORS)
        prod_set = resolve_odata_set(TABLE_PRODUCTS)

        # Lookup Unicare
        params = {"$filter": "contains(ma_distributorname,'Unicare')", "$top": 5}
        resp = http_requests.get(f"{API_URL}/{dist_set}", headers=get_headers(), params=params, timeout=15)
        dist_guid = None
        if resp.status_code == 200 and resp.json().get("value"):
            dist_guid = resp.json()["value"][0]["ma_distributorsid"]

        if not dist_guid:
            results["status"] = "error"
            results["errors"].append("Unicare not found")
            return results

        # Lookup products
        resp = http_requests.get(f"{API_URL}/{prod_set}?$top=5000", headers=get_headers(), timeout=30)
        all_products = resp.json().get("value", []) if resp.status_code == 200 else []

        brand_guids = {}
        for info in UNICARE_PRODUCT_MAP.values():
            brand_lower = info["brand"].lower()
            for p in all_products:
                for k, v in p.items():
                    if k.startswith("ma_") and isinstance(v, str) and brand_lower in v.lower():
                        brand_guids[info["brand"]] = p["ma_imsproductsid"]
                        break
                if info["brand"] in brand_guids:
                    break

        valid_columns, primary_name = discover_valid_columns(TABLE_SKU_BATCHES)

        # Find existing vs missing
        found_batches = {}
        missing_batches = []
        for batch_nb in sorted(all_batch_numbers):
            url = f"{API_URL}/{batch_set}?$filter=ma_batchnb eq '{batch_nb}'&$top=100"
            resp = http_requests.get(url, headers=get_headers(), timeout=15)
            if resp.status_code == 200:
                records = resp.json().get("value", [])
                if records:
                    found_batches[batch_nb] = records
                else:
                    missing_batches.append(batch_nb)
            time.sleep(0.05)

        def build_payload(batch_nb):
            payload = {
                "ma_Distributor@odata.bind": f"/{dist_set}({dist_guid})",
                "ma_month": stock_data.get(batch_nb, {}).get("month", 9),
                "ma_year": stock_data.get(batch_nb, {}).get("year", 2025),
            }
            if batch_nb in stock_data:
                sd = stock_data[batch_nb]
                payload["ma_quantity"] = sd["closing"]
                payload["ma_openingqty"] = sd["opening"]
                payload["ma_closingqty"] = sd["closing"]
                nbp_info = sd.get("nbp_info", {})
                if nbp_info:
                    payload["ma_itemno"] = nbp_info["nbp_desc"]
                if sd["expiry"]:
                    payload["ma_expirydate"] = sd["expiry"]
                brand = nbp_info.get("brand", "")
                if brand in brand_guids:
                    payload["ma_Product@odata.bind"] = f"/{prod_set}({brand_guids[brand]})"
            elif batch_nb in mtd_data:
                md = mtd_data[batch_nb]
                payload["ma_quantity"] = int(md["qty"])
                payload["ma_itemno"] = md["name"]
                payload["ma_month"] = 9
                payload["ma_year"] = 2025
                if md["expiry"]:
                    payload["ma_expirydate"] = md["expiry"]
                brand = md.get("brand", "")
                if brand in brand_guids:
                    payload["ma_Product@odata.bind"] = f"/{prod_set}({brand_guids[brand]})"
            return payload

        # PATCHes
        for batch_nb, records in found_batches.items():
            payload = safe_payload(build_payload(batch_nb), valid_columns)
            for rec in records:
                resp = http_requests.patch(
                    f"{API_URL}/{batch_set}({rec['ma_imsskubatchesid']})",
                    headers=get_headers(), json=payload, timeout=30,
                )
                if resp.status_code in (200, 204):
                    results["patched"] += 1
                else:
                    results["failed"] += 1
                    results["errors"].append(f"PATCH {batch_nb}: {resp.status_code}")
                time.sleep(0.1)

        # CREATEs
        for batch_nb in missing_batches:
            payload = build_payload(batch_nb)
            payload["ma_batchnb"] = batch_nb
            brand = ""
            if batch_nb in stock_data:
                brand = stock_data[batch_nb].get("nbp_info", {}).get("brand", "")
            elif batch_nb in mtd_data:
                brand = mtd_data[batch_nb].get("brand", "")
            if primary_name:
                payload[primary_name] = f"{brand}_{batch_nb}_UNI_{payload.get('ma_year', 2025)}-{payload.get('ma_month', 9):02d}"
            payload = safe_payload(payload, valid_columns)
            resp = http_requests.post(f"{API_URL}/{batch_set}", headers=get_headers(), json=payload, timeout=30)
            if resp.status_code in (200, 201, 204):
                results["created"] += 1
            else:
                results["failed"] += 1
                results["errors"].append(f"CREATE {batch_nb}: {resp.status_code}")
            time.sleep(0.1)

        results["status"] = "success"
        results["steps"].append(f"Done: {results['patched']} patched, {results['created']} created")

    except Exception as e:
        results["status"] = "error"
        results["errors"].append(str(e))

    return results


# =============================================================================
# VERCEL HANDLER (Flask-like routing via BaseHTTPRequestHandler)
# =============================================================================

from flask import Flask, request as flask_request, jsonify

app = Flask(__name__)


@app.route("/", methods=["GET"])
def index():
    return jsonify({
        "service": "NBPharma IMS Batch Enrichment",
        "status": "running",
        "endpoints": {
            "POST /api/process": "Upload Excel files",
            "GET /api/health": "Health check",
        },
    })


@app.route("/api/health", methods=["GET"])
def health():
    try:
        get_token()
        resp = http_requests.get(f"{API_URL}/WhoAmI", headers=get_headers(), timeout=10)
        connected = resp.status_code == 200
    except Exception:
        connected = False
    return jsonify({"status": "healthy" if connected else "degraded", "dataverse_connected": connected})


@app.route("/api/process", methods=["POST"])
def process():
    api_key = flask_request.headers.get("X-API-Key", "")
    if api_key != API_SECRET_KEY:
        return jsonify({"error": "Unauthorized"}), 401

    if not flask_request.files and not flask_request.data:
        return jsonify({"error": "No files uploaded"}), 400

    files_dict = {}
    file_info = []

    # Handle multipart form data
    if flask_request.files:
        for key in flask_request.files:
            f = flask_request.files[key]
            filename = f.filename or key
            file_bytes = f.read()
            file_type = detect_file_type(filename)
            file_info.append({"name": filename, "type": file_type, "size": len(file_bytes)})
            if file_type in ("closing_stock", "mtd_sales"):
                files_dict[file_type] = file_bytes
            else:
                files_dict.setdefault("closing_stock", file_bytes)

    # Handle raw binary with filename in header
    elif flask_request.data:
        filename = flask_request.headers.get("X-Filename", "file.xlsx")
        file_bytes = flask_request.data
        file_type = detect_file_type(filename)
        file_info.append({"name": filename, "type": file_type, "size": len(file_bytes)})
        if file_type in ("closing_stock", "mtd_sales"):
            files_dict[file_type] = file_bytes
        else:
            files_dict["closing_stock"] = file_bytes

    results = process_files(files_dict)
    results["files_received"] = file_info
    return jsonify(results), 200 if results["status"] == "success" else 500


# Vercel expects this
handler = app
