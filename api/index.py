"""
NBPharma IMS — Vercel Serverless Function (Complete Pipeline)
==============================================================
Accepts distributor Excel files via HTTP POST and runs the FULL pipeline:

  MTD Sales Report ->
    1. Enrich ma_imsskubatches (batch-level: PATCH/CREATE)
    2. Upload ma_imsproductdata (metrics: Private, Tender, IMS Total, FOC)

  Closing Stock Report ->
    1. Enrich ma_imsskubatches (batch-level: PATCH/CREATE)
    2. Upload ma_imsproductdata (metrics: Stock Open, Stock Close)

Auth: Refresh token (no App Registration needed).
"""

import os, io, json, time, logging
from datetime import datetime
from collections import defaultdict
import msal
import requests as http_requests
import openpyxl

DATAVERSE_URL = os.environ.get("DATAVERSE_URL", "https://ppcustomprojects.crm4.dynamics.com")
API_URL = f"{DATAVERSE_URL}/api/data/v9.2"
REFRESH_TOKEN = os.environ.get("REFRESH_TOKEN", "")
CLIENT_ID = "51f81489-12ee-4a9e-aaae-a2591f45987d"
AUTHORITY = "https://login.microsoftonline.com/organizations"
API_SECRET_KEY = os.environ.get("API_SECRET_KEY", "nbpharma-ims-poc-2025")
SCOPES = [f"{DATAVERSE_URL}/.default"]

TABLE_SKU_BATCHES = "ma_imsskubatches"
TABLE_PRODUCT_DATA = "ma_imsproductdata"
TABLE_DISTRIBUTORS = "ma_distributors"
TABLE_PRODUCTS = "ma_imsproducts"

UNICARE_PRODUCT_MAP = {
    "U190001": {"nbp_desc": "CIMZIA 200mg/ml 2 Prefilled Syringes ME", "brand": "CIMZIA"},
    "U190003": {"nbp_desc": "Orladeyo 150mg Caps 28's US1", "brand": "ORLADEYO"},
    "U190004": {"nbp_desc": "Livmarli oral solution 30ml bottle, 1's", "brand": "LIVMARLI"},
    "U190005": {"nbp_desc": "Bimzelx Inj 160mg/ml 2 Prefilled Injectors ME1", "brand": "BIMZELX"},
}
MATERIAL_TO_BRAND = {k: v["brand"] for k, v in UNICARE_PRODUCT_MAP.items()}

logging.basicConfig(level=logging.INFO)
log = logging.getLogger(__name__)

# ── AUTH ──
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
    return {"Authorization": f"Bearer {get_token()}", "Content-Type": "application/json",
            "OData-MaxVersion": "4.0", "OData-Version": "4.0", "Accept": "application/json"}

# ── DATAVERSE HELPERS ──
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
                    if "batchid" in ln or "dataid" in ln or "name" in ln:
                        primary_name = ln
    return columns, primary_name

def discover_metric_picklist():
    url = (f"{API_URL}/EntityDefinitions(LogicalName='{TABLE_PRODUCT_DATA}')"
           f"/Attributes/Microsoft.Dynamics.CRM.PicklistAttributeMetadata"
           f"?$filter=LogicalName eq 'ma_metric'&$expand=OptionSet")
    resp = http_requests.get(url, headers=get_headers(), timeout=15)
    metric_map = {}
    if resp.status_code == 200:
        data = resp.json().get("value", [])
        if data:
            for opt in data[0].get("OptionSet", {}).get("Options", []):
                lbl = (opt.get("Label", {}).get("UserLocalizedLabel") or {}).get("Label", "")
                metric_map[lbl.upper()] = opt["Value"]
                metric_map[opt["Value"]] = lbl
    return metric_map

def safe_payload(payload, valid_columns):
    clean = {}
    for k, v in payload.items():
        if "@odata.bind" in k:
            clean[k] = v
        elif k.startswith("ma_") and k in valid_columns:
            clean[k] = v
    return clean

def find_product_guid(brand_name, all_products):
    bl = brand_name.lower()
    for p in all_products:
        for k, v in p.items():
            if k.startswith("ma_") and isinstance(v, str) and bl in v.lower():
                return p["ma_imsproductsid"]
    return None

# ── PARSERS ──

def parse_closing_stock(file_bytes):
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes), read_only=True)
    ws = wb.active
    batches = {}
    product_agg = defaultdict(lambda: {"opening": 0, "closing": 0, "description": "", "month": 9, "year": 2025})
    for row in ws.iter_rows(min_row=2, values_only=True):
        material = str(row[0]).strip() if row[0] else None
        if not material:
            continue
        batch_nb = str(row[4]).strip() if row[4] else ""
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
        opening = int(row[7]) if row[7] is not None else 0
        closing = int(row[8]) if row[8] is not None else 0
        if batch_nb:
            batches[batch_nb] = {"material": material, "description": str(row[1] or ""),
                "opening": opening, "closing": closing, "expiry": expiry_str,
                "month": month, "year": year, "nbp_info": UNICARE_PRODUCT_MAP.get(material, {})}
        product_agg[material]["opening"] += opening
        product_agg[material]["closing"] += closing
        product_agg[material]["description"] = str(row[1] or "")
        product_agg[material]["month"] = month
        product_agg[material]["year"] = year
    wb.close()
    return batches, dict(product_agg)

def parse_mtd_sales(file_bytes):
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes), read_only=True)
    ws = wb.active
    batches = defaultdict(lambda: {"qty": 0, "value": 0, "material": "", "name": "", "brand": "", "expiry": None})
    product_metrics = defaultdict(lambda: {"private_qty": 0, "private_val": 0, "tender_qty": 0, "tender_val": 0,
                                           "foc_qty": 0, "foc_val": 0, "total_qty": 0, "total_val": 0})
    for row in ws.iter_rows(min_row=2, values_only=True):
        brand = str(row[21] or "").strip()
        batch = str(row[18] or "").strip()
        custgroup = str(row[2] or "").strip().lower()
        billt = str(row[12] or "").strip()
        try:
            qty = float(row[10]) if row[10] is not None else 0
            val = float(row[11]) if row[11] is not None else 0
        except (ValueError, TypeError):
            continue
        if not brand:
            continue
        if batch:
            batches[batch]["qty"] += qty
            batches[batch]["value"] += val
            batches[batch]["material"] = str(row[7] or "")
            batches[batch]["name"] = str(row[8] or "")
            batches[batch]["brand"] = brand
            if row[19] and isinstance(row[19], datetime):
                batches[batch]["expiry"] = row[19].strftime("%Y-%m-%d")
        if billt in ("ZRE5", "ZUFA"):
            product_metrics[brand]["foc_qty"] += abs(qty)
            product_metrics[brand]["foc_val"] += abs(val)
            continue
        if custgroup == "private":
            product_metrics[brand]["private_qty"] += qty
            product_metrics[brand]["private_val"] += val
        elif custgroup == "institutional":
            product_metrics[brand]["tender_qty"] += qty
            product_metrics[brand]["tender_val"] += val
        else:
            product_metrics[brand]["private_qty"] += qty
            product_metrics[brand]["private_val"] += val
        product_metrics[brand]["total_qty"] += qty
        product_metrics[brand]["total_val"] += val
    wb.close()
    return dict(batches), dict(product_metrics)

def detect_file_type(filename):
    fn = filename.lower()
    if "closing" in fn or "stock" in fn:
        return "closing_stock"
    elif "mtd" in fn or "sales" in fn:
        return "mtd_sales"
    return "unknown"

# ── PIPELINE A: ENRICH BATCHES ──

def pipeline_enrich_batches(stock_batches, mtd_batches, dist_guid, brand_guids,
                            batch_set, dist_set, prod_set, valid_columns, primary_name):
    results = {"patched": 0, "created": 0, "failed": 0, "errors": []}
    all_batch_numbers = set(stock_batches.keys()) | set(mtd_batches.keys())
    if not all_batch_numbers:
        return results
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

    def build_batch_payload(batch_nb):
        payload = {"ma_Distributor@odata.bind": f"/{dist_set}({dist_guid})"}
        if batch_nb in stock_batches:
            sd = stock_batches[batch_nb]
            payload.update({"ma_month": sd.get("month", 9), "ma_year": sd.get("year", 2025),
                "ma_quantity": sd["closing"], "ma_openingqty": sd["opening"], "ma_closingqty": sd["closing"]})
            nbp_info = sd.get("nbp_info", {})
            if nbp_info:
                payload["ma_itemno"] = nbp_info["nbp_desc"]
            if sd.get("expiry"):
                payload["ma_expirydate"] = sd["expiry"]
            brand = nbp_info.get("brand", "")
            if brand in brand_guids:
                payload["ma_Product@odata.bind"] = f"/{prod_set}({brand_guids[brand]})"
        elif batch_nb in mtd_batches:
            md = mtd_batches[batch_nb]
            payload.update({"ma_month": 9, "ma_year": 2025, "ma_quantity": int(md["qty"]), "ma_itemno": md["name"]})
            if md.get("expiry"):
                payload["ma_expirydate"] = md["expiry"]
            brand = md.get("brand", "")
            if brand in brand_guids:
                payload["ma_Product@odata.bind"] = f"/{prod_set}({brand_guids[brand]})"
        return payload

    for batch_nb, records in found_batches.items():
        payload = safe_payload(build_batch_payload(batch_nb), valid_columns)
        for rec in records:
            resp = http_requests.patch(f"{API_URL}/{batch_set}({rec['ma_imsskubatchesid']})",
                headers=get_headers(), json=payload, timeout=30)
            if resp.status_code in (200, 204):
                results["patched"] += 1
            else:
                results["failed"] += 1
                results["errors"].append(f"PATCH {batch_nb}: {resp.status_code}")
            time.sleep(0.1)

    for batch_nb in missing_batches:
        payload = build_batch_payload(batch_nb)
        payload["ma_batchnb"] = batch_nb
        brand = ""
        if batch_nb in stock_batches:
            brand = stock_batches[batch_nb].get("nbp_info", {}).get("brand", "")
        elif batch_nb in mtd_batches:
            brand = mtd_batches[batch_nb].get("brand", "")
        if primary_name:
            yr = payload.get("ma_year", 2025)
            mn = payload.get("ma_month", 9)
            payload[primary_name] = f"{brand}_{batch_nb}_UNI_{yr}-{mn:02d}"
        payload = safe_payload(payload, valid_columns)
        resp = http_requests.post(f"{API_URL}/{batch_set}", headers=get_headers(), json=payload, timeout=30)
        if resp.status_code in (200, 201, 204):
            results["created"] += 1
        else:
            results["failed"] += 1
            results["errors"].append(f"CREATE {batch_nb}: {resp.status_code}")
        time.sleep(0.1)
    return results

# ── PIPELINE B: UPLOAD PRODUCT METRICS ──

def pipeline_upload_metrics(file_type, stock_product_agg, mtd_product_metrics,
                            dist_guid, brand_guids, dist_set, prod_set):
    results = {"metrics_created": 0, "metrics_failed": 0, "errors": []}
    metric_map = discover_metric_picklist()
    if not metric_map:
        results["errors"].append("Could not discover metric picklist")
        return results
    pd_columns, pd_primary = discover_valid_columns(TABLE_PRODUCT_DATA)
    if not pd_primary:
        pd_primary = "ma_dataid"
    pd_set = resolve_odata_set(TABLE_PRODUCT_DATA)

    METRICS = {}
    for label_upper, val in metric_map.items():
        if not isinstance(label_upper, str):
            continue
        if "DISTRIBUTORS OPENING STOCK" in label_upper:
            METRICS["stock_open"] = val
        elif "REPORTED CLOSING STOCK" in label_upper:
            METRICS["stock_close"] = val
        elif "PRIVATE MARKET SALES" in label_upper:
            METRICS["private_sales"] = val
        elif "SUPPLY TO LOCAL TENDERS" in label_upper:
            METRICS["tender_sales"] = val
        elif "IMS ACT" in label_upper and "QTTY" in label_upper:
            METRICS["ims_total_qty"] = val
        elif "FOC" in label_upper and "SAMPLE" in label_upper:
            METRICS["foc_samples"] = val
        elif "TOTAL NON SALES" in label_upper:
            METRICS["total_non_sales"] = val

    log.info(f"Mapped {len(METRICS)} metric keys: {list(METRICS.keys())}")
    records = []

    def add_metric(brand, metric_key, value, month, year):
        if metric_key not in METRICS or brand not in brand_guids or value == 0:
            return
        records.append({
            "ma_Distributor@odata.bind": f"/{dist_set}({dist_guid})",
            "ma_Product@odata.bind": f"/{prod_set}({brand_guids[brand]})",
            "ma_metric": METRICS[metric_key], "ma_month": month, "ma_year": year,
            "ma_value": int(round(value)),
            pd_primary: f"{metric_key}_{brand}_UNI_{year}-{month:02d}",
        })

    if file_type == "closing_stock" and stock_product_agg:
        for material, vals in stock_product_agg.items():
            brand = MATERIAL_TO_BRAND.get(material)
            if not brand:
                continue
            month = vals.get("month", 9)
            year = vals.get("year", 2025)
            add_metric(brand, "stock_open", vals["opening"], month, year)
            add_metric(brand, "stock_close", vals["closing"], month, year)

    if file_type == "mtd_sales" and mtd_product_metrics:
        for brand, v in mtd_product_metrics.items():
            add_metric(brand, "private_sales", v["private_qty"], 9, 2025)
            add_metric(brand, "tender_sales", v["tender_qty"], 9, 2025)
            add_metric(brand, "ims_total_qty", v["total_qty"], 9, 2025)
            add_metric(brand, "foc_samples", v["foc_qty"], 9, 2025)
            add_metric(brand, "total_non_sales", v["foc_qty"], 9, 2025)

    log.info(f"Prepared {len(records)} metric records for ma_imsproductdata")
    for rec in records:
        clean = safe_payload(rec, pd_columns)
        resp = http_requests.post(f"{API_URL}/{pd_set}", headers=get_headers(), json=clean, timeout=30)
        if resp.status_code in (200, 201, 204):
            results["metrics_created"] += 1
        else:
            results["metrics_failed"] += 1
            results["errors"].append(f"METRIC {rec.get(pd_primary, '?')}: {resp.status_code}")
        time.sleep(0.1)
    return results

# ── MAIN PROCESSING ──

def process_file(file_type, file_bytes):
    results = {"status": "processing", "file_type": file_type, "steps": [],
        "batches_patched": 0, "batches_created": 0, "metrics_created": 0, "failed": 0, "errors": []}
    try:
        stock_batches, stock_product_agg, mtd_batches, mtd_product_metrics = {}, {}, {}, {}
        if file_type == "closing_stock":
            stock_batches, stock_product_agg = parse_closing_stock(file_bytes)
            results["steps"].append(f"Parsed Closing Stock: {len(stock_batches)} batches, {len(stock_product_agg)} products")
        elif file_type == "mtd_sales":
            mtd_batches, mtd_product_metrics = parse_mtd_sales(file_bytes)
            results["steps"].append(f"Parsed MTD Sales: {len(mtd_batches)} batches, {len(mtd_product_metrics)} products")

        batch_set = resolve_odata_set(TABLE_SKU_BATCHES)
        dist_set = resolve_odata_set(TABLE_DISTRIBUTORS)
        prod_set = resolve_odata_set(TABLE_PRODUCTS)

        params = {"$filter": "contains(ma_distributorname,'Unicare')", "$top": 5}
        resp = http_requests.get(f"{API_URL}/{dist_set}", headers=get_headers(), params=params, timeout=15)
        dist_guid = None
        if resp.status_code == 200 and resp.json().get("value"):
            dist_guid = resp.json()["value"][0]["ma_distributorsid"]
        if not dist_guid:
            results["status"] = "error"
            results["errors"].append("Unicare not found")
            return results

        resp = http_requests.get(f"{API_URL}/{prod_set}?$top=5000", headers=get_headers(), timeout=30)
        all_products = resp.json().get("value", []) if resp.status_code == 200 else []
        brand_guids = {}
        for info in UNICARE_PRODUCT_MAP.values():
            guid = find_product_guid(info["brand"], all_products)
            if guid:
                brand_guids[info["brand"]] = guid
        for brand in mtd_product_metrics.keys():
            if brand not in brand_guids:
                guid = find_product_guid(brand, all_products)
                if guid:
                    brand_guids[brand] = guid
        results["steps"].append(f"Resolved: Unicare + {len(brand_guids)} products")

        valid_columns, primary_name = discover_valid_columns(TABLE_SKU_BATCHES)

        # PIPELINE A: Enrich batches
        br = pipeline_enrich_batches(stock_batches, mtd_batches, dist_guid, brand_guids,
            batch_set, dist_set, prod_set, valid_columns, primary_name)
        results["batches_patched"] = br["patched"]
        results["batches_created"] = br["created"]
        results["failed"] += br["failed"]
        results["errors"].extend(br["errors"])
        results["steps"].append(f"Batches: {br['patched']} patched, {br['created']} created")

        # PIPELINE B: Upload product metrics
        mr = pipeline_upload_metrics(file_type, stock_product_agg, mtd_product_metrics,
            dist_guid, brand_guids, dist_set, prod_set)
        results["metrics_created"] = mr["metrics_created"]
        results["failed"] += mr["metrics_failed"]
        results["errors"].extend(mr["errors"])
        results["steps"].append(f"Metrics: {mr['metrics_created']} created in ma_imsproductdata")

        results["status"] = "success"
    except Exception as e:
        results["status"] = "error"
        results["errors"].append(str(e))
        log.exception("Processing failed")
    return results

# ── FLASK APP ──
from flask import Flask, request as flask_request, jsonify
app = Flask(__name__)

@app.route("/", methods=["GET"])
def index():
    return jsonify({"service": "NBPharma IMS - Complete Pipeline", "status": "running",
        "supported_files": {"closing_stock": "batches + stock metrics", "mtd_sales": "batches + sales metrics"}})

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

    def is_excel(fn):
        return fn.lower().endswith((".xlsx", ".xls", ".xlsm"))

    files_to_process = []
    if flask_request.files:
        for key in flask_request.files:
            f = flask_request.files[key]
            fn = f.filename or key
            if not is_excel(fn):
                continue
            ft = detect_file_type(fn)
            if ft != "unknown":
                files_to_process.append((ft, f.read(), fn))
    elif flask_request.data:
        fn = flask_request.headers.get("X-Filename", "file.xlsx")
        if not is_excel(fn):
            return jsonify({"status": "skipped", "reason": f"Not Excel: {fn}"}), 200
        ft = detect_file_type(fn)
        if ft == "unknown":
            ft = "closing_stock"
        files_to_process.append((ft, flask_request.data, fn))

    if not files_to_process:
        return jsonify({"status": "skipped", "reason": "No processable Excel files"}), 200

    all_results = []
    for ft, fb, fn in files_to_process:
        log.info(f"Processing: {fn} as {ft}")
        result = process_file(ft, fb)
        result["filename"] = fn
        all_results.append(result)

    tp = sum(r.get("batches_patched", 0) for r in all_results)
    tc = sum(r.get("batches_created", 0) for r in all_results)
    tm = sum(r.get("metrics_created", 0) for r in all_results)
    tf = sum(r.get("failed", 0) for r in all_results)
    errs = []
    for r in all_results:
        errs.extend(r.get("errors", []))

    st = "success" if all(r["status"] == "success" for r in all_results) else "partial"
    if all(r["status"] == "error" for r in all_results):
        st = "error"

    response = {"status": st, "summary": {"files_processed": len(files_to_process),
        "batches_patched": tp, "batches_created": tc, "metrics_created": tm, "failed": tf},
        "files": all_results, "errors": errs[:10]}
    return jsonify(response), 200 if st in ("success", "partial") else 500

handler = app
