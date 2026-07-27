import cgi
import io
import json
import math
import os
import re
import shutil
import subprocess
import sys
import tempfile
import unicodedata
import uuid
from difflib import SequenceMatcher
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import unquote, urlparse

import pandas as pd
import pdfplumber


ROOT = Path(__file__).resolve().parent
OUTPUT_DIR = ROOT / "outputs"
CACHE_DIR = ROOT / ".cache"
DEFAULT_PRODUCT_FILE = CACHE_DIR / "last_opc_site_product.xlsx"
LAST_INVOICE_FILE = CACHE_DIR / "last_invoice.pdf"
LAST_PO_FILE = CACHE_DIR / "last_po.xlsx"
RUNTIME_PYTHON_PACKAGES = ROOT / ".runtime" / "python-packages"
RUNTIME_JDK_BIN = ROOT / ".runtime" / "jdk17" / "jdk-17.0.19+10" / "bin"
OPENDATALOADER_SRC = ROOT / "vendor_pdf_parser" / "opendataloader-pdf-main" / "python" / "opendataloader-pdf" / "src"
OUTPUT_DIR.mkdir(exist_ok=True)
CACHE_DIR.mkdir(exist_ok=True)

PRODUCT_NAME = "\u4ea7\u54c1\u540d"
COLUMNS = [
    "EAN",
    PRODUCT_NAME,
    "sku",
    "invoice qty",
    "PO Actual received qty",
    "qty gap",
    "invoice price excl",
    "PO price excl",
    "price gap",
]

HTML = """<!doctype html>
<html lang="zh-CN">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>Invoice PO 自动对账</title>
  <style>
    :root { --bg:#f7f7f4; --panel:#fff; --text:#242826; --muted:#66706a; --line:#d9ded8; --accent:#0f766e; --accent-dark:#115e59; --warn:#b45309; --bad:#b91c1c; }
    * { box-sizing: border-box; }
    body { margin:0; font-family:"Segoe UI",Arial,sans-serif; background:var(--bg); color:var(--text); }
    header { padding:24px 28px 14px; border-bottom:1px solid var(--line); background:#fbfbf8; }
    h1 { margin:0; font-size:24px; font-weight:700; letter-spacing:0; }
    main { display:grid; grid-template-columns:minmax(280px,360px) 1fr; gap:18px; padding:18px 28px 28px; }
    section { background:var(--panel); border:1px solid var(--line); border-radius:8px; }
    .upload-panel { padding:18px; align-self:start; }
    .field { margin-bottom:16px; }
    label { display:block; font-size:14px; font-weight:650; margin-bottom:8px; }
    .hint { display:block; color:var(--muted); font-size:12px; line-height:1.4; margin-top:6px; }
    input[type=file] { width:100%; padding:10px; border:1px solid var(--line); border-radius:6px; background:#fff; font-size:13px; }
    .actions { display:flex; gap:10px; align-items:center; margin-top:18px; }
    button,.download { border:0; border-radius:6px; background:var(--accent); color:white; padding:10px 14px; font-weight:700; cursor:pointer; text-decoration:none; font-size:14px; line-height:1.2; }
    button:hover,.download:hover { background:var(--accent-dark); }
    button:disabled { opacity:.6; cursor:wait; }
    .status { color:var(--muted); font-size:13px; line-height:1.45; margin-top:12px; min-height:20px; }
    .results { overflow:hidden; }
    .summary { display:grid; grid-template-columns:repeat(5,minmax(100px,1fr)); gap:1px; background:var(--line); border-bottom:1px solid var(--line); }
    .metric { background:#fff; padding:14px 16px; }
    .metric span { display:block; color:var(--muted); font-size:12px; margin-bottom:5px; }
    .metric strong { font-size:22px; }
    .table-wrap { overflow:auto; max-height:calc(100vh - 220px); }
    table { width:100%; border-collapse:collapse; min-width:1180px; font-size:13px; }
    th,td { padding:9px 10px; border-bottom:1px solid var(--line); text-align:left; vertical-align:top; white-space:nowrap; }
    th { position:sticky; top:0; background:#eef4f2; z-index:1; font-size:12px; }
    td.name { white-space:normal; min-width:260px; }
    .gap-bad { color:var(--bad); font-weight:700; }
    .gap-warn { color:var(--warn); font-weight:700; }
    .empty { padding:40px; color:var(--muted); text-align:center; }
    .error { color:var(--bad); }
    @media (max-width:900px) { main { grid-template-columns:1fr; padding:14px; } header { padding:18px 14px 12px; } .summary { grid-template-columns:repeat(2,minmax(100px,1fr)); } }
  </style>
</head>
<body>
  <header><h1>Invoice PO 自动对账</h1></header>
  <main>
    <section class="upload-panel">
      <form id="form">
        <div class="field">
          <label for="product">OPC_site product</label>
          <input id="product" name="product" type="file" accept=".xlsx,.xls">
          <span class="hint">可选。上传后会保存为默认文件；不上传则自动使用上次保存的 OPC 文件。</span>
        </div>
        <div class="field"><label for="invoice">invoice PDF</label><input id="invoice" name="invoice" type="file" accept=".pdf" required></div>
        <div class="field"><label for="po">PO detail</label><input id="po" name="po" type="file" accept=".xlsx,.xls" required></div>
        <div class="actions"><button id="submit" type="submit">开始对账</button><a id="download" class="download" href="#" style="display:none">下载结果</a></div>
        <div id="status" class="status">上传 invoice 和 PO detail 即可对账；OPC_site product 可沿用上次上传文件。</div>
      </form>
    </section>
    <section class="results">
      <div class="summary">
        <div class="metric"><span>发票行数</span><strong id="mInvoice">0</strong></div>
        <div class="metric"><span>差异行数</span><strong id="mDiff">0</strong></div>
        <div class="metric"><span>数量差异</span><strong id="mQty">0</strong></div>
        <div class="metric"><span>价格差异</span><strong id="mPrice">0</strong></div>
        <div class="metric"><span>金额总差异</span><strong id="mAmountGap">0</strong></div>
      </div>
      <div class="table-wrap">
        <table>
          <thead><tr><th>EAN</th><th>产品名</th><th>sku</th><th>invoice qty</th><th>PO Actual received qty</th><th>qty gap</th><th>invoice price excl</th><th>PO price excl</th><th>price gap</th></tr></thead>
          <tbody id="tbody"><tr><td class="empty" colspan="9">等待上传文件</td></tr></tbody>
        </table>
      </div>
    </section>
  </main>
  <script>
    const form = document.getElementById("form");
    const statusEl = document.getElementById("status");
    const button = document.getElementById("submit");
    const download = document.getElementById("download");
    const tbody = document.getElementById("tbody");
    const fmt = (v) => v === null || v === undefined || v === "" ? "" : v;
    const esc = (s) => String(s ?? "").replace(/[&<>"']/g, c => ({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;"}[c]));
    form.addEventListener("submit", async (event) => {
      event.preventDefault();
      button.disabled = true;
      download.style.display = "none";
      statusEl.className = "status";
      statusEl.textContent = "正在识别和匹配，请稍等...";
      tbody.innerHTML = `<tr><td class="empty" colspan="9">处理中</td></tr>`;
      try {
        const res = await fetch("/reconcile", { method: "POST", body: new FormData(form) });
        const data = await res.json();
        if (!res.ok || !data.ok) throw new Error(data.error || "处理失败");
        document.getElementById("mInvoice").textContent = data.metrics.invoice_rows;
        document.getElementById("mDiff").textContent = data.metrics.diff_rows;
        document.getElementById("mQty").textContent = data.metrics.qty_diff_rows;
        document.getElementById("mPrice").textContent = data.metrics.price_diff_rows;
        document.getElementById("mAmountGap").textContent = data.metrics.amount_gap_total;
        if (!data.rows.length) {
          tbody.innerHTML = `<tr><td class="empty" colspan="9">没有发现数量或价格差异</td></tr>`;
        } else {
          tbody.innerHTML = data.rows.map(row => `
            <tr>
              <td>${esc(fmt(row.EAN))}</td>
              <td class="name">${esc(fmt(row["产品名"]))}</td>
              <td>${esc(fmt(row.sku))}</td>
              <td>${esc(fmt(row["invoice qty"]))}</td>
              <td>${esc(fmt(row["PO Actual received qty"]))}</td>
              <td class="${Number(row["qty gap"] || 0) !== 0 ? "gap-bad" : ""}">${esc(fmt(row["qty gap"]))}</td>
              <td>${esc(fmt(row["invoice price excl"]))}</td>
              <td>${esc(fmt(row["PO price excl"]))}</td>
              <td class="${Number(row["price gap"] || 0) !== 0 ? "gap-warn" : ""}">${esc(fmt(row["price gap"]))}</td>
            </tr>`).join("");
        }
        download.href = data.download_url;
        download.style.display = "inline-block";
        statusEl.textContent = data.product_source === "uploaded"
          ? "对账完成。本次 OPC 文件已保存为默认文件。"
          : "对账完成。本次使用的是上次保存的 OPC 文件。";
      } catch (err) {
        statusEl.className = "status error";
        statusEl.textContent = err.message;
        tbody.innerHTML = `<tr><td class="empty" colspan="9">处理失败</td></tr>`;
      } finally {
        button.disabled = false;
      }
    });
  </script>
</body>
</html>
"""

MONEY_RE = re.compile(r"(?:EUR|€)?\s*[-+]?\d{1,3}(?:[.,]\d{3})*[.,]\d{2}\s*(?:EUR|€)?|(?:EUR|€)\s*[-+]?\d+(?:[.,]\d{2})?", re.I)
EAN_RE = re.compile(r"(?<!\d)0?\d{12,14}(?!\d)")
SUMMARY_RE = re.compile(
    r"\b(total|subtotal|summe|gesamt|betrag|nettobetrag|endbetrag|mwst|vat|tax|discount|rabatt|shipping|versand|"
    r"totaal|subtotaal|bedrag|btw|korting|verzending|"
    r"montant|total|sous total|remise|livraison|tva)\b",
    re.I,
)


def clean_text(value):
    if value is None or (isinstance(value, float) and math.isnan(value)):
        return ""
    text = str(value).strip()
    return "" if text.lower() == "nan" else text


def normalize_ean(value):
    text = clean_text(value)
    if re.fullmatch(r"\d+\.0", text):
        text = text[:-2]
    digits = re.sub(r"\D", "", text)
    return digits.lstrip("0") or digits


def display_ean(value):
    text = clean_text(value)
    if re.fullmatch(r"\d+\.0", text):
        text = text[:-2]
    return re.sub(r"\D", "", text) or text


def to_number(value):
    text = clean_text(value)
    if not text:
        return None
    text = text.replace("€", "").replace("EUR", "").replace("eur", "").replace(" ", "")
    if "," in text and "." in text:
        text = text.replace(".", "").replace(",", ".")
    else:
        text = text.replace(",", ".")
    try:
        return float(text)
    except ValueError:
        return None


def fmt_num(value, decimals=2):
    if value is None or value == "":
        return ""
    if abs(value - round(value)) < 0.0000001:
        return str(int(round(value)))
    return f"{value:.{decimals}f}".rstrip("0").rstrip(".")


def find_col(columns, candidates):
    lowered = {str(col).strip().lower(): col for col in columns}
    for candidate in candidates:
        key = candidate.lower()
        if key in lowered:
            return lowered[key]
    for col in columns:
        name = str(col).strip().lower()
        if any(candidate.lower() in name for candidate in candidates):
            return col
    return None


def field_has_file(field):
    return bool(field is not None and getattr(field, "filename", "") and getattr(field, "file", None))


def save_uploaded_product(field):
    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
        shutil.copyfileobj(field.file, tmp)
        tmp_path = Path(tmp.name)
    try:
        DEFAULT_PRODUCT_FILE.parent.mkdir(exist_ok=True)
        shutil.move(str(tmp_path), DEFAULT_PRODUCT_FILE)
    finally:
        tmp_path.unlink(missing_ok=True)
    return DEFAULT_PRODUCT_FILE


def resolve_product_file(form):
    product_field = form["product"] if "product" in form else None
    if field_has_file(product_field):
        return save_uploaded_product(product_field), "uploaded"
    if DEFAULT_PRODUCT_FILE.exists():
        return DEFAULT_PRODUCT_FILE, "cached"
    bundled_default = ROOT / "OPC_site product.xlsx"
    if bundled_default.exists():
        shutil.copyfile(bundled_default, DEFAULT_PRODUCT_FILE)
        return DEFAULT_PRODUCT_FILE, "cached"
    raise ValueError("请先上传一次 OPC_site product 文件；之后可不再重复上传")


def read_excel_upload(field):
    return io.BytesIO(field.file.read())


def read_po_upload(field):
    data = field.file.read()
    LAST_PO_FILE.write_bytes(data)
    return io.BytesIO(data)


def build_product_map(file_obj):
    mapping = {}
    excel = pd.ExcelFile(file_obj)
    for sheet in excel.sheet_names:
        df = read_table_with_detected_header(excel, sheet)
        sku_col = find_col(df.columns, ["SKU ID", "SKU"])
        ean_col = find_col(df.columns, ["UPC/EAN Code", "UPC/EAN", "EAN"])
        if sku_col is None or ean_col is None:
            continue
        for _, row in df.iterrows():
            sku = clean_text(row.get(sku_col))
            if not sku or not re.search(r"\d", sku):
                continue
            for part in re.split(r"[,;，\s]+", clean_text(row.get(ean_col))):
                key = normalize_ean(part)
                if key and key not in mapping:
                    mapping[key] = sku
    return mapping


def read_table_with_detected_header(excel, sheet_name):
    raw = pd.read_excel(excel, sheet_name=sheet_name, header=None, dtype=str)
    header_idx = None
    for idx in range(min(20, len(raw))):
        values = [clean_text(v).lower() for v in raw.iloc[idx].tolist()]
        joined = "\n".join(values)
        if ("sku id" in joined or "sku" in joined) and ("upc/ean" in joined or "ean" in joined):
            header_idx = idx
            break
    if header_idx is None:
        return pd.read_excel(excel, sheet_name=sheet_name, dtype=str)
    headers = [clean_text(v) or f"Column {i+1}" for i, v in enumerate(raw.iloc[header_idx].tolist())]
    data = raw.iloc[header_idx + 1 :].copy()
    data.columns = headers
    return data


def parse_po(file_obj):
    raw = pd.read_excel(file_obj, sheet_name=0, header=None, dtype=str)
    header_idx = None
    for idx in range(min(20, len(raw))):
        joined = "\n".join(clean_text(v).lower() for v in raw.iloc[idx].tolist())
        if "upc/ean" in joined and "actual received qty" in joined:
            header_idx = idx
            break
    if header_idx is None:
        raise ValueError("PO detail 中未找到 UPC/EAN 和 Actual received qty 表头")

    headers = [clean_text(v) or f"Column {i+1}" for i, v in enumerate(raw.iloc[header_idx].tolist())]
    data = raw.iloc[header_idx + 1 :].copy()
    data.columns = headers

    ean_col = find_col(data.columns, ["UPC/EAN"])
    name_col = find_col(data.columns, ["SKU name"])
    price_col = find_col(data.columns, ["Price(excl.VAT)", "Price excl.VAT"])
    qty_col = find_col(data.columns, ["Actual received qty"])
    if None in (ean_col, price_col, qty_col):
        raise ValueError("PO detail 缺少 UPC/EAN、Price(excl.VAT) 或 Actual received qty 列")

    grouped = {}
    for _, row in data.iterrows():
        key = normalize_ean(row.get(ean_col))
        if not key:
            continue
        item = grouped.setdefault(key, {"ean": display_ean(row.get(ean_col)), "po_qty": 0.0, "prices": [], "sku_name": clean_text(row.get(name_col)) if name_col else ""})
        item["po_qty"] += to_number(row.get(qty_col)) or 0
        price = to_number(row.get(price_col))
        if price is not None:
            item["prices"].append(price)
        if not item["sku_name"] and name_col:
            item["sku_name"] = clean_text(row.get(name_col))

    for item in grouped.values():
        unique_prices = []
        for price in item["prices"]:
            if not any(abs(price - seen) < 0.0001 for seen in unique_prices):
                unique_prices.append(price)
        item["po_price"] = unique_prices[0] if unique_prices else None
        item["po_price_display"] = "; ".join(fmt_num(p, 4) for p in unique_prices)
    return grouped


def parse_invoice(file_field, po_map=None):
    with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
        invoice_bytes = file_field.file.read()
        tmp.write(invoice_bytes)
        tmp_path = Path(tmp.name)
    LAST_INVOICE_FILE.write_bytes(invoice_bytes)
    try:
        od_text = extract_text_with_opendataloader(tmp_path)
        with pdfplumber.open(tmp_path) as pdf:
            rows = parse_invoice_rows_by_headers(pdf.pages)
            text = "\n".join(page.extract_text(x_tolerance=1, y_tolerance=3) or "" for page in pdf.pages)
            if od_text:
                text = text + "\n" + od_text
    finally:
        tmp_path.unlink(missing_ok=True)

    lines = [" ".join(line.split()) for line in text.splitlines() if line.strip()]
    if not rows:
        rows = parse_invoice_rows_with_inline_ean(lines)
    if not rows:
        rows = parse_invoice_rows_with_nedis_multiline(lines)
    if not rows:
        rows = parse_invoice_rows_with_separate_ean(lines, text)
    if not rows:
        raise ValueError("未能从发票中识别产品行，请确认 PDF 可复制文字")
    # Some invoices print an article number on the item line and the actual
    # 13-digit EAN on a following `EAN:` line. Prefer that explicit labelled
    # value when every parsed item has one.
    labelled_eans = extract_labelled_eans(text)
    if len(labelled_eans) >= len(rows):
        rows = [{**row, "ean": labelled_eans[idx]} for idx, row in enumerate(rows)]
    model_names = extract_model_product_names(text)
    if len(model_names) >= len(rows):
        rows = [{**row, "name": model_names[idx]} for idx, row in enumerate(rows)]
    normalized_rows = [{"ean": display_ean(row.get("ean")), "key": normalize_ean(row.get("ean")), "name": row["name"], "qty": row["qty"], "price": row["price"], "price_candidates": row.get("price_candidates", [])} for row in rows]
    if po_map:
        normalized_rows = align_invoice_prices_with_po(normalized_rows, po_map)
    return normalized_rows


def parse_invoice_rows_by_headers(pages):
    rows = []
    for page in pages:
        words = page.extract_words(x_tolerance=1, y_tolerance=3, keep_blank_chars=False)
        if not words:
            continue
        lines = group_words_by_line(words)
        header = find_invoice_header(lines)
        if not header:
            continue
        anchors, header_bottom = header
        page_rows = parse_rows_under_header(lines, anchors, header_bottom)
        rows.extend(page_rows)
    return rows


def extract_text_with_opendataloader(pdf_path):
    if RUNTIME_JDK_BIN.exists():
        os.environ["PATH"] = str(RUNTIME_JDK_BIN) + os.pathsep + os.environ.get("PATH", "")
        os.environ.setdefault("JAVA_HOME", str(RUNTIME_JDK_BIN.parent))
    try:
        if RUNTIME_PYTHON_PACKAGES.exists() and str(RUNTIME_PYTHON_PACKAGES) not in sys.path:
            sys.path.insert(0, str(RUNTIME_PYTHON_PACKAGES))
        if OPENDATALOADER_SRC.exists() and str(OPENDATALOADER_SRC) not in sys.path and not (RUNTIME_PYTHON_PACKAGES / "opendataloader_pdf").exists():
            sys.path.insert(0, str(OPENDATALOADER_SRC))
        import opendataloader_pdf
    except Exception:
        return ""

    package_root = Path(opendataloader_pdf.__file__).resolve().parent
    jar_dir = package_root / "jar"
    has_jar = jar_dir.exists() and any(jar_dir.glob("*.jar"))
    if not has_jar and shutil.which("opendataloader-pdf") is None:
        return ""

    out_dir = CACHE_DIR / "opendataloader_output"
    if out_dir.exists():
        shutil.rmtree(out_dir, ignore_errors=True)
    out_dir.mkdir(parents=True, exist_ok=True)
    try:
        opendataloader_pdf.convert(input_path=[str(pdf_path)], output_dir=str(out_dir), format="markdown", quiet=True)
    except (FileNotFoundError, subprocess.CalledProcessError, RuntimeError, Exception):
        return ""

    parts = []
    for path in out_dir.rglob("*.md"):
        try:
            parts.append(path.read_text(encoding="utf-8", errors="ignore"))
        except OSError:
            continue
    return "\n".join(parts)


def group_words_by_line(words, tolerance=3):
    sorted_words = sorted(words, key=lambda w: (float(w["top"]), float(w["x0"])))
    lines = []
    for word in sorted_words:
        top = float(word["top"])
        if not lines or abs(lines[-1]["top"] - top) > tolerance:
            lines.append({"top": top, "bottom": float(word["bottom"]), "words": [word]})
        else:
            lines[-1]["words"].append(word)
            lines[-1]["bottom"] = max(lines[-1]["bottom"], float(word["bottom"]))
    for line in lines:
        line["words"].sort(key=lambda w: float(w["x0"]))
        line["text"] = " ".join(w["text"] for w in line["words"])
    return lines


def find_invoice_header(lines):
    for idx, line in enumerate(lines):
        band = line["words"][:]
        if idx + 1 < len(lines) and lines[idx + 1]["top"] - line["top"] <= 18:
            band.extend(lines[idx + 1]["words"])
        band_text = normalize_header_text(" ".join(w["text"] for w in band))
        if not looks_like_invoice_header(band_text):
            continue
        anchors = detect_header_anchors(band)
        if "product" in anchors and "price_excl" in anchors:
            bottom = max(float(w["bottom"]) for w in band)
            return anchors, bottom
    return None


def normalize_header_text(text):
    text = unicodedata.normalize("NFKD", text.lower())
    text = "".join(ch for ch in text if not unicodedata.combining(ch))
    text = text.replace("ä", "ae").replace("ö", "oe").replace("ü", "ue").replace("ß", "ss")
    return re.sub(r"[^a-z0-9]+", " ", text)


def looks_like_invoice_header(text):
    has_product = any(token in text for token in [
        "artikel", "bezeichnung", "description", "product", "produkt", "item", "sku name",
        "omschrijving", "artikelomschrijving", "beschrijving", "productnaam",
        "designation", "descriptif", "article", "libelle", "produit",
    ])
    has_price = any(token in text for token in [
        "preis", "price", "netto", "net", "excl", "exkl", "without", "ohne",
        "prijs", "eenheidsprijs", "stuksprijs", "btw",
        "prix", "unitaire", "ht", "tva", "hors taxe", "hors taxes",
    ])
    return has_product and has_price


def detect_header_anchors(words):
    anchors = {}
    ordered = sorted(words, key=lambda w: float(w["x0"]))
    for idx, word in enumerate(ordered):
        local = ordered[max(0, idx - 2): idx + 3]
        context = normalize_header_text(" ".join(w["text"] for w in local))
        text = normalize_header_text(word["text"])
        center = (float(word["x0"]) + float(word["x1"])) / 2
        if any(token in context for token in [
            "artikelbezeichnung", "description", "product", "produkt", "item description", "sku name",
            "omschrijving", "artikelomschrijving", "beschrijving", "productnaam",
            "designation", "descriptif", "article", "libelle", "produit",
        ]):
            anchors.setdefault("product", center)
        elif text in {"artikel", "bezeichnung", "omschrijving", "beschrijving", "designation", "article", "libelle", "produit"}:
            anchors.setdefault("product", center)
        if any(token in context for token in ["ean", "gtin", "upc", "barcode"]):
            anchors.setdefault("ean", center)
        if any(token in context for token in [
            "menge", "qty", "quantity", "anzahl", "stueck", "pcs",
            "aantal", "hoeveelheid", "stuks",
            "quantite", "qte", "qté", "nombre",
        ]):
            anchors.setdefault("qty", center)
        if is_price_excl_header(context):
            anchors.setdefault("price_excl", center)
    return anchors


def is_price_excl_header(text):
    if any(token in text for token in ["gesamt", "total", "amount", "betrag", "summe"]):
        return False
    if any(token in text for token in ["totaal", "bedrag", "montant", "totale", "total hors"]):
        return False
    explicit = [
        "price excl", "preis exkl", "excl vat", "excl mwst", "without vat", "ohne mwst", "net price", "nettpreis", "nettopreis",
        "prijs excl", "excl btw", "btw excl", "zonder btw", "nettoprijs", "eenheidsprijs excl", "prijs zonder",
        "prix ht", "prix hors taxe", "prix hors taxes", "prix unitaire ht", "ht", "hors tva", "tva exclue", "prix net",
    ]
    if any(token in text for token in explicit):
        return True
    return any(token in text for token in ["netto", "net ", "ht"]) and any(token in text for token in ["preis", "price", "unit", "einzel", "prijs", "prix", "unitaire"])


def parse_rows_under_header(lines, anchors, header_bottom):
    rows = []
    boundaries = build_column_boundaries(anchors)
    for line in lines:
        if line["top"] <= header_bottom + 2:
            continue
        text = line["text"]
        if SUMMARY_RE.search(text):
            continue
        cells = assign_words_to_header_columns(line["words"], boundaries)
        name = clean_invoice_name(cells.get("product", ""))
        qty = first_number(cells.get("qty", ""))
        if qty is None:
            name, qty = strip_trailing_qty_from_name(name)
        price = first_number(cells.get("price_excl", ""))
        ean = first_ean(cells.get("ean", "")) or first_ean(text)
        if is_valid_invoice_product_name(name) and qty is not None and price is not None:
            rows.append({"ean": ean or "", "name": name, "qty": qty, "price": price, "price_candidates": extract_money_numbers(text)})
    return rows


def build_column_boundaries(anchors):
    ordered = sorted((x, role) for role, x in anchors.items())
    boundaries = {}
    for idx, (center, role) in enumerate(ordered):
        left = -math.inf if idx == 0 else (ordered[idx - 1][0] + center) / 2
        right = math.inf if idx == len(ordered) - 1 else (center + ordered[idx + 1][0]) / 2
        boundaries[role] = (left, right)
    return boundaries


def assign_words_to_header_columns(words, boundaries):
    cells = {role: [] for role in boundaries}
    for word in words:
        center = (float(word["x0"]) + float(word["x1"])) / 2
        for role, (left, right) in boundaries.items():
            if left <= center < right:
                cells[role].append(word["text"])
                break
    return {role: " ".join(parts) for role, parts in cells.items()}


def first_number(text):
    match = re.search(r"-?\d{1,3}(?:[.,]\d{3})*[.,]\d+|-?\d+", clean_text(text))
    return to_number(match.group(0)) if match else None


def extract_money_numbers(text):
    numbers = []
    for match in MONEY_RE.finditer(clean_text(text)):
        value = to_number(match.group(0))
        if value is not None:
            numbers.append(value)
    return numbers


def first_ean(text):
    match = EAN_RE.search(clean_text(text))
    return match.group(0) if match else ""


def extract_labelled_eans(text):
    """Extract explicit 13-digit EAN values printed with an EAN label."""
    return [match.group(1) for match in re.finditer(r"\bEAN\s*:\s*([0-9]{13})\b", clean_text(text), flags=re.I)]


def extract_model_product_names(text):
    """Extract model-number product lines such as `X3003/00 SHAVER ...`."""
    names = []
    model_line = re.compile(r"^([A-Z][A-Z0-9.-]{2,}/[A-Z0-9]{2})\s+(.+)$", re.I)
    for raw_line in text.splitlines():
        line = clean_text(raw_line)
        match = model_line.match(line)
        if not match:
            continue
        description = re.split(r"\s+(?:Beheerbijdrage|btw\b|VAT\b|Tax\b)", match.group(2), maxsplit=1, flags=re.I)[0]
        description = clean_invoice_name(description)
        if description:
            names.append(f"{match.group(1)} {description}".strip())
    return names


def strip_trailing_qty_from_name(name):
    match = re.search(r"\s+(\d+(?:[.,]\d+)?)\s*$", name)
    if not match:
        return name, None
    qty = to_number(match.group(1))
    return name[: match.start()].strip(), qty


def parse_invoice_rows_with_inline_ean(lines):
    rows = []
    for line in lines:
        ean_match = EAN_RE.search(line)
        money_matches = list(MONEY_RE.finditer(line))
        if not ean_match or not money_matches:
            continue
        nums = [to_number(match.group(0)) for match in money_matches]
        nums = [num for num in nums if num is not None]
        if not nums:
            continue
        price = nums[-2] if len(nums) >= 2 else nums[-1]
        total = nums[-1] if len(nums) >= 2 else None
        qty = infer_qty_from_line(line, ean_match, money_matches, price, total)
        name = clean_invoice_name(line[: ean_match.start()] + " " + line[ean_match.end() : money_matches[0].start()])
        name, name_qty = strip_trailing_qty_from_name(name)
        if name_qty is not None:
            qty = name_qty
        if is_valid_invoice_product_name(name) and price is not None and qty is not None:
            rows.append({"ean": ean_match.group(0), "name": name, "qty": qty, "price": price, "price_candidates": nums})
    return rows


def parse_invoice_rows_with_nedis_multiline(lines):
    header_text = " ".join(lines[:80]).lower()
    if not ("artikelnummer" in header_text and "omschrijving" in header_text and "prijs per" in header_text):
        return []

    rows = []
    current = None
    item_re = re.compile(
        r"^\s*(\d{1,4})\s+([A-Z0-9][A-Z0-9\-/.]+)\s+(.+?)\s+"
        r"(\d+(?:[.,]\d+)?)\s+(\d+(?:[.,]\d+)?)\s+(\d+(?:[.,]\d+)?)\s+"
        r"(\d+(?:[.,]\d+)?)\s+(\d+(?:[.,]\d+)?)%\s+(\S+)\s+(\d+(?:[.,]\d+)?)\s*$",
        re.I,
    )
    ean_re = re.compile(r"^\s*/\s*(\d{12,14})\b")
    origin_re = re.compile(r"^\s*\d{8,12}\s*/\s*[A-Z]{2,3}\b", re.I)

    for line in lines:
        match = item_re.match(line)
        if match:
            if current and current.get("ean"):
                rows.append(finalize_nedis_row(current))
            current = {
                "article": match.group(2),
                "qty": to_number(match.group(5)),
                "price": to_number(match.group(7)),
                "amount": to_number(match.group(10)),
                "ean": "",
                "description": "",
                "price_candidates": [to_number(match.group(i)) for i in (4, 5, 6, 7, 10)],
            }
            current["price_candidates"] = [value for value in current["price_candidates"] if value is not None]
            continue

        if not current:
            continue
        ean_match = ean_re.match(line)
        if ean_match:
            current["ean"] = ean_match.group(1)
            continue
        if origin_re.match(line):
            description = re.sub(r"^\s*\d{8,12}\s*/\s*[A-Z]{2,3}\s+", "", line, flags=re.I).strip()
            if description:
                current["description"] = description
            continue

    if current and current.get("ean"):
        rows.append(finalize_nedis_row(current))
    return [row for row in rows if row]


def finalize_nedis_row(item):
    name = clean_invoice_name(f"{item.get('article', '')} {item.get('description', '')}".strip())
    if not is_valid_invoice_product_name(name):
        name = clean_invoice_name(item.get("article", ""))
    if not item.get("ean") or item.get("qty") is None or item.get("price") is None:
        return None
    return {
        "ean": item["ean"],
        "name": name,
        "qty": item["qty"],
        "price": item["price"],
        "price_candidates": item.get("price_candidates", []),
    }


def parse_invoice_rows_with_separate_ean(lines, text):
    product_rows = []
    amount = r"(\d{1,3}(?:\.\d{3})*,\d{2}|\d+,\d{2}|\d+\.\d{2})"
    patterns = [
        re.compile(rf"^\s*\d+\s+(.+?)\s+{amount}\s*(?:€|EUR)?\s+{amount}\s*(?:€|EUR)?\s*$", re.I),
        re.compile(rf"^(.+?)\s+{amount}\s*(?:€|EUR)?\s+{amount}\s*(?:€|EUR)?\s*$", re.I),
    ]
    skip = re.compile(r"\b(MwSt|VAT|tax|total|subtotal|netto|endbetrag|nettobetrag)\b", re.I)
    for line in lines:
        if skip.search(line):
            continue
        for pattern in patterns:
            match = pattern.match(line)
            if not match:
                continue
            name = clean_invoice_name(match.group(1))
            price = to_number(match.group(2))
            total = to_number(match.group(3))
            qty = round(total / price) if price else None
            name, name_qty = strip_trailing_qty_from_name(name)
            if name_qty is not None:
                qty = name_qty
            if is_valid_invoice_product_name(name) and price is not None and qty is not None:
                product_rows.append({"ean": "", "name": name, "price": price, "qty": qty, "price_candidates": extract_money_numbers(line)})
            break

    eans = []
    for match in re.finditer(r"(?:EAN|UPC|GTIN|Barcode|Bar code|Art\.?Nr\.?)[:\s]*([0-9\s]{12,18})", text, flags=re.I):
        ean = re.sub(r"\D", "", match.group(1))
        if 12 <= len(ean) <= 14:
            eans.append(ean)
    if not eans:
        eans = [m.group(0) for m in EAN_RE.finditer(text)]

    for idx, row in enumerate(product_rows):
        if idx < len(eans):
            row["ean"] = eans[idx]
    return product_rows


def infer_qty_from_line(line, ean_match, money_matches, price, total):
    if price and total and abs(price) > 0.000001:
        qty = total / price
        if abs(qty - round(qty)) < 0.0001:
            return round(qty)
    before_price = line[ean_match.end() : money_matches[0].start()]
    qty_candidates = re.findall(r"(?<![A-Z0-9])\d+(?:[.,]\d+)?(?![A-Z0-9])", before_price, flags=re.I)
    return to_number(qty_candidates[-1]) if qty_candidates else None


def clean_invoice_name(text):
    text = re.sub(r"\b(?:EAN|UPC|GTIN|Barcode|Bar code|Art\.?Nr\.?)\b[:\s]*", " ", text, flags=re.I)
    text = re.sub(r"^\s*\d+\s+", "", text)
    return re.sub(r"\s+", " ", text).strip(" -:;")


def is_valid_invoice_product_name(name):
    text = clean_text(name)
    if not text:
        return False
    if not re.search(r"[A-Za-z]", text):
        return False
    if re.fullmatch(r"[\d\s.,€EUR+-]+", text, flags=re.I):
        return False
    summary_words = r"\b(total|subtotal|summe|gesamt|betrag|nettobetrag|endbetrag|mwst|vat|tax|discount|rabatt|shipping|versand)\b"
    if re.search(summary_words, text, flags=re.I):
        return False
    return True


def match_invoice_to_po(inv, po_map):
    if inv.get("key") and inv["key"] in po_map:
        return {"key": inv["key"], "po": po_map[inv["key"]], "score": 1.0}
    best = None
    for key, po in po_map.items():
        score = fuzzy_name_score(inv.get("name") or "", po.get("sku_name") or "")
        if best is None or score > best["score"]:
            best = {"key": key, "po": po, "score": score}
    if best and best["score"] >= 0.55:
        return best
    return None


def align_invoice_prices_with_po(invoice_rows, po_map):
    aligned = []
    for row in invoice_rows:
        match = match_invoice_to_po(row, po_map)
        po_price = match["po"].get("po_price") if match else None
        candidates = [value for value in row.get("price_candidates", []) if value is not None]
        if po_price is not None and candidates:
            current_gap = abs((row.get("price") or 0) - po_price)
            best = min(candidates, key=lambda value: abs(value - po_price))
            best_gap = abs(best - po_price)
            # Use PO only as a sanity check: switch when another extracted price is materially closer.
            if best_gap < 0.01 and current_gap > 0.01:
                row = {**row, "price": best}
        aligned.append(row)
    return aligned


def fuzzy_name_score(left, right):
    left_tokens = name_tokens(left)
    right_tokens = name_tokens(right)
    if not left_tokens or not right_tokens:
        return 0.0
    overlap = len(left_tokens & right_tokens) / max(1, min(len(left_tokens), len(right_tokens)))
    ratio = SequenceMatcher(None, " ".join(sorted(left_tokens)), " ".join(sorted(right_tokens))).ratio()
    left_models = {token for token in left_tokens if re.search(r"[A-Z]\d|\d[A-Z]", token)}
    right_models = {token for token in right_tokens if re.search(r"[A-Z]\d|\d[A-Z]", token)}
    model_bonus = 0.25 if left_models and right_models and left_models & right_models else 0.0
    return min(1.0, max(overlap, ratio) + model_bonus)


def name_tokens(text):
    normalized = re.sub(r"[^A-Z0-9]+", " ", str(text).upper())
    stop = {"THE", "AND", "WITH", "FOR", "INCLUDES", "MULTIPLE", "ATTACHMENTS", "BLACK", "WHITE", "GREY", "GRAY", "SILVER", "STAINLESS", "FINISH", "SN", "DE", "EU", "EUR", "NEW"}
    return {token for token in normalized.split() if len(token) >= 2 and token not in stop}


def reconcile(product_file, invoice_file, po_file):
    product_map = build_product_map(product_file)
    po_map = parse_po(po_file)
    invoice_rows = parse_invoice(invoice_file, po_map)

    output = []
    qty_diff_rows = 0
    price_diff_rows = 0
    invoice_amount_total = 0.0
    po_amount_total = 0.0
    for inv in invoice_rows:
        match = match_invoice_to_po(inv, po_map)
        po = match["po"] if match else None
        matched_key = match["key"] if match else inv["key"]
        ean_display = inv["ean"] or (po.get("ean") if po else "")
        sku = product_map.get(matched_key, "")
        invoice_amount_total += amount_or_zero(inv.get("qty"), inv.get("price"))

        if not po:
            output.append({"EAN": ean_display, PRODUCT_NAME: inv["name"], "sku": sku, "invoice qty": fmt_num(inv["qty"], 0), "PO Actual received qty": "", "qty gap": "", "invoice price excl": fmt_num(inv["price"], 4), "PO price excl": "", "price gap": ""})
            qty_diff_rows += 1
            continue

        po_amount_total += amount_or_zero(po.get("po_qty"), po.get("po_price"))
        qty_gap = po["po_qty"] - (inv["qty"] or 0)
        price_gap = None if po["po_price"] is None or inv["price"] is None else po["po_price"] - inv["price"]
        has_qty_gap = abs(qty_gap) > 0.0001
        has_price_gap = price_gap is not None and abs(price_gap) > 0.0001
        if not (has_qty_gap or has_price_gap):
            continue
        qty_diff_rows += 1 if has_qty_gap else 0
        price_diff_rows += 1 if has_price_gap else 0
        output.append({"EAN": ean_display, PRODUCT_NAME: inv["name"], "sku": sku, "invoice qty": fmt_num(inv["qty"], 0), "PO Actual received qty": fmt_num(po["po_qty"], 0), "qty gap": fmt_num(qty_gap, 0) if has_qty_gap else "0", "invoice price excl": fmt_num(inv["price"], 4), "PO price excl": po["po_price_display"], "price gap": fmt_num(price_gap, 4) if has_price_gap else "0"})

    return output, {
        "invoice_rows": len(invoice_rows),
        "diff_rows": len(output),
        "qty_diff_rows": qty_diff_rows,
        "price_diff_rows": price_diff_rows,
        "invoice_amount_total": fmt_num(invoice_amount_total, 2),
        "po_amount_total": fmt_num(po_amount_total, 2),
        "amount_gap_total": fmt_num(po_amount_total - invoice_amount_total, 2),
    }


def amount_or_zero(qty, price):
    return (qty or 0) * (price or 0)


def save_result(rows):
    result_id = uuid.uuid4().hex
    path = OUTPUT_DIR / f"reconciliation_{result_id}.xlsx"
    pd.DataFrame(rows, columns=COLUMNS).to_excel(path, index=False)
    return result_id, path


class Handler(BaseHTTPRequestHandler):
    def log_message(self, format, *args):
        return

    def send_bytes(self, status, body, content_type):
        self.send_response(status)
        self.send_header("Content-Type", content_type)
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def send_json(self, status, payload):
        self.send_bytes(status, json.dumps(payload, ensure_ascii=False).encode("utf-8"), "application/json; charset=utf-8")

    def do_GET(self):
        parsed = urlparse(self.path)
        if parsed.path == "/":
            self.send_bytes(200, HTML.encode("utf-8"), "text/html; charset=utf-8")
            return
        if parsed.path.startswith("/download/"):
            result_id = re.sub(r"[^a-f0-9]", "", unquote(parsed.path.rsplit("/", 1)[-1]))
            matches = list(OUTPUT_DIR.glob(f"reconciliation_{result_id}.xlsx"))
            if not matches:
                self.send_json(404, {"ok": False, "error": "下载文件不存在"})
                return
            body = matches[0].read_bytes()
            self.send_response(200)
            self.send_header("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            self.send_header("Content-Disposition", f'attachment; filename="{matches[0].name}"')
            self.send_header("Content-Length", str(len(body)))
            self.end_headers()
            self.wfile.write(body)
            return
        self.send_json(404, {"ok": False, "error": "页面不存在"})

    def do_POST(self):
        if urlparse(self.path).path != "/reconcile":
            self.send_json(404, {"ok": False, "error": "接口不存在"})
            return
        try:
            form = cgi.FieldStorage(fp=self.rfile, headers=self.headers, environ={"REQUEST_METHOD": "POST", "CONTENT_TYPE": self.headers.get("Content-Type")})
            for name in ["invoice", "po"]:
                if name not in form or not field_has_file(form[name]):
                    raise ValueError(f"缺少文件：{name}")
            product_path, product_source = resolve_product_file(form)
            with product_path.open("rb") as product_file:
                rows, metrics = reconcile(product_file, form["invoice"], read_po_upload(form["po"]))
            result_id, _ = save_result(rows)
            self.send_json(200, {"ok": True, "rows": rows, "metrics": metrics, "download_url": f"/download/{result_id}", "product_source": product_source})
        except Exception as exc:
            self.send_json(400, {"ok": False, "error": str(exc)})


def main():
    port = int(os.environ.get("PORT", "8799"))
    host = os.environ.get("HOST", "0.0.0.0")
    server = ThreadingHTTPServer((host, port), Handler)
    print(f"Invoice PO tool: http://{host}:{port}")
    server.serve_forever()


if __name__ == "__main__":
    main()
