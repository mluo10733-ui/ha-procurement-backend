from __future__ import annotations

import argparse
import html
import json
import mimetypes
import os
import re
import sys
import threading
import webbrowser
from datetime import datetime
from email.parser import BytesParser
from email.policy import default
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import parse_qs, quote, unquote, urlparse

from generate_po_offer import BASE_DIR, build_outputs, get_reference_paths, load_config, read_error_rows
from generate_po_offer import OUTPUT_DIR


HOST = os.getenv("HOST", "0.0.0.0")
DEFAULT_PORT = 18888
UPLOAD_DIR = Path(os.getenv("PO_UPLOAD_DIR", BASE_DIR / "web_uploads")).resolve()
TOOL_VERSION = "2026-07-16 sku-missing-site-product-allowed"


def safe_filename(name: str) -> str:
    name = Path(name or "uploaded_file").name
    return re.sub(r'[<>:"/\\|?*\x00-\x1F]', "_", name).strip() or "uploaded_file"


def rel_link(path: Path) -> str:
    rel = path.resolve().relative_to(BASE_DIR)
    return f"/download?path={quote(str(rel).replace(os.sep, '/'))}"


def current_default(key: str) -> str:
    value = load_config().get(key, "")
    return html.escape(value or "Not set")


def page(content: str, status: str = "") -> bytes:
    status_html = f'<div class="status">{html.escape(status)}</div>' if status else ""
    return f"""<!doctype html>
<html lang="zh-CN">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>PO / Offer Generator</title>
  <style>
    :root {{
      --bg: #f6f7f9;
      --panel: #ffffff;
      --text: #1f2933;
      --muted: #697586;
      --line: #d8dee8;
      --accent: #1264a3;
      --accent-dark: #0f4f82;
      --ok: #116d4e;
      --error: #b42318;
    }}
    * {{ box-sizing: border-box; }}
    body {{
      margin: 0;
      font-family: "Segoe UI", Arial, sans-serif;
      background: var(--bg);
      color: var(--text);
    }}
    header {{
      background: #263444;
      color: #fff;
      padding: 18px 28px;
      display: flex;
      justify-content: space-between;
      align-items: center;
      gap: 16px;
    }}
    header h1 {{
      margin: 0;
      font-size: 20px;
      font-weight: 650;
    }}
    header span {{
      color: #d8dee8;
      font-size: 13px;
    }}
    main {{
      max-width: 1040px;
      margin: 24px auto;
      padding: 0 20px 32px;
    }}
    form, .result, .defaults {{
      background: var(--panel);
      border: 1px solid var(--line);
      border-radius: 8px;
      padding: 22px;
      margin-bottom: 18px;
    }}
    .grid {{
      display: grid;
      grid-template-columns: repeat(2, minmax(0, 1fr));
      gap: 18px;
    }}
    label {{
      display: block;
      font-size: 14px;
      font-weight: 650;
      margin-bottom: 8px;
    }}
    input[type="file"] {{
      width: 100%;
      border: 1px solid var(--line);
      border-radius: 6px;
      padding: 10px;
      background: #fbfcfe;
    }}
    .hint {{
      margin-top: 6px;
      color: var(--muted);
      font-size: 12px;
      min-height: 17px;
    }}
    .actions {{
      margin-top: 22px;
      display: flex;
      justify-content: flex-end;
    }}
    button {{
      border: 0;
      border-radius: 6px;
      padding: 11px 18px;
      background: var(--accent);
      color: #fff;
      font-size: 14px;
      font-weight: 650;
      cursor: pointer;
    }}
    button:hover {{ background: var(--accent-dark); }}
    .defaults h2, .result h2 {{
      margin: 0 0 14px;
      font-size: 16px;
    }}
    .defaults dl {{
      display: grid;
      grid-template-columns: 210px 1fr;
      gap: 8px 14px;
      margin: 0;
      font-size: 13px;
    }}
    .defaults dt {{ color: var(--muted); }}
    .defaults dd {{ margin: 0; overflow-wrap: anywhere; }}
    .links {{
      display: grid;
      grid-template-columns: repeat(auto-fit, minmax(220px, 1fr));
      gap: 10px;
    }}
    .links a {{
      border: 1px solid var(--line);
      border-radius: 6px;
      padding: 12px;
      color: var(--accent-dark);
      text-decoration: none;
      background: #fbfcfe;
      font-weight: 650;
    }}
    .status {{
      border-left: 4px solid var(--error);
      background: #fff5f5;
      padding: 12px 14px;
      margin-bottom: 18px;
      color: var(--error);
      border-radius: 6px;
    }}
    @media (max-width: 760px) {{
      header {{ align-items: flex-start; flex-direction: column; }}
      .grid, .defaults dl {{ grid-template-columns: 1fr; }}
    }}
  </style>
</head>
<body>
  <header>
    <h1>PO / Offer Generator</h1>
    <span>Local browser tool</span>
  </header>
  <main>
    <section class="defaults">
      <h2>Storage paths</h2>
      <dl>
        <dt>Logic version</dt><dd>{html.escape(TOOL_VERSION)}</dd>
        <dt>Working folder</dt><dd>{html.escape(str(BASE_DIR))}</dd>
        <dt>Output folder</dt><dd>{html.escape(str(OUTPUT_DIR))}</dd>
        <dt>Upload cache</dt><dd>{html.escape(str(UPLOAD_DIR))}</dd>
      </dl>
    </section>
    {status_html}
    {content}
  </main>
</body>
</html>""".encode("utf-8")


def index_page(status: str = "") -> bytes:
    content = f"""
<form action="/generate" method="post" enctype="multipart/form-data">
  <div class="grid">
    <div>
      <label for="input_file">Input Excel file</label>
      <input id="input_file" name="input_file" type="file" accept=".xlsx,.xlsm,.xls">
      <div class="hint">Leave blank to reuse last saved input.</div>
    </div>
    <div>
      <label for="supplier_offer">supplier offer information</label>
      <input id="supplier_offer" name="supplier_offer" type="file" accept=".xlsx,.xlsm,.xls">
      <div class="hint">Leave blank to reuse last saved file.</div>
    </div>
    <div>
      <label for="supplier_po">supplier PO information</label>
      <input id="supplier_po" name="supplier_po" type="file" accept=".xlsx,.xlsm,.xls">
      <div class="hint">Leave blank to reuse last saved file.</div>
    </div>
    <div>
      <label for="site_product">OPC_site product</label>
      <input id="site_product" name="site_product" type="file" accept=".xlsx,.xlsm,.xls">
      <div class="hint">Leave blank to reuse last saved file.</div>
    </div>
  </div>
  <div class="actions">
    <button type="submit">Generate files</button>
  </div>
</form>
<section class="defaults">
  <h2>Saved defaults</h2>
  <dl>
    <dt>Input Excel</dt><dd>{current_default("last_input")}</dd>
    <dt>supplier offer information</dt><dd>{current_default("supplier_offer")}</dd>
    <dt>supplier PO information</dt><dd>{current_default("supplier_po")}</dd>
    <dt>OPC_site product</dt><dd>{current_default("site_product")}</dd>
  </dl>
</section>
"""
    return page(content, status)


def parse_uploads(headers, body: bytes) -> dict[str, tuple[str, bytes]]:
    content_type = headers.get("Content-Type", "")
    if "multipart/form-data" not in content_type:
        return {}
    raw_message = (
        f"Content-Type: {content_type}\r\nMIME-Version: 1.0\r\n\r\n".encode("utf-8")
        + body
    )
    message = BytesParser(policy=default).parsebytes(raw_message)
    uploads = {}
    for part in message.iter_parts():
        params = dict(part.get_params(header="content-disposition") or [])
        name = params.get("name")
        filename = params.get("filename")
        if not name or not filename:
            continue
        uploads[name] = (filename, part.get_payload(decode=True) or b"")
    return uploads


def save_uploaded(uploads: dict[str, tuple[str, bytes]], field_name: str, run_dir: Path) -> str | None:
    if field_name not in uploads:
        return None
    original_name, data = uploads[field_name]
    if not original_name:
        return None
    filename = safe_filename(original_name)
    target = run_dir / filename
    with target.open("wb") as fh:
        fh.write(data)
    return str(target)


class ToolHandler(BaseHTTPRequestHandler):
    protocol_version = "HTTP/1.1"

    def handle_expect_100(self) -> bool:
        self.send_response_only(100)
        self.end_headers()
        return True

    def send_html(self, body: bytes, status_code: int = 200) -> None:
        self.send_response(status_code)
        self.send_header("Content-Type", "text/html; charset=utf-8")
        self.send_header("Cache-Control", "no-store, no-cache, must-revalidate, max-age=0")
        self.send_header("Pragma", "no-cache")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def do_GET(self) -> None:
        parsed = urlparse(self.path)
        if parsed.path == "/":
            self.send_html(index_page())
            return
        if parsed.path == "/download":
            self.send_download(parsed.query)
            return
        self.send_html(index_page("Page not found."), 404)

    def do_POST(self) -> None:
        parsed = urlparse(self.path)
        if parsed.path != "/generate":
            self.send_html(index_page("Page not found."), 404)
            return

        try:
            content_length = int(self.headers.get("Content-Length", "0") or "0")
            uploads = parse_uploads(self.headers, self.rfile.read(content_length))
            stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            run_dir = UPLOAD_DIR / stamp
            run_dir.mkdir(parents=True, exist_ok=True)

            input_file = save_uploaded(uploads, "input_file", run_dir)
            supplier_offer = save_uploaded(uploads, "supplier_offer", run_dir)
            supplier_po = save_uploaded(uploads, "supplier_po", run_dir)
            site_product = save_uploaded(uploads, "site_product", run_dir)

            input_path, offer_info_path, po_info_path, site_product_path = get_reference_paths(
                supplier_offer=supplier_offer,
                supplier_po=supplier_po,
                site_product=site_product,
                input_path=input_file,
            )
            if not input_path:
                raise ValueError("Input Excel file is required for the first run.")

            offer_path, po_path, error_path = build_outputs(
                input_path=input_path,
                supplier_offer_path=offer_info_path,
                supplier_po_path=po_info_path,
                site_product_path=site_product_path,
                run_stamp=f"web_{stamp}",
            )

            links = []
            if offer_path:
                links.append(f'<a download href="{rel_link(offer_path)}">Download Offer</a>')
            links.append(f'<a download href="{rel_link(po_path)}">Download PO</a>')
            links.append(f'<a download href="{rel_link(error_path)}">Download Error Report</a>')

            if self.headers.get("Accept", "").find("application/json") >= 0:
                self.send_json(200, {
                    "ok": True,
                    "offer_download_url": rel_link(offer_path) if offer_path else None,
                    "po_download_url": rel_link(po_path),
                    "error_download_url": rel_link(error_path),
                    "errors": read_error_rows(error_path),
                })
                return

            content = f"""
<section class="result">
  <h2>Generated files</h2>
  <div class="links">{''.join(links)}</div>
</section>
<section class="defaults">
  <h2>Files used</h2>
  <dl>
    <dt>Input Excel</dt><dd>{html.escape(str(input_path))}</dd>
    <dt>supplier offer information</dt><dd>{html.escape(str(offer_info_path))}</dd>
    <dt>supplier PO information</dt><dd>{html.escape(str(po_info_path))}</dd>
    <dt>OPC_site product</dt><dd>{html.escape(str(site_product_path))}</dd>
  </dl>
</section>
"""
            self.send_html(page(content))
        except Exception as exc:
            self.send_html(index_page(f"Generate failed: {exc}"), 400)

    def send_json(self, status: int, payload: dict) -> None:
        data = json.dumps(payload, ensure_ascii=False).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)

    def send_download(self, query: str) -> None:
        params = parse_qs(query)
        requested = params.get("path", [""])[0]
        if not requested:
            self.send_error(404)
            return
        try:
            target = (BASE_DIR / unquote(requested)).resolve()
            target.relative_to(BASE_DIR)
        except ValueError:
            self.send_error(403)
            return
        if not target.exists() or not target.is_file():
            self.send_error(404)
            return
        content_type = mimetypes.guess_type(target.name)[0] or "application/octet-stream"
        data = target.read_bytes()
        self.send_response(200)
        self.send_header("Content-Type", content_type)
        self.send_header("Content-Disposition", f'attachment; filename="{target.name}"')
        self.send_header("Cache-Control", "no-store")
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)

    def log_message(self, format: str, *args) -> None:
        print(f"[web] {self.address_string()} - {format % args}")


def run(open_browser: bool = True) -> None:
    port = DEFAULT_PORT
    server = None
    for candidate in range(DEFAULT_PORT, DEFAULT_PORT + 20):
        try:
            server = ThreadingHTTPServer((HOST, candidate), ToolHandler)
            port = candidate
            break
        except OSError:
            continue
    if server is None:
        raise OSError("No available local port found.")

    url = f"http://{HOST}:{port}"
    print(f"PO / Offer Generator is running: {url}")
    print("Keep this window open while using the web page.")
    if open_browser and os.getenv("NO_BROWSER") != "1":
        threading.Timer(0.8, lambda: webbrowser.open(url)).start()
    server.serve_forever()


if __name__ == "__main__":
    try:
        parser = argparse.ArgumentParser(description="Run the local PO / Offer web tool.")
        parser.add_argument("--no-browser", action="store_true", help="Start the server without opening a browser.")
        args = parser.parse_args()
        run(open_browser=not args.no_browser)
    except KeyboardInterrupt:
        print("\nServer stopped.")
        sys.exit(0)
