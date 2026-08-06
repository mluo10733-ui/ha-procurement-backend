import json
import os
import re
import sys
import threading
import webbrowser
from email.parser import BytesParser
from email.policy import default
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import parse_qs, quote, urlparse

from generate_inquiry import create_supplier_inventory_template_bytes, generate


ROOT = Path(__file__).resolve().parent
WORKSPACE = Path(os.environ.get("INQUIRY_TOOL_WORKSPACE", r"D:\PO\InquiryTool"))
PORT = int(os.environ.get("PORT", os.environ.get("INQUIRY_TOOL_PORT", "8788")))
HOST = os.environ.get("HOST", "0.0.0.0")
UPLOAD_DIR = WORKSPACE / "work" / "inquiry_tool_uploads"
OUTPUT_DIR = WORKSPACE / "outputs" / "inquiry_tool"
LAST_FILES_PATH = WORKSPACE / "work" / "last_uploaded_files.json"


def safe_filename(name):
    cleaned = re.sub(r"[^A-Za-z0-9._() -]+", "_", name or "upload.xlsx")
    return cleaned[:120] or "upload.xlsx"


def parse_multipart(headers, body):
    content_type = headers.get("Content-Type", "")
    raw = f"Content-Type: {content_type}\r\nMIME-Version: 1.0\r\n\r\n".encode("utf-8") + body
    message = BytesParser(policy=default).parsebytes(raw)
    fields = {}
    files = {}
    if not message.is_multipart():
        return fields, files
    UPLOAD_DIR.mkdir(parents=True, exist_ok=True)
    for part in message.iter_parts():
        name = part.get_param("name", header="content-disposition")
        filename = part.get_filename()
        payload = part.get_payload(decode=True) or b""
        if not name:
            continue
        if filename:
            target = UPLOAD_DIR / safe_filename(filename or f"{name}.xlsx")
            target.write_bytes(payload)
            files[name] = str(target.resolve())
        elif name in {"autoFile", "haFile", "inventoryFile"}:
            continue
        else:
            fields[name] = payload.decode(part.get_content_charset() or "utf-8", errors="replace").strip()
    return fields, files


def load_last_files():
    if not LAST_FILES_PATH.exists():
        return {}
    try:
        data = json.loads(LAST_FILES_PATH.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return {}
    return {
        key: value
        for key, value in data.items()
        if key in {"auto_path", "ha_path"} and value and Path(value).exists()
    }


def save_last_files(auto_path, ha_path):
    LAST_FILES_PATH.parent.mkdir(parents=True, exist_ok=True)
    payload = {
        "auto_path": str(Path(auto_path).resolve()),
        "ha_path": str(Path(ha_path).resolve()),
    }
    LAST_FILES_PATH.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")


def is_safe_download_path(file_path):
    try:
        resolved = Path(file_path).resolve()
        return resolved.is_file() and resolved.is_relative_to(OUTPUT_DIR.resolve())
    except Exception:
        return False


class Handler(BaseHTTPRequestHandler):
    def _send(self, status, body, content_type="application/json; charset=utf-8"):
        payload = body.encode("utf-8") if isinstance(body, str) else body
        self.send_response(status)
        self.send_header("Content-Type", content_type)
        self.send_header("Content-Length", str(len(payload)))
        self.end_headers()
        self.wfile.write(payload)

    def do_GET(self):
        parsed = urlparse(self.path)
        path = parsed.path
        if path == "/":
            self._send(200, (ROOT / "index.html").read_text(encoding="utf-8"), "text/html; charset=utf-8")
            return
        if path == "/app.css":
            self._send(200, (ROOT / "app.css").read_text(encoding="utf-8"), "text/css; charset=utf-8")
            return
        if path == "/supplier_inventory_template":
            payload = create_supplier_inventory_template_bytes()
            self.send_response(200)
            self.send_header("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            self.send_header("Content-Length", str(len(payload)))
            self.send_header("Content-Disposition", "attachment; filename=Supplier_Inventory_Template.xlsx")
            self.end_headers()
            self.wfile.write(payload)
            return
        if path == "/download":
            query = parse_qs(parsed.query)
            requested = query.get("file", [""])[0]
            if not is_safe_download_path(requested):
                self._send(404, json.dumps({"error": "File not found"}, ensure_ascii=False))
                return
            file_path = Path(requested).resolve()
            payload = file_path.read_bytes()
            self.send_response(200)
            self.send_header("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            self.send_header("Content-Length", str(len(payload)))
            self.send_header("Content-Disposition", f"attachment; filename*=UTF-8''{quote(file_path.name)}")
            self.end_headers()
            self.wfile.write(payload)
            return
        self._send(404, json.dumps({"error": "Not found"}, ensure_ascii=False))

    def do_POST(self):
        if urlparse(self.path).path != "/generate":
            self._send(404, json.dumps({"error": "Not found"}, ensure_ascii=False))
            return
        try:
            length = int(self.headers.get("Content-Length", "0"))
            body = self.rfile.read(length)
            content_type = self.headers.get("Content-Type", "")
            if content_type.startswith("multipart/form-data"):
                fields, files = parse_multipart(self.headers, body)
                auto_path = files.get("autoFile")
                ha_path = files.get("haFile")
                inventory_path = files.get("inventoryFile")
                supplier_offer_path = files.get("supplierOfferFile")
                supplier_po_path = files.get("supplierPoFile")
                site_product_path = files.get("siteProductFile")
                supplier_code = fields.get("supplierCode")
                mode = fields.get("mode") or "vendor"
                output_dir = str(OUTPUT_DIR)
            else:
                payload = json.loads(body.decode("utf-8"))
                auto_path = payload.get("autoPath")
                ha_path = payload.get("haPath")
                inventory_path = payload.get("inventoryPath")
                supplier_offer_path = payload.get("supplierOfferPath")
                supplier_po_path = payload.get("supplierPoPath")
                site_product_path = payload.get("siteProductPath")
                supplier_code = payload.get("supplierCode")
                mode = payload.get("mode") or "vendor"
                output_dir = payload.get("outputDir") or str(OUTPUT_DIR)
            last_files = load_last_files()
            auto_path = auto_path or last_files.get("auto_path")
            ha_path = ha_path or last_files.get("ha_path")
            if not auto_path or not ha_path or not supplier_code:
                self._send(400, json.dumps({"error": "请上传表 A、表 B 并输入供应商编码。之后再次生成时可不重复上传表 A 和表 B。"}, ensure_ascii=False))
                return
            result = generate(auto_path, ha_path, supplier_code, output_dir, inventory_path, mode, supplier_offer_path, supplier_po_path, site_product_path)
            save_last_files(auto_path, ha_path)
            result["used_auto_path"] = str(Path(auto_path).resolve())
            result["used_ha_path"] = str(Path(ha_path).resolve())
            result["download_url"] = "/download?file=" + quote(result["output_path"])
            if result.get("po_offer"):
                po_offer = result["po_offer"]
                po_offer["po_download_url"] = "/download?file=" + quote(po_offer["po_path"])
                po_offer["error_download_url"] = "/download?file=" + quote(po_offer["error_path"])
                po_offer["po_input_download_url"] = "/download?file=" + quote(po_offer["po_input_path"])
                if po_offer.get("offer_path"):
                    po_offer["offer_download_url"] = "/download?file=" + quote(po_offer["offer_path"])
            result["workspace"] = str(WORKSPACE)
            result["output_dir"] = str(OUTPUT_DIR)
            self._send(200, json.dumps(result, ensure_ascii=False, indent=2))
        except Exception as exc:
            self._send(500, json.dumps({"error": str(exc)}, ensure_ascii=False, indent=2))

    def log_message(self, format, *args):
        return


if __name__ == "__main__":
    server = ThreadingHTTPServer((HOST, PORT), Handler)
    url = f"http://{HOST}:{PORT}/"
    if "--open" in sys.argv:
        threading.Timer(0.8, lambda: webbrowser.open(url)).start()
    print(f"Inquiry tool is running at {url}")
    server.serve_forever()
