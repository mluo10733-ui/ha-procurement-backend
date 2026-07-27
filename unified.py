import os
import signal
import subprocess
import sys
import threading
import urllib.error
import urllib.request
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer


PUBLIC_PORT = int(os.environ.get("PORT", "8080"))
HOST = os.environ.get("HOST", "0.0.0.0")
SERVICES = {
    "/inquiry": (8788, [sys.executable, "server.py"], "inquiry"),
    "/po-offer": (18888, [sys.executable, "web_tool.py", "--no-browser"], "po-offer"),
    "/reconciliation": (8799, [sys.executable, "reconcile_app_v2.py"], "reconciliation"),
}
processes = []


def start_services():
    for port, command, cwd in SERVICES.values():
        env = os.environ.copy()
        env["PORT"] = str(port)
        env["HOST"] = "127.0.0.1"
        processes.append(subprocess.Popen(command, cwd=os.path.join(os.path.dirname(__file__), cwd), env=env))


class Gateway(BaseHTTPRequestHandler):
    def _cors(self):
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Methods", "GET, POST, PUT, OPTIONS")
        self.send_header("Access-Control-Allow-Headers", "Content-Type")

    def _target(self):
        for prefix, (port, _, _) in SERVICES.items():
            if self.path == prefix or self.path.startswith(prefix + "/"):
                return f"http://127.0.0.1:{port}{self.path[len(prefix):] or '/'}"
        return None

    def _proxy(self):
        target = self._target()
        if not target:
            self.send_error(404, "Unknown service path")
            return
        length = int(self.headers.get("Content-Length", "0"))
        body = self.rfile.read(length) if length else None
        headers = {k: v for k, v in self.headers.items() if k.lower() not in {"host", "content-length"}}
        request = urllib.request.Request(target, data=body, headers=headers, method=self.command)
        try:
            with urllib.request.urlopen(request, timeout=600) as response:
                payload = response.read()
                self.send_response(response.status)
                for key, value in response.headers.items():
                    if key.lower() not in {"content-length", "transfer-encoding", "connection"}:
                        self.send_header(key, value)
                self._cors()
                self.send_header("Content-Length", str(len(payload)))
                self.end_headers()
                self.wfile.write(payload)
        except urllib.error.HTTPError as error:
            payload = error.read()
            self.send_response(error.code)
            self.send_header("Content-Type", error.headers.get("Content-Type", "text/plain"))
            self._cors()
            self.send_header("Content-Length", str(len(payload)))
            self.end_headers()
            self.wfile.write(payload)
        except Exception as error:
            self.send_response(502)
            self._cors()
            payload = f"Backend unavailable: {error}".encode()
            self.send_header("Content-Type", "text/plain; charset=utf-8")
            self.send_header("Content-Length", str(len(payload)))
            self.end_headers()
            self.wfile.write(payload)

    do_GET = _proxy
    do_POST = _proxy
    do_PUT = _proxy
    def do_OPTIONS(self):
        self.send_response(204)
        self._cors()
        self.send_header("Content-Length", "0")
        self.end_headers()


def shutdown(*_):
    for process in processes:
        process.terminate()
    raise SystemExit(0)


if __name__ == "__main__":
    signal.signal(signal.SIGTERM, shutdown)
    signal.signal(signal.SIGINT, shutdown)
    start_services()
    ThreadingHTTPServer((HOST, PUBLIC_PORT), Gateway).serve_forever()
