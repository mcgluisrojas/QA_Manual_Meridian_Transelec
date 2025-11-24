"""Simple HTTP API to expose QA plan data from the XLSX file.

Usage:
    python api_server.py --host 0.0.0.0 --port 8000

Endpoints:
    GET /health      -> basic readiness check
    GET /api/tests   -> returns headers and rows from sheet 1 of the plan

No external dependencies are required; everything relies on the Python
standard library.
"""
from __future__ import annotations

import argparse
import json
from http.server import BaseHTTPRequestHandler, HTTPServer
from typing import Tuple

from excel_reader import EXCEL_FILE, load_sheet


class RequestHandler(BaseHTTPRequestHandler):
    server_version = "QAPlanAPI/1.0"

    def _send_json(self, payload, status: int = 200) -> None:
        body = json.dumps(payload, ensure_ascii=False, indent=2).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def do_GET(self):  # noqa: N802 - HTTPServer interface requires this name
        if self.path == "/health":
            return self._send_json({"status": "ok"})

        if self.path == "/api/tests":
            try:
                headers, rows = load_sheet()
                return self._send_json({"headers": headers, "rows": rows})
            except FileNotFoundError as exc:
                return self._send_json({"error": str(exc)}, status=500)
            except Exception as exc:  # pragma: no cover - safety net
                return self._send_json({"error": f"Error interno: {exc}"}, status=500)

        return self._send_json({"error": "Ruta no encontrada"}, status=404)

    def log_message(self, format: str, *args) -> None:  # noqa: A003
        # Redirect default logging to stdout to keep container logs tidy
        print("%s - - [%s] %s" % (self.client_address[0], self.log_date_time_string(), format % args))


def parse_args() -> Tuple[str, int]:
    parser = argparse.ArgumentParser(description="Inicia la API del plan de QA")
    parser.add_argument("--host", default="0.0.0.0", help="Host de escucha (default: 0.0.0.0)")
    parser.add_argument("--port", type=int, default=8000, help="Puerto de escucha (default: 8000)")
    args = parser.parse_args()
    return args.host, args.port


def run_server(host: str, port: int) -> None:
    if not EXCEL_FILE.exists():
        raise SystemExit(f"No se encontró el archivo de datos: {EXCEL_FILE}")

    server = HTTPServer((host, port), RequestHandler)
    print(f"API escuchando en http://{host}:{port}\nArchivo de datos: {EXCEL_FILE}")
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\nDeteniendo el servidor...")
    finally:
        server.server_close()


if __name__ == "__main__":
    host, port = parse_args()
    run_server(host, port)
