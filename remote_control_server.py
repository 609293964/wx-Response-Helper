import datetime
import hmac
import json
import threading
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from urllib.parse import urlparse


API_VERSION = 1
SERVICE_NAME = "EasyChat Momo"


COMMAND_ROUTES = {
    "/api/narrator/start": "narrator_start",
    "/api/narrator/stop": "narrator_stop",
    "/api/monitoring/start": "monitoring_start",
    "/api/monitoring/stop": "monitoring_stop",
    "/api/scheduled/start": "scheduled_start",
    "/api/scheduled/stop": "scheduled_stop",
    "/api/check": "run_check",
}


class RemoteControlServer:
    def __init__(
        self,
        host,
        port,
        token,
        status_provider,
        command_handler,
        logger=None,
    ):
        self.host = host
        self.port = int(port)
        self.token = token
        self.status_provider = status_provider
        self.command_handler = command_handler
        self.logger = logger
        self._server = None
        self._thread = None

    def start(self):
        if self._server is not None:
            return

        owner = self

        class RequestHandler(BaseHTTPRequestHandler):
            server_version = "EasyChatMomoRemote/1.0"

            def do_GET(self):
                path = urlparse(self.path).path
                if path == "/api/health":
                    self._send_json(200, owner.get_health_payload())
                    return
                if not self._is_authorized():
                    self._send_json(401, {"ok": False, "error": "unauthorized"})
                    return
                if path != "/api/status":
                    self._send_json(404, {"ok": False, "error": "not_found"})
                    return
                try:
                    status = owner.status_provider()
                    status["ok"] = True
                    status.setdefault("service", SERVICE_NAME)
                    status.setdefault("api_version", API_VERSION)
                    self._send_json(200, status)
                except Exception as exc:
                    owner._log(f"远程状态读取失败: {exc}")
                    self._send_json(500, {"ok": False, "error": str(exc)})

            def do_POST(self):
                path = urlparse(self.path).path
                if not self._is_authorized():
                    self._send_json(401, {"ok": False, "error": "unauthorized"})
                    return
                command = COMMAND_ROUTES.get(path)
                if command is None:
                    self._send_json(404, {"ok": False, "error": "not_found"})
                    return
                try:
                    owner.command_handler(command)
                    self._send_json(202, {"ok": True, "command": command, "accepted": True})
                except Exception as exc:
                    owner._log(f"远程命令提交失败: {exc}")
                    self._send_json(500, {"ok": False, "error": str(exc)})

            def log_message(self, format_string, *args):
                owner._log("远程请求: " + (format_string % args))

            def _is_authorized(self):
                supplied = self.headers.get("X-API-Token", "")
                return bool(owner.token) and hmac.compare_digest(supplied, owner.token)

            def _send_json(self, status_code, payload):
                body = json.dumps(payload, ensure_ascii=False).encode("utf-8")
                self.send_response(status_code)
                self.send_header("Content-Type", "application/json; charset=utf-8")
                self.send_header("Cache-Control", "no-store")
                self.send_header("Content-Length", str(len(body)))
                self.end_headers()
                self.wfile.write(body)

        self._server = ThreadingHTTPServer((self.host, self.port), RequestHandler)
        self._server.daemon_threads = True
        self._thread = threading.Thread(
            target=self._server.serve_forever,
            name="momo-remote-control",
            daemon=True,
        )
        self._thread.start()

    def stop(self):
        server = self._server
        thread = self._thread
        self._server = None
        self._thread = None
        if server is not None:
            server.shutdown()
            server.server_close()
        if thread is not None and thread.is_alive():
            thread.join(timeout=2)

    def get_health_payload(self):
        return {
            "ok": True,
            "service": SERVICE_NAME,
            "api_version": API_VERSION,
            "server_time": datetime.datetime.now().astimezone().isoformat(timespec="seconds"),
            "remote_enabled": True,
        }

    def _log(self, message):
        if self.logger is not None:
            self.logger(message)
