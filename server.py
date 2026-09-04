#!/usr/bin/env python3
import csv
import io
import json
import re
import sys
import urllib.error
import urllib.parse
import urllib.request
from http.server import BaseHTTPRequestHandler, SimpleHTTPRequestHandler, ThreadingHTTPServer


REPORTS = {
    "m365-apps": "/v1.0/reports/getM365AppUserDetail(period='{period}')?$format=text/csv",
    "copilot": "/v1.0/copilot/reports/getMicrosoft365CopilotUsageUserDetail(period='{period}',version='v2')",
    "sharepoint-activity": "/v1.0/reports/getSharePointActivityUserDetail(period='{period}')?$format=text/csv",
    "onedrive-activity": "/v1.0/reports/getOneDriveActivityUserDetail(period='{period}')?$format=text/csv",
    "viva-engage": "/v1.0/reports/getYammerActivityUserDetail(period='{period}')?$format=text/csv",
}
REPORT_PERIODS = {"D7", "D30", "D90", "D180"}
REPORT_DOWNLOAD_HOST = re.compile(r"^reports[a-z0-9-]*\.office\.com$")
MAX_REPORT_BYTES = 20 * 1024 * 1024
MAX_ERROR_BYTES = 64 * 1024
MAX_AUTHORIZATION_BYTES = 16 * 1024


class NoRedirect(urllib.request.HTTPRedirectHandler):
    def redirect_request(self, request, file_pointer, code, message, headers, new_url):
        return None


def field_name(header):
    value = re.sub(r"\(([^)]+)\)", r" \1", header.strip())
    value = re.sub(r"[^a-zA-Z0-9]+(.)?", lambda match: (match.group(1) or "").upper(), value)
    return value[:1].lower() + value[1:]


def parse_report(content):
    reader = csv.DictReader(io.StringIO(content.decode("utf-8-sig")))
    return [{field_name(key): value.lower() == "true" if isinstance(value, str) and value.lower() in ("true", "false") else value or "" for key, value in row.items()} for row in reader]


def is_report_download_url(location):
    try:
        target = urllib.parse.urlparse(location)
        return target.scheme == "https" and target.port in (None, 443) and not target.username and REPORT_DOWNLOAD_HOST.fullmatch(target.hostname or "") and target.path.startswith("/data/")
    except ValueError:
        return False


def report_url(report, period):
    if report not in REPORTS or period not in REPORT_PERIODS:
        raise ValueError("Unzulässiger Report oder Zeitraum.")
    return "https://graph.microsoft.com" + REPORTS[report].format(period=period)


def download_report(report, period, authorization):
    request = urllib.request.Request(report_url(report, period), headers={"Authorization": authorization, "Accept": "text/csv"})
    try:
        response = urllib.request.build_opener(NoRedirect).open(request, timeout=120)
    except urllib.error.HTTPError as error:
        if error.code not in (301, 302, 303, 307, 308):
            raise
        location = error.headers.get("Location", "")
        if not is_report_download_url(location):
            raise ValueError("Microsoft Graph lieferte ein unerwartetes Download-Ziel.")
        response = urllib.request.build_opener(NoRedirect).open(urllib.request.Request(location, headers={"Accept": "text/csv"}), timeout=120)
    with response:
        content = response.read(MAX_REPORT_BYTES + 1)
        if len(content) > MAX_REPORT_BYTES:
            raise ValueError("Der Microsoft-Report überschreitet das 20-MB-Sicherheitslimit.")
        return parse_report(content)


class Handler(BaseHTTPRequestHandler):
    server_version = "TenantScope"
    sys_version = ""

    def do_GET(self):
        if self.path == "/healthz":
            return self.send_json(200, {"status": "ok"})
        self.send_json(405, {"error": {"code": "methodNotAllowed", "message": "Methode nicht erlaubt."}})

    def do_POST(self):
        target = urllib.parse.urlsplit(self.path)
        report = target.path.removeprefix("/api/reports/")
        query = urllib.parse.parse_qs(target.query, keep_blank_values=True)
        period = query.get("period", ["D90"])[0]
        authorization = self.headers.get("Authorization", "")
        if target.path != f"/api/reports/{report}" or report not in REPORTS or set(query) - {"period"}:
            return self.send_json(404, {"error": {"code": "notFound", "message": "Report nicht gefunden."}})
        if len(query.get("period", [])) > 1 or period not in REPORT_PERIODS:
            return self.send_json(400, {"error": {"code": "invalidPeriod", "message": "Zeitraum muss D7, D30, D90 oder D180 sein."}})
        if not authorization.startswith("Bearer ") or not 32 <= len(authorization) <= MAX_AUTHORIZATION_BYTES:
            return self.send_json(401, {"error": {"code": "unauthorized", "message": "Microsoft-Graph-Token fehlt."}})
        try:
            rows = download_report(report, period, authorization)
            self.send_json(200, {"value": rows})
        except urllib.error.HTTPError as error:
            self.audit_error(report, f"graph-http-{error.code}")
            try:
                payload = json.loads(error.read(MAX_ERROR_BYTES).decode("utf-8"))
            except (UnicodeDecodeError, json.JSONDecodeError):
                payload = {"error": {"code": "graphError", "message": f"Microsoft Graph antwortete mit HTTP {error.code}."}}
            self.send_json(error.code, payload)
        except ValueError as error:
            self.audit_error(report, "validation")
            self.send_json(502, {"error": {"code": "reportDownloadFailed", "message": str(error)}})
        except (urllib.error.URLError, TimeoutError):
            self.audit_error(report, "upstream-unavailable")
            self.send_json(502, {"error": {"code": "reportDownloadFailed", "message": "Microsoft Reports ist vorübergehend nicht erreichbar."}})
        except Exception:
            self.audit_error(report, "processing")
            self.send_json(502, {"error": {"code": "reportDownloadFailed", "message": "Der Report konnte nicht sicher verarbeitet werden."}})

    def audit_error(self, report, code):
        print(f"report={report} error={code}", file=sys.stderr, flush=True)

    def send_json(self, status, payload):
        content = json.dumps(payload, ensure_ascii=False).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Cache-Control", "no-store")
        self.send_header("X-Content-Type-Options", "nosniff")
        self.send_header("Content-Length", str(len(content)))
        self.end_headers()
        self.wfile.write(content)

    def log_message(self, *_):
        pass


class DevHandler(Handler, SimpleHTTPRequestHandler):
    def do_GET(self):
        Handler.do_GET(self) if self.path == "/healthz" else SimpleHTTPRequestHandler.do_GET(self)


if __name__ == "__main__":
    if "--self-test" in sys.argv:
        assert parse_report(b"User Principal Name,PowerPoint (Web)\r\nada@example.com,True\r\n") == [{"userPrincipalName": "ada@example.com", "powerPointWeb": True}]
        assert is_report_download_url("https://reports.office.com/data/download/one")
        assert is_report_download_url("https://reportsweu.office.com/data/v1.0/download?id=one")
        assert not is_report_download_url("https://reports.office.com.evil.example/data/download/one")
        assert not is_report_download_url("http://reportsweu.office.com/data/download/one")
        assert not is_report_download_url("https://reports.office.com/private/download/one")
        assert report_url("sharepoint-activity", "D30") == "https://graph.microsoft.com/v1.0/reports/getSharePointActivityUserDetail(period='D30')?$format=text/csv"
        try:
            report_url("m365-apps", "D365")
            raise AssertionError("invalid report period accepted")
        except ValueError:
            pass
        assert MAX_REPORT_BYTES == 20 * 1024 * 1024
        print("Report proxy self-check passed")
    else:
        port = int(sys.argv[sys.argv.index("--port") + 1]) if "--port" in sys.argv else 4173
        handler = DevHandler if "--serve-static" in sys.argv else Handler
        ThreadingHTTPServer(("127.0.0.1", port), handler).serve_forever()
