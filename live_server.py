import os, json, time, webbrowser
from http.server import HTTPServer, SimpleHTTPRequestHandler
import generate_dashboard

PORT = 5050
last_mtime = 0
cached_data = []

def check_excel():
    global last_mtime, cached_data
    efile = generate_dashboard.get_excel_file()
    if not efile: return []
    mtime = os.path.getmtime(efile)
    if mtime > last_mtime:
        last_mtime = mtime
        cached_data = generate_dashboard.read_data(efile)
        html = generate_dashboard.build_html(cached_data, os.path.basename(efile))
        with open(os.path.join(generate_dashboard.folder, "DVPR_Live_Dashboard.html"), "w", encoding="utf-8") as f:
            f.write(html)
        print(f"Excel updated! Loaded {len(cached_data)} vehicles.")
    return cached_data

class LiveHandler(SimpleHTTPRequestHandler):
    def do_GET(self):
        if self.path in ("/", "/index.html"):
            check_excel()
            self.path = "/DVPR_Live_Dashboard.html"
            return super().do_GET()
        elif self.path == "/api/data":
            data = check_excel()
            payload = json.dumps({"data": data, "count": len(data), "ts": time.time()}).encode("utf-8")
            self.send_response(200)
            self.send_header("Content-Type", "application/json; charset=utf-8")
            self.send_header("Access-Control-Allow-Origin", "*")
            self.end_headers()
            self.wfile.write(payload)
        else:
            return super().do_GET()

    def log_message(self, format, *args):
        pass

def run():
    check_excel()
    os.chdir(generate_dashboard.folder)
    server = HTTPServer(("0.0.0.0", PORT), LiveHandler)
    url = f"http://localhost:{PORT}"
    print(f"DVPR Live Server running at {url}")
    webbrowser.open(url)
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        server.server_close()

if __name__ == "__main__":
    run()
