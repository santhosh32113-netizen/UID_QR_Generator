"""KUIN-G desktop launcher for the packaged dashboard."""

import threading
import webbrowser

from tools.dashboard_server import DashboardHandler, PORT, ThreadingHTTPServer, dashboard_root, prepare_runtime_data, refresh_dashboard_data


def main():
	prepare_runtime_data()
	dashboard_root()
	refresh_dashboard_data()
	server = ThreadingHTTPServer(("127.0.0.1", PORT), DashboardHandler)
	threading.Timer(0.8, lambda: webbrowser.open(f"http://127.0.0.1:{PORT}/index.html")).start()
	print(f"KUIN-G dashboard: http://127.0.0.1:{PORT}/index.html")
	server.serve_forever()


if __name__ == "__main__":
	main()
