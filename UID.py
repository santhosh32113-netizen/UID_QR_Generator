"""KUIN-G desktop launcher for the packaged dashboard."""

import threading
import webbrowser

from tools.dashboard_server import DashboardHandler, PORT, ThreadingHTTPServer


def main():
	server = ThreadingHTTPServer(("127.0.0.1", PORT), DashboardHandler)
	threading.Timer(0.8, lambda: webbrowser.open(f"http://127.0.0.1:{PORT}/index.html")).start()
	print(f"KUIN-G dashboard: http://127.0.0.1:{PORT}/index.html")
	server.serve_forever()


if __name__ == "__main__":
	main()
