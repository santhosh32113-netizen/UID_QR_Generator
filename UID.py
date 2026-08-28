"""KUIN-G desktop launcher for the packaged dashboard."""

import threading
import webbrowser
import traceback
import socket
from pathlib import Path

from tools.dashboard_server import DashboardHandler, PORT, ThreadingHTTPServer, dashboard_root, prepare_runtime_data, refresh_dashboard_data


def main():
	prepare_runtime_data()
	dashboard_root()
	refresh_dashboard_data()
	url = f"http://127.0.0.1:{PORT}/index.html"
	try:
		with socket.create_connection(("127.0.0.1", PORT), timeout=0.25):
			print(f"KUIN-G is already running: {url}")
			webbrowser.open(url)
			return
	except OSError:
		pass
	server = ThreadingHTTPServer(("127.0.0.1", PORT), DashboardHandler)
	threading.Timer(1.0, lambda: webbrowser.open(url)).start()
	print(f"KUIN-G dashboard: {url}")
	server.serve_forever()


if __name__ == "__main__":
	try:
		main()
	except Exception as error:
		log_path = Path(__file__).resolve().parent / "startup_error.log"
		log_path.write_text(traceback.format_exc(), encoding="utf-8")
		message = f"KUIN-G could not start.\n\n{error}\n\nDetails: {log_path}"
		try:
			import tkinter.messagebox
			tkinter.messagebox.showerror("KUIN-G startup error", message)
		except Exception:
			print(message)
		input("Press Enter to close...")
