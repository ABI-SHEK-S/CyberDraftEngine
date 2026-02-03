import subprocess
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
import time
import os
import signal

WATCH_PATHS = ['main_app.py', 'db/', 'templates/', 'logic/']

class ReloadHandler(FileSystemEventHandler):
    def __init__(self):
        self.process = None
        self.start_app()

    def start_app(self):
        if self.process:
            try:
                os.kill(self.process.pid, signal.SIGTERM)
            except Exception:
                pass
        print("🔁 Restarting app...\n")
        self.process = subprocess.Popen(["python", "main.py"])

    def on_modified(self, event):
        print(f"📝 Detected change in {event.src_path}")
        self.start_app()

if __name__ == "__main__":
    observer = Observer()
    handler = ReloadHandler()
    for path in WATCH_PATHS:
        observer.schedule(handler, path='.', recursive=True)

    observer.start()
    print("🔍 Watching for changes... Press Ctrl+C to stop.")
    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        print("🛑 Stopping...")
        observer.stop()
        if handler.process:
            handler.process.terminate()
    observer.join()
