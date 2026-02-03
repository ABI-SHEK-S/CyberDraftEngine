import time
import sys
import os
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler

class ChangeHandler(FileSystemEventHandler):
    def on_modified(self, event):
        if event.src_path.endswith('main.py') or event.src_path.endswith('.py') and ('gui/' in event.src_path or 'db/' in event.src_path or 'utils/' in event.src_path):
            print("File changed, restarting...")
            python = sys.executable
            os.execl(python, python, *sys.argv)

if __name__ == "__main__":
    event_handler = ChangeHandler()
    observer = Observer()
    observer.schedule(event_handler, path='.', recursive=True)
    observer.start()

    try:
        print("Watching for changes... Run your script with: python main.py")
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        observer.stop()
    observer.join()