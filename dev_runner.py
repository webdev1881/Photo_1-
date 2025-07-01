import subprocess
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
import time
import os

class RestartOnChange(FileSystemEventHandler):
    def __init__(self, script):
        self.script = script
        self.process = subprocess.Popen(["python", self.script])

    def on_any_event(self, event):
        if event.src_path.endswith(".py"):
            print(f"🔄 Изменения в {event.src_path}, перезапуск...")
            self.process.kill()
            self.process = subprocess.Popen(["python", self.script])

    def run(self):
        observer = Observer()
        observer.schedule(self, ".", recursive=True)
        observer.start()
        try:
            while True:
                time.sleep(1)
        except KeyboardInterrupt:
            self.process.kill()
            observer.stop()
        observer.join()

if __name__ == "__main__":
    watcher = RestartOnChange("main.py")
    watcher.run()
