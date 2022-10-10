import tkinter as tk
from tkinter import ttk
from watcher import Watcher

class Window(tk.Tk):
    is_watching = False # state of the watcher if its following the folder or not
    def __init__(self, *args, **kwargs) -> None:
        tk.Tk.__init__(self, *args, **kwargs)
        self.protocol("WM_DELETE_WINDOW", self.on_close)
        self.wm_title("Watcher")
        # creating a frame and assigning it to container
        container = tk.Frame(self, height=400, width=600)
        # specifying the region where the frame is packed in root
        container.pack(side="top", fill="both", expand=True)
        frame = ttk.Frame(container, padding=10)
        frame.grid()
        ttk.Label(frame, text="Folder to be watched:   ").grid(column=0, row=0, pady=20)
        self.text = tk.StringVar(value="./lookup")
        self.entry = ttk.Entry(frame, textvariable=self.text)
        self.entry.grid(column=1, row=0, pady=20)
        self.button = ttk.Button(frame, text="Start", command=self.watch)
        self.button.grid(column=0, row=1, columnspan=2)

        self.watcher = Watcher()

    def on_close(self):
        """actions to be done if the window is closed, if it's still watching then stop the watcher and then destroy"""
        if self.is_watching:
            self.watcher.pause()
            self.watcher.stop()
        self.destroy()

    def watch(self):
        """Start watching the folder stated in the Entry if it's not watching or stop if it is"""
        if self.is_watching:
            self.button['text'] = "Start"
            self.entry['state'] = "enabled"
            self.watcher.pause()
            self.watcher.stop()
        else:
            self.button['text'] = "Stop"
            self.entry['state'] = "disabled"
            path = self.text.get()
            self.watcher.observe(path)

        self.is_watching = not self.is_watching


if __name__ == "__main__":
    w = Window()
    w.mainloop()