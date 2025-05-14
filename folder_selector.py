import json
from tkinter import Tk, filedialog

from ctypes import windll
windll.shcore.SetProcessDpiAwareness(1)

def select_folder():
    root = Tk()
    root.withdraw()  # Hide the main window
    root.attributes("-topmost", True)  # Bring dialog to front
    folder_path = filedialog.askdirectory(title='Select your folder here.')  # Open the dialog to select a folder
    if folder_path:
        print(json.dumps({"folder_path": folder_path}))
    root.destroy()

if __name__ == "__main__":
    select_folder()