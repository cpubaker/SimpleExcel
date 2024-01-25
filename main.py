# main.py
import tkinter as tk
from excel import SimpleExcel

if __name__ == "__main__":
    root = tk.Tk()
    app = SimpleExcel(root)
    root.mainloop()
