import tkinter as tk
import sqlite3
import openpyxl as xl
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font
import func as f
from ttkthemes import ThemedStyle

root = tk.Tk()
style = ThemedStyle(root)
style.set_theme("arc")

screen_width = root.winfo_screenwidth()
screen_height = root.winfo_screenheight()
window_width = 900
window_height = 500
x = (screen_width - window_width) // 2
y = (screen_height - window_height) // 2

root.geometry(f"{window_width}x{window_height}+{x}+{y}")
root.title("Rejestracja Pracowników")
isExisting = f.BaseCheck()
f.DrawList(root, isExisting)
root.mainloop()

