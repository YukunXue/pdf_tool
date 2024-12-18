# utils.py

from tkinter import ttk
from tkinter import messagebox

def update_progress(progress_var, value, root, message=None):
    """
    更新进度条，并在需要时显示消息框。
    
    Parameters:
        - progress_var (ttk.DoubleVar): Tkinter DoubleVar for the progressbar.
        - value (float): Progress value (0.0 to 100.0).
        - root (tk.Tk): Tkinter root window.
        - message (str): Optional message to display in a messagebox.
    """
    progress_var.set(value)
    root.update_idletasks()
    
    if message:
        messagebox.showinfo("Info", message)



def get_version():
    return "0.0.1"


