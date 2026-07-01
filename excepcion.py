import os
import tkinter as tk
from tkinter import messagebox

EXCEPCION = {
    "700172932",
    "700172943"
}
def mostrar_error(mensaje):
    root = tk.Tk()
    root.withdraw()
    messagebox.showerror("Error", mensaje)
    root.destroy()
    return False    


def mostrar_info(mensaje):
    root = tk.Tk()
    root.withdraw()
    messagebox.showinfo("Información", mensaje)
    root.destroy()
    return True

