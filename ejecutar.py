import subprocess
import tkinter as tk
import os

MAIN_SCRIPT_PATH = r"C:\Users\Alessandro\PycharmProjects\automated_sending_mail\main.py"
LIBR_SCRIPT_PATH = r"C:\Users\Alessandro\PycharmProjects\automated_sending_mail\.venv\Scripts\activate.bat"

def ejecutar_script():
    command = f'cmd /c "{LIBR_SCRIPT_PATH} && python {MAIN_SCRIPT_PATH}"'
    subprocess.run(command, shell=True)

root = tk.Tk()
root.title("Ejecutar Script")

boton = tk.Button(root, text="Enviar Correos", command=ejecutar_script, font=("Arial", 12), padx=20, pady=10)
boton.pack(pady=20)

root.mainloop()