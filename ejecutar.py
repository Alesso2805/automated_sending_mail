import subprocess
import tkinter as tk

MAIN_SCRIPT_PATH = r"C:\Users\Alessandro\PycharmProjects\automated_sending_mail\main.py"
LIBR_SCRIPT_PATH = r"C:\Users\Alessandro\PycharmProjects\automated_sending_mail\.venv\Scripts\activate"

def ejecutar_script():
    subprocess.run(["python", LIBR_SCRIPT_PATH], shell=True)
    subprocess.run(["python", MAIN_SCRIPT_PATH], shell=True)

root = tk.Tk()
root.title("Ejecutar Script")

boton = tk.Button(root, text="Enviar Correos", command=ejecutar_script, font=("Arial", 12), padx=20, pady=10)
boton.pack(pady=20)

root.mainloop()
