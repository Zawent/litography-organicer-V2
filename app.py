import os
import tkinter as tk
from tkinter import filedialog
import webview
import pandas as pd
import subprocess

class Api:
    def __init__(self):
        self.ruta_excel = ""

    def exportar_excel(self, datos):
        # Ocultar la ventana de diálogo raíz de Tkinter
        root = tk.Tk()
        root.withdraw()

        carpeta = filedialog.askdirectory(title="Selecciona una carpeta para guardar el Excel")

        if not carpeta:
            return "Exportación cancelada por el usuario."

        self.ruta_excel = os.path.join(carpeta, "datos_exportados.xlsx")

        try:
            df = pd.DataFrame(datos)
            df.to_excel(self.ruta_excel, index=False)
            return f"Archivo guardado exitosamente en:\n{self.ruta_excel}"
        except Exception as e:
            return f"Ocurrió un error al guardar:\n{e}"

    def ejecutar_renombrado(self, ruta_excel):
        if not ruta_excel:
            return "No se ha exportado un archivo Excel aún."

        try:
            script_exe = os.path.join(os.path.dirname(__file__), "renombre_fotos_seleccionadas.exe")
            # Usar Popen para abrir el CMD solo cuando se ejecute el proceso
            subprocess.Popen([script_exe, ruta_excel])
            return "Renombrado iniciado con éxito."
        except Exception as e:
            return f"Error inesperado:\n{e}"


if __name__ == '__main__':
    api = Api()
    window = webview.create_window("Renombrador de Archivos", "index.html", js_api=api)
    webview.start(http_server=True)
