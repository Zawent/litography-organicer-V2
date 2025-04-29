import os
import tkinter as tk
from tkinter import filedialog
import webview
import pandas as pd

class Api:

    def exportar_excel(self, datos):
        # Ocultar la ventana de diálogo raíz de Tkinter
        root = tk.Tk()
        root.withdraw()

        carpeta = filedialog.askdirectory(title="Selecciona una carpeta para guardar el Excel")

        if not carpeta:
            return "Exportación cancelada por el usuario."

        ruta = os.path.join(carpeta, "datos_exportados.xlsx")

        try:
            df = pd.DataFrame(datos)
            df.to_excel(ruta, index=False)
            return f"Archivo guardado exitosamente en:\n{ruta}"
        except Exception as e:
            return f"Ocurrió un error al guardar:\n{e}"

if __name__ == '__main__':
    api = Api()
    window = webview.create_window("Renombrador de Archivos", "index.html", js_api=api)
    webview.start(http_server=True)
