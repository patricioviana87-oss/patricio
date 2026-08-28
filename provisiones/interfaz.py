"""
Ventana simple para generar los asientos de provisiones sin usar la terminal:
elegís el archivo de provisiones, indicás el año y generás con un botón.

Usa la misma lógica que generar.py (línea de comandos); esto es solo una
capa visual arriba. Se abre haciendo doble clic en Ejecutar.bat.
"""
from __future__ import annotations

import os
import tkinter as tk
import traceback
from datetime import date
from pathlib import Path
from tkinter import filedialog, messagebox

from generar import generar_asientos


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Asientos de Provisiones")
        self.geometry("560x400")
        self.resizable(False, False)

        self.archivo = tk.StringVar()
        self.anio = tk.StringVar(value=str(date.today().year))

        tk.Label(self, text="1. Archivo de provisiones (.xlsx):", anchor="w").pack(fill="x", padx=16, pady=(16, 4))
        frame_archivo = tk.Frame(self)
        frame_archivo.pack(fill="x", padx=16)
        tk.Entry(frame_archivo, textvariable=self.archivo, state="readonly").pack(side="left", fill="x", expand=True)
        tk.Button(frame_archivo, text="Seleccionar...", command=self.elegir_archivo).pack(side="left", padx=(8, 0))

        tk.Label(self, text="2. Año del ejercicio:", anchor="w").pack(fill="x", padx=16, pady=(16, 4))
        tk.Entry(self, textvariable=self.anio, width=10).pack(anchor="w", padx=16)

        tk.Button(
            self, text="Generar asientos", command=self.generar,
            bg="#1F4E78", fg="white", font=("Segoe UI", 11, "bold"), height=2,
        ).pack(fill="x", padx=16, pady=20)

        tk.Label(self, text="Resultado:", anchor="w").pack(fill="x", padx=16)
        self.texto = tk.Text(self, height=9, state="disabled")
        self.texto.pack(fill="both", expand=True, padx=16, pady=(4, 16))

    def elegir_archivo(self):
        path = filedialog.askopenfilename(
            title="Elegí el archivo de provisiones",
            filetypes=[("Excel", "*.xlsx")],
        )
        if path:
            self.archivo.set(path)

    def escribir(self, linea: str):
        self.texto.configure(state="normal")
        self.texto.insert("end", linea + "\n")
        self.texto.configure(state="disabled")
        self.texto.see("end")
        self.update_idletasks()

    def generar(self):
        self.texto.configure(state="normal")
        self.texto.delete("1.0", "end")
        self.texto.configure(state="disabled")

        archivo = self.archivo.get().strip()
        if not archivo:
            messagebox.showwarning("Falta el archivo", "Primero elegí el archivo de provisiones.")
            return
        try:
            anio = int(self.anio.get().strip())
        except ValueError:
            messagebox.showwarning("Año inválido", "El año tiene que ser un número, ej: 2026.")
            return

        carpeta_salida = Path(archivo).parent / "Asientos_generados"
        try:
            archivos = generar_asientos(archivo, anio, carpeta_salida, on_linea=self.escribir)
        except Exception as e:
            traceback.print_exc()
            messagebox.showerror("Error", f"No se pudo generar los asientos:\n\n{e}")
            return

        self.escribir(f"\nListo: {len(archivos)} archivo(s) generados.")
        messagebox.showinfo("Listo", f"Se generaron {len(archivos)} archivo(s) en:\n{carpeta_salida}")
        try:
            os.startfile(carpeta_salida)  # type: ignore[attr-defined]
        except Exception:
            pass


if __name__ == "__main__":
    App().mainloop()
