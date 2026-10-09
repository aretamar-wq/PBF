#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Interfaz gráfica (tkinter, librería estándar) para pbf_agregar_cuentas.py.

Deben estar en la misma carpeta. Ejecutar con:
    python pbf_gui.py        (en Windows también: doble clic o "py pbf_gui.py")
"""
import io
import sys
import threading
import traceback
from contextlib import redirect_stderr, redirect_stdout
from datetime import date, timedelta
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import pbf_agregar_cuentas as core


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("PBF BASE24 - Agregar cuentas")
        self.minsize(720, 540)

        hoy = date.today().strftime("%y%m%d")
        ayer = (date.today() - timedelta(days=1)).strftime("%y%m%d")

        self.v_entrada = tk.StringVar()
        self.v_salida = tk.StringVar()
        self.v_cantidad = tk.StringVar(value="100000")
        self.v_inicial = tk.StringVar()
        self.v_umbral = tk.StringVar(value="1000000")
        self.v_fh = tk.StringVar(value=hoy)
        self.v_fd = tk.StringVar(value=ayer)

        frm = ttk.Frame(self, padding=10)
        frm.pack(fill="both", expand=True)
        frm.columnconfigure(1, weight=1)

        # Archivos
        ttk.Label(frm, text="Archivo PBF de entrada:").grid(row=0, column=0, sticky="w", pady=3)
        ttk.Entry(frm, textvariable=self.v_entrada).grid(row=0, column=1, sticky="ew", padx=5)
        ttk.Button(frm, text="Examinar...", command=self.elegir_entrada).grid(row=0, column=2)

        ttk.Label(frm, text="Archivo de salida:").grid(row=1, column=0, sticky="w", pady=3)
        ttk.Entry(frm, textvariable=self.v_salida).grid(row=1, column=1, sticky="ew", padx=5)
        ttk.Button(frm, text="Examinar...", command=self.elegir_salida).grid(row=1, column=2)

        # Parámetros
        par = ttk.LabelFrame(frm, text="Parámetros", padding=8)
        par.grid(row=2, column=0, columnspan=3, sticky="ew", pady=8)
        campos = [
            ("Cuentas a agregar por moneda:", self.v_cantidad),
            ("Cuenta inicial (vacío = automática):", self.v_inicial),
            ("Umbral de serie (fuera de serie >=):", self.v_umbral),
            ("Fecha header (AAMMDD):", self.v_fh),
            ("Fecha detalle (AAMMDD):", self.v_fd),
        ]
        for i, (txt, var) in enumerate(campos):
            ttk.Label(par, text=txt).grid(row=i, column=0, sticky="w", pady=2)
            ttk.Entry(par, textvariable=var, width=16).grid(row=i, column=1, sticky="w", padx=8)

        self.btn = ttk.Button(frm, text="Procesar", command=self.procesar)
        self.btn.grid(row=3, column=0, columnspan=3, pady=4)

        # Log
        self.log = tk.Text(frm, height=16, wrap="word", state="disabled", font=("Consolas", 9))
        sb = ttk.Scrollbar(frm, command=self.log.yview)
        self.log.configure(yscrollcommand=sb.set)
        self.log.grid(row=4, column=0, columnspan=2, sticky="nsew")
        sb.grid(row=4, column=2, sticky="ns")
        frm.rowconfigure(4, weight=1)

    # --- selección de archivos ---
    def elegir_entrada(self):
        ruta = filedialog.askopenfilename(title="Archivo PBF de entrada")
        if ruta:
            self.v_entrada.set(ruta)
            if not self.v_salida.get():
                self.v_salida.set(ruta + "_out")

    def elegir_salida(self):
        ruta = filedialog.asksaveasfilename(title="Archivo de salida")
        if ruta:
            self.v_salida.set(ruta)

    def escribir(self, texto):
        self.log.configure(state="normal")
        self.log.insert("end", texto)
        self.log.see("end")
        self.log.configure(state="disabled")

    # --- ejecución ---
    def procesar(self):
        if not self.v_entrada.get().strip():
            messagebox.showwarning("Falta dato", "Elegí el archivo PBF de entrada.")
            return
        argv = ["pbf_agregar_cuentas.py", self.v_entrada.get().strip(),
                "--cantidad", self.v_cantidad.get().strip(),
                "--umbral-serie", self.v_umbral.get().strip(),
                "--fecha-header", self.v_fh.get().strip(),
                "--fecha-detalle", self.v_fd.get().strip()]
        if self.v_salida.get().strip():
            argv += ["--salida", self.v_salida.get().strip()]
        if self.v_inicial.get().strip():
            argv += ["--cuenta-inicial", self.v_inicial.get().strip()]

        self.btn.configure(state="disabled")
        self.log.configure(state="normal")
        self.log.delete("1.0", "end")
        self.log.configure(state="disabled")
        self.escribir("Procesando...\n")
        threading.Thread(target=self.trabajar, args=(argv,), daemon=True).start()

    def trabajar(self, argv):
        """Corre core.main() en un hilo, capturando su salida y errores."""
        buf = io.StringIO()
        ok = True
        viejo = sys.argv
        sys.argv = argv
        try:
            with redirect_stdout(buf), redirect_stderr(buf):
                core.main()
        except SystemExit as e:
            # error() usa sys.exit("mensaje"); argparse usa código 2 y escribe a stderr
            if e.code not in (None, 0):
                ok = False
                if isinstance(e.code, str):
                    buf.write(e.code + "\n")
        except Exception:
            ok = False
            buf.write(traceback.format_exc())
        finally:
            sys.argv = viejo
        self.after(0, self.terminar, buf.getvalue(), ok)

    def terminar(self, salida, ok):
        self.escribir(salida)
        self.btn.configure(state="normal")
        if ok:
            messagebox.showinfo("Listo", "Archivo generado correctamente.")
        else:
            messagebox.showerror("Error", "El proceso falló. Revisá el detalle en pantalla.")


if __name__ == "__main__":
    App().mainloop()
