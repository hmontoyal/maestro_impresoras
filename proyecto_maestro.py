import pandas as pd
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import tkinter.font as tkFont
import matplotlib.pyplot as plt

SINONIMOS = {
    'marca': 'Marca',
    'modelo': 'Modelo',
    'n° serie': 'Serie',
    'numero de serie': 'Serie',
    'serie': 'Serie',
    'ubicacion': 'Ubicación',
    'ubicación': 'Ubicación',
    'estado': 'Estado',
    'bandeja': 'Bandeja'
}

def estandarizar_nombre(nombre):
    nombre = str(nombre).strip().lower()
    return SINONIMOS.get(nombre, nombre.capitalize())

def cargar_datos(path):
    hojas = pd.read_excel(path, sheet_name=None)
    datos_estandarizados = {}
    for nombre_hoja, df in hojas.items():
        df = df.loc[:, ~df.columns.str.contains('^Unnamed', na=False)]
        columnas = {col: estandarizar_nombre(col) for col in df.columns}
        df.rename(columns=columnas, inplace=True)
        df["Cliente"] = nombre_hoja
        datos_estandarizados[nombre_hoja] = df
    return datos_estandarizados

class BuscadorSerieApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Buscador por Serie (por cliente)")
        self.datos = {}
        self.df_global = pd.DataFrame()
        self.df_actual = pd.DataFrame()
        self.resultado_filtrado = pd.DataFrame()

        self.crear_widgets()

    def crear_widgets(self):
        tk.Button(self.root, text="Cargar archivo Excel", command=self.cargar_excel).pack(pady=10)

        self.cliente_var = tk.StringVar()
        self.combo_clientes = ttk.Combobox(self.root, textvariable=self.cliente_var)
        self.combo_clientes.bind("<<ComboboxSelected>>", self.cargar_cliente)
        self.combo_clientes.pack(pady=5)

        tk.Label(self.root, text="Buscar por N° de Serie:").pack()
        self.entrada_serie = tk.Entry(self.root)
        self.entrada_serie.insert(0, "Ingrese serie aquí...")
        self.entrada_serie.bind("<FocusIn>", lambda e: self.entrada_serie.delete(0, 'end'))
        self.entrada_serie.pack(pady=5)

        botones_frame = tk.Frame(self.root)
        botones_frame.pack(pady=5)

        tk.Button(botones_frame, text="Buscar", command=self.buscar_serie).pack(side="left", padx=5)
        tk.Button(botones_frame, text="Exportar resultado", command=self.exportar_resultado).pack(side="left", padx=5)
        tk.Button(botones_frame, text="Exportar consolidado", command=self.exportar_consolidado).pack(side="left", padx=5)

        resumen_frame = tk.Frame(self.root)
        resumen_frame.pack(pady=5)

        tk.Button(resumen_frame, text="Resumen por Marca (Total)", command=self.mostrar_resumen_marca_total).pack(side="left", padx=5)
        tk.Button(resumen_frame, text="Resumen por Cliente", command=self.mostrar_resumen_por_cliente).pack(side="left", padx=5)
        tk.Button(resumen_frame, text="Resumen por Modelo (Total)", command=self.mostrar_resumen_modelo_total).pack(side="left", padx=5)
        tk.Button(resumen_frame, text="Resumen de Marcas por Cliente", command=self.mostrar_marcas_por_cliente).pack(side="left", padx=5)

        self.tabla = ttk.Treeview(self.root)
        self.tabla.pack(expand=True, fill='both', padx=10, pady=10)

        self.scroll_y = ttk.Scrollbar(self.root, orient="vertical", command=self.tabla.yview)
        self.tabla.configure(yscroll=self.scroll_y.set)
        self.scroll_y.pack(side='right', fill='y')

    def cargar_excel(self):
        archivo = filedialog.askopenfilename(title="Selecciona archivo Excel", filetypes=[("Excel", "*.xlsx")])
        if not archivo:
            return

        self.datos = cargar_datos(archivo)
        self.df_global = pd.concat(self.datos.values(), ignore_index=True)
        self.combo_clientes['values'] = list(self.datos.keys())
        messagebox.showinfo("Archivo cargado", f"Se cargaron {len(self.df_global)} registros en total.")

    def cargar_cliente(self, event=None):
        cliente = self.cliente_var.get()
        if cliente not in self.datos:
            return
        self.df_actual = self.datos[cliente]
        self.mostrar_datos(self.df_actual, self.tabla)

    def mostrar_datos(self, df, tabla):
        tabla.delete(*tabla.get_children())
        if df.empty:
            return

        if "Estado" in df.columns:
            df = df[df["Estado"].astype(str).str.upper().str.strip() != "RETIRADA"]

        if df.empty:
            return

        tabla["columns"] = list(df.columns)
        tabla["show"] = "headings"

        font = tkFont.Font()
        style = ttk.Style()
        style.configure("Treeview.Heading", anchor="center")
        style.configure("Treeview", rowheight=25, font=("Arial", 10))

        for col in df.columns:
            tabla.heading(col, text=col, anchor="center")

            max_ancho = font.measure(col)
            for valor in df[col].astype(str):
                ancho = font.measure(valor)
                if ancho > max_ancho:
                    max_ancho = ancho

            tabla.column(col, width=max_ancho + 30, anchor="center")

        for _, row in df.iterrows():
            tabla.insert("", "end", values=list(row))

    def buscar_serie(self):
        valor = self.entrada_serie.get().strip()
        if not valor or self.df_actual.empty:
            messagebox.showwarning("Advertencia", "Debe seleccionar cliente y escribir una serie.")
            return
        if "Serie" not in self.df_actual.columns:
            messagebox.showerror("Error", "No se encontró la columna 'Serie'.")
            return
        resultado = self.df_actual[self.df_actual["Serie"].astype(str).str.contains(valor, case=False, na=False)]
        if resultado.empty:
            messagebox.showinfo("Sin resultados", "No se encontró esa serie.")
        else:
            self.resultado_filtrado = resultado
            self.mostrar_datos(resultado, self.tabla)

    def exportar_resultado(self):
        if self.resultado_filtrado.empty:
            messagebox.showwarning("Nada que exportar", "Realiza una búsqueda válida primero.")
            return
        ruta = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                            initialfile="resultado_busqueda.xlsx",
                                            filetypes=[("Excel", "*.xlsx")])
        if ruta:
            self.resultado_filtrado.to_excel(ruta, index=False)
            messagebox.showinfo("Exportado", f"Resultado guardado en:\n{ruta}")

    def exportar_consolidado(self):
        if self.df_global.empty:
            messagebox.showwarning("Nada que exportar", "Primero carga un archivo.")
            return
        ruta = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                            initialfile="consolidado_clientes.xlsx",
                                            filetypes=[("Excel", "*.xlsx")])
        if ruta:
            self.df_global.to_excel(ruta, index=False)
            messagebox.showinfo("Exportado", f"Archivo consolidado guardado en:\n{ruta}")

    def mostrar_resumen_marca_total(self):
        df = self.df_global.copy()
        if df.empty or "Marca" not in df.columns:
            messagebox.showwarning("Datos insuficientes", "Debe haber datos y columna 'Marca'.")
            return
        if "Estado" in df.columns:
            df = df[df["Estado"].astype(str).str.upper().str.strip() != "RETIRADA"]

        resumen = df["Marca"].value_counts().reset_index()
        resumen.columns = ["Marca", "Cantidad"]
        resumen["Porcentaje"] = (resumen["Cantidad"] / resumen["Cantidad"].sum() * 100).round(2)
        self.mostrar_resumen_en_ventana(resumen, "Resumen por Marca")

        plt.figure(figsize=(6, 6))
        plt.pie(resumen["Cantidad"], labels=resumen["Marca"], autopct='%1.1f%%', startangle=140)
        plt.title("Distribución Total de Marcas (Todos los Clientes)")
        plt.axis('equal')
        plt.tight_layout()
        plt.show()

    def mostrar_resumen_modelo_total(self):
        df = self.df_global.copy()
        if df.empty or "Modelo" not in df.columns:
            messagebox.showwarning("Datos insuficientes", "Debe haber datos y columna 'Modelo'.")
            return
        if "Estado" in df.columns:
            df = df[df["Estado"].astype(str).str.upper().str.strip() != "RETIRADA"]

        resumen = df["Modelo"].value_counts().reset_index()
        resumen.columns = ["Modelo", "Cantidad"]
        resumen["Porcentaje"] = (resumen["Cantidad"] / resumen["Cantidad"].sum() * 100).round(2)
        self.mostrar_resumen_en_ventana(resumen, "Resumen por Modelo")

    def mostrar_resumen_por_cliente(self):
        df = self.df_global.copy()
        if df.empty or "Cliente" not in df.columns:
            messagebox.showwarning("Datos insuficientes", "Debe haber datos y columna 'Cliente'.")
            return
        if "Estado" in df.columns:
            df = df[df["Estado"].astype(str).str.upper().str.strip() != "RETIRADA"]

        resumen = df["Cliente"].value_counts().reset_index()
        resumen.columns = ["Cliente", "Cantidad"]
        resumen["Porcentaje"] = (resumen["Cantidad"] / resumen["Cantidad"].sum() * 100).round(2)
        self.mostrar_resumen_en_ventana(resumen, "Resumen por Cliente")

    def mostrar_resumen_en_ventana(self, df_resumen, titulo):
        top = tk.Toplevel(self.root)
        top.title(titulo)

        tree = ttk.Treeview(top)
        tree.pack(expand=True, fill="both", padx=10, pady=10)

        scrollbar = ttk.Scrollbar(top, orient="vertical", command=tree.yview)
        tree.configure(yscroll=scrollbar.set)
        scrollbar.pack(side='right', fill='y')

        self.mostrar_datos(df_resumen, tree)

    def mostrar_marcas_por_cliente(self):
        df = self.df_global.copy()
        if df.empty or "Cliente" not in df.columns or "Marca" not in df.columns:
            messagebox.showwarning("Datos insuficientes", "Debe haber datos con columnas 'Cliente' y 'Marca'.")
            return
        if "Estado" in df.columns:
            df = df[df["Estado"].astype(str).str.upper().str.strip() != "RETIRADA"]

        resumen = df.groupby(["Cliente", "Marca"]).size().reset_index(name="Cantidad")
        self.mostrar_resumen_en_ventana(resumen, "Marcas por Cliente")

# Ejecutar app
if __name__ == "__main__":
    root = tk.Tk()
    app = BuscadorSerieApp(root)
    root.mainloop()
