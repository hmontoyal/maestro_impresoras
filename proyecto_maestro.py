import pandas as pd
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
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
        df["Cliente"] = nombre_hoja  # Añadir columna con nombre de hoja
        datos_estandarizados[nombre_hoja] = df

    return datos_estandarizados

class BuscadorSerieApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Buscador por Serie (por cliente)")
        self.datos = {}
        self.df_global = pd.DataFrame()
        self.df_actual = pd.DataFrame()

        self.crear_widgets()

    def crear_widgets(self):
        tk.Button(self.root, text="Cargar archivo Excel", command=self.cargar_excel).pack(pady=10)

        self.cliente_var = tk.StringVar()
        self.combo_clientes = ttk.Combobox(self.root, textvariable=self.cliente_var, state="normal")
        self.combo_clientes.bind("<<ComboboxSelected>>", self.cargar_cliente)
        self.combo_clientes.pack(pady=5)

        tk.Label(self.root, text="Buscar por N° de Serie:").pack()
        self.entrada_serie = tk.Entry(self.root)
        self.entrada_serie.pack(pady=5)

        botones_frame = tk.Frame(self.root)
        botones_frame.pack(pady=5)

        tk.Button(botones_frame, text="Buscar", command=self.buscar_serie).pack(side="left", padx=5)
        tk.Button(botones_frame, text="Exportar resultado", command=self.exportar_resultado).pack(side="left", padx=5)
        tk.Button(botones_frame, text="Resumen por Marca (Total)", command=self.mostrar_resumen_marca_total).pack(side="left", padx=5)

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
        self.mostrar_datos(self.df_actual)

    def mostrar_datos(self, df):
        self.tabla.delete(*self.tabla.get_children())
        self.tabla["columns"] = list(df.columns)
        self.tabla["show"] = "headings"
        for col in df.columns:
            self.tabla.heading(col, text=col)
            self.tabla.column(col, width=100)

        for _, row in df.iterrows():
            self.tabla.insert("", "end", values=list(row))

    def buscar_serie(self):
        valor = self.entrada_serie.get().strip()
        if not valor:
            messagebox.showwarning("Sin valor", "Ingrese un número de serie.")
            return
        if self.df_actual.empty or "Serie" not in self.df_actual.columns:
            messagebox.showerror("Error", "No hay datos cargados o falta columna 'Serie'.")
            return

        resultado = self.df_actual[self.df_actual["Serie"].astype(str).str.contains(valor, case=False, na=False)]
        if resultado.empty:
            messagebox.showinfo("Sin resultados", "No se encontró esa serie.")
        else:
            self.mostrar_datos(resultado)
            self.resultado_filtrado = resultado

    def exportar_resultado(self):
        if not hasattr(self, 'resultado_filtrado') or self.resultado_filtrado.empty:
            messagebox.showwarning("Nada que exportar", "Realiza una búsqueda válida primero.")
            return

        ruta = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel", "*.xlsx")])
        if ruta:
            self.resultado_filtrado.to_excel(ruta, index=False)
            messagebox.showinfo("Exportado", f"Resultado guardado en:\n{ruta}")

    def mostrar_resumen_marca_total(self):
        if self.df_global.empty:
            messagebox.showwarning("Datos no cargados", "Carga un archivo primero.")
            return
        if "Marca" not in self.df_global.columns:
            messagebox.showerror("Error", "No se encontró la columna 'Marca'.")
            return

        resumen = self.df_global["Marca"].value_counts().reset_index()
        resumen.columns = ["Marca", "Cantidad"]
        resumen["Porcentaje"] = (resumen["Cantidad"] / resumen["Cantidad"].sum() * 100).round(2)

        self.mostrar_datos(resumen)

        # Gráfico
        plt.figure(figsize=(6, 6))
        plt.pie(resumen["Cantidad"], labels=resumen["Marca"], autopct='%1.1f%%', startangle=140)
        plt.title("Distribución Total de Marcas (Todos los Clientes)")
        plt.axis('equal')
        plt.tight_layout()
        plt.show()

# Ejecutar app
if __name__ == "__main__":
    root = tk.Tk()
    app = BuscadorSerieApp(root)
    root.mainloop()
