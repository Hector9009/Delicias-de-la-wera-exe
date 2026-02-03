"""
Delicias de la Wera - app local (todo en uno)
Guardar como: delicias_de_la_wera.py

Requisitos:
    pip install pandas openpyxl

Probar:
    python delicias_de_la_wera.py

Compilar a .exe (opcional):
    pip install pyinstaller
    pyinstaller --onefile delicias_de_la_wera.py
    pyinstaller --onefile --noconsole delicias_de_la_wera.py
"""
import os
import shutil
import pandas as pd
from datetime import datetime, date
import tkinter as tk
from tkinter import ttk, messagebox, simpledialog, filedialog

# Archivo único con varias hojas
DATA_FILE = "delicias_de_la_wera.xlsx"
BACKUP_DIR = "backups"

# Nombres de hojas
SHEET_INV = "Inventario"
SHEET_VEN = "Ventas"
SHEET_DEU = "Deudas"
SHEET_TRA = "Transferencias"
SHEET_RES = "ResumenPagos"
SHEET_GAN = "Ganancias"  # resumen mensual (ventas + ganancia)

INV_COLS = ["Código", "Nombre", "PrecioCompra", "PrecioVenta", "Stock", "Categoría"]


# -------------------- Helpers para archivos (todo en uno) --------------------
def asegurarmisarchivos():
    """Crea el excel con hojas si no existe."""
    if not os.path.exists(DATA_FILE):
        df_inv = pd.DataFrame(columns=INV_COLS)

        # Ventas: ahora guarda PrecioCompra y Ganancia
        df_ven = pd.DataFrame(columns=[
            "Fecha","Código","Nombre","Cantidad","PrecioVenta","PrecioCompra","Total","Ganancia",
            "Persona","Tipo","Descripción"
        ])
        df_deu = pd.DataFrame(columns=["Persona","Adeuda","Pagado", "TotalDeuda", "Estado"])
        df_tra = pd.DataFrame(columns=["Fecha","Código","Nombre","Cantidad","Precio","Total","Persona","Cuenta","Descripción"])
        df_res = pd.DataFrame(columns=["Persona", "TotalEfectivo", "TotalTransferencia", "TotalFiado", "TotalPagado", "DeudaActual", "UltimaActualizacion"])
        df_gan = pd.DataFrame(columns=["Mes", "TotalVentasMes", "TotalGananciaMes", "UltimaActualizacion"])

        with pd.ExcelWriter(DATA_FILE, engine="openpyxl") as w:
            df_inv.to_excel(w, sheet_name=SHEET_INV, index=False)
            df_ven.to_excel(w, sheet_name=SHEET_VEN, index=False)
            df_deu.to_excel(w, sheet_name=SHEET_DEU, index=False)
            df_tra.to_excel(w, sheet_name=SHEET_TRA, index=False)
            df_res.to_excel(w, sheet_name=SHEET_RES, index=False)
            df_gan.to_excel(w, sheet_name=SHEET_GAN, index=False)

    if not os.path.exists(BACKUP_DIR):
        os.makedirs(BACKUP_DIR, exist_ok=True)


def _df_vacio_por_hoja(sheet):
    if sheet == SHEET_INV:
        return pd.DataFrame(columns=INV_COLS)
    if sheet == SHEET_VEN:
        return pd.DataFrame(columns=[
            "Fecha","Código","Nombre","Cantidad","PrecioVenta","PrecioCompra","Total","Ganancia",
            "Persona","Tipo","Descripción"
        ])
    if sheet == SHEET_DEU:
        return pd.DataFrame(columns=["Persona","Adeuda","Pagado","TotalDeuda","Estado"])
    if sheet == SHEET_TRA:
        return pd.DataFrame(columns=["Fecha","Código","Nombre","Cantidad","Precio","Total","Persona","Cuenta","Descripción"])
    if sheet == SHEET_RES:
        return pd.DataFrame(columns=["Persona", "TotalEfectivo", "TotalTransferencia", "TotalFiado", "TotalPagado", "DeudaActual", "UltimaActualizacion"])
    if sheet == SHEET_GAN:
        return pd.DataFrame(columns=["Mes", "TotalVentasMes", "TotalGananciaMes", "UltimaActualizacion"])
    return pd.DataFrame()


def cargar_hoja(sheet):
    """
    Carga una hoja del Excel sin recursión (evita RecursionError).
    Si falla leer, devuelve DF vacío con columnas correctas.
    """
    asegurarmisarchivos()
    try:
        xls = pd.ExcelFile(DATA_FILE, engine="openpyxl")
        if sheet not in xls.sheet_names:
            return _df_vacio_por_hoja(sheet)

        df = pd.read_excel(DATA_FILE, sheet_name=sheet, dtype=str, engine="openpyxl")
    except Exception:
        return _df_vacio_por_hoja(sheet)

    # Normalización por hoja
    if sheet == SHEET_INV:
        for c in INV_COLS:
            if c not in df.columns:
                df[c] = ""
        df["PrecioCompra"] = pd.to_numeric(df["PrecioCompra"], errors="coerce").fillna(0.0)
        df["PrecioVenta"] = pd.to_numeric(df["PrecioVenta"], errors="coerce").fillna(0.0)
        df["Stock"] = pd.to_numeric(df["Stock"], errors="coerce").fillna(0).astype(int)

    elif sheet == SHEET_VEN:
        for c in ["Fecha","Código","Nombre","Cantidad","PrecioVenta","PrecioCompra","Total","Ganancia","Persona","Tipo","Descripción"]:
            if c not in df.columns:
                df[c] = ""
        df["Cantidad"] = pd.to_numeric(df["Cantidad"], errors="coerce").fillna(0).astype(int)
        df["PrecioVenta"] = pd.to_numeric(df["PrecioVenta"], errors="coerce").fillna(0.0)
        df["PrecioCompra"] = pd.to_numeric(df["PrecioCompra"], errors="coerce").fillna(0.0)
        df["Total"] = pd.to_numeric(df["Total"], errors="coerce").fillna(0.0)
        df["Ganancia"] = pd.to_numeric(df["Ganancia"], errors="coerce").fillna(0.0)

    elif sheet == SHEET_DEU:
        for c in ["Persona","Adeuda","Pagado","TotalDeuda", "Estado"]:
            if c not in df.columns:
                df[c] = ""
        df["Adeuda"] = pd.to_numeric(df["Adeuda"], errors="coerce").fillna(0.0)
        df["Pagado"] = pd.to_numeric(df["Pagado"], errors="coerce").fillna(0.0)
        df["TotalDeuda"] = pd.to_numeric(df["TotalDeuda"], errors="coerce").fillna(0.0)

    elif sheet == SHEET_RES:
        for c in ["Persona", "TotalEfectivo", "TotalTransferencia", "TotalFiado", "TotalPagado", "DeudaActual", "UltimaActualizacion"]:
            if c not in df.columns:
                df[c] = ""
        df["TotalEfectivo"] = pd.to_numeric(df["TotalEfectivo"], errors="coerce").fillna(0.0)
        df["TotalTransferencia"] = pd.to_numeric(df["TotalTransferencia"], errors="coerce").fillna(0.0)
        df["TotalFiado"] = pd.to_numeric(df["TotalFiado"], errors="coerce").fillna(0.0)
        df["TotalPagado"] = pd.to_numeric(df["TotalPagado"], errors="coerce").fillna(0.0)
        df["DeudaActual"] = pd.to_numeric(df["DeudaActual"], errors="coerce").fillna(0.0)

    elif sheet == SHEET_GAN:
        for c in ["Mes", "TotalVentasMes", "TotalGananciaMes", "UltimaActualizacion"]:
            if c not in df.columns:
                df[c] = ""
        df["TotalVentasMes"] = pd.to_numeric(df["TotalVentasMes"], errors="coerce").fillna(0.0)
        df["TotalGananciaMes"] = pd.to_numeric(df["TotalGananciaMes"], errors="coerce").fillna(0.0)

    return df


def guardar_todo(df_inv, df_ven, df_deu, df_tra, df_res, df_gan):
    with pd.ExcelWriter(DATA_FILE, engine="openpyxl") as w:
        df_inv.to_excel(w, sheet_name=SHEET_INV, index=False)
        df_ven.to_excel(w, sheet_name=SHEET_VEN, index=False)
        df_deu.to_excel(w, sheet_name=SHEET_DEU, index=False)
        df_tra.to_excel(w, sheet_name=SHEET_TRA, index=False)
        df_res.to_excel(w, sheet_name=SHEET_RES, index=False)
        df_gan.to_excel(w, sheet_name=SHEET_GAN, index=False)


def hacer_backup():
    t = datetime.now().strftime("%Y%m%d_%H%M%S")
    dest = os.path.join(BACKUP_DIR, f"backup_{t}")
    os.makedirs(dest, exist_ok=True)
    try:
        pd.read_excel(DATA_FILE, sheet_name=None, engine="openpyxl")
        shutil.copy(DATA_FILE, os.path.join(dest, DATA_FILE))
        return dest
    except Exception as e:
        return f"Error backup: {e}"


# -------------------- App --------------------
class DeliciasApp:
    def __init__(self, root):
        self.root = root
        root.title("Delicias de la Wera")
        root.geometry("1050x720")
        root.configure(bg="#faf7ff")

        asegurarmisarchivos()
        self.load_dataframes()

        # Top bar
        top = ttk.Frame(root, padding=8)
        top.pack(fill="x")
        ttk.Label(top, text="Buscar (código o nombre):").pack(side="left")
        self.search_var = tk.StringVar()
        ent = ttk.Entry(top, textvariable=self.search_var, width=36)
        ent.pack(side="left", padx=6)
        ent.bind("<Return>", lambda e: self.refresh_table())
        ttk.Button(top, text="Buscar", command=self.refresh_table).pack(side="left")
        ttk.Button(top, text="Refrescar", command=self.reload).pack(side="left", padx=6)
        ttk.Button(top, text="Exportar / Guardar", command=self.exportar).pack(side="right", padx=6)
        ttk.Button(top, text="Respaldar", command=self.ui_backup).pack(side="right", padx=6)

        style = ttk.Style()
        style.theme_use("default")
        style.configure("TButton", padding=6)

        # NOTEBOOK: Inventario + Reportes
        self.nb = ttk.Notebook(root)
        self.nb.pack(fill="both", expand=True, padx=10, pady=8)

        self.tab_inv = ttk.Frame(self.nb)
        self.tab_rep = ttk.Frame(self.nb)
        self.nb.add(self.tab_inv, text="Inventario")
        self.nb.add(self.tab_rep, text="Reportes / Ganancias")

        # ------- TAB Inventario -------
        cols = ("Código","Nombre","PrecioVenta","Stock","Categoría")
        self.tree = ttk.Treeview(self.tab_inv, columns=cols, show="headings", height=18)
        for c in cols:
            self.tree.heading(c, text=c)
            self.tree.column(c, anchor="center", width=170)
        self.tree.pack(fill="both", expand=True, padx=10, pady=8)

        self.tree_data = {}
        self.tree.bind("<Double-1>", lambda e: self.open_edit_selected())

        btn_frame = ttk.Frame(self.tab_inv, padding=8)
        btn_frame.pack(fill="x")
        ttk.Button(btn_frame, text="Agregar producto", command=self.ui_add).pack(side="left", padx=4)
        ttk.Button(btn_frame, text="Editar / Abastecer", command=self.open_edit_selected).pack(side="left", padx=4)
        ttk.Button(btn_frame, text="Eliminar producto", command=self.ui_delete_product).pack(side="left", padx=4)
        ttk.Button(btn_frame, text="Venta - Efectivo", command=lambda: self.ui_sale("Efectivo")).pack(side="left", padx=4)
        ttk.Button(btn_frame, text="Venta - Fiado", command=lambda: self.ui_sale("Fiado")).pack(side="left", padx=4)
        ttk.Button(btn_frame, text="Venta - Transferencia", command=lambda: self.ui_sale("Transferencia")).pack(side="left", padx=4)
        ttk.Button(btn_frame, text="Registrar pago", command=self.ui_register_payment).pack(side="left", padx=4)
        ttk.Button(btn_frame, text="Ver deudores", command=self.ui_view_debtors).pack(side="right", padx=4)
        ttk.Button(btn_frame, text="Ver resumen pagos", command=self.ui_view_resumen_pagos).pack(side="right", padx=4)

        # ------- TAB Reportes / Ganancias -------
        self.build_report_tab()

        # status
        self.status_var = tk.StringVar()
        ttk.Label(root, textvariable=self.status_var).pack(side="bottom", fill="x")

        self.refresh_table()
        self.refresh_reports()

    # ---------------- Data load/save ----------------
    def load_dataframes(self):
        self.df_inv = cargar_hoja(SHEET_INV)
        self.df_ven = cargar_hoja(SHEET_VEN)
        self.df_deu = cargar_hoja(SHEET_DEU)
        self.df_tra = cargar_hoja(SHEET_TRA)
        self.df_res = cargar_hoja(SHEET_RES)
        self.df_gan = cargar_hoja(SHEET_GAN)

        # Deudas: recalcular total/estado si aplica
        if "TotalDeuda" not in self.df_deu.columns:
            self.df_deu["TotalDeuda"] = 0.0

        self.df_deu["TotalDeuda"] = pd.to_numeric(self.df_deu["TotalDeuda"], errors="coerce").fillna(0.0)
        self.df_deu["Adeuda"] = pd.to_numeric(self.df_deu["Adeuda"], errors="coerce").fillna(0.0)
        self.df_deu["Pagado"] = pd.to_numeric(self.df_deu["Pagado"], errors="coerce").fillna(0.0)

        # Si hay filas viejas sin TotalDeuda, lo re-arma
        self.df_deu["TotalDeuda"] = self.df_deu["Adeuda"] - self.df_deu["Pagado"]
        self.df_deu["Estado"] = self.df_deu["TotalDeuda"].apply(
            lambda x: "AL DÍA" if x == 0 else f"A FAVOR ${-x:.2f}" if x < 0 else f"ADEUDA ${x:.2f}"
        )

        # recalcula mensual
        self.recalcular_ganancias_mensuales()

    def reload(self):
        self.load_dataframes()
        self.refresh_table()
        self.refresh_reports()
        self.update_status("Datos recargados")

    def update_status(self, text):
        self.status_var.set(text)

    # ---------------- Inventario table ----------------
    def refresh_table(self):
        q = self.search_var.get().strip().lower()
        df = self.df_inv.copy()
        if q:
            df = df[
                df["Código"].astype(str).str.lower().str.contains(q) |
                df["Nombre"].astype(str).str.lower().str.contains(q)
            ]
        df = df.sort_values(by="Stock")

        for r in self.tree.get_children():
            self.tree.delete(r)
        self.tree_data.clear()

        for _, row in df.iterrows():
            pv = float(row.get("PrecioVenta", 0.0))
            st = int(row.get("Stock", 0))
            cat = row.get("Categoría", "")
            code = str(row["Código"])
            item_id = self.tree.insert("", "end", values=(code, row["Nombre"], f"{pv:.2f}", st, cat))
            self.tree_data[item_id] = str(row["Código"])

        self.update_status(f"{len(df)} producto(s) mostrados")

    # ---------------- Reportes tab ----------------
    def build_report_tab(self):
        top = ttk.Frame(self.tab_rep, padding=8)
        top.pack(fill="x")

        ttk.Label(top, text="Filtro tabla por producto:").pack(side="left")
        self.rep_filter_var = tk.StringVar(value="Este mes")

        ttk.Button(top, text="Hoy", command=lambda: self.set_report_filter("Hoy")).pack(side="left", padx=4)
        ttk.Button(top, text="Este mes", command=lambda: self.set_report_filter("Este mes")).pack(side="left", padx=4)
        ttk.Button(top, text="Todo", command=lambda: self.set_report_filter("Todo")).pack(side="left", padx=4)

        ttk.Button(top, text="Refrescar reportes", command=self.refresh_reports).pack(side="right", padx=4)

        # KPIs (3 líneas: hoy / semana / mes)
        kpi = ttk.Frame(self.tab_rep, padding=8)
        kpi.pack(fill="x")

        self.lbl_hoy = ttk.Label(kpi, text="HOY | Unidades: 0 | Ventas: $0.00 | Ganancia: $0.00",
                                 font=("Arial", 10, "bold"))
        self.lbl_hoy.pack(anchor="w", pady=2)

        self.lbl_sem = ttk.Label(kpi, text="SEMANA | Unidades: 0 | Ventas: $0.00 | Ganancia: $0.00",
                                 font=("Arial", 10, "bold"))
        self.lbl_sem.pack(anchor="w", pady=2)

        self.lbl_mes = ttk.Label(kpi, text="MES | Unidades: 0 | Ventas: $0.00 | Ganancia: $0.00",
                                 font=("Arial", 10, "bold"))
        self.lbl_mes.pack(anchor="w", pady=2)

        # Tree: ventas por producto
        mid = ttk.Frame(self.tab_rep, padding=8)
        mid.pack(fill="both", expand=True)

        cols = ("Código", "Nombre", "Cantidad", "Ventas", "Ganancia")
        self.rep_tree = ttk.Treeview(mid, columns=cols, show="headings", height=18)
        for c in cols:
            self.rep_tree.heading(c, text=c)
            self.rep_tree.column(c, anchor="center", width=160)
        self.rep_tree.pack(fill="both", expand=True)

        footer = ttk.Frame(self.tab_rep, padding=8)
        footer.pack(fill="x")
        ttk.Label(
            footer,
            text="Tip: Ganancia = (PrecioVenta - PrecioCompra) * Cantidad. Pagos no cuentan como ventas."
        ).pack(side="left")

    def set_report_filter(self, value):
        self.rep_filter_var.set(value)
        self.refresh_reports()

    def _ventas_filtradas_para_reportes(self):
        df = self.df_ven.copy()
        df = df[df["Tipo"].astype(str) != "Pago"].copy()

        # Fix definitivo: datetime real
        df["_dt"] = pd.to_datetime(df["Fecha"], errors="coerce")
        df = df[df["_dt"].notna()].copy()

        filtro = self.rep_filter_var.get()
        hoy = date.today()

        if filtro == "Hoy":
            df = df[df["_dt"].dt.date == hoy]
        elif filtro == "Este mes":
            df = df[(df["_dt"].dt.year == hoy.year) & (df["_dt"].dt.month == hoy.month)]
        elif filtro == "Todo":
            pass

        return df

    def refresh_reports(self):
        hoy = date.today()

        df = self.df_ven.copy()
        df = df[df["Tipo"].astype(str) != "Pago"].copy()

        # Fix definitivo: datetime real
        df["_dt"] = pd.to_datetime(df["Fecha"], errors="coerce")
        df = df[df["_dt"].notna()].copy()

        # HOY
        df_hoy = df[df["_dt"].dt.date == hoy].copy()

        # SEMANA (lunes a hoy)
        start_week = hoy.fromordinal(hoy.toordinal() - hoy.weekday())  # lunes
        df_sem = df[(df["_dt"].dt.date >= start_week) & (df["_dt"].dt.date <= hoy)].copy()

        # MES
        df_mes = df[(df["_dt"].dt.year == hoy.year) & (df["_dt"].dt.month == hoy.month)].copy()

        # Totales
        ven_hoy = float(df_hoy["Total"].sum()) if not df_hoy.empty else 0.0
        gan_hoy = float(df_hoy["Ganancia"].sum()) if not df_hoy.empty else 0.0
        uni_hoy = int(df_hoy["Cantidad"].sum()) if not df_hoy.empty else 0

        ven_sem = float(df_sem["Total"].sum()) if not df_sem.empty else 0.0
        gan_sem = float(df_sem["Ganancia"].sum()) if not df_sem.empty else 0.0
        uni_sem = int(df_sem["Cantidad"].sum()) if not df_sem.empty else 0

        ven_mes = float(df_mes["Total"].sum()) if not df_mes.empty else 0.0
        gan_mes = float(df_mes["Ganancia"].sum()) if not df_mes.empty else 0.0
        uni_mes = int(df_mes["Cantidad"].sum()) if not df_mes.empty else 0

        self.lbl_hoy.config(text=f"HOY | Unidades: {uni_hoy} | Ventas: ${ven_hoy:.2f} | Ganancia: ${gan_hoy:.2f}")
        self.lbl_sem.config(text=f"SEMANA (desde {start_week.strftime('%d/%m')}) | Unidades: {uni_sem} | Ventas: ${ven_sem:.2f} | Ganancia: ${gan_sem:.2f}")
        self.lbl_mes.config(text=f"MES | Unidades: {uni_mes} | Ventas: ${ven_mes:.2f} | Ganancia: ${gan_mes:.2f}")

        # Tabla por producto según filtro seleccionado
        df_f = self._ventas_filtradas_para_reportes()

        for r in self.rep_tree.get_children():
            self.rep_tree.delete(r)

        if df_f.empty:
            return

        grp = df_f.groupby(["Código", "Nombre"], dropna=False).agg(
            Cantidad=("Cantidad", "sum"),
            Ventas=("Total", "sum"),
            Ganancia=("Ganancia", "sum"),
        ).reset_index()

        grp = grp.sort_values(by="Ganancia", ascending=False)

        for _, r in grp.iterrows():
            self.rep_tree.insert("", "end", values=(
                str(r["Código"]),
                str(r["Nombre"]),
                int(r["Cantidad"]),
                f"{float(r['Ventas']):.2f}",
                f"{float(r['Ganancia']):.2f}",
            ))

    def recalcular_ganancias_mensuales(self):
        """
        Recalcula hoja Ganancias (mensual) a partir de Ventas.
        Ignora Tipo == Pago.
        """
        df = self.df_ven.copy()
        df = df[df["Tipo"].astype(str) != "Pago"].copy()

        df["_dt"] = pd.to_datetime(df["Fecha"], errors="coerce")
        df = df[df["_dt"].notna()].copy()

        if df.empty:
            self.df_gan = pd.DataFrame(columns=["Mes", "TotalVentasMes", "TotalGananciaMes", "UltimaActualizacion"])
            return

        df["Mes"] = df["_dt"].dt.strftime("%Y-%m")
        monthly = df.groupby("Mes", dropna=False).agg(
            TotalVentasMes=("Total", "sum"),
            TotalGananciaMes=("Ganancia", "sum"),
        ).reset_index()

        monthly["UltimaActualizacion"] = datetime.now().isoformat()
        self.df_gan = monthly

    # ---------------- UI: Add product ----------------
    def ui_add(self):
        win = tk.Toplevel(self.root)
        win.title("Agregar producto - Delicias de la Wera")
        win.geometry("380x340")
        pad = {"padx": 8, "pady": 6}

        fields = [
            ("Código", "Código"),
            ("Nombre", "Nombre"),
            ("PrecioCompra", "Precio compra"),
            ("PrecioVenta", "Precio venta"),
            ("Stock", "Stock"),
            ("Categoría", "Categoría"),
        ]
        entries = {}
        for i, (k, label) in enumerate(fields):
            ttk.Label(win, text=label).grid(row=i, column=0, sticky="w", **pad)
            e = ttk.Entry(win)
            e.grid(row=i, column=1, **pad)
            entries[k] = e

        def save():
            vals = {k: entries[k].get().strip() for k in entries}
            if not vals["Código"] or not vals["Nombre"]:
                messagebox.showwarning("Faltan datos", "Código y nombre son obligatorios")
                return
            try:
                vals["PrecioCompra"] = float(vals["PrecioCompra"]) if vals["PrecioCompra"] else 0.0
                vals["PrecioVenta"] = float(vals["PrecioVenta"]) if vals["PrecioVenta"] else 0.0
                vals["Stock"] = int(vals["Stock"]) if vals["Stock"] else 0
            except Exception:
                messagebox.showwarning("Error", "Precio o Stock con formato inválido")
                return

            if vals["Código"] in self.df_inv["Código"].astype(str).values:
                messagebox.showwarning("Duplicado", "Ya existe un producto con ese código")
                return

            self.df_inv = pd.concat([self.df_inv, pd.DataFrame([vals])], ignore_index=True)

            self.recalcular_ganancias_mensuales()
            guardar_todo(self.df_inv, self.df_ven, self.df_deu, self.df_tra, self.df_res, self.df_gan)
            self.reload()
            messagebox.showinfo("OK", "Producto agregado")
            win.destroy()

        ttk.Button(win, text="Guardar", command=save).grid(row=len(fields), column=0, columnspan=2, pady=12)

    # ---------------- Edit / Restock ----------------
    def get_selected_code(self):
        sel = self.tree.selection()
        if not sel:
            messagebox.showinfo("Seleccione", "Seleccione un producto de la lista")
            return None
        item_id = sel[0]
        return self.tree_data.get(item_id)

    def open_edit_selected(self):
        code = self.get_selected_code()
        if not code:
            return

        row = self.df_inv[self.df_inv["Código"].astype(str) == code]
        if row.empty:
            messagebox.showerror("Error", f"Producto no encontrado. Código: '{code}'")
            return

        data = row.iloc[0].to_dict()

        win = tk.Toplevel(self.root)
        win.title("Editar / Abastecer - Delicias de la Wera")
        win.geometry("380x360")
        pad = {"padx": 8, "pady": 6}

        fields = [
            ("Código", "Código"),
            ("Nombre", "Nombre"),
            ("PrecioCompra", "Precio compra"),
            ("PrecioVenta", "Precio venta"),
            ("Stock", "Stock"),
            ("Categoría", "Categoría"),
        ]
        entries = {}

        for i, (k, label) in enumerate(fields):
            ttk.Label(win, text=label).grid(row=i, column=0, sticky="w", **pad)
            e = ttk.Entry(win)
            e.grid(row=i, column=1, **pad)
            e.insert(0, str(data.get(k, "")))
            entries[k] = e

        entries["Código"].config(state="readonly")

        def save_edit():
            vals = {k: entries[k].get().strip() for k in entries if k != "Código"}
            try:
                vals["PrecioCompra"] = float(vals["PrecioCompra"]) if vals["PrecioCompra"] else 0.0
                vals["PrecioVenta"] = float(vals["PrecioVenta"]) if vals["PrecioVenta"] else 0.0
                vals["Stock"] = int(vals["Stock"]) if vals["Stock"] else 0
            except Exception:
                messagebox.showwarning("Error", "Precio o Stock con formato inválido")
                return

            idxs = self.df_inv[self.df_inv["Código"].astype(str) == code].index
            if len(idxs) == 0:
                messagebox.showerror("Error", "No se pudo encontrar el producto para editar")
                return

            idx = idxs[0]
            for k in ["Nombre", "PrecioCompra", "PrecioVenta", "Stock", "Categoría"]:
                self.df_inv.at[idx, k] = vals[k]

            self.recalcular_ganancias_mensuales()
            guardar_todo(self.df_inv, self.df_ven, self.df_deu, self.df_tra, self.df_res, self.df_gan)
            self.reload()
            messagebox.showinfo("OK", "Producto actualizado")
            win.destroy()

        def restock():
            try:
                add = simpledialog.askinteger("Abastecer", "Cantidad a agregar:", parent=win, minvalue=1)
                if not add:
                    return
                idxs = self.df_inv[self.df_inv["Código"].astype(str) == code].index
                if len(idxs) == 0:
                    messagebox.showerror("Error", "No se pudo encontrar el producto")
                    return
                idx = idxs[0]
                current_stock = int(self.df_inv.at[idx, "Stock"])
                self.df_inv.at[idx, "Stock"] = current_stock + int(add)

                self.recalcular_ganancias_mensuales()
                guardar_todo(self.df_inv, self.df_ven, self.df_deu, self.df_tra, self.df_res, self.df_gan)
                self.reload()
                messagebox.showinfo("OK", f"Stock actualizado: {current_stock} + {add} = {current_stock + add}")
            except Exception as e:
                messagebox.showerror("Error", str(e))

        ttk.Button(win, text="Guardar cambios", command=save_edit).grid(row=7, column=0, pady=10)
        ttk.Button(win, text="Abastecer (añadir)", command=restock).grid(row=7, column=1, pady=10)

    # ---------------- Delete product ----------------
    def ui_delete_product(self):
        code = self.get_selected_code()
        if not code:
            return

        row = self.df_inv[self.df_inv["Código"].astype(str) == code]
        if row.empty:
            messagebox.showerror("Error", "Producto no encontrado")
            return

        producto_info = row.iloc[0]
        nombre = producto_info["Nombre"]
        stock = producto_info["Stock"]

        confirmacion = messagebox.askyesno(
            "Confirmar eliminación",
            f"¿Está seguro de que desea eliminar el producto?\n\n"
            f"Código: {code}\n"
            f"Nombre: {nombre}\n"
            f"Stock actual: {stock}\n\n"
            f"Esta acción no se puede deshacer."
        )

        if confirmacion:
            try:
                self.df_inv = self.df_inv[self.df_inv["Código"].astype(str) != code]
                self.recalcular_ganancias_mensuales()
                guardar_todo(self.df_inv, self.df_ven, self.df_deu, self.df_tra, self.df_res, self.df_gan)
                self.reload()
                messagebox.showinfo("Producto eliminado", f"El producto '{nombre}' ha sido eliminado correctamente.")
            except Exception as e:
                messagebox.showerror("Error", f"No se pudo eliminar el producto: {str(e)}")

    # ---------------- Resumen pagos ----------------
    def actualizar_resumen_pagos(self, persona, monto, tipo_pago):
        persona = persona.strip()
        if not persona:
            return

        existe = self.df_res[self.df_res["Persona"] == persona]
        deuda_actual = 0.0
        if tipo_pago in ("Fiado", "Pago"):
            deuda_existe = self.df_deu[self.df_deu["Persona"] == persona]
            if not deuda_existe.empty:
                deuda_actual = float(deuda_existe.iloc[0]["TotalDeuda"])

        if existe.empty:
            nuevo = {
                "Persona": persona,
                "TotalEfectivo": 0.0,
                "TotalTransferencia": 0.0,
                "TotalFiado": 0.0,
                "TotalPagado": 0.0,
                "DeudaActual": deuda_actual,
                "UltimaActualizacion": datetime.now().isoformat()
            }
            if tipo_pago == "Pago":
                nuevo["TotalPagado"] = float(monto)
            else:
                col = f"Total{tipo_pago}"
                if col in nuevo:
                    nuevo[col] = float(monto)

            self.df_res = pd.concat([self.df_res, pd.DataFrame([nuevo])], ignore_index=True)
        else:
            idx = existe.index[0]
            if tipo_pago == "Pago":
                self.df_res.at[idx, "TotalPagado"] = float(self.df_res.at[idx, "TotalPagado"]) + float(monto)
            else:
                col = f"Total{tipo_pago}"
                if col in self.df_res.columns:
                    self.df_res.at[idx, col] = float(self.df_res.at[idx, col]) + float(monto)

            self.df_res.at[idx, "DeudaActual"] = deuda_actual
            self.df_res.at[idx, "UltimaActualizacion"] = datetime.now().isoformat()

    # ---------------- Sale UI ----------------
    def ui_sale(self, tipo):
        win = tk.Toplevel(self.root)
        win.title(f"Registrar venta - {tipo}")
        win.geometry("420x320")

        pad = {"padx": 8, "pady": 6}
        ttk.Label(win, text="Código:").grid(row=0, column=0, sticky="w", **pad)
        code_var = tk.StringVar()
        ttk.Entry(win, textvariable=code_var).grid(row=0, column=1, **pad)
        ttk.Button(win, text="Buscar", command=lambda: self.fill_from_code(code_var)).grid(row=0, column=2, padx=4)

        ttk.Label(win, text="Cantidad:").grid(row=1, column=0, sticky="w", **pad)
        qty_var = tk.IntVar(value=1)
        ttk.Entry(win, textvariable=qty_var).grid(row=1, column=1, **pad)

        ttk.Label(win, text="Persona:").grid(row=2, column=0, sticky="w", **pad)
        person_var = tk.StringVar()
        ttk.Entry(win, textvariable=person_var).grid(row=2, column=1, **pad)

        ttk.Label(win, text="Descripción (opcional):").grid(row=3, column=0, sticky="w", **pad)
        desc_var = tk.StringVar()
        ttk.Entry(win, textvariable=desc_var, width=36).grid(row=3, column=1, columnspan=2, **pad)

        account_var = tk.StringVar()
        if tipo == "Transferencia":
            ttk.Label(win, text="Cuenta (opcional):").grid(row=4, column=0, sticky="w", **pad)
            ttk.Entry(win, textvariable=account_var).grid(row=4, column=1, **pad)

        def register():
            code = code_var.get().strip()
            if not code:
                messagebox.showwarning("Error", "Ingrese código")
                return

            try:
                qty = int(qty_var.get())
                if qty <= 0:
                    raise ValueError
            except Exception:
                messagebox.showwarning("Error", "Cantidad inválida")
                return

            person = person_var.get().strip() or "Cliente"
            desc = desc_var.get().strip()

            row = self.df_inv[self.df_inv["Código"].astype(str) == code]
            if row.empty:
                messagebox.showerror("No existe", "Producto no encontrado")
                return

            idx = row.index[0]
            stock = int(self.df_inv.at[idx, "Stock"])
            if stock < qty:
                messagebox.showerror("Stock insuficiente", f"Stock actual: {stock}")
                return

            precio_venta = float(self.df_inv.at[idx, "PrecioVenta"])
            precio_compra = float(self.df_inv.at[idx, "PrecioCompra"])
            total = precio_venta * qty
            ganancia = (precio_venta - precio_compra) * qty

            # restar stock
            self.df_inv.at[idx, "Stock"] = stock - qty

            # registrar venta
            venta_row = {
                "Fecha": datetime.now().isoformat(),
                "Código": code,
                "Nombre": self.df_inv.at[idx, "Nombre"],
                "Cantidad": qty,
                "PrecioVenta": precio_venta,
                "PrecioCompra": precio_compra,
                "Total": total,
                "Ganancia": ganancia,
                "Persona": person,
                "Tipo": tipo,
                "Descripción": desc
            }
            self.df_ven = pd.concat([self.df_ven, pd.DataFrame([venta_row])], ignore_index=True)

            # transferencias
            if tipo == "Transferencia":
                tra = {
                    "Fecha": datetime.now().isoformat(),
                    "Código": code,
                    "Nombre": self.df_inv.at[idx, "Nombre"],
                    "Cantidad": qty,
                    "Precio": precio_venta,
                    "Total": total,
                    "Persona": person,
                    "Cuenta": account_var.get().strip(),
                    "Descripción": desc
                }
                self.df_tra = pd.concat([self.df_tra, pd.DataFrame([tra])], ignore_index=True)

            # deudas si fiado
            if tipo == "Fiado":
                existe = self.df_deu[self.df_deu["Persona"] == person]
                if existe.empty:
                    new_deu = {"Persona": person, "Adeuda": total, "Pagado": 0.0, "TotalDeuda": total, "Estado": f"ADEUDA ${total:.2f}"}
                    self.df_deu = pd.concat([self.df_deu, pd.DataFrame([new_deu])], ignore_index=True)
                else:
                    ix = existe.index[0]
                    current_adeuda = float(self.df_deu.at[ix, "Adeuda"])
                    current_pagado = float(self.df_deu.at[ix, "Pagado"])
                    new_adeuda = current_adeuda + total
                    new_total = new_adeuda - current_pagado
                    self.df_deu.at[ix, "Adeuda"] = new_adeuda
                    self.df_deu.at[ix, "TotalDeuda"] = new_total
                    self.df_deu.at[ix, "Estado"] = f"ADEUDA ${new_total:.2f}" if new_total > 0 else "AL DÍA"

            # resumen pagos (por tipo)
            self.actualizar_resumen_pagos(person, total, tipo)

            # recalcular ganancias mensuales y guardar todo
            self.recalcular_ganancias_mensuales()
            guardar_todo(self.df_inv, self.df_ven, self.df_deu, self.df_tra, self.df_res, self.df_gan)

            self.reload()

            msg = f"Venta registrada:\nTotal: ${total:.2f}\nGanancia: ${ganancia:.2f}\nTipo: {tipo}\nPersona: {person}"
            if tipo == "Fiado":
                deuda_existe = self.df_deu[self.df_deu["Persona"] == person]
                if not deuda_existe.empty:
                    deuda_actual = float(deuda_existe.iloc[0]["TotalDeuda"])
                    msg += f"\nDeuda actual: ${deuda_actual:.2f}"
            messagebox.showinfo("Venta registrada", msg)
            win.destroy()

        ttk.Button(win, text="Registrar venta", command=register).grid(row=6, column=0, columnspan=3, pady=12)

    def fill_from_code(self, code_var):
        code = code_var.get().strip()
        if not code:
            return
        row = self.df_inv[self.df_inv["Código"].astype(str) == code]
        if row.empty:
            messagebox.showwarning("No encontrado", "Código no existe")
            return
        r = row.iloc[0]
        messagebox.showinfo("Encontrado", f"{r['Nombre']} - Precio: {float(r['PrecioVenta']):.2f} - Stock: {int(r['Stock'])}")

    # ---------------- Register payment ----------------
    def ui_register_payment(self):
        win = tk.Toplevel(self.root)
        win.title("Registrar pago - Delicias de la Wera")
        win.geometry("380x240")

        pad = {"padx": 8, "pady": 6}
        ttk.Label(win, text="Persona:").grid(row=0, column=0, sticky="w", **pad)
        person_var = tk.StringVar()
        ttk.Entry(win, textvariable=person_var).grid(row=0, column=1, **pad)

        ttk.Label(win, text="Cantidad:").grid(row=1, column=0, sticky="w", **pad)
        amount_var = tk.DoubleVar(value=0.0)
        ttk.Entry(win, textvariable=amount_var).grid(row=1, column=1, **pad)

        ttk.Label(win, text="Descripción (ej: transferencia, efectivo):").grid(row=2, column=0, sticky="w", **pad)
        desc_var = tk.StringVar()
        ttk.Entry(win, textvariable=desc_var, width=38).grid(row=2, column=1, columnspan=2, **pad)

        def save_payment():
            person = person_var.get().strip()
            amt = float(amount_var.get())
            desc = desc_var.get().strip()

            if not person or amt <= 0:
                messagebox.showwarning("Error", "Persona y cantidad válida son requeridas")
                return

            existe = self.df_deu[self.df_deu["Persona"] == person]
            if existe.empty:
                new_deu = {
                    "Persona": person,
                    "Adeuda": 0.0,
                    "Pagado": amt,
                    "TotalDeuda": -amt,
                    "Estado": f"A FAVOR ${amt:.2f}"
                }
                self.df_deu = pd.concat([self.df_deu, pd.DataFrame([new_deu])], ignore_index=True)
                new_total = -amt
            else:
                ix = existe.index[0]
                current_pagado = float(self.df_deu.at[ix, "Pagado"])
                current_adeuda = float(self.df_deu.at[ix, "Adeuda"])

                new_pagado = current_pagado + amt
                new_total = current_adeuda - new_pagado

                self.df_deu.at[ix, "Pagado"] = new_pagado
                self.df_deu.at[ix, "TotalDeuda"] = new_total
                self.df_deu.at[ix, "Estado"] = "AL DÍA" if new_total == 0 else (f"A FAVOR ${-new_total:.2f}" if new_total < 0 else f"ADEUDA ${new_total:.2f}")

            # resumen pagos
            self.actualizar_resumen_pagos(person, amt, "Pago")

            # registrar pago en ventas como movimiento (no cuenta para reportes)
            pago_record = {
                "Fecha": datetime.now().isoformat(),
                "Código": "",
                "Nombre": "Pago de deuda",
                "Cantidad": 1,
                "PrecioVenta": amt,
                "PrecioCompra": 0.0,
                "Total": amt,
                "Ganancia": 0.0,
                "Persona": person,
                "Tipo": "Pago",
                "Descripción": desc
            }
            self.df_ven = pd.concat([self.df_ven, pd.DataFrame([pago_record])], ignore_index=True)

            self.recalcular_ganancias_mensuales()
            guardar_todo(self.df_inv, self.df_ven, self.df_deu, self.df_tra, self.df_res, self.df_gan)
            self.reload()

            if new_total > 0:
                msg = f"Pago de ${amt:.2f} registrado\nDeuda restante: ${new_total:.2f}"
            elif new_total < 0:
                msg = f"Pago de ${amt:.2f} registrado\nSaldo a favor: ${-new_total:.2f}"
            else:
                msg = f"Pago de ${amt:.2f} registrado\n¡Deuda saldada completamente!"
            messagebox.showinfo("Pago registrado", msg)
            win.destroy()

        ttk.Button(win, text="Registrar pago", command=save_payment).grid(row=4, column=0, columnspan=3, pady=8)

    # ---------------- View debtors ----------------
    def ui_view_debtors(self):
        win = tk.Toplevel(self.root)
        win.title("Deudores - Delicias de la Wera")
        win.geometry("720x420")

        cols = ("Persona","Adeuda","Pagado","TotalDeuda", "Estado")
        tree = ttk.Treeview(win, columns=cols, show="headings", height=18)
        for c in cols:
            tree.heading(c, text=c)
            tree.column(c, anchor="center", width=130)
        tree.pack(fill="both", expand=True, padx=8, pady=8)

        df = self.df_deu.copy()
        if df.empty:
            return

        df["Adeuda"] = pd.to_numeric(df["Adeuda"], errors="coerce").fillna(0.0)
        df["Pagado"] = pd.to_numeric(df["Pagado"], errors="coerce").fillna(0.0)
        df["TotalDeuda"] = df["Adeuda"] - df["Pagado"]
        df["Estado"] = df["TotalDeuda"].apply(lambda x: "AL DÍA" if x == 0 else f"A FAVOR ${-x:.2f}" if x < 0 else f"ADEUDA ${x:.2f}")

        total_general_deuda = float(df[df["TotalDeuda"] > 0]["TotalDeuda"].sum()) if not df.empty else 0.0
        total_a_favor = float((-df[df["TotalDeuda"] < 0]["TotalDeuda"]).sum()) if not df.empty else 0.0

        for _, r in df.iterrows():
            tree.insert("", "end", values=(
                str(r["Persona"]),
                f"{float(r['Adeuda']):.2f}",
                f"{float(r['Pagado']):.2f}",
                f"{float(r['TotalDeuda']):.2f}",
                str(r["Estado"])
            ))

        footer = ttk.Frame(win)
        footer.pack(fill="x", padx=8, pady=4)

        if total_general_deuda > 0:
            ttk.Label(footer, text=f"TOTAL ADEUDADO: ${total_general_deuda:.2f}",
                      font=("Arial", 10, "bold"), foreground="red").pack(side="left", padx=10)
        if total_a_favor > 0:
            ttk.Label(footer, text=f"TOTAL A FAVOR: ${total_a_favor:.2f}",
                      font=("Arial", 10, "bold"), foreground="green").pack(side="left", padx=10)

    # ---------------- View resumen pagos ----------------
    def ui_view_resumen_pagos(self):
        win = tk.Toplevel(self.root)
        win.title("Resumen de Pagos - Delicias de la Wera")
        win.geometry("860x470")

        cols = ("Persona", "TotalEfectivo", "TotalTransferencia", "TotalFiado", "TotalPagado", "DeudaActual", "UltimaActualizacion")
        tree = ttk.Treeview(win, columns=cols, show="headings", height=18)

        headings = {
            "Persona": "Persona",
            "TotalEfectivo": "Total Efectivo",
            "TotalTransferencia": "Total Transferencia",
            "TotalFiado": "Total Fiado",
            "TotalPagado": "Total Pagado",
            "DeudaActual": "Deuda Actual",
            "UltimaActualizacion": "Última Actualización"
        }
        for c in cols:
            tree.heading(c, text=headings[c])
            tree.column(c, anchor="center", width=120)
        tree.pack(fill="both", expand=True, padx=8, pady=8)

        df = self.df_res.copy().fillna(0)
        if df.empty:
            return

        total_efectivo = 0.0
        total_transferencia = 0.0
        total_fiado = 0.0
        total_pagado = 0.0
        total_deuda_actual = 0.0

        for _, r in df.iterrows():
            efectivo = float(r["TotalEfectivo"])
            transferencia = float(r["TotalTransferencia"])
            fiado = float(r["TotalFiado"])
            pagado = float(r["TotalPagado"])
            deuda_actual = float(r["DeudaActual"])

            total_efectivo += efectivo
            total_transferencia += transferencia
            total_fiado += fiado
            total_pagado += pagado
            total_deuda_actual += deuda_actual

            fecha = r["UltimaActualizacion"]
            if pd.isna(fecha) or str(fecha).strip() == "" or fecha == 0:
                fecha_str = "Nunca"
            else:
                try:
                    fecha_dt = datetime.fromisoformat(str(fecha))
                    fecha_str = fecha_dt.strftime("%d/%m/%Y %H:%M")
                except Exception:
                    fecha_str = str(fecha)

            tree.insert("", "end", values=(
                str(r["Persona"]),
                f"{efectivo:.2f}",
                f"{transferencia:.2f}",
                f"{fiado:.2f}",
                f"{pagado:.2f}",
                f"{deuda_actual:.2f}",
                fecha_str
            ))

        footer = ttk.Frame(win)
        footer.pack(fill="x", padx=8, pady=4)

        ttk.Label(footer, text=f"EFECTIVO: ${total_efectivo:.2f}", font=("Arial", 8, "bold"), foreground="green").pack(side="left", padx=4)
        ttk.Label(footer, text=f"TRANSFERENCIA: ${total_transferencia:.2f}", font=("Arial", 8, "bold"), foreground="blue").pack(side="left", padx=4)
        ttk.Label(footer, text=f"FIADO: ${total_fiado:.2f}", font=("Arial", 8, "bold"), foreground="orange").pack(side="left", padx=4)
        ttk.Label(footer, text=f"PAGADO: ${total_pagado:.2f}", font=("Arial", 8, "bold"), foreground="purple").pack(side="left", padx=4)
        ttk.Label(footer, text=f"DEUDA ACTUAL: ${total_deuda_actual:.2f}", font=("Arial", 8, "bold"), foreground="red").pack(side="left", padx=4)

    # ---------------- Export / Backup ----------------
    def exportar(self):
        folder = filedialog.askdirectory(title="Selecciona carpeta para exportar el archivo .xlsx")
        if not folder:
            return
        try:
            self.recalcular_ganancias_mensuales()
            guardar_todo(self.df_inv, self.df_ven, self.df_deu, self.df_tra, self.df_res, self.df_gan)
            shutil.copy(DATA_FILE, os.path.join(folder, DATA_FILE))
            messagebox.showinfo("Exportado", f"Archivo exportado a:\n{os.path.join(folder, DATA_FILE)}")
        except Exception as e:
            messagebox.showerror("Error", str(e))

    def ui_backup(self):
        dest = hacer_backup()
        if str(dest).startswith("Error"):
            messagebox.showerror("Error backup", dest)
        else:
            messagebox.showinfo("Backup", f"Copia guardada en:\n{dest}")


def main():
    root = tk.Tk()
    app = DeliciasApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
