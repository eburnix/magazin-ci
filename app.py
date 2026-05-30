import csv
import html
import json
import os
import platform
import sqlite3
import sys
from datetime import datetime, date
from pathlib import Path
import tkinter as tk
from tkinter import ttk, messagebox, filedialog, simpledialog

try:
    import win32print
except ImportError:
    win32print = None

APP_NAME = "Magazin-ci"
APP_VERSION = "1.0.0"
INSTALL_PIN = "05535350"
ACCESS_PIN = "0714"
DEVELOPER_CREDIT = "Developpe par Datadev-ci Tel : +225 0714351471 Abidjan - datadev.wps@gmail.com"

if getattr(sys, "frozen", False):
    BASE_DIR = Path(sys.executable).resolve().parent
    RESOURCE_DIR = Path(getattr(sys, "_MEIPASS", BASE_DIR))
else:
    BASE_DIR = Path(__file__).resolve().parent
    RESOURCE_DIR = BASE_DIR

DATA_DIR = BASE_DIR / "data"
DB_PATH = DATA_DIR / "magazin_ci.db"
RECEIPTS_DIR = DATA_DIR / "recus"
REPORTS_DIR = DATA_DIR / "rapports"
FNE_DIR = DATA_DIR / "fne"
INVOICES_DIR = DATA_DIR / "factures"


class Database:
    def __init__(self, path: Path):
        DATA_DIR.mkdir(exist_ok=True)
        self.conn = sqlite3.connect(path)
        self.conn.row_factory = sqlite3.Row
        self.conn.execute("PRAGMA foreign_keys = ON")
        self.setup()
        self.seed()

    def setup(self):
        cur = self.conn.cursor()
        cur.executescript(
            """
            PRAGMA foreign_keys = ON;

            CREATE TABLE IF NOT EXISTS settings (
                key TEXT PRIMARY KEY,
                value TEXT NOT NULL
            );

            CREATE TABLE IF NOT EXISTS products (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT NOT NULL,
                category TEXT NOT NULL DEFAULT 'General',
                barcode TEXT,
                unit TEXT NOT NULL DEFAULT 'piece',
                purchase_price REAL NOT NULL DEFAULT 0,
                sale_price REAL NOT NULL DEFAULT 0,
                stock REAL NOT NULL DEFAULT 0,
                alert_stock REAL NOT NULL DEFAULT 0,
                active INTEGER NOT NULL DEFAULT 1,
                created_at TEXT NOT NULL
            );

            CREATE TABLE IF NOT EXISTS customers (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT NOT NULL,
                phone TEXT,
                type TEXT NOT NULL DEFAULT 'Particulier',
                balance REAL NOT NULL DEFAULT 0,
                created_at TEXT NOT NULL
            );

            CREATE TABLE IF NOT EXISTS suppliers (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT NOT NULL,
                phone TEXT,
                address TEXT,
                created_at TEXT NOT NULL
            );

            CREATE TABLE IF NOT EXISTS sales (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                sale_no TEXT NOT NULL UNIQUE,
                customer_id INTEGER,
                customer_name TEXT NOT NULL,
                total REAL NOT NULL,
                paid REAL NOT NULL,
                payment_method TEXT NOT NULL,
                status TEXT NOT NULL,
                invoice_no TEXT,
                fne_status TEXT NOT NULL DEFAULT 'A exporter',
                fne_reference TEXT,
                fne_exported_at TEXT,
                note TEXT,
                created_at TEXT NOT NULL,
                FOREIGN KEY(customer_id) REFERENCES customers(id)
            );

            CREATE TABLE IF NOT EXISTS sale_items (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                sale_id INTEGER NOT NULL,
                product_id INTEGER NOT NULL,
                product_name TEXT NOT NULL,
                quantity REAL NOT NULL,
                unit_price REAL NOT NULL,
                total REAL NOT NULL,
                FOREIGN KEY(sale_id) REFERENCES sales(id) ON DELETE CASCADE,
                FOREIGN KEY(product_id) REFERENCES products(id)
            );

            CREATE TABLE IF NOT EXISTS stock_movements (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                product_id INTEGER NOT NULL,
                product_name TEXT NOT NULL,
                movement_type TEXT NOT NULL,
                quantity REAL NOT NULL,
                note TEXT,
                created_at TEXT NOT NULL,
                FOREIGN KEY(product_id) REFERENCES products(id)
            );

            CREATE TABLE IF NOT EXISTS cash_entries (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                entry_type TEXT NOT NULL,
                label TEXT NOT NULL,
                amount REAL NOT NULL,
                payment_method TEXT NOT NULL,
                created_at TEXT NOT NULL
            );
            """
        )
        self.ensure_column("customers", "email", "TEXT")
        self.ensure_column("customers", "ncc", "TEXT")
        self.ensure_column("customers", "address", "TEXT")
        self.ensure_column("suppliers", "email", "TEXT")
        self.ensure_column("suppliers", "active", "INTEGER NOT NULL DEFAULT 1")
        self.ensure_column("products", "supplier_id", "INTEGER")
        self.ensure_column("products", "description", "TEXT")
        self.ensure_column("sales", "invoice_no", "TEXT")
        self.ensure_column("sales", "fne_status", "TEXT NOT NULL DEFAULT 'A exporter'")
        self.ensure_column("sales", "fne_reference", "TEXT")
        self.ensure_column("sales", "fne_exported_at", "TEXT")
        self.ensure_column("sales", "note", "TEXT")
        indexes = [
            "CREATE INDEX IF NOT EXISTS idx_products_supplier ON products(supplier_id)",
            "CREATE INDEX IF NOT EXISTS idx_products_barcode ON products(barcode)",
            "CREATE INDEX IF NOT EXISTS idx_sales_created_at ON sales(created_at)",
            "CREATE INDEX IF NOT EXISTS idx_sales_customer ON sales(customer_id)",
            "CREATE INDEX IF NOT EXISTS idx_sales_fne_status ON sales(fne_status)",
            "CREATE INDEX IF NOT EXISTS idx_sale_items_sale ON sale_items(sale_id)",
            "CREATE INDEX IF NOT EXISTS idx_sale_items_product ON sale_items(product_id)",
            "CREATE INDEX IF NOT EXISTS idx_stock_movements_product ON stock_movements(product_id)",
            "CREATE INDEX IF NOT EXISTS idx_cash_entries_date ON cash_entries(created_at)",
        ]
        for query in indexes:
            self.conn.execute(query)
        self.conn.commit()

    def ensure_column(self, table, column, definition):
        columns = {row["name"] for row in self.conn.execute(f"PRAGMA table_info({table})").fetchall()}
        if column not in columns:
            self.conn.execute(f"ALTER TABLE {table} ADD COLUMN {column} {definition}")

    def seed(self):
        defaults = {
            "shop_name": "Magazin-ci",
            "shop_phone": "+225 0714351471",
            "shop_email": "datadev.wps@gmail.com",
            "shop_address": "Abidjan, Cote d'Ivoire",
            "shop_ncc": "",
            "shop_rccm": "",
            "shop_tax_regime": "Regime reel simplifie",
            "invoice_footer": "Merci pour votre confiance. Exigez votre facture normalisee FNE pour justificatif fiscal.",
            "fne_note": "Export de preparation FNE. La certification officielle doit etre effectuee via la plateforme DGI/FNE ou un TERNE.",
            "currency": "FCFA",
            "theme": "Light",
        }
        for key, value in defaults.items():
            self.conn.execute(
                "INSERT OR IGNORE INTO settings(key, value) VALUES (?, ?)",
                (key, value),
            )
        count = self.conn.execute("SELECT COUNT(*) FROM products").fetchone()[0]
        if count == 0:
            now = datetime.now().isoformat(timespec="seconds")
            products = [
                ("Riz 5kg", "Alimentation", "", "sac", 2800, 3500, 20, 5),
                ("Huile 1L", "Alimentation", "", "bouteille", 950, 1200, 30, 8),
                ("Savon", "Menage", "", "piece", 250, 400, 50, 10),
                ("Credit mobile", "Service", "", "operation", 0, 1000, 0, 0),
            ]
            self.conn.executemany(
                """
                INSERT INTO products(name, category, barcode, unit, purchase_price, sale_price, stock, alert_stock, created_at)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                [(*p, now) for p in products],
            )
        self.conn.commit()

    def all(self, query, params=()):
        return self.conn.execute(query, params).fetchall()

    def one(self, query, params=()):
        return self.conn.execute(query, params).fetchone()

    def execute(self, query, params=()):
        cur = self.conn.execute(query, params)
        self.conn.commit()
        return cur


class SplashScreen(tk.Toplevel):
    def __init__(self, parent, logo_image=None):
        super().__init__(parent)
        self.configure(bg="#0b1220")
        self.overrideredirect(True)
        self.resizable(False, False)
        self.status = tk.StringVar(value="Chargement...")
        width, height = 560, 340
        x = max(0, (self.winfo_screenwidth() - width) // 2)
        y = max(0, (self.winfo_screenheight() - height) // 2)
        self.geometry(f"{width}x{height}+{x}+{y}")

        body = tk.Frame(self, bg="#0b1220", padx=34, pady=30)
        body.pack(fill="both", expand=True)
        if logo_image:
            tk.Label(body, image=logo_image, bg="#0b1220", borderwidth=0).pack(pady=(0, 14))
        else:
            tk.Label(body, text="M", bg="#0b1220", fg="#34d399", font=("Segoe UI", 42, "bold")).pack(pady=(0, 14))
        tk.Label(body, text=APP_NAME, bg="#0b1220", fg="#ffffff", font=("Segoe UI", 26, "bold")).pack()
        tk.Label(body, text="Gestion commerciale et caisse tactile", bg="#0b1220", fg="#cbd5e1", font=("Segoe UI", 11)).pack(pady=(4, 22))
        bar = tk.Frame(body, bg="#1f2937", height=10)
        bar.pack(fill="x", pady=(0, 12))
        bar.pack_propagate(False)
        self.fill = tk.Frame(bar, bg="#34d399", height=10)
        self.fill.place(x=0, y=0, width=1, relheight=1)
        tk.Label(body, textvariable=self.status, bg="#0b1220", fg="#e5e7eb", font=("Segoe UI", 10)).pack(anchor="w")
        tk.Label(body, text=DEVELOPER_CREDIT, bg="#0b1220", fg="#94a3b8", font=("Segoe UI", 9), wraplength=490).pack(side="bottom", pady=(18, 0))
        self.update_idletasks()

    def set_progress(self, percent, text):
        self.status.set(text)
        width = max(1, self.fill.master.winfo_width())
        self.fill.place_configure(width=int(width * max(0, min(100, percent)) / 100))
        self.update_idletasks()


class MagazinApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.withdraw()
        self.logo_image = self.load_logo_image(max_width=180, max_height=120)
        self.splash_logo_image = self.load_logo_image(max_width=180, max_height=120)
        self.apply_app_icon()
        self.splash = SplashScreen(self, self.splash_logo_image)
        self.splash.set_progress(16, "Verification de l'acces...")
        if not self.prompt_access_pin():
            self.splash.destroy()
            self.destroy()
            return
        self.splash.set_progress(42, "Preparation de la base de donnees...")
        self.db = Database(DB_PATH)
        self.splash.set_progress(70, "Chargement de l'interface...")
        self.theme = self.setting("theme").strip().lower() or "light"
        self.title(f"{APP_NAME} - Gestion commerciale")
        self.geometry("1180x720")
        self.minsize(1060, 640)
        self.rowconfigure(0, weight=1)
        self.columnconfigure(0, weight=1)
        self.cart = []
        self.current_view = None
        self.build_style()
        self.build_shell()
        self.show_dashboard()
        self.splash.set_progress(100, "Pret.")
        self.after(250, self.splash.destroy)
        self.deiconify()

    def load_logo_image(self, max_width=180, max_height=120):
        logo_path = BASE_DIR / "LOGO_.png"
        if not logo_path.exists():
            logo_path = RESOURCE_DIR / "LOGO_.png"
        if not logo_path.exists():
            return None
        try:
            image = tk.PhotoImage(file=str(logo_path))
            if image.width() > max_width or image.height() > max_height:
                factor = max((image.width() + max_width - 1) // max_width, (image.height() + max_height - 1) // max_height)
                image = image.subsample(factor, factor)
            return image
        except Exception:
            return None

    def apply_app_icon(self):
        icon_path = BASE_DIR / "LOGO_.ico"
        if not icon_path.exists():
            icon_path = RESOURCE_DIR / "LOGO_.ico"
        if icon_path.exists():
            try:
                self.iconbitmap(str(icon_path))
            except Exception:
                pass
        if self.logo_image:
            try:
                self.iconphoto(True, self.logo_image)
            except Exception:
                pass

    def prompt_access_pin(self):
        result = {"ok": False, "attempts": 0}
        window = tk.Toplevel(self)
        window.title(APP_NAME)
        window.configure(bg="#0b1220")
        window.resizable(False, False)
        window.grab_set()
        window.geometry("420x300")
        window.update_idletasks()
        x = max(0, (window.winfo_screenwidth() - 420) // 2)
        y = max(0, (window.winfo_screenheight() - 300) // 2)
        window.geometry(f"420x300+{x}+{y}")
        if self.logo_image:
            try:
                window.iconphoto(True, self.logo_image)
            except Exception:
                pass

        tk.Label(window, text=APP_NAME, bg="#0b1220", fg="#ffffff", font=("Segoe UI", 22, "bold")).pack(pady=(28, 4))
        tk.Label(window, text="Acces securise", bg="#0b1220", fg="#cbd5e1", font=("Segoe UI", 11)).pack()
        pin_var = tk.StringVar()
        entry = ttk.Entry(window, textvariable=pin_var, show="*", width=24, justify="center")
        entry.pack(pady=(26, 10), ipady=6)
        status = tk.Label(window, text="", bg="#0b1220", fg="#fca5a5", font=("Segoe UI", 9))
        status.pack()

        def validate(_event=None):
            if pin_var.get() == ACCESS_PIN:
                result["ok"] = True
                window.destroy()
                return
            result["attempts"] += 1
            pin_var.set("")
            remaining = max(0, 3 - result["attempts"])
            status.config(text=f"PIN incorrect. Tentatives restantes: {remaining}")
            if result["attempts"] >= 3:
                window.destroy()

        buttons = tk.Frame(window, bg="#0b1220")
        buttons.pack(fill="x", padx=48, pady=(12, 0))
        ttk.Button(buttons, text="Entrer", style="Accent.TButton", command=validate).pack(side="left", fill="x", expand=True, padx=(0, 6))
        ttk.Button(buttons, text="Fermer", command=window.destroy).pack(side="left", fill="x", expand=True, padx=(6, 0))
        tk.Label(window, text=DEVELOPER_CREDIT, bg="#0b1220", fg="#94a3b8", font=("Segoe UI", 8), wraplength=360).pack(side="bottom", pady=14)
        entry.bind("<Return>", validate)
        entry.focus_set()
        self.wait_window(window)
        return result["ok"]

    def build_style(self):
        style = ttk.Style(self)
        style.theme_use("clam")
        dark = self.theme == "dark"
        bg = "#0f172a" if dark else "#eef2f7"
        surface = "#111827" if dark else "#ffffff"
        surface2 = "#1e293b" if dark else "#f8fafc"
        fg = "#f8fafc" if dark else "#111827"
        muted = "#94a3b8" if dark else "#6b7280"
        card_border = "#334155" if dark else "#e2e8f0"
        button_bg = "#1f2937" if dark else "#0f172a"
        button_active = "#0f766e" if dark else "#115e59"

        self.configure(bg=bg)
        style.configure("TFrame", background=bg)
        style.configure("Card.TFrame", background=surface, relief="flat", borderwidth=1)
        style.configure("Sidebar.TFrame", background="#0f172a" if dark else "#111827")
        style.configure("Sidebar.TButton", background="#0f172a" if dark else "#111827", foreground="#f8fafc", padding=12, anchor="w", font=("Segoe UI", 10, "bold"), relief="flat")
        style.map("Sidebar.TButton", background=[("active", "#111827" if dark else "#1f2937")], foreground=[("active", "#ffffff")])
        style.configure("Title.TLabel", background=bg, foreground=fg, font=("Segoe UI", 24, "bold"))
        style.configure("Sub.TLabel", background=bg, foreground=muted, font=("Segoe UI", 10))
        style.configure("CardTitle.TLabel", background=surface, foreground=fg, font=("Segoe UI", 12, "bold"))
        style.configure("Metric.TLabel", background=surface, foreground="#22c55e" if dark else "#0f766e", font=("Segoe UI", 20, "bold"))
        style.configure("TButton", background=button_bg, foreground="#f8fafc", font=("Segoe UI", 10), padding=8, relief="flat")
        style.map("TButton", background=[("active", button_active)])
        style.configure("Accent.TButton", background="#0f766e", foreground="#ffffff", relief="flat")
        style.map("Accent.TButton", background=[("active", "#115e59")])
        style.configure("Touch.TButton", font=("Segoe UI", 12, "bold"), padding=16, relief="flat")
        style.configure("ProductTile.TButton", background=surface2, foreground=fg, anchor="center", justify="center", padding=16, font=("Segoe UI", 10, "bold"), relief="solid", borderwidth=1, wraplength=180)
        style.map("ProductTile.TButton", background=[("active", "#334155" if dark else "#e2e8f0"), ("pressed", "#475569" if dark else "#cbd5e1")])
        style.configure("Danger.TButton", background="#dc2626", foreground="#ffffff", relief="flat")
        style.map("Danger.TButton", background=[("active", "#991b1b")])
        style.configure("TEntry", padding=6)
        style.configure("TLabel", background=bg, foreground=fg)
        style.configure("Treeview", rowheight=30, font=("Segoe UI", 9), fieldbackground=surface, background=surface, foreground=fg)
        style.configure("Treeview.Heading", font=("Segoe UI", 9, "bold"), background=card_border, foreground="#f8fafc" if dark else "#111827")

    def build_shell(self):
        self.sidebar = ttk.Frame(self, style="Sidebar.TFrame", width=240)
        self.sidebar.pack(side="left", fill="y")
        self.sidebar.pack_propagate(False)

        if self.logo_image:
            logo_label = tk.Label(self.sidebar, image=self.logo_image, bg="#111827")
            logo_label.pack(fill="x", pady=(16, 8), padx=12)

        brand = tk.Label(
            self.sidebar,
            text="Magazin-ci",
            bg="#111827",
            fg="#ffffff",
            font=("Segoe UI", 20, "bold"),
            pady=12,
        )
        brand.pack(fill="x")

        buttons = [
            ("Dashboard", self.show_dashboard),
            ("POS tactile", self.show_sales),
            ("Produits", self.show_products),
            ("Clients", self.show_customers),
            ("Fournisseurs", self.show_suppliers),
            ("Stock", self.show_stock),
            ("Inventaire", self.show_inventory),
            ("Caisse", self.show_cash),
            ("Statistiques", self.show_statistics),
            ("Rapports", self.show_reports),
            ("Parametres", self.show_settings),
        ]
        for text, command in buttons:
            ttk.Button(self.sidebar, text=text, command=command, style="Sidebar.TButton").pack(fill="x", padx=10, pady=3)

        footer = tk.Label(
            self.sidebar,
            text=DEVELOPER_CREDIT,
            bg="#111827",
            fg="#d1d5db",
            font=("Segoe UI", 8),
            justify="left",
            wraplength=200,
        )
        footer.pack(side="bottom", fill="x", padx=10, pady=12)

        self.main = ttk.Frame(self)
        self.main.pack(side="left", fill="both", expand=True)

    def clear(self, title, subtitle=""):
        for child in self.main.winfo_children():
            child.destroy()
        header = ttk.Frame(self.main)
        header.pack(fill="x", padx=24, pady=(22, 12))
        ttk.Label(header, text=title, style="Title.TLabel").pack(anchor="w")
        ttk.Label(header, text=subtitle, style="Sub.TLabel").pack(anchor="w", pady=(2, 0))
        self.current_view = title

    def card(self, parent, padx=14, pady=12):
        frame = ttk.Frame(parent, style="Card.TFrame", padding=(padx, pady))
        return frame

    def setting(self, key):
        row = self.db.one("SELECT value FROM settings WHERE key = ?", (key,))
        return row["value"] if row else ""

    def money(self, amount):
        return f"{amount:,.0f} {self.setting('currency')}".replace(",", " ")

    def export_excel_xml(self, path, sheet_name, headers, rows):
        def cell(value):
            if isinstance(value, (int, float)):
                return f'<Cell><Data ss:Type="Number">{value}</Data></Cell>'
            return f'<Cell><Data ss:Type="String">{html.escape(str(value or ""))}</Data></Cell>'

        content = [
            '<?xml version="1.0"?>',
            '<?mso-application progid="Excel.Sheet"?>',
            '<Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet"',
            ' xmlns:o="urn:schemas-microsoft-com:office:office"',
            ' xmlns:x="urn:schemas-microsoft-com:office:excel"',
            ' xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet">',
            f'<Worksheet ss:Name="{html.escape(sheet_name[:31])}">',
            '<Table>',
            "<Row>" + "".join(cell(h) for h in headers) + "</Row>",
        ]
        for row in rows:
            content.append("<Row>" + "".join(cell(v) for v in row) + "</Row>")
        content.extend(["</Table>", "</Worksheet>", "</Workbook>"])
        Path(path).write_text("\n".join(content), encoding="utf-8")

    def export_csv(self, path, headers, rows):
        with open(path, mode="w", newline="", encoding="utf-8") as csvfile:
            writer = csv.writer(csvfile)
            writer.writerow(headers)
            writer.writerows(rows)

    def print_text_file(self, path):
        if platform.system() == "Windows":
            try:
                os.startfile(path, "print")
                messagebox.showinfo(APP_NAME, "Le document a ete envoye a l'imprimante.")
            except OSError as exc:
                messagebox.showerror(APP_NAME, f"Impossible d'imprimer le document : {exc}")
        else:
            messagebox.showwarning(APP_NAME, "Impression uniquement disponible sur Windows.")

    def open_cash_drawer(self):
        if win32print is None:
            messagebox.showwarning(APP_NAME, "Module win32print non installe. Ouvrir le tiroir n'est pas disponible.")
            return
        try:
            printer_name = win32print.GetDefaultPrinter()
            handle = win32print.OpenPrinter(printer_name)
            raw_data = b"\x1b\x70\x00\x19\xc8"
            job = win32print.StartDocPrinter(handle, 1, ("Tiroir", None, "RAW"))
            win32print.StartPagePrinter(handle)
            win32print.WritePrinter(handle, raw_data)
            win32print.EndPagePrinter(handle)
            win32print.EndDocPrinter(handle)
            win32print.ClosePrinter(handle)
            messagebox.showinfo(APP_NAME, "Commande d'ouverture du tiroir envoyee a l'imprimante.")
        except Exception as exc:
            messagebox.showerror(APP_NAME, f"Impossible d'ouvrir le tiroir : {exc}")

    def show_dashboard(self):
        self.clear("Dashboard", "Vue rapide des ventes, du stock et de la caisse du jour.")
        today = date.today().isoformat()
        total_sales = self.db.one("SELECT COALESCE(SUM(total), 0) v FROM sales WHERE date(created_at) = ?", (today,))["v"]
        paid = self.db.one("SELECT COALESCE(SUM(paid), 0) v FROM sales WHERE date(created_at) = ?", (today,))["v"]
        sales_count = self.db.one("SELECT COUNT(*) v FROM sales WHERE date(created_at) = ?", (today,))["v"]
        low_stock = self.db.one("SELECT COUNT(*) v FROM products WHERE active = 1 AND alert_stock > 0 AND stock <= alert_stock")["v"]

        grid = ttk.Frame(self.main)
        grid.pack(fill="x", padx=24, pady=8)
        metrics = [
            ("Ventes du jour", self.money(total_sales)),
            ("Encaisse", self.money(paid)),
            ("Tickets", str(sales_count)),
            ("Alertes stock", str(low_stock)),
        ]
        for i, (label, value) in enumerate(metrics):
            c = self.card(grid)
            c.grid(row=0, column=i, sticky="ew", padx=(0 if i == 0 else 12, 0))
            grid.columnconfigure(i, weight=1)
            ttk.Label(c, text=label, style="CardTitle.TLabel").pack(anchor="w")
            ttk.Label(c, text=value, style="Metric.TLabel").pack(anchor="w", pady=(8, 0))

        body = ttk.Frame(self.main)
        body.pack(fill="both", expand=True, padx=24, pady=10)
        body.columnconfigure(0, weight=1)
        body.columnconfigure(1, weight=1)
        recent = self.card(body)
        recent.grid(row=0, column=0, sticky="nsew", padx=(0, 12), pady=0)
        ttk.Label(recent, text="Dernieres ventes", style="CardTitle.TLabel").pack(anchor="w", pady=(0, 8))
        self.table(
            recent,
            ("N", "Client", "Total", "Paye", "Statut", "Date"),
            [
                (r["sale_no"], r["customer_name"], self.money(r["total"]), self.money(r["paid"]), r["status"], r["created_at"][:16])
                for r in self.db.all("SELECT * FROM sales ORDER BY id DESC LIMIT 12")
            ],
        )

        alerts = self.card(body)
        alerts.grid(row=0, column=1, sticky="nsew", pady=0)
        ttk.Label(alerts, text="Produits en alerte", style="CardTitle.TLabel").pack(anchor="w", pady=(0, 8))
        self.table(
            alerts,
            ("Produit", "Categorie", "Stock", "Seuil"),
            [
                (r["name"], r["category"], r["stock"], r["alert_stock"])
                for r in self.db.all("SELECT * FROM products WHERE active = 1 AND alert_stock > 0 AND stock <= alert_stock ORDER BY stock ASC LIMIT 12")
            ],
        )

    def table(self, parent, columns, rows):
        tree = ttk.Treeview(parent, columns=columns, show="headings", height=12)
        for col in columns:
            tree.heading(col, text=col)
            tree.column(col, width=120, anchor="w", stretch=True)
        for row in rows:
            tree.insert("", "end", values=row)
        tree.pack(fill="both", expand=True)
        return tree

    def form_row(self, parent, label, row, var, width=28):
        ttk.Label(parent, text=label, background="#ffffff").grid(row=row, column=0, sticky="w", pady=5)
        entry = ttk.Entry(parent, textvariable=var, width=width)
        entry.grid(row=row, column=1, sticky="ew", pady=5, padx=(10, 0))
        return entry

    def show_products(self):
        self.clear("Produits", "Catalogue, prix, categories et stock minimum.")
        body = ttk.Frame(self.main)
        body.pack(fill="both", expand=True, padx=24, pady=8)
        body.columnconfigure(0, weight=0)
        body.columnconfigure(1, weight=1)
        form = self.card(body)
        form.grid(row=0, column=0, sticky="nsew", padx=(0, 12), pady=0)
        form.columnconfigure(1, weight=1)

        selected_product_id = tk.IntVar(value=0)
        vars_ = {k: tk.StringVar() for k in ["name", "category", "barcode", "unit", "purchase", "sale", "stock", "alert"]}
        defaults = {"category": "General", "unit": "piece", "purchase": "0", "sale": "0", "stock": "0", "alert": "0"}
        for key, value in defaults.items():
            vars_[key].set(value)

        labels = [
            ("Nom", "name"),
            ("Categorie", "category"),
            ("Code barre", "barcode"),
            ("Unite", "unit"),
            ("Prix achat", "purchase"),
            ("Prix vente", "sale"),
            ("Stock", "stock"),
            ("Seuil alerte", "alert"),
        ]
        title_label = ttk.Label(form, text="Nouveau produit", style="CardTitle.TLabel")
        title_label.grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 10))
        for idx, (label, key) in enumerate(labels, 1):
            self.form_row(form, label, idx, vars_[key])

        def clear_selection():
            selected_product_id.set(0)
            title_label.config(text="Nouveau produit")
            for key, value in defaults.items():
                vars_[key].set(value)
            vars_["barcode"].set("")

        def load_product_selection(_=None):
            selected = product_tree.selection()
            if not selected:
                return
            product_id = int(product_tree.item(selected[0])["values"][0])
            product = self.db.one("SELECT * FROM products WHERE id = ?", (product_id,))
            if not product:
                return
            selected_product_id.set(product["id"])
            title_label.config(text="Modifier produit")
            vars_["name"].set(product["name"])
            vars_["category"].set(product["category"])
            vars_["barcode"].set(product["barcode"])
            vars_["unit"].set(product["unit"])
            vars_["purchase"].set(str(product["purchase_price"]))
            vars_["sale"].set(str(product["sale_price"]))
            vars_["stock"].set(str(product["stock"]))
            vars_["alert"].set(str(product["alert_stock"]))

        def save_product():
            if not vars_["name"].get().strip():
                messagebox.showwarning(APP_NAME, "Le nom du produit est obligatoire.")
                return
            try:
                values = (
                    vars_["name"].get().strip(),
                    vars_["category"].get().strip() or "General",
                    vars_["barcode"].get().strip(),
                    vars_["unit"].get().strip() or "piece",
                    float(vars_["purchase"].get() or 0),
                    float(vars_["sale"].get() or 0),
                    float(vars_["stock"].get() or 0),
                    float(vars_["alert"].get() or 0),
                )
            except ValueError:
                messagebox.showerror(APP_NAME, "Les prix et stocks doivent etre des nombres.")
                return
            if selected_product_id.get():
                self.db.execute(
                    """
                    UPDATE products
                    SET name = ?, category = ?, barcode = ?, unit = ?, purchase_price = ?, sale_price = ?, stock = ?, alert_stock = ?
                    WHERE id = ?
                    """,
                    (*values, selected_product_id.get()),
                )
            else:
                self.db.execute(
                    """
                    INSERT INTO products(name, category, barcode, unit, purchase_price, sale_price, stock, alert_stock, created_at)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """,
                    (*values, datetime.now().isoformat(timespec="seconds")),
                )
            self.show_products()

        def delete_product():
            if not selected_product_id.get():
                messagebox.showwarning(APP_NAME, "Selectionnez un produit pour le supprimer.")
                return
            try:
                self.db.execute("DELETE FROM products WHERE id = ?", (selected_product_id.get(),))
                messagebox.showinfo(APP_NAME, "Produit supprime.")
            except sqlite3.IntegrityError:
                self.db.execute("UPDATE products SET active = 0 WHERE id = ?", (selected_product_id.get(),))
                messagebox.showinfo(APP_NAME, "Produit desactive car des references existent.")
            self.show_products()

        ttk.Button(form, text="Enregistrer", style="Accent.TButton", command=save_product).grid(row=10, column=0, columnspan=2, sticky="ew", pady=(12, 0))
        ttk.Button(form, text="Supprimer", style="Danger.TButton", command=delete_product).grid(row=11, column=0, columnspan=2, sticky="ew", pady=(6, 0))
        ttk.Button(form, text="Nouveau", command=clear_selection).grid(row=12, column=0, columnspan=2, sticky="ew", pady=(6, 0))

        listing = self.card(body)
        listing.grid(row=0, column=1, sticky="nsew", pady=0)
        listing.columnconfigure(0, weight=1)
        ttk.Label(listing, text="Catalogue", style="CardTitle.TLabel").pack(anchor="w", pady=(0, 8))
        product_tree = self.table(
            listing,
            ("ID", "Produit", "Categorie", "Unite", "Achat", "Vente", "Stock", "Alerte"),
            [
                (r["id"], r["name"], r["category"], r["unit"], self.money(r["purchase_price"]), self.money(r["sale_price"]), r["stock"], r["alert_stock"])
                for r in self.db.all("SELECT * FROM products WHERE active = 1 ORDER BY name")
            ],
        )
        product_tree.bind("<<TreeviewSelect>>", load_product_selection)

    def show_sales(self):
        self.clear("POS tactile", "Caisse rapide avec tuiles produits, panier, paiement et recu.")
        body = ttk.Frame(self.main)
        body.pack(fill="both", expand=True, padx=24, pady=8)
        body.columnconfigure(0, weight=1)
        body.columnconfigure(1, weight=0)
        body.rowconfigure(0, weight=1)

        product_panel = self.card(body, padx=12, pady=12)
        product_panel.grid(row=0, column=0, sticky="nsew", padx=(0, 12), pady=0)
        cart_panel = self.card(body, padx=12, pady=12)
        cart_panel.grid(row=0, column=1, sticky="nsew", pady=0)
        cart_panel.configure(width=420)
        cart_panel.pack_propagate(False)

        search = tk.StringVar()
        category = tk.StringVar(value="Tous")
        total_var = tk.StringVar(value=self.money(0))
        customer = tk.StringVar(value="Client comptoir")
        paid = tk.StringVar(value="0")
        payment = tk.StringVar(value="Especes")
        customers = self.db.all("SELECT * FROM customers ORDER BY name")
        customer_values = ["Client comptoir"] + [f"{c['id']} - {c['name']}" for c in customers]

        toolbar = ttk.Frame(product_panel, style="Card.TFrame")
        toolbar.pack(fill="x", pady=(0, 10))
        toolbar.columnconfigure(1, weight=1)
        toolbar.columnconfigure(2, weight=0)
        toolbar.columnconfigure(3, weight=0)
        ttk.Label(toolbar, text="Recherche", background="#ffffff").grid(row=0, column=0, sticky="w")
        search_entry = ttk.Entry(toolbar, textvariable=search)
        search_entry.grid(row=0, column=1, sticky="ew", padx=8)
        categories = ["Tous"] + [r["category"] for r in self.db.all("SELECT DISTINCT category FROM products WHERE active = 1 ORDER BY category")]
        ttk.Combobox(toolbar, textvariable=category, values=categories, width=18, state="readonly").grid(row=0, column=2, padx=(0, 8))
        ttk.Button(toolbar, text="Actualiser", command=self.show_sales).grid(row=0, column=3)

        tiles_outer = ttk.Frame(product_panel, style="Card.TFrame")
        tiles_outer.pack(fill="both", expand=True)
        canvas = tk.Canvas(tiles_outer, bg="#ffffff", highlightthickness=0)
        scrollbar = ttk.Scrollbar(tiles_outer, orient="vertical", command=canvas.yview)
        tiles = ttk.Frame(canvas, style="Card.TFrame")
        tiles.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=tiles, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        ttk.Label(cart_panel, text="Panier client", style="CardTitle.TLabel").pack(anchor="w", pady=(0, 8))
        cart_tree = self.table(cart_panel, ("Produit", "Qte", "Total"), [])

        def refresh_cart():
            for item in cart_tree.get_children():
                cart_tree.delete(item)
            for idx, line in enumerate(self.cart, 1):
                cart_tree.insert("", "end", iid=str(idx - 1), values=(line["name"], line["qty"], self.money(line["total"])))
            total = sum(line["total"] for line in self.cart)
            total_var.set(self.money(total))
            paid.set(str(int(total)))

        def add_product(product, qty=1):
            for line in self.cart:
                if line["product_id"] == product["id"]:
                    line["qty"] += qty
                    line["total"] = line["qty"] * line["price"]
                    refresh_cart()
                    return
            self.cart.append(
                {
                    "product_id": product["id"],
                    "name": product["name"],
                    "qty": qty,
                    "price": product["sale_price"],
                    "total": qty * product["sale_price"],
                }
            )
            refresh_cart()

        def remove_selected():
            selected = cart_tree.selection()
            if selected:
                index = int(selected[0])
                if 0 <= index < len(self.cart):
                    self.cart.pop(index)
                    refresh_cart()

        def edit_cart_quantity(event=None):
            selected = cart_tree.selection()
            if not selected:
                return
            index = int(selected[0])
            if index < 0 or index >= len(self.cart):
                return
            line = self.cart[index]
            new_qty = simpledialog.askfloat(APP_NAME, f"Quantite pour {line['name']}", initialvalue=line['qty'], minvalue=0)
            if new_qty is None:
                return
            if new_qty <= 0:
                self.cart.pop(index)
            else:
                line['qty'] = new_qty
                line['total'] = line['qty'] * line['price']
            refresh_cart()

        def render_products(*_):
            for child in tiles.winfo_children():
                child.destroy()
            query = search.get().strip().lower()
            selected_category = category.get()
            products = self.db.all("SELECT * FROM products WHERE active = 1 ORDER BY category, name")
            shown = 0
            for product in products:
                if selected_category != "Tous" and product["category"] != selected_category:
                    continue
                if query and query not in product["name"].lower() and query not in (product["barcode"] or "").lower():
                    continue
                text = f"{product['name']}\n{self.money(product['sale_price'])}\nStock: {product['stock']}"
                btn = ttk.Button(tiles, text=text, style="ProductTile.TButton", command=lambda p=product: add_product(p))
                btn.grid(row=shown // 3, column=shown % 3, sticky="nsew", padx=6, pady=6)
                tiles.columnconfigure(shown % 3, weight=1)
                shown += 1
            if shown == 0:
                ttk.Label(tiles, text="Aucun produit trouve", background="#ffffff").grid(row=0, column=0, padx=20, pady=20)

        search.trace_add("write", render_products)
        category.trace_add("write", render_products)

        cart_tree.bind("<Double-1>", edit_cart_quantity)

        footer = ttk.Frame(cart_panel, style="Card.TFrame")
        footer.pack(fill="x", pady=(12, 0))
        ttk.Label(footer, text="Client", background="#ffffff").grid(row=0, column=0, sticky="w", pady=5)
        ttk.Combobox(footer, textvariable=customer, values=customer_values, width=28, state="readonly").grid(row=0, column=1, sticky="ew", pady=5, padx=(10, 0))
        ttk.Label(footer, text="Paiement", background="#ffffff").grid(row=1, column=0, sticky="w", pady=5)
        ttk.Combobox(footer, textvariable=payment, values=("Especes", "Mobile Money", "Carte", "Credit"), width=25, state="readonly").grid(row=1, column=1, sticky="ew", pady=5, padx=(10, 0))
        self.form_row(footer, "Montant paye", 2, paid, 28)
        ttk.Label(footer, text="Total", background="#ffffff").grid(row=3, column=0, sticky="w", pady=8)
        ttk.Label(footer, textvariable=total_var, style="Metric.TLabel").grid(row=3, column=1, sticky="w", padx=(10, 0))

        pay_buttons = ttk.Frame(cart_panel, style="Card.TFrame")
        pay_buttons.pack(fill="x", pady=(8, 0))

        def quick_payment(method):
            payment.set(method)
            paid.set(str(int(sum(line["total"] for line in self.cart))))

        for method in ("Especes", "Mobile Money", "Carte", "Credit"):
            ttk.Button(pay_buttons, text=method, command=lambda m=method: quick_payment(m)).pack(side="left", fill="x", expand=True, padx=2)

        def validate_sale(print_receipt=True):
            selected_customer_id = None
            selected_customer_name = customer.get()
            if customer.get() != "Client comptoir" and " - " in customer.get():
                selected_customer_id = int(customer.get().split(" - ")[0])
                selected_customer_name = customer.get().split(" - ", 1)[1]
            total = sum(line["total"] for line in self.cart)
            paid_amount = 0
            try:
                paid_amount = float(paid.get() or 0)
            except ValueError:
                messagebox.showerror(APP_NAME, "Montant paye invalide.")
                return
            status = "Paye" if paid_amount >= total else "Credit"
            if status == "Credit" and not selected_customer_id:
                messagebox.showwarning(APP_NAME, "Pour un credit, choisissez un client existant.")
                return
            sale = self.record_sale(selected_customer_name, paid.get(), payment.get(), customer_id=selected_customer_id)
            if not sale:
                return
            receipt_path = self.create_receipt(sale["sale_no"], sale["customer"], sale["total"], sale["paid"], sale["payment"], sale["status"], sale["items"])
            self.cart = []
            if print_receipt:
                self.show_receipt_window(receipt_path)
            else:
                messagebox.showinfo(APP_NAME, f"Vente enregistree : {sale['sale_no']}")
            self.show_sales()

        ttk.Button(cart_panel, text="Valider et recu", style="Accent.TButton", command=validate_sale).pack(fill="x", pady=(10, 0))
        ttk.Button(cart_panel, text="Retirer ligne", command=remove_selected).pack(fill="x", pady=(6, 0))
        ttk.Button(cart_panel, text="Vider panier", style="Danger.TButton", command=lambda: (self.cart.clear(), refresh_cart())).pack(fill="x", pady=(6, 0))

        render_products()
        refresh_cart()

    def record_sale(self, customer_name, paid_value, payment_method, customer_id=None):
        if not self.cart:
            messagebox.showwarning(APP_NAME, "Le panier est vide.")
            return None
        total = sum(line["total"] for line in self.cart)
        try:
            paid_amount = float(paid_value or 0)
        except ValueError:
            messagebox.showerror(APP_NAME, "Montant paye invalide.")
            return None
        now = datetime.now().isoformat(timespec="seconds")
        sale_no = datetime.now().strftime("V%Y%m%d%H%M%S")
        invoice_no = sale_no.replace("V", "FAC-", 1)
        status = "Paye" if paid_amount >= total else "Credit"
        customer_label = customer_name.strip() or "Client comptoir"
        cur = self.db.execute(
            """
            INSERT INTO sales(sale_no, customer_id, customer_name, total, paid, payment_method, status, invoice_no, fne_status, created_at)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, 'A exporter', ?)
            """,
            (sale_no, customer_id, customer_label, total, paid_amount, payment_method, status, invoice_no, now),
        )
        sale_id = cur.lastrowid
        items = list(self.cart)
        for line in items:
            self.db.execute(
                "INSERT INTO sale_items(sale_id, product_id, product_name, quantity, unit_price, total) VALUES (?, ?, ?, ?, ?, ?)",
                (sale_id, line["product_id"], line["name"], line["qty"], line["price"], line["total"]),
            )
            self.db.execute("UPDATE products SET stock = stock - ? WHERE id = ?", (line["qty"], line["product_id"]))
            self.db.execute(
                "INSERT INTO stock_movements(product_id, product_name, movement_type, quantity, note, created_at) VALUES (?, ?, ?, ?, ?, ?)",
                (line["product_id"], line["name"], "Sortie vente", -line["qty"], sale_no, now),
            )
        if status == "Credit" and customer_id:
            due = total - paid_amount
            self.db.execute("UPDATE customers SET balance = balance + ? WHERE id = ?", (due, customer_id))
        self.db.execute(
            "INSERT INTO cash_entries(entry_type, label, amount, payment_method, created_at) VALUES (?, ?, ?, ?, ?)",
            ("Recette", f"Vente {sale_no}", paid_amount, payment_method, now),
        )
        return {
            "sale_no": sale_no,
            "sale_id": sale_id,
            "invoice_no": invoice_no,
            "customer": customer_label,
            "total": total,
            "paid": paid_amount,
            "payment": payment_method,
            "status": status,
            "items": items,
        }

    def sale_payload(self, sale_id=None, sale_no=None):
        if sale_id is not None:
            sale = self.db.one("SELECT * FROM sales WHERE id = ?", (sale_id,))
        else:
            sale = self.db.one("SELECT * FROM sales WHERE sale_no = ?", (sale_no,))
        if not sale:
            return None
        items = self.db.all(
            """
            SELECT si.*, p.barcode, p.unit, p.purchase_price
            FROM sale_items si
            LEFT JOIN products p ON p.id = si.product_id
            WHERE si.sale_id = ?
            ORDER BY si.id
            """,
            (sale["id"],),
        )
        customer = self.db.one("SELECT * FROM customers WHERE id = ?", (sale["customer_id"],)) if sale["customer_id"] else None
        return {"sale": sale, "items": items, "customer": customer}

    def create_receipt(self, sale_no, customer, total, paid, payment, status, items):
        RECEIPTS_DIR.mkdir(parents=True, exist_ok=True)
        path = RECEIPTS_DIR / f"{sale_no}.txt"
        change = max(0, paid - total)
        lines = [
            self.setting("shop_name"),
            self.setting("shop_phone"),
            self.setting("shop_address"),
            "-" * 32,
            f"Recu: {sale_no}",
            f"Date: {datetime.now().strftime('%d/%m/%Y %H:%M')}",
            f"Client: {customer}",
            "-" * 32,
        ]
        for item in items:
            lines.append(f"{item['name'][:20]}")
            lines.append(f"  {item['qty']} x {self.money(item['price'])} = {self.money(item['total'])}")
        lines.extend(
            [
                "-" * 32,
                f"Total: {self.money(total)}",
                f"Paye: {self.money(paid)}",
                f"Rendu: {self.money(change)}",
                f"Mode: {payment}",
                f"Statut: {status}",
                "-" * 32,
                "Merci pour votre achat.",
            ]
        )
        path.write_text("\n".join(lines), encoding="utf-8")
        return path

    def create_professional_invoice(self, sale_id):
        payload = self.sale_payload(sale_id=sale_id)
        if not payload:
            messagebox.showwarning(APP_NAME, "Vente introuvable.")
            return None
        sale = payload["sale"]
        items = payload["items"]
        customer = payload["customer"]
        INVOICES_DIR.mkdir(parents=True, exist_ok=True)
        invoice_no = sale["invoice_no"] or sale["sale_no"].replace("V", "FAC-", 1)
        path = INVOICES_DIR / f"{invoice_no}.txt"
        remaining = max(0, float(sale["total"]) - float(sale["paid"]))
        lines = [
            self.setting("shop_name"),
            f"Tel: {self.setting('shop_phone')} | Email: {self.setting('shop_email')}",
            self.setting("shop_address"),
            f"NCC: {self.setting('shop_ncc') or '-'} | RCCM: {self.setting('shop_rccm') or '-'}",
            f"Regime fiscal: {self.setting('shop_tax_regime') or '-'}",
            "=" * 72,
            "FACTURE PRO FORMA / DOCUMENT INTERNE",
            f"Numero facture: {invoice_no}",
            f"Reference vente: {sale['sale_no']}",
            f"Date: {sale['created_at'][:16]}",
            f"Client: {sale['customer_name']}",
            f"Contact client: {(customer['phone'] if customer else '') or '-'}",
            f"NCC client: {(customer['ncc'] if customer else '') or '-'}",
            "-" * 72,
            f"{'Designation':28} {'Qte':>8} {'PU':>14} {'Total':>14}",
            "-" * 72,
        ]
        for item in items:
            lines.append(
                f"{item['product_name'][:28]:28} {item['quantity']:8.2f} {self.money(item['unit_price']):>14} {self.money(item['total']):>14}"
            )
        lines.extend(
            [
                "-" * 72,
                f"Total TTC: {self.money(sale['total'])}",
                f"Montant paye: {self.money(sale['paid'])}",
                f"Reste a payer: {self.money(remaining)}",
                f"Mode paiement: {sale['payment_method']}",
                f"Statut paiement: {sale['status']}",
                f"Statut FNE: {sale['fne_status'] or 'A exporter'}",
                f"Reference FNE: {sale['fne_reference'] or '-'}",
                "-" * 72,
                self.setting("invoice_footer"),
                self.setting("fne_note"),
                DEVELOPER_CREDIT,
            ]
        )
        path.write_text("\n".join(lines), encoding="utf-8")
        return path

    def fne_payload_for_sale(self, sale_id):
        payload = self.sale_payload(sale_id=sale_id)
        if not payload:
            return None
        sale = payload["sale"]
        customer = payload["customer"]
        items = []
        for item in payload["items"]:
            items.append(
                {
                    "product_id": item["product_id"],
                    "designation": item["product_name"],
                    "barcode": item["barcode"] or "",
                    "quantity": float(item["quantity"]),
                    "unit": item["unit"] or "piece",
                    "unit_price": float(item["unit_price"]),
                    "total": float(item["total"]),
                }
            )
        return {
            "source": APP_NAME,
            "version": APP_VERSION,
            "document_type": "facture_vente",
            "internal_number": sale["invoice_no"] or sale["sale_no"],
            "sale_no": sale["sale_no"],
            "issued_at": sale["created_at"],
            "currency": self.setting("currency") or "FCFA",
            "seller": {
                "name": self.setting("shop_name"),
                "phone": self.setting("shop_phone"),
                "email": self.setting("shop_email"),
                "address": self.setting("shop_address"),
                "ncc": self.setting("shop_ncc"),
                "rccm": self.setting("shop_rccm"),
                "tax_regime": self.setting("shop_tax_regime"),
            },
            "buyer": {
                "name": sale["customer_name"],
                "phone": (customer["phone"] if customer else "") or "",
                "email": (customer["email"] if customer else "") or "",
                "ncc": (customer["ncc"] if customer else "") or "",
                "address": (customer["address"] if customer else "") or "",
            },
            "items": items,
            "totals": {
                "total_amount": float(sale["total"]),
                "paid_amount": float(sale["paid"]),
                "remaining_amount": max(0, float(sale["total"]) - float(sale["paid"])),
            },
            "payment": {"method": sale["payment_method"], "status": sale["status"]},
            "fne": {
                "status": sale["fne_status"] or "A exporter",
                "reference": sale["fne_reference"] or "",
                "note": self.setting("fne_note"),
            },
        }

    def export_sale_fne(self, sale_id):
        payload = self.fne_payload_for_sale(sale_id)
        if not payload:
            messagebox.showwarning(APP_NAME, "Vente introuvable.")
            return None
        FNE_DIR.mkdir(parents=True, exist_ok=True)
        path = FNE_DIR / f"fne-{payload['internal_number']}.json"
        path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
        self.db.execute(
            "UPDATE sales SET fne_status = 'Exportee', fne_exported_at = ? WHERE id = ?",
            (datetime.now().isoformat(timespec="seconds"), sale_id),
        )
        return path

    def show_receipt_window(self, path):
        win = tk.Toplevel(self)
        win.title("Recu client")
        win.geometry("420x620")
        text = tk.Text(win, font=("Consolas", 11), wrap="word")
        text.pack(fill="both", expand=True, padx=12, pady=12)
        text.insert("1.0", path.read_text(encoding="utf-8"))
        text.configure(state="disabled")
        button_frame = ttk.Frame(win)
        button_frame.pack(fill="x", padx=12, pady=(0, 12))
        ttk.Button(button_frame, text="Imprimer", command=lambda: self.print_text_file(path)).pack(side="left", expand=True, fill="x", padx=(0, 4))
        ttk.Button(button_frame, text="Ouvrir tiroir", command=self.open_cash_drawer).pack(side="left", expand=True, fill="x", padx=(4, 0))
        ttk.Button(win, text="Fermer", command=win.destroy).pack(fill="x", padx=12, pady=(0, 12))

    def show_customers(self):
        self.simple_directory(
            "Clients",
            "Carnet clients, contacts, NCC et soldes.",
            "customers",
            ["Nom", "Telephone", "Email", "NCC", "Adresse", "Type", "Solde"],
            ["name", "phone", "email", "ncc", "address", "type", "balance"],
        )
    def show_suppliers(self):
        self.simple_directory("Fournisseurs", "Contacts fournisseurs et adresses.", "suppliers", ["Nom", "Telephone", "Email", "Adresse"], ["name", "phone", "email", "address"])

    def simple_directory(self, title, subtitle, table_name, labels, fields):
        self.clear(title, subtitle)
        body = ttk.Frame(self.main)
        body.pack(fill="both", expand=True, padx=24, pady=8)
        form = self.card(body)
        form.pack(side="left", fill="y", padx=(0, 12))
        selected_id = tk.IntVar(value=0)
        vars_ = {field: tk.StringVar() for field in fields}
        title_label = ttk.Label(form, text=f"Nouveau {title.lower()[:-1]}", style="CardTitle.TLabel")
        title_label.grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 10))
        for idx, (label, field) in enumerate(zip(labels, fields), 1):
            self.form_row(form, label, idx, vars_[field])

        def clear_selection():
            selected_id.set(0)
            title_label.config(text=f"Nouveau {title.lower()[:-1]}")
            for field in fields:
                vars_[field].set("")

        def load_selection(_=None):
            selected = directory_tree.selection()
            if not selected:
                return
            row_id = int(directory_tree.item(selected[0])["values"][0])
            row = self.db.one(f"SELECT * FROM {table_name} WHERE id = ?", (row_id,))
            if not row:
                return
            selected_id.set(row_id)
            title_label.config(text=f"Modifier {title.lower()[:-1]}")
            for field in fields:
                vars_[field].set(row[field] or "")

        def save():
            values = [vars_[field].get().strip() for field in fields]
            if not values[0]:
                messagebox.showwarning(APP_NAME, "Le nom est obligatoire.")
                return
            if selected_id.get():
                set_clause = ", ".join([f"{field} = ?" for field in fields])
                self.db.execute(
                    f"UPDATE {table_name} SET {set_clause} WHERE id = ?",
                    (*values, selected_id.get()),
                )
            else:
                columns = ", ".join(fields) + ", created_at"
                placeholders = ", ".join(["?"] * (len(fields) + 1))
                self.db.execute(
                    f"INSERT INTO {table_name}({columns}) VALUES ({placeholders})",
                    (*values, datetime.now().isoformat(timespec="seconds")),
                )
            self.simple_directory(title, subtitle, table_name, labels, fields)

        def delete_selected():
            if not selected_id.get():
                messagebox.showwarning(APP_NAME, "Selectionnez un enregistrement à supprimer.")
                return
            if table_name == "customers":
                existing = self.db.one("SELECT COUNT(*) v FROM sales WHERE customer_id = ?", (selected_id.get(),))["v"]
                if existing > 0:
                    messagebox.showwarning(APP_NAME, "Impossible de supprimer un client ayant des ventes enregistrees.")
                    return
            self.db.execute(f"DELETE FROM {table_name} WHERE id = ?", (selected_id.get(),))
            self.simple_directory(title, subtitle, table_name, labels, fields)

        save_button_row = len(fields) + 2
        ttk.Button(form, text="Enregistrer", style="Accent.TButton", command=save).grid(row=save_button_row, column=0, columnspan=2, sticky="ew", pady=(12, 0))
        ttk.Button(form, text="Supprimer", style="Danger.TButton", command=delete_selected).grid(row=save_button_row + 1, column=0, columnspan=2, sticky="ew", pady=(6, 0))
        ttk.Button(form, text="Nouveau", command=clear_selection).grid(row=save_button_row + 2, column=0, columnspan=2, sticky="ew", pady=(6, 0))

        listing = self.card(body)
        listing.pack(side="left", fill="both", expand=True)
        columns = ("ID", *labels, "Date")
        rows = [
            (r["id"], *[r[field] for field in fields], r["created_at"][:10])
            for r in self.db.all(f"SELECT * FROM {table_name} ORDER BY id DESC")
        ]
        directory_tree = self.table(listing, columns, rows)
        directory_tree.bind("<<TreeviewSelect>>", load_selection)

    def show_stock(self):
        self.clear("Stock", "Entrees, sorties, corrections et historique.")
        body = ttk.Frame(self.main)
        body.pack(fill="both", expand=True, padx=24, pady=8)
        form = self.card(body)
        form.pack(side="left", fill="y", padx=(0, 12))
        products = self.db.all("SELECT id, name FROM products WHERE active = 1 ORDER BY name")
        product_var = tk.StringVar()
        movement = tk.StringVar(value="Entree stock")
        quantity = tk.StringVar(value="1")
        note = tk.StringVar()

        ttk.Label(form, text="Mouvement de stock", style="CardTitle.TLabel").grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 10))
        ttk.Label(form, text="Produit", background="#ffffff").grid(row=1, column=0, sticky="w", pady=5)
        combo = ttk.Combobox(form, textvariable=product_var, values=[f"{p['id']} - {p['name']}" for p in products], width=28)
        combo.grid(row=1, column=1, sticky="ew", pady=5, padx=(10, 0))
        ttk.Label(form, text="Type", background="#ffffff").grid(row=2, column=0, sticky="w", pady=5)
        ttk.Combobox(form, textvariable=movement, values=("Entree stock", "Sortie manuelle", "Correction"), width=28).grid(row=2, column=1, sticky="ew", pady=5, padx=(10, 0))
        self.form_row(form, "Quantite", 3, quantity)
        self.form_row(form, "Note", 4, note)

        def save_movement():
            if " - " not in product_var.get():
                messagebox.showwarning(APP_NAME, "Choisissez un produit.")
                return
            product_id = int(product_var.get().split(" - ")[0])
            product = self.db.one("SELECT * FROM products WHERE id = ?", (product_id,))
            try:
                qty = float(quantity.get() or 0)
            except ValueError:
                messagebox.showerror(APP_NAME, "Quantite invalide.")
                return
            signed = qty if movement.get() == "Entree stock" else -qty
            now = datetime.now().isoformat(timespec="seconds")
            self.db.execute("UPDATE products SET stock = stock + ? WHERE id = ?", (signed, product_id))
            self.db.execute(
                "INSERT INTO stock_movements(product_id, product_name, movement_type, quantity, note, created_at) VALUES (?, ?, ?, ?, ?, ?)",
                (product_id, product["name"], movement.get(), signed, note.get(), now),
            )
            self.show_stock()

        ttk.Button(form, text="Enregistrer", style="Accent.TButton", command=save_movement).grid(row=6, column=0, columnspan=2, sticky="ew", pady=(12, 0))

        listing = self.card(body)
        listing.pack(side="left", fill="both", expand=True)
        ttk.Label(listing, text="Historique stock", style="CardTitle.TLabel").pack(anchor="w", pady=(0, 8))
        self.table(
            listing,
            ("Produit", "Type", "Quantite", "Note", "Date"),
            [(r["product_name"], r["movement_type"], r["quantity"], r["note"], r["created_at"][:16]) for r in self.db.all("SELECT * FROM stock_movements ORDER BY id DESC LIMIT 80")],
        )

    def show_inventory(self):
        self.clear("Inventaire", "Controle physique du stock et ajustement des ecarts.")
        body = ttk.Frame(self.main)
        body.pack(fill="both", expand=True, padx=24, pady=8)
        left = self.card(body)
        left.pack(side="left", fill="both", expand=True, padx=(0, 12))
        right = self.card(body)
        right.pack(side="left", fill="y")

        ttk.Label(left, text="Etat du stock", style="CardTitle.TLabel").pack(anchor="w", pady=(0, 8))
        products = self.db.all("SELECT * FROM products WHERE active = 1 ORDER BY category, name")
        tree = self.table(
            left,
            ("ID", "Produit", "Categorie", "Theorique", "Seuil", "Valeur achat"),
            [(p["id"], p["name"], p["category"], p["stock"], p["alert_stock"], self.money(p["stock"] * p["purchase_price"])) for p in products],
        )

        selected_product = tk.StringVar()
        counted = tk.StringVar(value="0")
        note = tk.StringVar(value="Inventaire physique")
        product_choices = [f"{p['id']} - {p['name']}" for p in products]
        ttk.Label(right, text="Ajustement inventaire", style="CardTitle.TLabel").grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 10))
        ttk.Label(right, text="Produit", background="#ffffff").grid(row=1, column=0, sticky="w", pady=5)
        ttk.Combobox(right, textvariable=selected_product, values=product_choices, width=30).grid(row=1, column=1, sticky="ew", pady=5, padx=(10, 0))
        self.form_row(right, "Stock compte", 2, counted, 30)
        self.form_row(right, "Note", 3, note, 30)

        def fill_from_table(_=None):
            selected = tree.selection()
            if not selected:
                return
            values = tree.item(selected[0], "values")
            selected_product.set(f"{values[0]} - {values[1]}")
            counted.set(str(values[3]))

        tree.bind("<<TreeviewSelect>>", fill_from_table)

        def adjust_inventory():
            if " - " not in selected_product.get():
                messagebox.showwarning(APP_NAME, "Choisissez un produit.")
                return
            product_id = int(selected_product.get().split(" - ")[0])
            product = self.db.one("SELECT * FROM products WHERE id = ?", (product_id,))
            try:
                counted_value = float(counted.get() or 0)
            except ValueError:
                messagebox.showerror(APP_NAME, "Stock compte invalide.")
                return
            diff = counted_value - float(product["stock"])
            now = datetime.now().isoformat(timespec="seconds")
            self.db.execute("UPDATE products SET stock = ? WHERE id = ?", (counted_value, product_id))
            self.db.execute(
                "INSERT INTO stock_movements(product_id, product_name, movement_type, quantity, note, created_at) VALUES (?, ?, ?, ?, ?, ?)",
                (product_id, product["name"], "Ajustement inventaire", diff, note.get(), now),
            )
            messagebox.showinfo(APP_NAME, f"Ecart applique : {diff}")
            self.show_inventory()

        ttk.Button(right, text="Appliquer ajustement", style="Accent.TButton", command=adjust_inventory).grid(row=5, column=0, columnspan=2, sticky="ew", pady=(12, 0))
        ttk.Button(right, text="Exporter inventaire Excel", style="Accent.TButton", command=self.export_inventory).grid(row=6, column=0, columnspan=2, sticky="ew", pady=(6, 0))
        ttk.Button(right, text="Exporter inventaire CSV", command=self.export_inventory_csv).grid(row=7, column=0, columnspan=2, sticky="ew", pady=(6, 0))
        ttk.Button(right, text="Imprimer inventaire", command=self.print_inventory).grid(row=8, column=0, columnspan=2, sticky="ew", pady=(6, 0))

    def export_inventory(self):
        path = filedialog.asksaveasfilename(defaultextension=".xls", filetypes=[("Microsoft Excel", "*.xls")], initialfile="inventaire-magazin-ci.xls")
        if not path:
            return
        rows = self.db.all("SELECT * FROM products WHERE active = 1 ORDER BY category, name")
        self.export_excel_xml(
            path,
            "Inventaire",
            ["Produit", "Categorie", "Unite", "Stock", "Seuil", "Prix achat", "Prix vente"],
            [[r["name"], r["category"], r["unit"], r["stock"], r["alert_stock"], r["purchase_price"], r["sale_price"]] for r in rows],
        )
        messagebox.showinfo(APP_NAME, "Inventaire exporte en fichier Microsoft Excel.")

    def export_inventory_csv(self):
        path = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV", "*.csv")], initialfile="inventaire-magazin-ci.csv")
        if not path:
            return
        rows = self.db.all("SELECT * FROM products WHERE active = 1 ORDER BY category, name")
        self.export_csv(
            path,
            ["Produit", "Categorie", "Unite", "Stock", "Seuil", "Prix achat", "Prix vente"],
            [[r["name"], r["category"], r["unit"], r["stock"], r["alert_stock"], r["purchase_price"], r["sale_price"]] for r in rows],
        )
        messagebox.showinfo(APP_NAME, "Inventaire exporte en fichier CSV.")

    def print_inventory(self):
        rows = self.db.all("SELECT * FROM products WHERE active = 1 ORDER BY category, name")
        if not rows:
            messagebox.showinfo(APP_NAME, "Aucun produit a imprimer.")
            return
        RECEIPTS_DIR.mkdir(parents=True, exist_ok=True)
        path = RECEIPTS_DIR / "inventaire-magazin-ci.txt"
        lines = [
            self.setting("shop_name"),
            self.setting("shop_phone"),
            self.setting("shop_address"),
            "Liste d'inventaire actualisee",
            "Date: " + datetime.now().strftime("%d/%m/%Y %H:%M"),
            "" + "=" * 52,
            f"{'Produit':30} {'Categorie':12} {'Stock':>6} {'Seuil':>6} {'P. Achat':>10} {'P. Vente':>10}",
            "" + "=" * 52,
        ]
        for r in rows:
            lines.append(
                f"{r['name'][:30]:30} {r['category'][:12]:12} {r['stock']:6.0f} {r['alert_stock']:6.0f} {r['purchase_price']:10.0f} {r['sale_price']:10.0f}"
            )
        lines.append("" + "=" * 52)
        path.write_text("\n".join(lines), encoding="utf-8")
        try:
            os.startfile(path, "print")
            messagebox.showinfo(APP_NAME, f"Inventaire envoye a l'imprimante via {path}.")
        except OSError as exc:
            messagebox.showerror(APP_NAME, f"Impossible d'imprimer l'inventaire : {exc}")

    def show_cash(self):
        self.clear("Caisse", "Recettes, depenses et mouvements de tresorerie.")
        body = ttk.Frame(self.main)
        body.pack(fill="both", expand=True, padx=24, pady=8)
        form = self.card(body)
        form.pack(side="left", fill="y", padx=(0, 12))
        entry_type = tk.StringVar(value="Depense")
        label = tk.StringVar()
        amount = tk.StringVar(value="0")
        payment = tk.StringVar(value="Especes")

        ttk.Label(form, text="Nouvelle ecriture", style="CardTitle.TLabel").grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 10))
        ttk.Label(form, text="Type", background="#ffffff").grid(row=1, column=0, sticky="w", pady=5)
        ttk.Combobox(form, textvariable=entry_type, values=("Recette", "Depense", "Apport", "Retrait"), width=28).grid(row=1, column=1, sticky="ew", pady=5, padx=(10, 0))
        self.form_row(form, "Libelle", 2, label)
        self.form_row(form, "Montant", 3, amount)
        ttk.Label(form, text="Paiement", background="#ffffff").grid(row=4, column=0, sticky="w", pady=5)
        ttk.Combobox(form, textvariable=payment, values=("Especes", "Mobile Money", "Carte", "Banque"), width=28).grid(row=4, column=1, sticky="ew", pady=5, padx=(10, 0))

        def save_entry():
            try:
                value = float(amount.get() or 0)
            except ValueError:
                messagebox.showerror(APP_NAME, "Montant invalide.")
                return
            if not label.get().strip():
                messagebox.showwarning(APP_NAME, "Libelle obligatoire.")
                return
            self.db.execute(
                "INSERT INTO cash_entries(entry_type, label, amount, payment_method, created_at) VALUES (?, ?, ?, ?, ?)",
                (entry_type.get(), label.get().strip(), value, payment.get(), datetime.now().isoformat(timespec="seconds")),
            )
            self.show_cash()

        ttk.Button(form, text="Enregistrer", style="Accent.TButton", command=save_entry).grid(row=6, column=0, columnspan=2, sticky="ew", pady=(12, 0))
        listing = self.card(body)
        listing.pack(side="left", fill="both", expand=True)
        self.table(
            listing,
            ("Type", "Libelle", "Montant", "Paiement", "Date"),
            [(r["entry_type"], r["label"], self.money(r["amount"]), r["payment_method"], r["created_at"][:16]) for r in self.db.all("SELECT * FROM cash_entries ORDER BY id DESC LIMIT 100")],
        )

    def draw_bar_chart(self, parent, title, rows, label_key, value_key, color="#0f766e"):
        frame = self.card(parent)
        frame.pack(fill="both", expand=True, padx=6, pady=6)
        ttk.Label(frame, text=title, style="CardTitle.TLabel").pack(anchor="w", pady=(0, 8))
        canvas = tk.Canvas(frame, height=220, bg="#ffffff", highlightthickness=0)
        canvas.pack(fill="both", expand=True)
        if not rows:
            canvas.create_text(160, 90, text="Aucune donnee", fill="#6b7280", font=("Segoe UI", 12))
            return
        max_value = max(float(row[value_key] or 0) for row in rows) or 1
        width = 520
        bar_h = 26
        gap = 12
        for idx, row in enumerate(rows[:7]):
            y = 18 + idx * (bar_h + gap)
            value = float(row[value_key] or 0)
            bar_w = int((value / max_value) * 300)
            label = str(row[label_key])[:22]
            canvas.create_text(8, y + 13, text=label, anchor="w", fill="#111827", font=("Segoe UI", 9))
            canvas.create_rectangle(170, y, 170 + bar_w, y + bar_h, fill=color, outline="")
            canvas.create_text(180 + bar_w, y + 13, text=self.money(value), anchor="w", fill="#111827", font=("Segoe UI", 9, "bold"))
        canvas.configure(scrollregion=(0, 0, width, 260))

    def show_statistics(self):
        self.clear("Statistiques", "Graphiques et indicateurs pour piloter le commerce.")
        today = date.today().isoformat()
        total_sales = self.db.one("SELECT COALESCE(SUM(total), 0) v FROM sales WHERE date(created_at) = ?", (today,))["v"]
        total_paid = self.db.one("SELECT COALESCE(SUM(paid), 0) v FROM sales WHERE date(created_at) = ?", (today,))["v"]
        margin = self.db.one(
            """
            SELECT COALESCE(SUM((si.unit_price - p.purchase_price) * si.quantity), 0) v
            FROM sale_items si
            JOIN products p ON p.id = si.product_id
            JOIN sales s ON s.id = si.sale_id
            WHERE date(s.created_at) = ?
            """,
            (today,),
        )["v"]
        low_stock = self.db.one("SELECT COUNT(*) v FROM products WHERE active = 1 AND alert_stock > 0 AND stock <= alert_stock")["v"]

        metrics = ttk.Frame(self.main)
        metrics.pack(fill="x", padx=24, pady=8)
        for i, (label, value) in enumerate(
            [
                ("Ventes aujourd'hui", self.money(total_sales)),
                ("Encaisse aujourd'hui", self.money(total_paid)),
                ("Marge estimee", self.money(margin)),
                ("Produits en alerte", str(low_stock)),
            ]
        ):
            card = self.card(metrics)
            card.grid(row=0, column=i, sticky="ew", padx=(0 if i == 0 else 12, 0))
            metrics.columnconfigure(i, weight=1)
            ttk.Label(card, text=label, style="CardTitle.TLabel").pack(anchor="w")
            ttk.Label(card, text=value, style="Metric.TLabel").pack(anchor="w", pady=(8, 0))

        charts = ttk.Frame(self.main)
        charts.pack(fill="both", expand=True, padx=18, pady=8)
        left = ttk.Frame(charts)
        left.pack(side="left", fill="both", expand=True)
        right = ttk.Frame(charts)
        right.pack(side="left", fill="both", expand=True)

        daily = self.db.all(
            """
            SELECT date(created_at) label, COALESCE(SUM(total), 0) value
            FROM sales
            GROUP BY date(created_at)
            ORDER BY label DESC
            LIMIT 7
            """
        )
        top_products = self.db.all(
            """
            SELECT product_name label, COALESCE(SUM(total), 0) value
            FROM sale_items
            GROUP BY product_id, product_name
            ORDER BY value DESC
            LIMIT 7
            """
        )
        payments = self.db.all(
            """
            SELECT payment_method label, COALESCE(SUM(paid), 0) value
            FROM sales
            GROUP BY payment_method
            ORDER BY value DESC
            """
        )
        stock_value = self.db.all(
            """
            SELECT name label, COALESCE(stock * purchase_price, 0) value
            FROM products
            WHERE active = 1
            ORDER BY value DESC
            LIMIT 7
            """
        )
        self.draw_bar_chart(left, "Ventes par jour", daily, "label", "value", "#0f766e")
        self.draw_bar_chart(left, "Top produits vendus", top_products, "label", "value", "#2563eb")
        self.draw_bar_chart(right, "Paiements", payments, "label", "value", "#7c3aed")
        self.draw_bar_chart(right, "Valeur du stock", stock_value, "label", "value", "#ea580c")

    def show_reports(self):
        self.clear("Rapports", "Filtres, exports, factures professionnelles et preparation FNE.")
        body = ttk.Frame(self.main)
        body.pack(fill="both", expand=True, padx=24, pady=8)
        filter_box = self.card(body)
        filter_box.pack(fill="x")
        filter_box.columnconfigure(7, weight=1)
        start_var = tk.StringVar(value=date.today().replace(day=1).isoformat())
        end_var = tk.StringVar(value=date.today().isoformat())
        payment_var = tk.StringVar(value="Tous")
        status_var = tk.StringVar(value="Tous")
        search_var = tk.StringVar()
        ttk.Label(filter_box, text="Debut", background="#ffffff").grid(row=0, column=0, sticky="w", padx=(0, 6))
        ttk.Entry(filter_box, textvariable=start_var, width=12).grid(row=0, column=1, sticky="w", padx=(0, 10))
        ttk.Label(filter_box, text="Fin", background="#ffffff").grid(row=0, column=2, sticky="w", padx=(0, 6))
        ttk.Entry(filter_box, textvariable=end_var, width=12).grid(row=0, column=3, sticky="w", padx=(0, 10))
        ttk.Label(filter_box, text="Paiement", background="#ffffff").grid(row=0, column=4, sticky="w", padx=(0, 6))
        ttk.Combobox(filter_box, textvariable=payment_var, values=["Tous", "Especes", "Mobile Money", "Carte", "Credit"], state="readonly", width=15).grid(row=0, column=5, sticky="w", padx=(0, 10))
        ttk.Label(filter_box, text="Statut", background="#ffffff").grid(row=0, column=6, sticky="w", padx=(0, 6))
        ttk.Combobox(filter_box, textvariable=status_var, values=["Tous", "Paye", "Credit"], state="readonly", width=12).grid(row=0, column=7, sticky="w", padx=(0, 10))
        ttk.Label(filter_box, text="Recherche", background="#ffffff").grid(row=1, column=0, sticky="w", padx=(0, 6), pady=(8, 0))
        ttk.Entry(filter_box, textvariable=search_var).grid(row=1, column=1, columnspan=5, sticky="ew", padx=(0, 10), pady=(8, 0))

        summary = self.card(body)
        summary.pack(fill="x", pady=(12, 0))
        summary_text = tk.StringVar()
        ttk.Label(summary, textvariable=summary_text, style="CardTitle.TLabel").pack(anchor="w")

        listing = self.card(body)
        listing.pack(fill="both", expand=True, pady=12)
        tree_holder = ttk.Frame(listing, style="Card.TFrame")
        tree_holder.pack(fill="both", expand=True)
        sales_tree = None

        def filtered_sales():
            clauses = ["date(created_at) BETWEEN ? AND ?"]
            params = [start_var.get().strip() or "1900-01-01", end_var.get().strip() or "2999-12-31"]
            if payment_var.get() != "Tous":
                clauses.append("payment_method = ?")
                params.append(payment_var.get())
            if status_var.get() != "Tous":
                clauses.append("status = ?")
                params.append(status_var.get())
            query = search_var.get().strip().lower()
            rows = self.db.all(f"SELECT * FROM sales WHERE {' AND '.join(clauses)} ORDER BY created_at DESC", params)
            if query:
                rows = [row for row in rows if query in row["sale_no"].lower() or query in row["customer_name"].lower()]
            return rows

        def refresh_report(*_):
            nonlocal sales_tree
            for child in tree_holder.winfo_children():
                child.destroy()
            rows = filtered_sales()
            total = sum(float(r["total"] or 0) for r in rows)
            paid = sum(float(r["paid"] or 0) for r in rows)
            credit = total - paid
            stock_value = self.db.one("SELECT COALESCE(SUM(stock * purchase_price), 0) v FROM products WHERE active = 1")["v"]
            summary_text.set(
                f"{len(rows)} vente(s) | Total {self.money(total)} | Encaisse {self.money(paid)} | Credit {self.money(credit)} | Stock achat {self.money(stock_value)}"
            )
            sales_tree = self.table(
                tree_holder,
                ("ID", "Numero", "Facture", "Client", "Total", "Paye", "Paiement", "Statut", "FNE", "Date"),
                [
                    (
                        r["id"],
                        r["sale_no"],
                        r["invoice_no"] or "",
                        r["customer_name"],
                        self.money(r["total"]),
                        self.money(r["paid"]),
                        r["payment_method"],
                        r["status"],
                        r["fne_status"] or "A exporter",
                        r["created_at"][:16],
                    )
                    for r in rows
                ],
            )

        actions = self.card(body)
        actions.pack(fill="x", pady=(0, 4))

        def selected_sale_id():
            if not sales_tree:
                return None
            selected = sales_tree.selection()
            if not selected:
                messagebox.showwarning(APP_NAME, "Selectionnez une vente dans le rapport.")
                return None
            return int(sales_tree.item(selected[0])["values"][0])

        def export_sales_excel():
            path = filedialog.asksaveasfilename(defaultextension=".xls", filetypes=[("Microsoft Excel", "*.xls")], initialfile="ventes-magazin-ci.xls")
            if not path:
                return
            rows = filtered_sales()
            self.export_excel_xml(
                path,
                "Ventes",
                ["Numero", "Facture", "Client", "Total", "Paye", "Paiement", "Statut", "FNE", "Date"],
                [[r["sale_no"], r["invoice_no"] or "", r["customer_name"], r["total"], r["paid"], r["payment_method"], r["status"], r["fne_status"], r["created_at"]] for r in rows],
            )
            messagebox.showinfo(APP_NAME, "Export Excel termine.")

        def export_sales_csv():
            path = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV", "*.csv")], initialfile="ventes-magazin-ci.csv")
            if not path:
                return
            rows = filtered_sales()
            self.export_csv(
                path,
                ["Numero", "Facture", "Client", "Total", "Paye", "Paiement", "Statut", "FNE", "Date"],
                [[r["sale_no"], r["invoice_no"] or "", r["customer_name"], r["total"], r["paid"], r["payment_method"], r["status"], r["fne_status"], r["created_at"]] for r in rows],
            )
            messagebox.showinfo(APP_NAME, "Export CSV termine.")

        def make_invoice():
            sale_id = selected_sale_id()
            if not sale_id:
                return
            path = self.create_professional_invoice(sale_id)
            if path:
                self.show_receipt_window(path)

        def print_invoice():
            sale_id = selected_sale_id()
            if not sale_id:
                return
            path = self.create_professional_invoice(sale_id)
            if path:
                self.print_text_file(path)

        def export_fne():
            sale_id = selected_sale_id()
            if not sale_id:
                return
            path = self.export_sale_fne(sale_id)
            if path:
                refresh_report()
                messagebox.showinfo(APP_NAME, f"Export FNE prepare:\n{path}")

        ttk.Button(actions, text="Filtrer", style="Accent.TButton", command=refresh_report).pack(side="left", padx=(0, 8))
        ttk.Button(actions, text="Exporter les ventes Excel", style="Accent.TButton", command=export_sales_excel).pack(side="left", padx=(0, 8))
        ttk.Button(actions, text="Exporter les ventes CSV", style="Accent.TButton", command=export_sales_csv).pack(side="left", padx=(0, 8))
        ttk.Button(actions, text="Facture pro", command=make_invoice).pack(side="left", padx=(0, 8))
        ttk.Button(actions, text="Imprimer facture", command=print_invoice).pack(side="left", padx=(0, 8))
        ttk.Button(actions, text="Export FNE", command=export_fne).pack(side="left")
        for var in (start_var, end_var, payment_var, status_var, search_var):
            var.trace_add("write", refresh_report)
        refresh_report()

    def show_settings(self):
        self.clear("Parametres", "Identite du commerce, fiscalite, FNE et monnaie.")
        body = self.card(self.main)
        body.pack(fill="x", padx=24, pady=8)
        keys = [
            ("Nom du commerce", "shop_name"),
            ("Telephone", "shop_phone"),
            ("Email", "shop_email"),
            ("Adresse", "shop_address"),
            ("NCC", "shop_ncc"),
            ("RCCM", "shop_rccm"),
            ("Regime fiscal", "shop_tax_regime"),
            ("Monnaie", "currency"),
            ("Pied facture", "invoice_footer"),
            ("Note FNE", "fne_note"),
        ]
        vars_ = {}
        for idx, (label, key) in enumerate(keys):
            vars_[key] = tk.StringVar(value=self.setting(key))
            self.form_row(body, label, idx, vars_[key], 40)

        theme_var = tk.StringVar(value=self.setting("theme") or "Light")
        ttk.Label(body, text="Theme", background="#ffffff").grid(row=len(keys), column=0, sticky="w", pady=5)
        ttk.Combobox(body, textvariable=theme_var, values=("Light", "Dark"), state="readonly", width=38).grid(row=len(keys), column=1, sticky="ew", pady=5, padx=(10, 0))

        def save():
            for key, var in vars_.items():
                self.db.execute("UPDATE settings SET value = ? WHERE key = ?", (var.get().strip(), key))
            self.db.execute("UPDATE settings SET value = ? WHERE key = ?", (theme_var.get().strip(), "theme"))
            self.theme = theme_var.get().strip().lower() or "light"
            self.build_style()
            messagebox.showinfo(APP_NAME, "Parametres enregistres.")
            self.show_settings()

        ttk.Button(body, text="Enregistrer", style="Accent.TButton", command=save).grid(row=len(keys) + 2, column=0, columnspan=2, sticky="ew", pady=(12, 0))


if __name__ == "__main__":
    app = MagazinApp()
    app.mainloop()
