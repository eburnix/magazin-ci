import csv
import hashlib
import html
import json
import os
import platform
import shutil
import sqlite3
import subprocess
import sys
import time
from datetime import datetime, date, timedelta
from pathlib import Path
import tkinter as tk
from tkinter import ttk, messagebox, filedialog, simpledialog

try:
    import win32print
except ImportError:
    win32print = None

try:
    from PIL import Image, ImageTk
except ImportError:
    Image = None
    ImageTk = None

APP_NAME = "Magazin-ci"
APP_VERSION = "1.1.2"
INSTALL_PIN = "05535350"
RENEWAL_PIN = "0707687068"
LICENSE_DAYS = 365
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
LABELS_DIR = DATA_DIR / "etiquettes"
PRODUCT_UPLOADS_DIR = DATA_DIR / "uploads" / "produits"
LOGO_UPLOADS_DIR = DATA_DIR / "uploads" / "logo"
QUOTES_DIR = DATA_DIR / "devis"
DELIVERIES_DIR = DATA_DIR / "livraisons"


def pin_hash(pin):
    return hashlib.sha256(str(pin or "").encode("utf-8")).hexdigest()


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

            CREATE TABLE IF NOT EXISTS users (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT NOT NULL UNIQUE,
                role TEXT NOT NULL DEFAULT 'Vendeur',
                pin_hash TEXT NOT NULL,
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
                invoice_status TEXT NOT NULL DEFAULT 'Active',
                fne_status TEXT NOT NULL DEFAULT 'A exporter',
                fne_reference TEXT,
                fne_exported_at TEXT,
                fne_certified_at TEXT,
                note TEXT,
                user_id INTEGER,
                cashier_name TEXT,
                cash_session_id INTEGER,
                created_at TEXT NOT NULL,
                FOREIGN KEY(customer_id) REFERENCES customers(id),
                FOREIGN KEY(user_id) REFERENCES users(id)
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
                user_id INTEGER,
                cash_session_id INTEGER,
                created_at TEXT NOT NULL
            );

            CREATE TABLE IF NOT EXISTS cash_sessions (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                opened_at TEXT NOT NULL,
                closed_at TEXT,
                opening_amount REAL NOT NULL DEFAULT 0,
                expected_amount REAL NOT NULL DEFAULT 0,
                real_amount REAL,
                gap_amount REAL NOT NULL DEFAULT 0,
                user_id INTEGER,
                cashier_name TEXT,
                note TEXT,
                status TEXT NOT NULL DEFAULT 'Ouverte',
                FOREIGN KEY(user_id) REFERENCES users(id)
            );

            CREATE TABLE IF NOT EXISTS accounting_entries (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                created_at TEXT NOT NULL,
                entry_type TEXT NOT NULL,
                category TEXT NOT NULL,
                label TEXT NOT NULL,
                amount REAL NOT NULL,
                payment_method TEXT NOT NULL DEFAULT 'Especes',
                user_id INTEGER,
                source_type TEXT,
                source_id INTEGER,
                note TEXT,
                FOREIGN KEY(user_id) REFERENCES users(id)
            );

            CREATE TABLE IF NOT EXISTS purchases (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                purchase_no TEXT NOT NULL UNIQUE,
                supplier_id INTEGER,
                supplier_name TEXT NOT NULL,
                total REAL NOT NULL DEFAULT 0,
                paid REAL NOT NULL DEFAULT 0,
                payment_method TEXT NOT NULL DEFAULT 'Especes',
                status TEXT NOT NULL DEFAULT 'Recu',
                user_id INTEGER,
                created_at TEXT NOT NULL,
                FOREIGN KEY(supplier_id) REFERENCES suppliers(id),
                FOREIGN KEY(user_id) REFERENCES users(id)
            );

            CREATE TABLE IF NOT EXISTS purchase_items (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                purchase_id INTEGER NOT NULL,
                product_id INTEGER NOT NULL,
                product_name TEXT NOT NULL,
                quantity REAL NOT NULL,
                unit_cost REAL NOT NULL,
                total REAL NOT NULL,
                FOREIGN KEY(purchase_id) REFERENCES purchases(id) ON DELETE CASCADE,
                FOREIGN KEY(product_id) REFERENCES products(id)
            );

            CREATE TABLE IF NOT EXISTS sale_returns (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                return_no TEXT NOT NULL UNIQUE,
                sale_id INTEGER NOT NULL,
                sale_item_id INTEGER NOT NULL,
                product_id INTEGER NOT NULL,
                product_name TEXT NOT NULL,
                quantity REAL NOT NULL,
                refund_amount REAL NOT NULL DEFAULT 0,
                reason TEXT,
                user_id INTEGER,
                created_at TEXT NOT NULL,
                FOREIGN KEY(sale_id) REFERENCES sales(id),
                FOREIGN KEY(sale_item_id) REFERENCES sale_items(id),
                FOREIGN KEY(product_id) REFERENCES products(id),
                FOREIGN KEY(user_id) REFERENCES users(id)
            );

            CREATE TABLE IF NOT EXISTS quotes (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                quote_no TEXT NOT NULL UNIQUE,
                customer_id INTEGER,
                customer_name TEXT NOT NULL,
                total REAL NOT NULL DEFAULT 0,
                status TEXT NOT NULL DEFAULT 'Brouillon',
                validity_date TEXT,
                note TEXT,
                converted_sale_id INTEGER,
                user_id INTEGER,
                created_at TEXT NOT NULL,
                FOREIGN KEY(customer_id) REFERENCES customers(id),
                FOREIGN KEY(converted_sale_id) REFERENCES sales(id),
                FOREIGN KEY(user_id) REFERENCES users(id)
            );

            CREATE TABLE IF NOT EXISTS quote_items (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                quote_id INTEGER NOT NULL,
                product_id INTEGER NOT NULL,
                product_name TEXT NOT NULL,
                quantity REAL NOT NULL,
                unit_price REAL NOT NULL,
                total REAL NOT NULL,
                FOREIGN KEY(quote_id) REFERENCES quotes(id) ON DELETE CASCADE,
                FOREIGN KEY(product_id) REFERENCES products(id)
            );

            CREATE TABLE IF NOT EXISTS delivery_notes (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                delivery_no TEXT NOT NULL UNIQUE,
                sale_id INTEGER,
                customer_id INTEGER,
                customer_name TEXT NOT NULL,
                status TEXT NOT NULL DEFAULT 'Prepare',
                note TEXT,
                delivered_at TEXT,
                user_id INTEGER,
                created_at TEXT NOT NULL,
                FOREIGN KEY(sale_id) REFERENCES sales(id),
                FOREIGN KEY(customer_id) REFERENCES customers(id),
                FOREIGN KEY(user_id) REFERENCES users(id)
            );

            CREATE TABLE IF NOT EXISTS delivery_note_items (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                delivery_id INTEGER NOT NULL,
                product_id INTEGER NOT NULL,
                product_name TEXT NOT NULL,
                quantity REAL NOT NULL,
                unit TEXT,
                FOREIGN KEY(delivery_id) REFERENCES delivery_notes(id) ON DELETE CASCADE,
                FOREIGN KEY(product_id) REFERENCES products(id)
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
        self.ensure_column("products", "part_number", "TEXT")
        self.ensure_column("products", "serial_number", "TEXT")
        self.ensure_column("products", "imei", "TEXT")
        self.ensure_column("products", "model", "TEXT")
        self.ensure_column("products", "color", "TEXT")
        self.ensure_column("products", "size", "TEXT")
        self.ensure_column("products", "image_path", "TEXT")
        self.ensure_column("sales", "invoice_no", "TEXT")
        self.ensure_column("sales", "invoice_status", "TEXT NOT NULL DEFAULT 'Active'")
        self.ensure_column("sales", "fne_status", "TEXT NOT NULL DEFAULT 'A exporter'")
        self.ensure_column("sales", "fne_reference", "TEXT")
        self.ensure_column("sales", "fne_exported_at", "TEXT")
        self.ensure_column("sales", "fne_certified_at", "TEXT")
        self.ensure_column("sales", "note", "TEXT")
        self.ensure_column("sales", "user_id", "INTEGER")
        self.ensure_column("sales", "cashier_name", "TEXT")
        self.ensure_column("sales", "cash_session_id", "INTEGER")
        self.ensure_column("cash_entries", "user_id", "INTEGER")
        self.ensure_column("cash_entries", "cash_session_id", "INTEGER")
        indexes = [
            "CREATE INDEX IF NOT EXISTS idx_users_role ON users(role)",
            "CREATE INDEX IF NOT EXISTS idx_products_supplier ON products(supplier_id)",
            "CREATE INDEX IF NOT EXISTS idx_products_barcode ON products(barcode)",
            "CREATE INDEX IF NOT EXISTS idx_sales_created_at ON sales(created_at)",
            "CREATE INDEX IF NOT EXISTS idx_sales_customer ON sales(customer_id)",
            "CREATE INDEX IF NOT EXISTS idx_sales_fne_status ON sales(fne_status)",
            "CREATE INDEX IF NOT EXISTS idx_sales_invoice_status ON sales(invoice_status)",
            "CREATE INDEX IF NOT EXISTS idx_sale_items_sale ON sale_items(sale_id)",
            "CREATE INDEX IF NOT EXISTS idx_sale_items_product ON sale_items(product_id)",
            "CREATE INDEX IF NOT EXISTS idx_stock_movements_product ON stock_movements(product_id)",
            "CREATE INDEX IF NOT EXISTS idx_cash_entries_date ON cash_entries(created_at)",
            "CREATE INDEX IF NOT EXISTS idx_cash_sessions_status ON cash_sessions(status)",
            "CREATE INDEX IF NOT EXISTS idx_accounting_entries_date ON accounting_entries(created_at)",
            "CREATE INDEX IF NOT EXISTS idx_accounting_entries_source ON accounting_entries(source_type, source_id)",
            "CREATE INDEX IF NOT EXISTS idx_purchases_date ON purchases(created_at)",
            "CREATE INDEX IF NOT EXISTS idx_purchase_items_product ON purchase_items(product_id)",
            "CREATE INDEX IF NOT EXISTS idx_sale_returns_date ON sale_returns(created_at)",
            "CREATE INDEX IF NOT EXISTS idx_quotes_date ON quotes(created_at)",
            "CREATE INDEX IF NOT EXISTS idx_quotes_customer ON quotes(customer_id)",
            "CREATE INDEX IF NOT EXISTS idx_quote_items_quote ON quote_items(quote_id)",
            "CREATE INDEX IF NOT EXISTS idx_delivery_notes_date ON delivery_notes(created_at)",
            "CREATE INDEX IF NOT EXISTS idx_delivery_notes_sale ON delivery_notes(sale_id)",
            "CREATE INDEX IF NOT EXISTS idx_delivery_items_delivery ON delivery_note_items(delivery_id)",
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
            "shop_logo_path": "",
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
        now = datetime.now().isoformat(timespec="seconds")
        users = [
            ("Admin", "Admin", "0714"),
            ("Vendeur", "Vendeur", "1234"),
        ]
        for name, role, pin in users:
            self.conn.execute(
                """
                INSERT OR IGNORE INTO users(name, role, pin_hash, active, created_at)
                VALUES (?, ?, ?, 1, ?)
                """,
                (name, role, pin_hash(pin), now),
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
        self.splash = None
        self.db = Database(DB_PATH)
        self.logo_image = self.load_logo_image(max_width=180, max_height=120)
        self.splash_logo_image = self.load_project_logo_image(max_width=180, max_height=120)
        self.apply_app_icon()
        self.theme = self.setting("theme").strip().lower() or "light"
        self.build_style()
        self.current_user = None
        if not self.ensure_installation_activated():
            self.destroy()
            return
        if not self.prompt_access_pin():
            self.destroy()
            return
        self.splash = SplashScreen(self, self.splash_logo_image)
        self.splash.set_progress(28, "Acces autorise...")
        self.splash.set_progress(42, "Preparation de la base de donnees...")
        self.splash.set_progress(70, "Chargement de l'interface...")
        self.title(f"{APP_NAME} - Gestion commerciale")
        self.geometry("1180x720")
        self.minsize(1060, 640)
        self.rowconfigure(0, weight=1)
        self.columnconfigure(0, weight=1)
        self.cart = []
        self.product_image_cache = {}
        self.current_view = None
        self.build_style()
        self.build_shell()
        self.show_dashboard()
        self.splash.set_progress(100, "Pret.")
        self.after(250, self.splash.destroy)
        self.deiconify()

    def load_logo_image(self, max_width=180, max_height=120):
        logo_path = self.logo_path()
        return self.load_image_from_path(logo_path, max_width, max_height)

    def load_project_logo_image(self, max_width=180, max_height=120):
        return self.load_image_from_path(self.project_logo_path(), max_width, max_height)

    def load_image_from_path(self, logo_path, max_width=180, max_height=120):
        if not logo_path:
            return None
        try:
            if Image is not None and ImageTk is not None:
                image = Image.open(logo_path).convert("RGBA")
                image.thumbnail((max_width, max_height))
                return ImageTk.PhotoImage(image)
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

    def ensure_installation_activated(self):
        if self.setting("installation_activated") == "1":
            return self.ensure_license_valid()
        title = f"Installation {APP_NAME}"
        for attempt in range(3):
            pin = simpledialog.askstring(
                title,
                "Entrez le PIN d'installation pour activer ce poste:",
                show="*",
                parent=self,
            )
            if pin is None:
                return False
            if pin == INSTALL_PIN:
                now = datetime.now()
                expires_at = now + timedelta(days=LICENSE_DAYS)
                self.db.execute(
                    "INSERT OR REPLACE INTO settings(key, value) VALUES ('installation_activated', '1')"
                )
                self.db.execute(
                    "INSERT OR REPLACE INTO settings(key, value) VALUES ('installation_activated_at', ?)",
                    (now.isoformat(timespec="seconds"),),
                )
                self.db.execute(
                    "INSERT OR REPLACE INTO settings(key, value) VALUES ('license_expires_at', ?)",
                    (expires_at.isoformat(timespec="seconds"),),
                )
                messagebox.showinfo(
                    title,
                    f"Installation activee avec succes.\nLicence valable jusqu'au {expires_at.strftime('%d/%m/%Y')}.",
                )
                return True
            remaining = 2 - attempt
            messagebox.showerror(title, f"PIN d'installation incorrect. Tentatives restantes: {remaining}")
        messagebox.showerror(title, "Activation refusee. Fermeture de l'application.")
        return False
        if self.logo_image:
            try:
                self.iconphoto(True, self.logo_image)
            except Exception:
                pass

    def ensure_license_valid(self):
        raw_expiry = self.setting("license_expires_at")
        now = datetime.now()
        if not raw_expiry:
            expires_at = now + timedelta(days=LICENSE_DAYS)
            self.db.execute(
                "INSERT OR REPLACE INTO settings(key, value) VALUES ('license_expires_at', ?)",
                (expires_at.isoformat(timespec="seconds"),),
            )
            return True
        try:
            expires_at = datetime.fromisoformat(raw_expiry)
        except ValueError:
            expires_at = now - timedelta(seconds=1)
        if now < expires_at:
            return True

        title = f"Licence {APP_NAME}"
        for attempt in range(3):
            pin = simpledialog.askstring(
                title,
                "Licence expiree.\nEntrez le PIN de reactivation pour prolonger 12 mois:",
                show="*",
                parent=self,
            )
            if pin is None:
                return False
            if pin == RENEWAL_PIN:
                new_expiry = now + timedelta(days=LICENSE_DAYS)
                self.db.execute(
                    "INSERT OR REPLACE INTO settings(key, value) VALUES ('license_expires_at', ?)",
                    (new_expiry.isoformat(timespec="seconds"),),
                )
                self.db.execute(
                    "INSERT OR REPLACE INTO settings(key, value) VALUES ('license_renewed_at', ?)",
                    (now.isoformat(timespec="seconds"),),
                )
                messagebox.showinfo(
                    title,
                    f"Licence reactivee avec succes.\nNouvelle validite jusqu'au {new_expiry.strftime('%d/%m/%Y')}.",
                )
                return True
            remaining = 2 - attempt
            messagebox.showerror(title, f"PIN de reactivation incorrect. Tentatives restantes: {remaining}")
        messagebox.showerror(title, "Licence expiree. Fermeture de l'application.")
        return False

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
            user = self.authenticate_user(pin_var.get())
            if user:
                self.current_user = user
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

    def authenticate_user(self, pin):
        return self.db.one(
            """
            SELECT * FROM users
            WHERE active = 1 AND pin_hash = ?
            ORDER BY CASE role WHEN 'Admin' THEN 0 ELSE 1 END, name
            LIMIT 1
            """,
            (pin_hash(pin),),
        )

    def user_role(self):
        return self.current_user["role"] if self.current_user else ""

    def is_admin(self):
        return self.user_role() == "Admin"

    def require_admin(self, action="Action reservee a l'administrateur"):
        if self.is_admin():
            return True
        pin = simpledialog.askstring(APP_NAME, f"{action}\n\nPIN admin:", show="*")
        user = self.db.one(
            "SELECT * FROM users WHERE active = 1 AND role = 'Admin' AND pin_hash = ?",
            (pin_hash(pin),),
        )
        if user:
            return True
        messagebox.showerror(APP_NAME, "Autorisation admin refusee.")
        return False

    def role_allowed_pages(self):
        if self.is_admin():
            return None
        if self.user_role() == "Comptable":
            return {"Dashboard", "Caisse", "Comptabilite", "Devis", "Factures/FNE", "Rapports"}
        if self.user_role() == "Magasinier":
            return {"Dashboard", "Produits", "Fournisseurs", "Achats", "Retours", "Stock", "Inventaire", "Bons livraison"}
        return {"Dashboard", "POS tactile", "Clients", "Devis", "Bons livraison", "Caisse", "Factures/FNE", "Rapports"}

    def build_style(self):
        style = ttk.Style(self)
        style.theme_use("clam")
        dark = self.theme == "dark"
        bg = "#0b1220" if dark else "#f3f6fb"
        surface = "#111827" if dark else "#ffffff"
        surface2 = "#1f2937" if dark else "#f8fafc"
        sidebar = "#08111f" if dark else "#111827"
        sidebar2 = "#0f1b2d" if dark else "#1f2937"
        fg = "#f8fafc" if dark else "#172033"
        muted = "#9ca3af" if dark else "#64748b"
        card_border = "#334155" if dark else "#dbe3ef"
        accent = "#10b981" if dark else "#0f766e"
        accent_active = "#059669" if dark else "#115e59"
        blue = "#38bdf8" if dark else "#2563eb"
        amber = "#fbbf24" if dark else "#d97706"
        danger = "#ef4444" if dark else "#dc2626"
        self.colors = {
            "bg": bg,
            "surface": surface,
            "surface2": surface2,
            "sidebar": sidebar,
            "sidebar2": sidebar2,
            "fg": fg,
            "muted": muted,
            "border": card_border,
            "accent": accent,
            "accent_active": accent_active,
            "blue": blue,
            "amber": amber,
            "danger": danger,
        }

        self.configure(bg=bg)
        style.configure(".", font=("Segoe UI", 10))
        style.configure("TFrame", background=bg)
        style.configure("Card.TFrame", background=surface, relief="solid", borderwidth=1)
        style.configure("Header.TFrame", background=bg)
        style.configure("Sidebar.TFrame", background=sidebar)
        style.configure("Sidebar.TButton", background=sidebar, foreground="#e5e7eb", padding=(14, 11), anchor="w", font=("Segoe UI", 10, "bold"), relief="flat")
        style.map("Sidebar.TButton", background=[("active", sidebar2), ("pressed", accent)], foreground=[("active", "#ffffff"), ("pressed", "#ffffff")])
        style.configure("Title.TLabel", background=bg, foreground=fg, font=("Segoe UI", 24, "bold"))
        style.configure("Sub.TLabel", background=bg, foreground=muted, font=("Segoe UI", 10))
        style.configure("Meta.TLabel", background=bg, foreground=muted, font=("Segoe UI", 9, "bold"))
        style.configure("CardTitle.TLabel", background=surface, foreground=fg, font=("Segoe UI", 12, "bold"))
        style.configure("MutedCard.TLabel", background=surface, foreground=muted, font=("Segoe UI", 9))
        style.configure("Metric.TLabel", background=surface, foreground=accent, font=("Segoe UI", 21, "bold"))
        style.configure("TButton", background="#172033" if not dark else "#1f2937", foreground="#f8fafc", font=("Segoe UI", 10, "bold"), padding=(10, 8), relief="flat")
        style.map("TButton", background=[("active", "#26344d" if not dark else "#334155")])
        style.configure("Accent.TButton", background=accent, foreground="#ffffff", relief="flat")
        style.map("Accent.TButton", background=[("active", accent_active), ("pressed", "#064e3b")])
        style.configure("Secondary.TButton", background=blue, foreground="#ffffff", relief="flat")
        style.map("Secondary.TButton", background=[("active", "#1d4ed8")])
        style.configure("Touch.TButton", font=("Segoe UI", 12, "bold"), padding=16, relief="flat")
        style.configure("ProductTile.TButton", background=surface2, foreground=fg, anchor="center", justify="center", padding=(16, 18), font=("Segoe UI", 10, "bold"), relief="solid", borderwidth=1, wraplength=180)
        style.map("ProductTile.TButton", background=[("active", "#e0f2fe" if not dark else "#334155"), ("pressed", "#ccfbf1" if not dark else "#475569")], foreground=[("active", fg)])
        style.configure("Danger.TButton", background=danger, foreground="#ffffff", relief="flat")
        style.map("Danger.TButton", background=[("active", "#991b1b")])
        style.configure("TEntry", padding=7, fieldbackground=surface, foreground=fg, bordercolor=card_border, lightcolor=card_border, darkcolor=card_border)
        style.configure("TCombobox", padding=6, fieldbackground=surface, foreground=fg)
        style.configure("TLabel", background=bg, foreground=fg)
        style.configure("Treeview", rowheight=32, font=("Segoe UI", 9), fieldbackground=surface, background=surface, foreground=fg, bordercolor=card_border, lightcolor=card_border, darkcolor=card_border)
        style.configure("Treeview.Heading", font=("Segoe UI", 9, "bold"), background="#e5edf7" if not dark else "#253244", foreground=fg, padding=(8, 8), relief="flat")
        style.map("Treeview", background=[("selected", accent)], foreground=[("selected", "#ffffff")])

    def build_shell(self):
        colors = getattr(self, "colors", {})
        sidebar_color = colors.get("sidebar", "#111827")
        sidebar2 = colors.get("sidebar2", "#1f2937")
        accent = colors.get("accent", "#0f766e")
        self.sidebar = ttk.Frame(self, style="Sidebar.TFrame", width=264)
        self.sidebar.pack(side="left", fill="y")
        self.sidebar.pack_propagate(False)

        if self.logo_image:
            logo_wrap = tk.Frame(self.sidebar, bg=sidebar_color)
            logo_wrap.pack(fill="x", pady=(18, 6), padx=18)
            logo_label = tk.Label(logo_wrap, image=self.logo_image, bg=sidebar_color)
            logo_label.pack(anchor="w")

        brand_box = tk.Frame(self.sidebar, bg=sidebar2, padx=14, pady=12)
        brand_box.pack(fill="x", padx=14, pady=(6, 14))
        tk.Label(
            brand_box,
            text=APP_NAME,
            bg=sidebar2,
            fg="#ffffff",
            font=("Segoe UI", 16, "bold"),
            anchor="w",
        ).pack(fill="x")
        tk.Label(
            brand_box,
            text=f"{self.current_user['name']} - {self.current_user['role']}",
            bg=sidebar2,
            fg="#cbd5e1",
            font=("Segoe UI", 9, "bold"),
            anchor="w",
        ).pack(fill="x", pady=(4, 0))
        tk.Frame(brand_box, bg=accent, height=3).pack(fill="x", pady=(10, 0))

        tk.Label(
            self.sidebar,
            text="NAVIGATION",
            bg=sidebar_color,
            fg="#94a3b8",
            font=("Segoe UI", 8, "bold"),
            anchor="w",
        ).pack(fill="x", padx=18, pady=(0, 6))

        buttons = [
            ("Dashboard", self.show_dashboard),
            ("POS tactile", self.show_sales),
            ("Produits", self.show_products),
            ("Clients", self.show_customers),
            ("Fournisseurs", self.show_suppliers),
            ("Achats", self.show_purchases),
            ("Retours", self.show_returns),
            ("Stock", self.show_stock),
            ("Inventaire", self.show_inventory),
            ("Caisse", self.show_cash),
            ("Comptabilite", self.show_accounting),
            ("Devis", self.show_quotes),
            ("Bons livraison", self.show_delivery_notes),
            ("Factures/FNE", self.show_invoices),
            ("Statistiques", self.show_statistics),
            ("Rapports", self.show_reports),
            ("Parametres", self.show_settings),
        ]
        allowed_pages = self.role_allowed_pages()
        for text, command in buttons:
            if allowed_pages is not None and text not in allowed_pages:
                continue
            ttk.Button(self.sidebar, text=text, command=command, style="Sidebar.TButton").pack(fill="x", padx=10, pady=3)

        ttk.Button(self.sidebar, text="Deconnexion", command=self.logout, style="Danger.TButton").pack(fill="x", padx=10, pady=(14, 3))

        footer = tk.Label(
            self.sidebar,
            text=DEVELOPER_CREDIT,
            bg=sidebar_color,
            fg="#94a3b8",
            font=("Segoe UI", 8),
            justify="left",
            wraplength=220,
        )
        footer.pack(side="bottom", fill="x", padx=18, pady=16)

        self.main = ttk.Frame(self)
        self.main.pack(side="left", fill="both", expand=True)

    def logout(self):
        if not messagebox.askyesno(APP_NAME, "Se deconnecter de la session actuelle ?"):
            return
        self.unbind_all("<Key>")
        self.withdraw()
        if self.sidebar.winfo_exists():
            self.sidebar.destroy()
        if self.main.winfo_exists():
            self.main.destroy()
        self.cart = []
        self.product_image_cache = {}
        self.current_user = None
        if not self.prompt_access_pin():
            self.destroy()
            return
        self.build_shell()
        self.show_dashboard()
        self.deiconify()

    def clear(self, title, subtitle=""):
        for child in self.main.winfo_children():
            child.destroy()
        header = ttk.Frame(self.main, style="Header.TFrame")
        header.pack(fill="x", padx=28, pady=(24, 14))
        header.columnconfigure(0, weight=1)
        ttk.Label(header, text=title, style="Title.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Label(header, text=subtitle, style="Sub.TLabel").grid(row=1, column=0, sticky="w", pady=(3, 0))
        meta = f"{date.today().strftime('%d/%m/%Y')}  |  {self.current_user['name']}  |  {self.current_user['role']}"
        ttk.Label(header, text=meta, style="Meta.TLabel").grid(row=0, column=1, rowspan=2, sticky="e")
        self.current_view = title

    def card(self, parent, padx=14, pady=12):
        frame = ttk.Frame(parent, style="Card.TFrame", padding=(padx + 2, pady + 2))
        return frame

    def setting(self, key):
        row = self.db.one("SELECT value FROM settings WHERE key = ?", (key,))
        return row["value"] if row else ""

    def money(self, amount):
        return f"{amount:,.0f} {self.setting('currency')}".replace(",", " ")

    def merchant_lines(self, include_tax=True):
        lines = [
            self.setting("shop_name"),
            f"Tel: {self.setting('shop_phone')} | Email: {self.setting('shop_email')}",
            self.setting("shop_address"),
        ]
        if include_tax:
            lines.extend(
                [
                    f"NCC: {self.setting('shop_ncc') or '-'} | RCCM: {self.setting('shop_rccm') or '-'}",
                    f"Regime fiscal: {self.setting('shop_tax_regime') or '-'}",
                ]
            )
        return [line for line in lines if str(line or "").strip()]

    def logo_path(self):
        custom_path = self.setting("shop_logo_path") if hasattr(self, "db") else ""
        if custom_path:
            path = Path(custom_path)
            if not path.is_absolute():
                path = BASE_DIR / path
            if path.exists():
                return path
        return self.project_logo_path()

    def project_logo_path(self):
        path = BASE_DIR / "LOGO_.png"
        if not path.exists():
            path = RESOURCE_DIR / "LOGO_.png"
        return path if path.exists() else None

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

    def open_document(self, path):
        try:
            if platform.system() == "Windows":
                os.startfile(path)
            elif platform.system() == "Darwin":
                subprocess.run(["open", str(path)], check=False)
            else:
                subprocess.run(["xdg-open", str(path)], check=False)
        except Exception as exc:
            messagebox.showwarning(APP_NAME, f"Document genere, mais ouverture automatique impossible:\n{path}\n\n{exc}")

    def generate_product_labels_pdf(self, products, copies_per_product=1, filename_prefix="etiquettes-produits"):
        try:
            from reportlab.graphics.barcode import code128
            from reportlab.lib import colors
            from reportlab.lib.pagesizes import A4
            from reportlab.lib.units import mm
            from reportlab.pdfgen import canvas
        except ImportError:
            messagebox.showerror(APP_NAME, "Le module reportlab est requis pour generer les etiquettes PDF.")
            return None

        LABELS_DIR.mkdir(parents=True, exist_ok=True)
        path = LABELS_DIR / f"{filename_prefix}-{datetime.now().strftime('%Y%m%d-%H%M%S')}.pdf"
        page_width, page_height = A4
        margin_x = 10 * mm
        margin_y = 11 * mm
        gap_x = 3 * mm
        gap_y = 2 * mm
        columns = 3
        rows = 8
        label_width = (page_width - 2 * margin_x - (columns - 1) * gap_x) / columns
        label_height = (page_height - 2 * margin_y - (rows - 1) * gap_y) / rows
        doc = canvas.Canvas(str(path), pagesize=A4)
        labels = []
        for product in products:
            for _ in range(max(1, int(copies_per_product))):
                labels.append(product)
        if not labels:
            messagebox.showwarning(APP_NAME, "Aucun produit actif a imprimer.")
            return None

        for index, product in enumerate(labels):
            slot = index % (columns * rows)
            if index and slot == 0:
                doc.showPage()
            col = slot % columns
            row = slot // columns
            x = margin_x + col * (label_width + gap_x)
            y = page_height - margin_y - (row + 1) * label_height - row * gap_y
            doc.setStrokeColor(colors.HexColor("#d1d5db"))
            doc.roundRect(x, y, label_width, label_height, 4, stroke=1, fill=0)
            doc.setFillColor(colors.HexColor("#111827"))
            doc.setFont("Helvetica-Bold", 8)
            doc.drawString(x + 4 * mm, y + label_height - 6 * mm, str(product["name"])[:28])
            doc.setFont("Helvetica", 6.8)
            doc.setFillColor(colors.HexColor("#4b5563"))
            doc.drawString(x + 4 * mm, y + label_height - 10 * mm, f"{product['category']} | {product['unit']}")
            doc.setFillColor(colors.HexColor("#047857"))
            doc.setFont("Helvetica-Bold", 10)
            doc.drawRightString(x + label_width - 4 * mm, y + label_height - 15 * mm, self.money(product["sale_price"]))
            barcode_value = (product["barcode"] or f"PRD{int(product['id']):06d}").strip()
            try:
                barcode = code128.Code128(barcode_value, barHeight=9 * mm, barWidth=0.38 * mm)
                barcode.drawOn(doc, x + 4 * mm, y + 8 * mm)
            except Exception:
                doc.setFont("Helvetica", 7)
                doc.setFillColor(colors.HexColor("#111827"))
                doc.drawString(x + 4 * mm, y + 12 * mm, barcode_value[:28])
            doc.setFont("Helvetica", 6)
            doc.setFillColor(colors.HexColor("#374151"))
            doc.drawString(x + 4 * mm, y + 4 * mm, f"Code: {barcode_value[:30]}")
            doc.drawRightString(x + label_width - 4 * mm, y + 4 * mm, APP_NAME)

        doc.save()
        return path

    def resolve_product_image_path(self, image_path):
        if not image_path:
            return None
        path = Path(image_path)
        if not path.is_absolute():
            path = BASE_DIR / path
        return path if path.exists() else None

    def product_thumbnail(self, product, size=(96, 65)):
        if Image is None or ImageTk is None:
            return None
        image_path = self.resolve_product_image_path(product["image_path"] if "image_path" in product.keys() else "")
        if not image_path:
            return None
        cache_key = (str(image_path), size)
        if cache_key in self.product_image_cache:
            return self.product_image_cache[cache_key]
        try:
            image = Image.open(image_path).convert("RGB")
            image.thumbnail(size)
            canvas_image = Image.new("RGB", size, "#f8fafc")
            x = (size[0] - image.width) // 2
            y = (size[1] - image.height) // 2
            canvas_image.paste(image, (x, y))
            photo = ImageTk.PhotoImage(canvas_image)
        except Exception:
            return None
        self.product_image_cache[cache_key] = photo
        return photo

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
        colors = getattr(self, "colors", {})
        for col in columns:
            tree.heading(col, text=col)
            width = 90 if col in {"ID", "Qte", "Stock", "Seuil"} else 135
            if col in {"Produit", "Client", "Libelle", "Designation"}:
                width = 190
            tree.column(col, width=width, anchor="w", stretch=True)
        tree.tag_configure("even", background=colors.get("surface", "#ffffff"))
        tree.tag_configure("odd", background=colors.get("surface2", "#f8fafc"))
        for index, row in enumerate(rows):
            tree.insert("", "end", values=row, tags=("odd" if index % 2 else "even",))
        tree.pack(fill="both", expand=True)
        return tree

    def form_row(self, parent, label, row, var, width=28):
        ttk.Label(parent, text=label, background=getattr(self, "colors", {}).get("surface", "#ffffff")).grid(row=row, column=0, sticky="w", pady=5)
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
        vars_ = {
            k: tk.StringVar()
            for k in [
                "name",
                "category",
                "barcode",
                "part_number",
                "serial_number",
                "imei",
                "model",
                "color",
                "size",
                "unit",
                "purchase",
                "sale",
                "stock",
                "alert",
                "image_path",
            ]
        }
        defaults = {"category": "General", "unit": "piece", "purchase": "0", "sale": "0", "stock": "0", "alert": "0"}
        for key, value in defaults.items():
            vars_[key].set(value)

        labels = [
            ("Nom", "name"),
            ("Categorie", "category"),
            ("Code barre", "barcode"),
            ("Part Number", "part_number"),
            ("S/N", "serial_number"),
            ("IMEI", "imei"),
            ("Modele", "model"),
            ("Couleur", "color"),
            ("Taille", "size"),
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
            for key in ("part_number", "serial_number", "imei", "model", "color", "size", "image_path"):
                vars_[key].set("")

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
            vars_["part_number"].set(product["part_number"] or "")
            vars_["serial_number"].set(product["serial_number"] or "")
            vars_["imei"].set(product["imei"] or "")
            vars_["model"].set(product["model"] or "")
            vars_["color"].set(product["color"] or "")
            vars_["size"].set(product["size"] or "")
            vars_["image_path"].set(product["image_path"] or "")
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
                    vars_["part_number"].get().strip(),
                    vars_["serial_number"].get().strip(),
                    vars_["imei"].get().strip(),
                    vars_["model"].get().strip(),
                    vars_["color"].get().strip(),
                    vars_["size"].get().strip(),
                    vars_["image_path"].get().strip(),
                )
            except ValueError:
                messagebox.showerror(APP_NAME, "Les prix et stocks doivent etre des nombres.")
                return
            values = (
                values[0],
                values[1],
                values[2] or f"PRD{datetime.now().strftime('%Y%m%d%H%M%S')}",
                *values[3:],
            )
            if selected_product_id.get():
                self.db.execute(
                    """
                    UPDATE products
                    SET name = ?, category = ?, barcode = ?, unit = ?, purchase_price = ?, sale_price = ?, stock = ?, alert_stock = ?,
                        part_number = ?, serial_number = ?, imei = ?, model = ?, color = ?, size = ?, image_path = ?
                    WHERE id = ?
                    """,
                    (*values, selected_product_id.get()),
                )
            else:
                self.db.execute(
                    """
                    INSERT INTO products(
                        name, category, barcode, unit, purchase_price, sale_price, stock, alert_stock,
                        part_number, serial_number, imei, model, color, size, image_path, created_at
                    )
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """,
                    (*values, datetime.now().isoformat(timespec="seconds")),
                )
            self.show_products()

        def upload_product_image():
            path = filedialog.askopenfilename(
                title="Choisir une image produit",
                filetypes=[
                    ("Images", "*.png *.jpg *.jpeg *.webp *.gif"),
                    ("Tous les fichiers", "*.*"),
                ],
            )
            if not path:
                return
            source = Path(path)
            PRODUCT_UPLOADS_DIR.mkdir(parents=True, exist_ok=True)
            safe_stem = "".join(ch if ch.isalnum() or ch in {"-", "_"} else "_" for ch in (vars_["name"].get().strip() or source.stem))[:40]
            filename = f"{safe_stem}-{datetime.now().strftime('%Y%m%d%H%M%S')}{source.suffix.lower()}"
            destination = PRODUCT_UPLOADS_DIR / filename
            try:
                shutil.copy2(source, destination)
            except OSError as exc:
                messagebox.showerror(APP_NAME, f"Upload image impossible: {exc}")
                return
            vars_["image_path"].set(str(destination.relative_to(BASE_DIR)))
            messagebox.showinfo(APP_NAME, "Image produit ajoutee.")

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

        def print_selected_label():
            if not selected_product_id.get():
                messagebox.showwarning(APP_NAME, "Selectionnez un produit pour generer son etiquette.")
                return
            copies = simpledialog.askinteger(APP_NAME, "Nombre d'etiquettes pour ce produit :", initialvalue=1, minvalue=1, maxvalue=240)
            if not copies:
                return
            product = self.db.one("SELECT * FROM products WHERE id = ?", (selected_product_id.get(),))
            path = self.generate_product_labels_pdf([product], copies_per_product=copies, filename_prefix="etiquette-produit")
            if path:
                messagebox.showinfo(APP_NAME, f"Etiquette generee:\n{path}")
                self.open_document(path)

        def print_batch_labels():
            copies = simpledialog.askinteger(APP_NAME, "Nombre d'etiquettes par produit actif :", initialvalue=1, minvalue=1, maxvalue=50)
            if not copies:
                return
            products = self.db.all("SELECT * FROM products WHERE active = 1 ORDER BY category, name")
            path = self.generate_product_labels_pdf(products, copies_per_product=copies, filename_prefix="lot-etiquettes-a4")
            if path:
                messagebox.showinfo(APP_NAME, f"Planche A4 generee:\n{path}")
                self.open_document(path)

        image_row = len(labels) + 1
        ttk.Button(form, text="Uploader image", command=upload_product_image).grid(row=image_row, column=0, columnspan=2, sticky="ew", pady=(10, 0))
        ttk.Label(form, textvariable=vars_["image_path"], style="MutedCard.TLabel", wraplength=260).grid(row=image_row + 1, column=0, columnspan=2, sticky="w", pady=(4, 0))
        ttk.Button(form, text="Enregistrer", style="Accent.TButton", command=save_product).grid(row=image_row + 2, column=0, columnspan=2, sticky="ew", pady=(12, 0))
        ttk.Button(form, text="Supprimer", style="Danger.TButton", command=delete_product).grid(row=image_row + 3, column=0, columnspan=2, sticky="ew", pady=(6, 0))
        ttk.Button(form, text="Nouveau", command=clear_selection).grid(row=image_row + 4, column=0, columnspan=2, sticky="ew", pady=(6, 0))
        ttk.Button(form, text="Etiquette produit", command=print_selected_label).grid(row=image_row + 5, column=0, columnspan=2, sticky="ew", pady=(12, 0))
        ttk.Button(form, text="Lot etiquettes A4", command=print_batch_labels).grid(row=image_row + 6, column=0, columnspan=2, sticky="ew", pady=(6, 0))

        listing = self.card(body)
        listing.grid(row=0, column=1, sticky="nsew", pady=0)
        listing.columnconfigure(0, weight=1)
        ttk.Label(listing, text="Catalogue", style="CardTitle.TLabel").pack(anchor="w", pady=(0, 8))
        product_tree = self.table(
            listing,
            ("ID", "Produit", "Part Number", "Modele", "Couleur", "Taille", "Achat", "Vente", "Stock"),
            [
                (r["id"], r["name"], r["part_number"] or "", r["model"] or "", r["color"] or "", r["size"] or "", self.money(r["purchase_price"]), self.money(r["sale_price"]), r["stock"])
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
        barcode_scan = tk.StringVar()
        scanner_status = tk.StringVar(value="Lecteur code-barres pret")
        scanner_buffer = {"value": "", "last": 0.0}
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
        toolbar.columnconfigure(3, weight=1)
        ttk.Label(toolbar, text="Recherche", background="#ffffff").grid(row=0, column=0, sticky="w")
        search_entry = ttk.Entry(toolbar, textvariable=search)
        search_entry.grid(row=0, column=1, sticky="ew", padx=8)
        ttk.Label(toolbar, text="Scan", background="#ffffff").grid(row=0, column=2, sticky="w")
        scan_entry = ttk.Entry(toolbar, textvariable=barcode_scan)
        scan_entry.grid(row=0, column=3, sticky="ew", padx=8)
        categories = ["Tous"] + [r["category"] for r in self.db.all("SELECT DISTINCT category FROM products WHERE active = 1 ORDER BY category")]
        ttk.Combobox(toolbar, textvariable=category, values=categories, width=18, state="readonly").grid(row=0, column=4, padx=(0, 8))
        ttk.Button(toolbar, text="Actualiser", command=self.show_sales).grid(row=0, column=5)
        ttk.Label(toolbar, textvariable=scanner_status, style="MutedCard.TLabel").grid(row=1, column=0, columnspan=6, sticky="w", pady=(6, 0))

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

        def add_scanned_product(_event=None, code=None):
            code = (code if code is not None else barcode_scan.get()).strip()
            if not code:
                return
            product = self.db.one(
                """
                SELECT * FROM products
                WHERE active = 1
                  AND (barcode = ? OR part_number = ? OR serial_number = ? OR imei = ?)
                """,
                (code, code, code, code),
            )
            if not product:
                scanner_status.set(f"Code introuvable: {code}")
                self.bell()
                barcode_scan.set("")
                scan_entry.focus_set()
                return
            add_product(product)
            scanner_status.set(f"Ajoute: {product['name']}")
            barcode_scan.set("")
            scan_entry.focus_set()

        def scanner_key_handler(event):
            if self.current_view != "POS tactile":
                return
            widget = self.focus_get()
            if widget in {search_entry, scan_entry}:
                return
            now = time.monotonic()
            if now - scanner_buffer["last"] > 0.12:
                scanner_buffer["value"] = ""
            scanner_buffer["last"] = now
            if event.keysym in {"Return", "KP_Enter", "Tab"}:
                code = scanner_buffer["value"].strip()
                scanner_buffer["value"] = ""
                if len(code) >= 3:
                    barcode_scan.set(code)
                    add_scanned_product(code=code)
                    return "break"
                return
            char = event.char or ""
            if char and char.isprintable():
                scanner_buffer["value"] += char
                if len(scanner_buffer["value"]) > 80:
                    scanner_buffer["value"] = scanner_buffer["value"][-80:]
                return "break"

        self.unbind_all("<Key>")
        self.bind_all("<Key>", scanner_key_handler)

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
                tile = tk.Frame(
                    tiles,
                    bg=self.colors["surface2"],
                    highlightthickness=1,
                    highlightbackground=self.colors["border"],
                    width=125,
                    height=135,
                    padx=6,
                    pady=6,
                )
                tile.grid(row=shown // 5, column=shown % 5, sticky="nsew", padx=5, pady=5)
                tile.grid_propagate(False)
                thumb = self.product_thumbnail(product)
                if thumb:
                    image_label = tk.Label(tile, image=thumb, bg=self.colors["surface2"], cursor="hand2")
                    image_label.image = thumb
                    image_label.pack(fill="x")
                else:
                    placeholder = tk.Label(
                        tile,
                        text=str(product["name"] or "?")[:1].upper(),
                        bg="#e5edf7",
                        fg=self.colors["accent"],
                        font=("Segoe UI", 20, "bold"),
                        height=2,
                        cursor="hand2",
                    )
                    placeholder.pack(fill="x")
                name_label = tk.Label(
                    tile,
                    text=str(product["name"])[:34],
                    bg=self.colors["surface2"],
                    fg=self.colors["fg"],
                    font=("Segoe UI", 8, "bold"),
                    wraplength=108,
                    justify="center",
                    cursor="hand2",
                )
                name_label.pack(fill="x", pady=(3, 1))
                price_label = tk.Label(
                    tile,
                    text=self.money(product["sale_price"]),
                    bg=self.colors["surface2"],
                    fg=self.colors["accent"],
                    font=("Segoe UI", 10, "bold"),
                    cursor="hand2",
                )
                price_label.pack(fill="x")
                stock_label = tk.Label(
                    tile,
                    text=f"Stock: {product['stock']}",
                    bg=self.colors["surface2"],
                    fg=self.colors["muted"],
                    font=("Segoe UI", 8),
                    cursor="hand2",
                )
                stock_label.pack(fill="x", pady=(2, 0))
                for widget in (tile, *tile.winfo_children()):
                    widget.bind("<Button-1>", lambda _event, p=product: add_product(p))
                tiles.columnconfigure(shown % 5, weight=1)
                shown += 1
            if shown == 0:
                ttk.Label(tiles, text="Aucun produit trouve", background="#ffffff").grid(row=0, column=0, padx=20, pady=20)

        search.trace_add("write", render_products)
        category.trace_add("write", render_products)
        scan_entry.bind("<Return>", add_scanned_product)

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
            if not self.open_cash_session() and not messagebox.askyesno(APP_NAME, "Aucune caisse ouverte. Enregistrer quand meme la vente ?"):
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
        cash_session = self.open_cash_session()
        cur = self.db.execute(
            """
            INSERT INTO sales(sale_no, customer_id, customer_name, total, paid, payment_method, status, invoice_no, invoice_status, fne_status, user_id, cashier_name, cash_session_id, created_at)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, 'Active', 'A exporter', ?, ?, ?, ?)
            """,
            (
                sale_no,
                customer_id,
                customer_label,
                total,
                paid_amount,
                payment_method,
                status,
                invoice_no,
                self.current_user["id"] if self.current_user else None,
                self.current_user["name"] if self.current_user else "",
                cash_session["id"] if cash_session else None,
                now,
            ),
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
            "INSERT INTO cash_entries(entry_type, label, amount, payment_method, user_id, cash_session_id, created_at) VALUES (?, ?, ?, ?, ?, ?, ?)",
            (
                "Recette",
                f"Vente {sale_no}",
                paid_amount,
                payment_method,
                self.current_user["id"] if self.current_user else None,
                cash_session["id"] if cash_session else None,
                now,
            ),
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
            *self.merchant_lines(include_tax=False),
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
            *self.merchant_lines(include_tax=True),
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
                "certified_at": sale["fne_certified_at"] or "",
                "invoice_status": sale["invoice_status"] or "Active",
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

    def create_professional_invoice_pdf(self, sale_id):
        try:
            from reportlab.lib import colors
            from reportlab.lib.pagesizes import A4
            from reportlab.lib.units import mm
            from reportlab.pdfgen import canvas
        except ImportError:
            messagebox.showerror(APP_NAME, "Le module reportlab est requis pour generer la facture PDF.")
            return None
        payload = self.sale_payload(sale_id=sale_id)
        if not payload:
            messagebox.showwarning(APP_NAME, "Vente introuvable.")
            return None
        sale = payload["sale"]
        customer = payload["customer"]
        items = payload["items"]
        INVOICES_DIR.mkdir(parents=True, exist_ok=True)
        invoice_no = sale["invoice_no"] or sale["sale_no"].replace("V", "FAC-", 1)
        path = INVOICES_DIR / f"{invoice_no}.pdf"
        doc = canvas.Canvas(str(path), pagesize=A4)
        width, height = A4
        margin = 16 * mm
        y = height - margin
        logo = self.logo_path()
        if logo and logo.exists():
            try:
                doc.drawImage(str(logo), margin, y - 22 * mm, width=28 * mm, height=22 * mm, preserveAspectRatio=True, mask="auto")
            except Exception:
                pass
        doc.setFont("Helvetica-Bold", 18)
        doc.setFillColor(colors.HexColor("#111827"))
        doc.drawRightString(width - margin, y - 4 * mm, "FACTURE PROFESSIONNELLE")
        doc.setFont("Helvetica", 9)
        doc.drawRightString(width - margin, y - 10 * mm, f"N: {invoice_no}")
        doc.drawRightString(width - margin, y - 15 * mm, f"Date: {sale['created_at'][:16]}")
        doc.drawRightString(width - margin, y - 20 * mm, f"FNE: {sale['fne_status'] or 'A exporter'}")
        y -= 34 * mm
        doc.setFont("Helvetica-Bold", 10)
        doc.drawString(margin, y, self.setting("shop_name"))
        doc.setFont("Helvetica", 8.5)
        for line in self.merchant_lines(include_tax=True)[1:]:
            y -= 5 * mm
            doc.drawString(margin, y, line[:115])
        y -= 8 * mm
        doc.setFillColor(colors.HexColor("#0f766e"))
        doc.rect(margin, y - 8 * mm, width - 2 * margin, 8 * mm, stroke=0, fill=1)
        doc.setFillColor(colors.white)
        doc.setFont("Helvetica-Bold", 9)
        doc.drawString(margin + 3 * mm, y - 5.5 * mm, "CLIENT")
        doc.setFillColor(colors.HexColor("#111827"))
        y -= 14 * mm
        buyer = [
            f"Nom: {sale['customer_name']}",
            f"Tel: {(customer['phone'] if customer else '') or '-'}",
            f"NCC: {(customer['ncc'] if customer else '') or '-'}",
            f"Adresse: {(customer['address'] if customer else '') or '-'}",
        ]
        doc.setFont("Helvetica", 8.5)
        for line in buyer:
            doc.drawString(margin, y, line[:100])
            y -= 5 * mm
        y -= 5 * mm
        headers = [("Designation", 0), ("Qte", 88), ("PU", 112), ("Total", 150)]
        doc.setFillColor(colors.HexColor("#e5edf7"))
        doc.rect(margin, y - 7 * mm, width - 2 * margin, 8 * mm, stroke=0, fill=1)
        doc.setFillColor(colors.HexColor("#111827"))
        doc.setFont("Helvetica-Bold", 8)
        for text, offset in headers:
            doc.drawString(margin + offset * mm, y - 4.7 * mm, text)
        y -= 11 * mm
        doc.setFont("Helvetica", 8)
        for item in items:
            if y < 45 * mm:
                doc.showPage()
                y = height - margin
            doc.drawString(margin, y, str(item["product_name"])[:42])
            doc.drawRightString(margin + 103 * mm, y, f"{float(item['quantity']):.2f}")
            doc.drawRightString(margin + 142 * mm, y, self.money(item["unit_price"]))
            doc.drawRightString(width - margin, y, self.money(item["total"]))
            y -= 7 * mm
        y -= 3 * mm
        doc.line(margin, y, width - margin, y)
        y -= 8 * mm
        remaining = max(0, float(sale["total"]) - float(sale["paid"]))
        totals = [
            ("Total TTC", sale["total"]),
            ("Montant paye", sale["paid"]),
            ("Reste a payer", remaining),
        ]
        doc.setFont("Helvetica-Bold", 10)
        for label, value in totals:
            doc.drawRightString(width - margin - 38 * mm, y, label)
            doc.drawRightString(width - margin, y, self.money(value))
            y -= 7 * mm
        y -= 5 * mm
        doc.setFont("Helvetica", 8)
        doc.drawString(margin, y, f"Mode paiement: {sale['payment_method']} | Statut paiement: {sale['status']} | Statut facture: {sale['invoice_status'] or 'Active'}")
        y -= 5 * mm
        doc.drawString(margin, y, f"Reference FNE: {sale['fne_reference'] or '-'}")
        y -= 8 * mm
        doc.setFillColor(colors.HexColor("#6b7280"))
        doc.setFont("Helvetica", 7)
        for line in [self.setting("invoice_footer"), self.setting("fne_note")]:
            doc.drawString(margin, y, str(line)[:130])
            y -= 4 * mm
        doc.save()
        return path

    def export_fne_batch(self, sales):
        records = []
        for sale in sales:
            payload = self.fne_payload_for_sale(sale_id=sale["id"])
            if payload:
                records.append(payload)
        if not records:
            return None
        FNE_DIR.mkdir(parents=True, exist_ok=True)
        path = FNE_DIR / f"fne-lot-{datetime.now().strftime('%Y%m%d-%H%M%S')}.json"
        path.write_text(json.dumps({"source": APP_NAME, "count": len(records), "records": records}, ensure_ascii=False, indent=2), encoding="utf-8")
        now = datetime.now().isoformat(timespec="seconds")
        for record in records:
            self.db.execute(
                "UPDATE sales SET fne_status = 'Exportee', fne_exported_at = ? WHERE sale_no = ?",
                (now, record["sale_no"]),
            )
        return path

    def create_quote_document(self, quote_id):
        quote = self.db.one("SELECT * FROM quotes WHERE id = ?", (quote_id,))
        if not quote:
            messagebox.showwarning(APP_NAME, "Devis introuvable.")
            return None
        items = self.db.all(
            """
            SELECT qi.*, p.barcode, p.unit
            FROM quote_items qi
            LEFT JOIN products p ON p.id = qi.product_id
            WHERE qi.quote_id = ?
            ORDER BY qi.id
            """,
            (quote_id,),
        )
        QUOTES_DIR.mkdir(parents=True, exist_ok=True)
        path = QUOTES_DIR / f"{quote['quote_no']}.txt"
        lines = [
            *self.merchant_lines(include_tax=True),
            "=" * 72,
            "DEVIS CLIENT",
            f"Numero devis: {quote['quote_no']}",
            f"Date: {quote['created_at'][:16]}",
            f"Validite: {quote['validity_date'] or '-'}",
            f"Client: {quote['customer_name']}",
            f"Statut: {quote['status']}",
            "-" * 72,
            f"{'Designation':30} {'Qte':>8} {'PU':>14} {'Total':>14}",
            "-" * 72,
        ]
        for item in items:
            lines.append(
                f"{item['product_name'][:30]:30} {float(item['quantity']):8.2f} {self.money(item['unit_price']):>14} {self.money(item['total']):>14}"
            )
        lines.extend(
            [
                "-" * 72,
                f"TOTAL DEVIS: {self.money(quote['total'])}",
                f"Note: {quote['note'] or '-'}",
                "-" * 72,
                "Ce devis est valable sous reserve de disponibilite du stock et confirmation commerciale.",
            ]
        )
        path.write_text("\n".join(lines), encoding="utf-8")
        return path

    def create_quote_pdf(self, quote_id):
        try:
            from reportlab.lib import colors
            from reportlab.lib.pagesizes import A4
            from reportlab.lib.units import mm
            from reportlab.pdfgen import canvas
        except ImportError:
            messagebox.showerror(APP_NAME, "Le module reportlab est requis pour generer le devis PDF.")
            return None
        quote = self.db.one("SELECT * FROM quotes WHERE id = ?", (quote_id,))
        if not quote:
            messagebox.showwarning(APP_NAME, "Devis introuvable.")
            return None
        items = self.db.all("SELECT * FROM quote_items WHERE quote_id = ? ORDER BY id", (quote_id,))
        QUOTES_DIR.mkdir(parents=True, exist_ok=True)
        path = QUOTES_DIR / f"{quote['quote_no']}.pdf"
        doc = canvas.Canvas(str(path), pagesize=A4)
        width, height = A4
        margin = 16 * mm
        y = height - margin
        logo = self.logo_path()
        if logo and logo.exists():
            try:
                doc.drawImage(str(logo), margin, y - 22 * mm, width=28 * mm, height=22 * mm, preserveAspectRatio=True, mask="auto")
            except Exception:
                pass
        doc.setFont("Helvetica-Bold", 18)
        doc.setFillColor(colors.HexColor("#111827"))
        doc.drawRightString(width - margin, y - 4 * mm, "DEVIS CLIENT")
        doc.setFont("Helvetica", 9)
        doc.drawRightString(width - margin, y - 10 * mm, f"N: {quote['quote_no']}")
        doc.drawRightString(width - margin, y - 15 * mm, f"Date: {quote['created_at'][:16]}")
        doc.drawRightString(width - margin, y - 20 * mm, f"Validite: {quote['validity_date'] or '-'}")
        y -= 34 * mm
        doc.setFont("Helvetica-Bold", 10)
        doc.drawString(margin, y, self.setting("shop_name"))
        doc.setFont("Helvetica", 8.5)
        for line in self.merchant_lines(include_tax=True)[1:]:
            y -= 5 * mm
            doc.drawString(margin, y, line[:115])
        y -= 8 * mm
        doc.setFillColor(colors.HexColor("#0f766e"))
        doc.rect(margin, y - 8 * mm, width - 2 * margin, 8 * mm, stroke=0, fill=1)
        doc.setFillColor(colors.white)
        doc.setFont("Helvetica-Bold", 9)
        doc.drawString(margin + 3 * mm, y - 5.5 * mm, "CLIENT")
        doc.setFillColor(colors.HexColor("#111827"))
        y -= 14 * mm
        doc.setFont("Helvetica", 8.5)
        doc.drawString(margin, y, f"Nom: {quote['customer_name']}")
        y -= 5 * mm
        doc.drawString(margin, y, f"Statut devis: {quote['status']}")
        y -= 10 * mm
        doc.setFillColor(colors.HexColor("#e5edf7"))
        doc.rect(margin, y - 7 * mm, width - 2 * margin, 8 * mm, stroke=0, fill=1)
        doc.setFillColor(colors.HexColor("#111827"))
        doc.setFont("Helvetica-Bold", 8)
        for text, offset in [("Designation", 0), ("Qte", 88), ("PU", 112), ("Total", 150)]:
            doc.drawString(margin + offset * mm, y - 4.7 * mm, text)
        y -= 11 * mm
        doc.setFont("Helvetica", 8)
        for item in items:
            if y < 45 * mm:
                doc.showPage()
                y = height - margin
            doc.drawString(margin, y, str(item["product_name"])[:42])
            doc.drawRightString(margin + 103 * mm, y, f"{float(item['quantity']):.2f}")
            doc.drawRightString(margin + 142 * mm, y, self.money(item["unit_price"]))
            doc.drawRightString(width - margin, y, self.money(item["total"]))
            y -= 7 * mm
        y -= 4 * mm
        doc.line(margin, y, width - margin, y)
        y -= 8 * mm
        doc.setFont("Helvetica-Bold", 11)
        doc.drawRightString(width - margin - 38 * mm, y, "Total devis")
        doc.drawRightString(width - margin, y, self.money(quote["total"]))
        y -= 10 * mm
        doc.setFont("Helvetica", 8)
        doc.drawString(margin, y, f"Note: {quote['note'] or '-'}"[:130])
        y -= 6 * mm
        doc.setFillColor(colors.HexColor("#6b7280"))
        doc.setFont("Helvetica", 7)
        doc.drawString(margin, y, "Ce devis est valable sous reserve de disponibilite du stock et confirmation commerciale."[:130])
        doc.save()
        return path

    def create_delivery_document(self, delivery_id):
        delivery = self.db.one("SELECT * FROM delivery_notes WHERE id = ?", (delivery_id,))
        if not delivery:
            messagebox.showwarning(APP_NAME, "Bon de livraison introuvable.")
            return None
        sale = self.db.one("SELECT * FROM sales WHERE id = ?", (delivery["sale_id"],)) if delivery["sale_id"] else None
        items = self.db.all(
            """
            SELECT dni.*, p.barcode
            FROM delivery_note_items dni
            LEFT JOIN products p ON p.id = dni.product_id
            WHERE dni.delivery_id = ?
            ORDER BY dni.id
            """,
            (delivery_id,),
        )
        DELIVERIES_DIR.mkdir(parents=True, exist_ok=True)
        path = DELIVERIES_DIR / f"{delivery['delivery_no']}.txt"
        lines = [
            *self.merchant_lines(include_tax=True),
            "=" * 72,
            "BON DE LIVRAISON",
            f"Numero BL: {delivery['delivery_no']}",
            f"Facture/Vente: {(sale['invoice_no'] or sale['sale_no']) if sale else '-'}",
            f"Date creation: {delivery['created_at'][:16]}",
            f"Date livraison: {delivery['delivered_at'] or '-'}",
            f"Client: {delivery['customer_name']}",
            f"Statut: {delivery['status']}",
            "-" * 72,
            f"{'Designation':42} {'Qte':>10} {'Unite':>10}",
            "-" * 72,
        ]
        for item in items:
            lines.append(f"{item['product_name'][:42]:42} {float(item['quantity']):10.2f} {(item['unit'] or 'piece')[:10]:>10}")
        lines.extend(
            [
                "-" * 72,
                f"Observation: {delivery['note'] or '-'}",
                "",
                "Signature client: ____________________    Signature magasin: ____________________",
            ]
        )
        path.write_text("\n".join(lines), encoding="utf-8")
        return path

    def create_delivery_pdf(self, delivery_id):
        try:
            from reportlab.lib import colors
            from reportlab.lib.pagesizes import A4
            from reportlab.lib.units import mm
            from reportlab.pdfgen import canvas
        except ImportError:
            messagebox.showerror(APP_NAME, "Le module reportlab est requis pour generer le bon de livraison PDF.")
            return None
        delivery = self.db.one("SELECT * FROM delivery_notes WHERE id = ?", (delivery_id,))
        if not delivery:
            messagebox.showwarning(APP_NAME, "Bon de livraison introuvable.")
            return None
        sale = self.db.one("SELECT * FROM sales WHERE id = ?", (delivery["sale_id"],)) if delivery["sale_id"] else None
        items = self.db.all("SELECT * FROM delivery_note_items WHERE delivery_id = ? ORDER BY id", (delivery_id,))
        DELIVERIES_DIR.mkdir(parents=True, exist_ok=True)
        path = DELIVERIES_DIR / f"{delivery['delivery_no']}.pdf"
        doc = canvas.Canvas(str(path), pagesize=A4)
        width, height = A4
        margin = 16 * mm
        y = height - margin
        logo = self.logo_path()
        if logo and logo.exists():
            try:
                doc.drawImage(str(logo), margin, y - 22 * mm, width=28 * mm, height=22 * mm, preserveAspectRatio=True, mask="auto")
            except Exception:
                pass
        doc.setFont("Helvetica-Bold", 18)
        doc.setFillColor(colors.HexColor("#111827"))
        doc.drawRightString(width - margin, y - 4 * mm, "BON DE LIVRAISON")
        doc.setFont("Helvetica", 9)
        doc.drawRightString(width - margin, y - 10 * mm, f"N: {delivery['delivery_no']}")
        doc.drawRightString(width - margin, y - 15 * mm, f"Date: {delivery['created_at'][:16]}")
        doc.drawRightString(width - margin, y - 20 * mm, f"Statut: {delivery['status']}")
        y -= 34 * mm
        doc.setFont("Helvetica-Bold", 10)
        doc.drawString(margin, y, self.setting("shop_name"))
        doc.setFont("Helvetica", 8.5)
        for line in self.merchant_lines(include_tax=True)[1:]:
            y -= 5 * mm
            doc.drawString(margin, y, line[:115])
        y -= 8 * mm
        doc.setFillColor(colors.HexColor("#0f766e"))
        doc.rect(margin, y - 8 * mm, width - 2 * margin, 8 * mm, stroke=0, fill=1)
        doc.setFillColor(colors.white)
        doc.setFont("Helvetica-Bold", 9)
        doc.drawString(margin + 3 * mm, y - 5.5 * mm, "LIVRAISON")
        doc.setFillColor(colors.HexColor("#111827"))
        y -= 14 * mm
        doc.setFont("Helvetica", 8.5)
        doc.drawString(margin, y, f"Client: {delivery['customer_name']}")
        y -= 5 * mm
        doc.drawString(margin, y, f"Facture/Vente: {(sale['invoice_no'] or sale['sale_no']) if sale else '-'}")
        y -= 5 * mm
        doc.drawString(margin, y, f"Date livraison: {delivery['delivered_at'] or '-'}")
        y -= 10 * mm
        doc.setFillColor(colors.HexColor("#e5edf7"))
        doc.rect(margin, y - 7 * mm, width - 2 * margin, 8 * mm, stroke=0, fill=1)
        doc.setFillColor(colors.HexColor("#111827"))
        doc.setFont("Helvetica-Bold", 8)
        for text, offset in [("Designation", 0), ("Qte", 125), ("Unite", 150)]:
            doc.drawString(margin + offset * mm, y - 4.7 * mm, text)
        y -= 11 * mm
        doc.setFont("Helvetica", 8)
        for item in items:
            if y < 55 * mm:
                doc.showPage()
                y = height - margin
            doc.drawString(margin, y, str(item["product_name"])[:58])
            doc.drawRightString(margin + 140 * mm, y, f"{float(item['quantity']):.2f}")
            doc.drawString(margin + 150 * mm, y, str(item["unit"] or "piece")[:12])
            y -= 7 * mm
        y -= 8 * mm
        doc.setFont("Helvetica", 8)
        doc.drawString(margin, y, f"Observation: {delivery['note'] or '-'}"[:130])
        y -= 22 * mm
        doc.line(margin, y, margin + 65 * mm, y)
        doc.line(width - margin - 65 * mm, y, width - margin, y)
        y -= 5 * mm
        doc.drawString(margin, y, "Signature client")
        doc.drawString(width - margin - 65 * mm, y, "Signature magasin")
        doc.save()
        return path

    def create_sale_from_quote(self, quote_id, paid_amount=0, payment_method="Credit"):
        quote = self.db.one("SELECT * FROM quotes WHERE id = ?", (quote_id,))
        if not quote:
            messagebox.showwarning(APP_NAME, "Devis introuvable.")
            return None
        if quote["converted_sale_id"]:
            messagebox.showwarning(APP_NAME, "Ce devis est deja transforme en vente.")
            return None
        items = self.db.all("SELECT * FROM quote_items WHERE quote_id = ? ORDER BY id", (quote_id,))
        if not items:
            messagebox.showwarning(APP_NAME, "Ce devis ne contient aucune ligne.")
            return None
        total = float(quote["total"] or 0)
        status = "Paye" if paid_amount >= total else "Credit"
        if status == "Credit" and not quote["customer_id"]:
            messagebox.showwarning(APP_NAME, "Pour convertir en credit, le devis doit etre lie a un client existant.")
            return None
        now = datetime.now().isoformat(timespec="seconds")
        sale_no = datetime.now().strftime("V%Y%m%d%H%M%S")
        invoice_no = sale_no.replace("V", "FAC-", 1)
        cash_session = self.open_cash_session()
        cur = self.db.execute(
            """
            INSERT INTO sales(sale_no, customer_id, customer_name, total, paid, payment_method, status, invoice_no, invoice_status, fne_status, user_id, cashier_name, cash_session_id, created_at)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, 'Active', 'A exporter', ?, ?, ?, ?)
            """,
            (
                sale_no,
                quote["customer_id"],
                quote["customer_name"],
                total,
                paid_amount,
                payment_method,
                status,
                invoice_no,
                self.current_user["id"] if self.current_user else None,
                self.current_user["name"] if self.current_user else "",
                cash_session["id"] if cash_session else None,
                now,
            ),
        )
        sale_id = cur.lastrowid
        for item in items:
            self.db.execute(
                "INSERT INTO sale_items(sale_id, product_id, product_name, quantity, unit_price, total) VALUES (?, ?, ?, ?, ?, ?)",
                (sale_id, item["product_id"], item["product_name"], item["quantity"], item["unit_price"], item["total"]),
            )
            self.db.execute("UPDATE products SET stock = stock - ? WHERE id = ?", (item["quantity"], item["product_id"]))
            self.db.execute(
                "INSERT INTO stock_movements(product_id, product_name, movement_type, quantity, note, created_at) VALUES (?, ?, ?, ?, ?, ?)",
                (item["product_id"], item["product_name"], "Sortie vente devis", -float(item["quantity"]), quote["quote_no"], now),
            )
        if status == "Credit" and quote["customer_id"]:
            self.db.execute("UPDATE customers SET balance = balance + ? WHERE id = ?", (total - paid_amount, quote["customer_id"]))
        if paid_amount > 0:
            self.db.execute(
                "INSERT INTO cash_entries(entry_type, label, amount, payment_method, user_id, cash_session_id, created_at) VALUES (?, ?, ?, ?, ?, ?, ?)",
                (
                    "Recette",
                    f"Vente {sale_no} depuis devis {quote['quote_no']}",
                    paid_amount,
                    payment_method,
                    self.current_user["id"] if self.current_user else None,
                    cash_session["id"] if cash_session else None,
                    now,
                ),
            )
        self.db.execute("UPDATE quotes SET status = 'Converti', converted_sale_id = ? WHERE id = ?", (sale_id, quote_id))
        return sale_id

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

    def show_purchases(self):
        self.clear("Achats", "Approvisionnement fournisseur, entree stock et charge comptable.")
        body = ttk.Frame(self.main)
        body.pack(fill="both", expand=True, padx=24, pady=8)
        form = self.card(body)
        form.pack(side="left", fill="y", padx=(0, 12))
        products = self.db.all("SELECT id, name, purchase_price FROM products WHERE active = 1 ORDER BY name")
        suppliers = self.db.all("SELECT id, name FROM suppliers WHERE active = 1 ORDER BY name")
        product_var = tk.StringVar()
        supplier_var = tk.StringVar(value="Comptoir fournisseur")
        quantity = tk.StringVar(value="1")
        unit_cost = tk.StringVar(value="0")
        paid = tk.StringVar(value="0")
        payment = tk.StringVar(value="Especes")
        ttk.Label(form, text="Nouvel achat", style="CardTitle.TLabel").grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 10))
        ttk.Label(form, text="Fournisseur", background=self.colors["surface"]).grid(row=1, column=0, sticky="w", pady=5)
        ttk.Combobox(form, textvariable=supplier_var, values=["Comptoir fournisseur"] + [f"{s['id']} - {s['name']}" for s in suppliers], width=30).grid(row=1, column=1, sticky="ew", padx=(10, 0), pady=5)
        ttk.Label(form, text="Produit", background=self.colors["surface"]).grid(row=2, column=0, sticky="w", pady=5)
        product_combo = ttk.Combobox(form, textvariable=product_var, values=[f"{p['id']} - {p['name']}" for p in products], width=30)
        product_combo.grid(row=2, column=1, sticky="ew", padx=(10, 0), pady=5)
        self.form_row(form, "Quantite", 3, quantity)
        self.form_row(form, "Cout unitaire", 4, unit_cost)
        self.form_row(form, "Montant paye", 5, paid)
        ttk.Label(form, text="Paiement", background=self.colors["surface"]).grid(row=6, column=0, sticky="w", pady=5)
        ttk.Combobox(form, textvariable=payment, values=("Especes", "Mobile Money", "Carte", "Banque", "Credit"), state="readonly", width=30).grid(row=6, column=1, sticky="ew", padx=(10, 0), pady=5)

        def fill_cost(_=None):
            if " - " not in product_var.get():
                return
            product_id = int(product_var.get().split(" - ")[0])
            product = self.db.one("SELECT * FROM products WHERE id = ?", (product_id,))
            if product:
                unit_cost.set(str(product["purchase_price"]))

        product_combo.bind("<<ComboboxSelected>>", fill_cost)

        def save_purchase():
            if " - " not in product_var.get():
                messagebox.showwarning(APP_NAME, "Choisissez un produit.")
                return
            product_id = int(product_var.get().split(" - ")[0])
            product = self.db.one("SELECT * FROM products WHERE id = ?", (product_id,))
            qty = self.read_float_value(quantity.get(), "Quantite")
            cost = self.read_float_value(unit_cost.get(), "Cout unitaire")
            paid_amount = self.read_float_value(paid.get(), "Montant paye")
            if None in {qty, cost, paid_amount}:
                return
            supplier_id = None
            supplier_name = supplier_var.get().strip() or "Comptoir fournisseur"
            if " - " in supplier_name:
                supplier_id = int(supplier_name.split(" - ")[0])
                supplier_name = supplier_name.split(" - ", 1)[1]
            now = datetime.now().isoformat(timespec="seconds")
            purchase_no = datetime.now().strftime("A%Y%m%d%H%M%S")
            total = qty * cost
            cur = self.db.execute(
                """
                INSERT INTO purchases(purchase_no, supplier_id, supplier_name, total, paid, payment_method, status, user_id, created_at)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (purchase_no, supplier_id, supplier_name, total, paid_amount, payment.get(), "Recu" if paid_amount >= total else "Credit fournisseur", self.current_user["id"], now),
            )
            purchase_id = cur.lastrowid
            self.db.execute(
                "INSERT INTO purchase_items(purchase_id, product_id, product_name, quantity, unit_cost, total) VALUES (?, ?, ?, ?, ?, ?)",
                (purchase_id, product_id, product["name"], qty, cost, total),
            )
            self.db.execute("UPDATE products SET stock = stock + ?, purchase_price = ? WHERE id = ?", (qty, cost, product_id))
            self.db.execute(
                "INSERT INTO stock_movements(product_id, product_name, movement_type, quantity, note, created_at) VALUES (?, ?, ?, ?, ?, ?)",
                (product_id, product["name"], "Entree achat", qty, purchase_no, now),
            )
            if paid_amount > 0:
                session = self.open_cash_session()
                self.db.execute(
                    "INSERT INTO cash_entries(entry_type, label, amount, payment_method, user_id, cash_session_id, created_at) VALUES (?, ?, ?, ?, ?, ?, ?)",
                    ("Depense", f"Achat {purchase_no}", paid_amount, payment.get(), self.current_user["id"], session["id"] if session else None, now),
                )
            self.show_purchases()

        ttk.Button(form, text="Enregistrer achat", style="Accent.TButton", command=save_purchase).grid(row=8, column=0, columnspan=2, sticky="ew", pady=(12, 0))
        listing = self.card(body)
        listing.pack(side="left", fill="both", expand=True)
        ttk.Label(listing, text="Derniers achats", style="CardTitle.TLabel").pack(anchor="w", pady=(0, 8))
        self.table(
            listing,
            ("Numero", "Fournisseur", "Total", "Paye", "Paiement", "Statut", "Date"),
            [(r["purchase_no"], r["supplier_name"], self.money(r["total"]), self.money(r["paid"]), r["payment_method"], r["status"], r["created_at"][:16]) for r in self.db.all("SELECT * FROM purchases ORDER BY id DESC LIMIT 120")],
        )

    def show_returns(self):
        self.clear("Retours", "Retour client, remise en stock et remboursement trace.")
        body = ttk.Frame(self.main)
        body.pack(fill="both", expand=True, padx=24, pady=8)
        form = self.card(body)
        form.pack(side="left", fill="y", padx=(0, 12))
        sale_no = tk.StringVar()
        item_var = tk.StringVar()
        quantity = tk.StringVar(value="1")
        refund = tk.StringVar(value="0")
        reason = tk.StringVar(value="Retour client")
        items_cache = {}
        ttk.Label(form, text="Nouveau retour", style="CardTitle.TLabel").grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 10))
        self.form_row(form, "N vente", 1, sale_no)
        ttk.Label(form, text="Article", background=self.colors["surface"]).grid(row=2, column=0, sticky="w", pady=5)
        item_combo = ttk.Combobox(form, textvariable=item_var, width=34)
        item_combo.grid(row=2, column=1, sticky="ew", padx=(10, 0), pady=5)
        self.form_row(form, "Quantite", 3, quantity)
        self.form_row(form, "Remboursement", 4, refund)
        self.form_row(form, "Motif", 5, reason)

        def load_sale_items(_=None):
            sale = self.db.one("SELECT * FROM sales WHERE sale_no = ?", (sale_no.get().strip(),))
            if not sale:
                item_combo["values"] = []
                items_cache.clear()
                return
            rows = self.db.all("SELECT * FROM sale_items WHERE sale_id = ? ORDER BY id", (sale["id"],))
            values = []
            items_cache.clear()
            for row in rows:
                label = f"{row['id']} - {row['product_name']} ({row['quantity']} x {self.money(row['unit_price'])})"
                values.append(label)
                items_cache[row["id"]] = row
            item_combo["values"] = values
            if values:
                item_var.set(values[0])
                refund.set(str(int(rows[0]["unit_price"])))

        sale_no.trace_add("write", load_sale_items)

        def save_return():
            sale = self.db.one("SELECT * FROM sales WHERE sale_no = ?", (sale_no.get().strip(),))
            if not sale or " - " not in item_var.get():
                messagebox.showwarning(APP_NAME, "Indiquez une vente valide et un article.")
                return
            sale_item_id = int(item_var.get().split(" - ")[0])
            item = self.db.one("SELECT * FROM sale_items WHERE id = ?", (sale_item_id,))
            qty = self.read_float_value(quantity.get(), "Quantite")
            refund_amount = self.read_float_value(refund.get(), "Remboursement")
            if None in {qty, refund_amount}:
                return
            if qty > float(item["quantity"]):
                messagebox.showwarning(APP_NAME, "La quantite retournee depasse la quantite vendue.")
                return
            now = datetime.now().isoformat(timespec="seconds")
            return_no = datetime.now().strftime("R%Y%m%d%H%M%S")
            self.db.execute(
                """
                INSERT INTO sale_returns(return_no, sale_id, sale_item_id, product_id, product_name, quantity, refund_amount, reason, user_id, created_at)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (return_no, sale["id"], sale_item_id, item["product_id"], item["product_name"], qty, refund_amount, reason.get().strip(), self.current_user["id"], now),
            )
            self.db.execute("UPDATE products SET stock = stock + ? WHERE id = ?", (qty, item["product_id"]))
            self.db.execute(
                "INSERT INTO stock_movements(product_id, product_name, movement_type, quantity, note, created_at) VALUES (?, ?, ?, ?, ?, ?)",
                (item["product_id"], item["product_name"], "Retour client", qty, return_no, now),
            )
            if refund_amount > 0:
                session = self.open_cash_session()
                self.db.execute(
                    "INSERT INTO cash_entries(entry_type, label, amount, payment_method, user_id, cash_session_id, created_at) VALUES (?, ?, ?, ?, ?, ?, ?)",
                    ("Depense", f"Retour {return_no}", refund_amount, sale["payment_method"], self.current_user["id"], session["id"] if session else None, now),
                )
            self.show_returns()

        ttk.Button(form, text="Enregistrer retour", style="Accent.TButton", command=save_return).grid(row=7, column=0, columnspan=2, sticky="ew", pady=(12, 0))
        listing = self.card(body)
        listing.pack(side="left", fill="both", expand=True)
        ttk.Label(listing, text="Historique retours", style="CardTitle.TLabel").pack(anchor="w", pady=(0, 8))
        self.table(
            listing,
            ("Numero", "Produit", "Qte", "Remboursement", "Motif", "Date"),
            [(r["return_no"], r["product_name"], r["quantity"], self.money(r["refund_amount"]), r["reason"], r["created_at"][:16]) for r in self.db.all("SELECT * FROM sale_returns ORDER BY id DESC LIMIT 120")],
        )

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
        self.clear("Caisse", "Ouverture, cloture, recettes, depenses et mouvements de tresorerie.")
        body = ttk.Frame(self.main)
        body.pack(fill="both", expand=True, padx=24, pady=8)
        form = self.card(body)
        form.pack(side="left", fill="y", padx=(0, 12))
        entry_type = tk.StringVar(value="Depense")
        label = tk.StringVar()
        amount = tk.StringVar(value="0")
        payment = tk.StringVar(value="Especes")
        opening_amount = tk.StringVar(value="0")
        real_amount = tk.StringVar(value="0")
        session = self.open_cash_session()

        ttk.Label(form, text="Session caisse", style="CardTitle.TLabel").grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 10))
        status_text = f"Ouverte depuis {session['opened_at'][:16]} par {session['cashier_name']}" if session else "Aucune caisse ouverte"
        ttk.Label(form, text=status_text, background="#ffffff", wraplength=260).grid(row=1, column=0, columnspan=2, sticky="w", pady=5)
        self.form_row(form, "Fond initial", 2, opening_amount)
        self.form_row(form, "Caisse reelle", 3, real_amount)

        def open_session():
            if self.open_cash_session():
                messagebox.showinfo(APP_NAME, "Une caisse est deja ouverte.")
                return
            value = self.read_float_value(opening_amount.get(), "Fond initial")
            if value is None:
                return
            self.db.execute(
                """
                INSERT INTO cash_sessions(opened_at, opening_amount, expected_amount, user_id, cashier_name, status)
                VALUES (?, ?, ?, ?, ?, 'Ouverte')
                """,
                (
                    datetime.now().isoformat(timespec="seconds"),
                    value,
                    value,
                    self.current_user["id"] if self.current_user else None,
                    self.current_user["name"] if self.current_user else "",
                ),
            )
            self.show_cash()

        def close_session():
            session_row = self.open_cash_session()
            if not session_row:
                messagebox.showwarning(APP_NAME, "Aucune caisse ouverte.")
                return
            if not self.require_admin("Cloturer la caisse"):
                return
            real = self.read_float_value(real_amount.get(), "Caisse reelle")
            if real is None:
                return
            expected = self.expected_cash_for_session(session_row["id"], session_row["opening_amount"])
            gap = real - expected
            self.db.execute(
                """
                UPDATE cash_sessions
                SET closed_at = ?, expected_amount = ?, real_amount = ?, gap_amount = ?, status = 'Fermee'
                WHERE id = ?
                """,
                (datetime.now().isoformat(timespec="seconds"), expected, real, gap, session_row["id"]),
            )
            report = self.create_cash_close_report(session_row["id"], expected, real, gap)
            messagebox.showinfo(APP_NAME, f"Caisse cloturee. Rapport:\n{report}")
            self.show_cash()

        ttk.Button(form, text="Ouvrir caisse", style="Accent.TButton", command=open_session).grid(row=4, column=0, columnspan=2, sticky="ew", pady=(10, 0))
        ttk.Button(form, text="Cloturer caisse", style="Danger.TButton", command=close_session).grid(row=5, column=0, columnspan=2, sticky="ew", pady=(6, 12))

        ttk.Label(form, text="Nouvelle ecriture", style="CardTitle.TLabel").grid(row=6, column=0, columnspan=2, sticky="w", pady=(8, 10))
        ttk.Label(form, text="Type", background="#ffffff").grid(row=7, column=0, sticky="w", pady=5)
        ttk.Combobox(form, textvariable=entry_type, values=("Recette", "Depense", "Apport", "Retrait"), width=28).grid(row=7, column=1, sticky="ew", pady=5, padx=(10, 0))
        self.form_row(form, "Libelle", 8, label)
        self.form_row(form, "Montant", 9, amount)
        ttk.Label(form, text="Paiement", background="#ffffff").grid(row=10, column=0, sticky="w", pady=5)
        ttk.Combobox(form, textvariable=payment, values=("Especes", "Mobile Money", "Carte", "Banque"), width=28).grid(row=10, column=1, sticky="ew", pady=5, padx=(10, 0))

        def save_entry():
            try:
                value = float(amount.get() or 0)
            except ValueError:
                messagebox.showerror(APP_NAME, "Montant invalide.")
                return
            if not label.get().strip():
                messagebox.showwarning(APP_NAME, "Libelle obligatoire.")
                return
            session_row = self.open_cash_session()
            self.db.execute(
                "INSERT INTO cash_entries(entry_type, label, amount, payment_method, user_id, cash_session_id, created_at) VALUES (?, ?, ?, ?, ?, ?, ?)",
                (
                    entry_type.get(),
                    label.get().strip(),
                    value,
                    payment.get(),
                    self.current_user["id"] if self.current_user else None,
                    session_row["id"] if session_row else None,
                    datetime.now().isoformat(timespec="seconds"),
                ),
            )
            self.show_cash()

        ttk.Button(form, text="Enregistrer", style="Accent.TButton", command=save_entry).grid(row=12, column=0, columnspan=2, sticky="ew", pady=(12, 0))
        listing = self.card(body)
        listing.pack(side="left", fill="both", expand=True)
        self.table(
            listing,
            ("Type", "Libelle", "Montant", "Paiement", "Date"),
            [(r["entry_type"], r["label"], self.money(r["amount"]), r["payment_method"], r["created_at"][:16]) for r in self.db.all("SELECT * FROM cash_entries ORDER BY id DESC LIMIT 100")],
        )

    def read_float_value(self, raw, label):
        try:
            value = float(str(raw or "0").replace(" ", "").replace(",", "."))
        except ValueError:
            messagebox.showerror(APP_NAME, f"{label} invalide.")
            return None
        if value < 0:
            messagebox.showerror(APP_NAME, f"{label} ne peut pas etre negatif.")
            return None
        return value

    def open_cash_session(self):
        return self.db.one("SELECT * FROM cash_sessions WHERE status = 'Ouverte' ORDER BY id DESC LIMIT 1")

    def expected_cash_for_session(self, session_id, opening_amount):
        cash_in = self.db.one(
            """
            SELECT COALESCE(SUM(amount), 0) v
            FROM cash_entries
            WHERE cash_session_id = ? AND payment_method = 'Especes' AND entry_type IN ('Recette', 'Apport')
            """,
            (session_id,),
        )["v"]
        cash_out = self.db.one(
            """
            SELECT COALESCE(SUM(amount), 0) v
            FROM cash_entries
            WHERE cash_session_id = ? AND payment_method = 'Especes' AND entry_type IN ('Depense', 'Retrait')
            """,
            (session_id,),
        )["v"]
        return float(opening_amount or 0) + float(cash_in or 0) - float(cash_out or 0)

    def create_cash_close_report(self, session_id, expected, real, gap):
        RECEIPTS_DIR.mkdir(parents=True, exist_ok=True)
        session = self.db.one("SELECT * FROM cash_sessions WHERE id = ?", (session_id,))
        rows = self.db.all("SELECT * FROM cash_entries WHERE cash_session_id = ? ORDER BY created_at", (session_id,))
        path = RECEIPTS_DIR / f"cloture-caisse-{session_id}-{datetime.now().strftime('%Y%m%d-%H%M%S')}.txt"
        lines = [
            *self.merchant_lines(include_tax=False),
            "=" * 48,
            "RAPPORT DE CLOTURE DE CAISSE",
            f"Caisse: #{session_id}",
            f"Ouverture: {session['opened_at'][:16]}",
            f"Cloture: {session['closed_at'][:16] if session['closed_at'] else datetime.now().strftime('%Y-%m-%dT%H:%M')}",
            f"Caissier: {session['cashier_name'] or ''}",
            f"Fond initial: {self.money(session['opening_amount'])}",
            "-" * 48,
        ]
        for row in rows:
            lines.append(f"{row['created_at'][:16]} | {row['entry_type'][:8]:8} | {self.money(row['amount']):>12} | {row['label'][:18]}")
        lines.extend(
            [
                "-" * 48,
                f"Caisse attendue: {self.money(expected)}",
                f"Caisse reelle:   {self.money(real)}",
                f"Ecart:           {self.money(gap)}",
                "=" * 48,
            ]
        )
        path.write_text("\n".join(lines), encoding="utf-8")
        return path

    def accounting_rows(self, start_date, end_date, entry_filter="Tous", payment_filter="Tous"):
        rows = []
        cash_rows = self.db.all(
            """
            SELECT cash_entries.*, users.name user_name
            FROM cash_entries
            LEFT JOIN users ON users.id = cash_entries.user_id
            WHERE date(cash_entries.created_at) BETWEEN ? AND ?
            ORDER BY cash_entries.created_at DESC
            """,
            (start_date, end_date),
        )
        for row in cash_rows:
            entry_type = "Recette" if row["entry_type"] in {"Recette", "Apport"} else "Depense"
            if entry_filter != "Tous" and entry_filter != entry_type:
                continue
            if payment_filter != "Tous" and row["payment_method"] != payment_filter:
                continue
            rows.append(
                {
                    "id": f"cash-{row['id']}",
                    "created_at": row["created_at"],
                    "type": entry_type,
                    "category": row["entry_type"],
                    "label": row["label"],
                    "amount": float(row["amount"] or 0),
                    "payment": row["payment_method"],
                    "user": row["user_name"] or "",
                    "source": "Caisse",
                }
            )
        manual_rows = self.db.all(
            """
            SELECT accounting_entries.*, users.name user_name
            FROM accounting_entries
            LEFT JOIN users ON users.id = accounting_entries.user_id
            WHERE date(accounting_entries.created_at) BETWEEN ? AND ?
            ORDER BY accounting_entries.created_at DESC
            """,
            (start_date, end_date),
        )
        for row in manual_rows:
            if entry_filter != "Tous" and entry_filter != row["entry_type"]:
                continue
            if payment_filter != "Tous" and row["payment_method"] != payment_filter:
                continue
            rows.append(
                {
                    "id": f"manual-{row['id']}",
                    "created_at": row["created_at"],
                    "type": row["entry_type"],
                    "category": row["category"],
                    "label": row["label"],
                    "amount": float(row["amount"] or 0),
                    "payment": row["payment_method"],
                    "user": row["user_name"] or "",
                    "source": "Manuel",
                }
            )
        rows.sort(key=lambda item: item["created_at"], reverse=True)
        return rows

    def show_accounting(self):
        self.clear("Comptabilite", "Journal financier, filtres, resultat et ecritures comptables.")
        body = ttk.Frame(self.main)
        body.pack(fill="both", expand=True, padx=24, pady=8)
        filters = self.card(body)
        filters.pack(fill="x")
        start_var = tk.StringVar(value=date.today().replace(day=1).isoformat())
        end_var = tk.StringVar(value=date.today().isoformat())
        type_var = tk.StringVar(value="Tous")
        payment_var = tk.StringVar(value="Tous")
        for col, (label, var, width) in enumerate(
            [
                ("Debut", start_var, 12),
                ("Fin", end_var, 12),
                ("Type", type_var, 14),
                ("Paiement", payment_var, 16),
            ]
        ):
            ttk.Label(filters, text=label, background=self.colors["surface"]).grid(row=0, column=col * 2, sticky="w", padx=(0, 6))
            if label == "Type":
                ttk.Combobox(filters, textvariable=var, values=("Tous", "Recette", "Depense"), state="readonly", width=width).grid(row=0, column=col * 2 + 1, sticky="w", padx=(0, 12))
            elif label == "Paiement":
                ttk.Combobox(filters, textvariable=var, values=("Tous", "Especes", "Mobile Money", "Carte", "Banque", "Credit"), state="readonly", width=width).grid(row=0, column=col * 2 + 1, sticky="w", padx=(0, 12))
            else:
                ttk.Entry(filters, textvariable=var, width=width).grid(row=0, column=col * 2 + 1, sticky="w", padx=(0, 12))

        summary = ttk.Frame(body)
        summary.pack(fill="x", pady=12)
        kpis = {}
        for index, key in enumerate(("recettes", "depenses", "resultat", "cash")):
            card = self.card(summary)
            card.grid(row=0, column=index, sticky="ew", padx=(0 if index == 0 else 12, 0))
            summary.columnconfigure(index, weight=1)
            ttk.Label(card, text={"recettes": "Recettes", "depenses": "Depenses", "resultat": "Resultat", "cash": "Solde especes"}[key], style="CardTitle.TLabel").pack(anchor="w")
            kpis[key] = ttk.Label(card, text=self.money(0), style="Metric.TLabel")
            kpis[key].pack(anchor="w", pady=(8, 0))

        middle = ttk.Frame(body)
        middle.pack(fill="both", expand=True)
        middle.columnconfigure(0, weight=0)
        middle.columnconfigure(1, weight=1)
        form = self.card(middle)
        form.grid(row=0, column=0, sticky="ns", padx=(0, 12))
        form.columnconfigure(1, weight=1)
        entry_type = tk.StringVar(value="Depense")
        category = tk.StringVar(value="Autre")
        label = tk.StringVar()
        amount = tk.StringVar(value="0")
        payment = tk.StringVar(value="Especes")
        note = tk.StringVar()
        ttk.Label(form, text="Ecriture manuelle", style="CardTitle.TLabel").grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 10))
        ttk.Label(form, text="Type", background=self.colors["surface"]).grid(row=1, column=0, sticky="w", pady=5)
        ttk.Combobox(form, textvariable=entry_type, values=("Depense", "Recette"), state="readonly", width=25).grid(row=1, column=1, sticky="ew", padx=(10, 0), pady=5)
        ttk.Label(form, text="Categorie", background=self.colors["surface"]).grid(row=2, column=0, sticky="w", pady=5)
        ttk.Combobox(form, textvariable=category, values=("Achat marchandises", "Loyer", "Salaire", "Electricite", "Transport", "Maintenance", "Impots", "Recette hors vente", "Autre"), width=25).grid(row=2, column=1, sticky="ew", padx=(10, 0), pady=5)
        self.form_row(form, "Libelle", 3, label)
        self.form_row(form, "Montant", 4, amount)
        ttk.Label(form, text="Paiement", background=self.colors["surface"]).grid(row=5, column=0, sticky="w", pady=5)
        ttk.Combobox(form, textvariable=payment, values=("Especes", "Mobile Money", "Carte", "Banque"), state="readonly", width=25).grid(row=5, column=1, sticky="ew", padx=(10, 0), pady=5)
        self.form_row(form, "Note", 6, note)

        journal_box = self.card(middle)
        journal_box.grid(row=0, column=1, sticky="nsew")
        journal_holder = ttk.Frame(journal_box, style="Card.TFrame")
        journal_holder.pack(fill="both", expand=True)
        journal_tree = {"widget": None}

        def refresh():
            for child in journal_holder.winfo_children():
                child.destroy()
            rows = self.accounting_rows(start_var.get(), end_var.get(), type_var.get(), payment_var.get())
            recettes = sum(row["amount"] for row in rows if row["type"] == "Recette")
            depenses = sum(row["amount"] for row in rows if row["type"] == "Depense")
            cash_in = sum(row["amount"] for row in rows if row["type"] == "Recette" and row["payment"] == "Especes")
            cash_out = sum(row["amount"] for row in rows if row["type"] == "Depense" and row["payment"] == "Especes")
            kpis["recettes"].configure(text=self.money(recettes))
            kpis["depenses"].configure(text=self.money(depenses))
            kpis["resultat"].configure(text=self.money(recettes - depenses))
            kpis["cash"].configure(text=self.money(cash_in - cash_out))
            journal_tree["widget"] = self.table(
                journal_holder,
                ("Date", "Type", "Categorie", "Libelle", "Montant", "Paiement", "Utilisateur", "Source"),
                [(r["created_at"][:16], r["type"], r["category"], r["label"], self.money(r["amount"]), r["payment"], r["user"], r["source"]) for r in rows],
            )

        def save_entry():
            value = self.read_float_value(amount.get(), "Montant")
            if value is None:
                return
            if not label.get().strip():
                messagebox.showwarning(APP_NAME, "Libelle obligatoire.")
                return
            self.db.execute(
                """
                INSERT INTO accounting_entries(created_at, entry_type, category, label, amount, payment_method, user_id, note)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    datetime.now().isoformat(timespec="seconds"),
                    entry_type.get(),
                    category.get().strip() or "Autre",
                    label.get().strip(),
                    value,
                    payment.get(),
                    self.current_user["id"] if self.current_user else None,
                    note.get().strip(),
                ),
            )
            label.set("")
            amount.set("0")
            note.set("")
            refresh()

        def export_accounting_csv():
            path = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV", "*.csv")], initialfile="comptabilite-magazin-ci.csv")
            if not path:
                return
            rows = self.accounting_rows(start_var.get(), end_var.get(), type_var.get(), payment_var.get())
            self.export_csv(
                path,
                ["Date", "Type", "Categorie", "Libelle", "Montant", "Paiement", "Utilisateur", "Source"],
                [[r["created_at"], r["type"], r["category"], r["label"], r["amount"], r["payment"], r["user"], r["source"]] for r in rows],
            )
            messagebox.showinfo(APP_NAME, "Export comptable CSV termine.")

        ttk.Button(form, text="Enregistrer ecriture", style="Accent.TButton", command=save_entry).grid(row=8, column=0, columnspan=2, sticky="ew", pady=(12, 0))
        ttk.Button(form, text="Exporter CSV", command=export_accounting_csv).grid(row=9, column=0, columnspan=2, sticky="ew", pady=(6, 0))
        ttk.Button(filters, text="Filtrer", style="Accent.TButton", command=refresh).grid(row=0, column=8, sticky="e")
        for variable in (start_var, end_var, type_var, payment_var):
            variable.trace_add("write", lambda *_: refresh())
        refresh()

    def show_quotes(self):
        self.clear("Devis", "Creation, suivi, impression et transformation en vente.")
        body = ttk.Frame(self.main)
        body.pack(fill="both", expand=True, padx=24, pady=8)
        body.columnconfigure(0, weight=0)
        body.columnconfigure(1, weight=1)
        quote_lines = []
        customers = self.db.all("SELECT id, name FROM customers ORDER BY name")
        products = self.db.all("SELECT id, name, sale_price, stock FROM products WHERE active = 1 ORDER BY name")
        customer_var = tk.StringVar(value="Client comptoir")
        validity_var = tk.StringVar(value=(date.today().replace(day=28) if date.today().day < 28 else date.today()).isoformat())
        note_var = tk.StringVar()
        product_var = tk.StringVar()
        qty_var = tk.StringVar(value="1")
        price_var = tk.StringVar(value="0")
        status_var = tk.StringVar(value="Brouillon")
        selected_quote_id = tk.IntVar(value=0)

        form = self.card(body)
        form.grid(row=0, column=0, sticky="ns", padx=(0, 12))
        form.columnconfigure(1, weight=1)
        ttk.Label(form, text="Nouveau devis", style="CardTitle.TLabel").grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 10))
        ttk.Label(form, text="Client", background=self.colors["surface"]).grid(row=1, column=0, sticky="w", pady=5)
        ttk.Combobox(form, textvariable=customer_var, values=["Client comptoir"] + [f"{c['id']} - {c['name']}" for c in customers], width=30).grid(row=1, column=1, sticky="ew", padx=(10, 0), pady=5)
        self.form_row(form, "Validite", 2, validity_var)
        self.form_row(form, "Note", 3, note_var)
        ttk.Label(form, text="Produit", background=self.colors["surface"]).grid(row=4, column=0, sticky="w", pady=5)
        product_combo = ttk.Combobox(form, textvariable=product_var, values=[f"{p['id']} - {p['name']}" for p in products], width=30)
        product_combo.grid(row=4, column=1, sticky="ew", padx=(10, 0), pady=5)
        self.form_row(form, "Quantite", 5, qty_var)
        self.form_row(form, "Prix unitaire", 6, price_var)

        listing = self.card(body)
        listing.grid(row=0, column=1, sticky="nsew")
        line_box = ttk.Frame(listing, style="Card.TFrame")
        line_box.pack(fill="x")
        line_tree_ref = {"tree": None}
        total_var = tk.StringVar(value=self.money(0))

        def refresh_lines():
            for child in line_box.winfo_children():
                child.destroy()
            line_tree_ref["tree"] = self.table(
                line_box,
                ("Produit", "Qte", "PU", "Total"),
                [(line["name"], line["qty"], self.money(line["price"]), self.money(line["total"])) for line in quote_lines],
            )
            total_var.set(self.money(sum(line["total"] for line in quote_lines)))

        def fill_price(_=None):
            if " - " not in product_var.get():
                return
            product_id = int(product_var.get().split(" - ")[0])
            product = self.db.one("SELECT * FROM products WHERE id = ?", (product_id,))
            if product:
                price_var.set(str(product["sale_price"]))

        def add_line():
            if " - " not in product_var.get():
                messagebox.showwarning(APP_NAME, "Choisissez un produit.")
                return
            product_id = int(product_var.get().split(" - ")[0])
            product = self.db.one("SELECT * FROM products WHERE id = ?", (product_id,))
            qty = self.read_float_value(qty_var.get(), "Quantite")
            price = self.read_float_value(price_var.get(), "Prix unitaire")
            if None in {qty, price}:
                return
            quote_lines.append({"product_id": product_id, "name": product["name"], "qty": qty, "price": price, "total": qty * price})
            refresh_lines()

        def remove_line():
            tree = line_tree_ref["tree"]
            if not tree or not tree.selection():
                return
            index = tree.index(tree.selection()[0])
            if 0 <= index < len(quote_lines):
                quote_lines.pop(index)
            refresh_lines()

        def clear_form():
            selected_quote_id.set(0)
            quote_lines.clear()
            customer_var.set("Client comptoir")
            note_var.set("")
            status_var.set("Brouillon")
            refresh_lines()

        def save_quote():
            if not quote_lines:
                messagebox.showwarning(APP_NAME, "Ajoutez au moins une ligne au devis.")
                return
            customer_id = None
            customer_name = customer_var.get().strip() or "Client comptoir"
            if " - " in customer_name:
                customer_id = int(customer_name.split(" - ")[0])
                customer_name = customer_name.split(" - ", 1)[1]
            now = datetime.now().isoformat(timespec="seconds")
            total = sum(line["total"] for line in quote_lines)
            if selected_quote_id.get():
                quote_id = selected_quote_id.get()
                existing_quote = self.db.one("SELECT converted_sale_id FROM quotes WHERE id = ?", (quote_id,))
                if existing_quote and existing_quote["converted_sale_id"]:
                    messagebox.showwarning(APP_NAME, "Ce devis est deja converti. Il ne peut plus etre modifie.")
                    return
                self.db.execute(
                    "UPDATE quotes SET customer_id = ?, customer_name = ?, total = ?, status = ?, validity_date = ?, note = ? WHERE id = ?",
                    (customer_id, customer_name, total, status_var.get(), validity_var.get().strip(), note_var.get().strip(), quote_id),
                )
                self.db.execute("DELETE FROM quote_items WHERE quote_id = ?", (quote_id,))
            else:
                quote_no = datetime.now().strftime("D%Y%m%d%H%M%S")
                cur = self.db.execute(
                    """
                    INSERT INTO quotes(quote_no, customer_id, customer_name, total, status, validity_date, note, user_id, created_at)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """,
                    (quote_no, customer_id, customer_name, total, status_var.get(), validity_var.get().strip(), note_var.get().strip(), self.current_user["id"] if self.current_user else None, now),
                )
                quote_id = cur.lastrowid
            for line in quote_lines:
                self.db.execute(
                    "INSERT INTO quote_items(quote_id, product_id, product_name, quantity, unit_price, total) VALUES (?, ?, ?, ?, ?, ?)",
                    (quote_id, line["product_id"], line["name"], line["qty"], line["price"], line["total"]),
                )
            self.show_quotes()

        product_combo.bind("<<ComboboxSelected>>", fill_price)
        ttk.Button(form, text="Ajouter ligne", style="Accent.TButton", command=add_line).grid(row=7, column=0, columnspan=2, sticky="ew", pady=(10, 0))
        ttk.Button(form, text="Retirer ligne", command=remove_line).grid(row=8, column=0, columnspan=2, sticky="ew", pady=(6, 0))
        ttk.Label(form, text="Statut", background=self.colors["surface"]).grid(row=9, column=0, sticky="w", pady=5)
        ttk.Combobox(form, textvariable=status_var, values=("Brouillon", "Envoye", "Accepte", "Refuse", "Converti"), state="readonly", width=30).grid(row=9, column=1, sticky="ew", padx=(10, 0), pady=5)
        ttk.Label(form, textvariable=total_var, style="Metric.TLabel").grid(row=10, column=0, columnspan=2, sticky="w", pady=(10, 0))
        ttk.Button(form, text="Enregistrer devis", style="Accent.TButton", command=save_quote).grid(row=11, column=0, columnspan=2, sticky="ew", pady=(10, 0))
        ttk.Button(form, text="Nouveau", command=clear_form).grid(row=12, column=0, columnspan=2, sticky="ew", pady=(6, 0))

        ttk.Label(listing, text="Lignes du devis", style="CardTitle.TLabel").pack(anchor="w", pady=(0, 8))
        refresh_lines()
        actions = ttk.Frame(listing, style="Card.TFrame")
        actions.pack(fill="x", pady=(10, 8))
        quotes_holder = ttk.Frame(listing, style="Card.TFrame")
        quotes_holder.pack(fill="both", expand=True)
        quote_tree = self.table(
            quotes_holder,
            ("ID", "Numero", "Client", "Total", "Statut", "Validite", "Date"),
            [(q["id"], q["quote_no"], q["customer_name"], self.money(q["total"]), q["status"], q["validity_date"] or "", q["created_at"][:16]) for q in self.db.all("SELECT * FROM quotes ORDER BY id DESC LIMIT 150")],
        )

        def selected_quote():
            selected = quote_tree.selection()
            if not selected:
                messagebox.showwarning(APP_NAME, "Selectionnez un devis.")
                return None
            return int(quote_tree.item(selected[0])["values"][0])

        def load_quote(_=None):
            quote_id = selected_quote()
            if not quote_id:
                return
            quote = self.db.one("SELECT * FROM quotes WHERE id = ?", (quote_id,))
            selected_quote_id.set(quote_id)
            customer_var.set(f"{quote['customer_id']} - {quote['customer_name']}" if quote["customer_id"] else quote["customer_name"])
            validity_var.set(quote["validity_date"] or "")
            note_var.set(quote["note"] or "")
            status_var.set(quote["status"])
            quote_lines.clear()
            for item in self.db.all("SELECT * FROM quote_items WHERE quote_id = ? ORDER BY id", (quote_id,)):
                quote_lines.append({"product_id": item["product_id"], "name": item["product_name"], "qty": item["quantity"], "price": item["unit_price"], "total": item["total"]})
            refresh_lines()

        def preview_quote():
            quote_id = selected_quote()
            if quote_id:
                path = self.create_quote_pdf(quote_id)
                if path:
                    self.open_document(path)

        def print_quote():
            quote_id = selected_quote()
            if quote_id:
                path = self.create_quote_document(quote_id)
                if path:
                    self.print_text_file(path)

        def convert_quote():
            quote_id = selected_quote()
            if not quote_id:
                return
            paid = simpledialog.askfloat(APP_NAME, "Montant paye a la conversion:", initialvalue=0.0, minvalue=0.0, parent=self)
            if paid is None:
                return
            sale_id = self.create_sale_from_quote(quote_id, paid, "Especes" if paid else "Credit")
            if sale_id:
                messagebox.showinfo(APP_NAME, "Devis transforme en vente et facture creee.")
                self.show_quotes()

        def delete_quote():
            quote_id = selected_quote()
            if not quote_id:
                return
            quote = self.db.one("SELECT * FROM quotes WHERE id = ?", (quote_id,))
            if quote["converted_sale_id"]:
                messagebox.showwarning(APP_NAME, "Impossible de supprimer un devis deja converti.")
                return
            if messagebox.askyesno(APP_NAME, "Supprimer ce devis ?"):
                self.db.execute("DELETE FROM quotes WHERE id = ?", (quote_id,))
                self.show_quotes()

        quote_tree.bind("<<TreeviewSelect>>", load_quote)
        ttk.Button(actions, text="Devis PDF", command=preview_quote).pack(side="left", padx=(0, 8))
        ttk.Button(actions, text="Imprimer devis", command=print_quote).pack(side="left", padx=(0, 8))
        ttk.Button(actions, text="Transformer en vente", style="Accent.TButton", command=convert_quote).pack(side="left", padx=(0, 8))
        ttk.Button(actions, text="Supprimer devis", style="Danger.TButton", command=delete_quote).pack(side="left")

    def show_delivery_notes(self):
        self.clear("Bons livraison", "Preparation, suivi et impression des livraisons client.")
        body = ttk.Frame(self.main)
        body.pack(fill="both", expand=True, padx=24, pady=8)
        form = self.card(body)
        form.pack(side="left", fill="y", padx=(0, 12))
        sales = self.db.all("SELECT * FROM sales WHERE invoice_status != 'Annulee' ORDER BY id DESC LIMIT 250")
        sale_var = tk.StringVar()
        status_var = tk.StringVar(value="Prepare")
        delivered_at_var = tk.StringVar()
        note_var = tk.StringVar()
        ttk.Label(form, text="Nouveau bon de livraison", style="CardTitle.TLabel").grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 10))
        ttk.Label(form, text="Vente/Facture", background=self.colors["surface"]).grid(row=1, column=0, sticky="w", pady=5)
        ttk.Combobox(form, textvariable=sale_var, values=[f"{s['id']} - {s['invoice_no'] or s['sale_no']} - {s['customer_name']}" for s in sales], width=42).grid(row=1, column=1, sticky="ew", padx=(10, 0), pady=5)
        ttk.Label(form, text="Statut", background=self.colors["surface"]).grid(row=2, column=0, sticky="w", pady=5)
        ttk.Combobox(form, textvariable=status_var, values=("Prepare", "En livraison", "Livre", "Annule"), state="readonly", width=30).grid(row=2, column=1, sticky="ew", padx=(10, 0), pady=5)
        self.form_row(form, "Date livraison", 3, delivered_at_var)
        self.form_row(form, "Observation", 4, note_var)

        def save_delivery():
            if " - " not in sale_var.get():
                messagebox.showwarning(APP_NAME, "Choisissez une vente/facture.")
                return
            sale_id = int(sale_var.get().split(" - ")[0])
            sale = self.db.one("SELECT * FROM sales WHERE id = ?", (sale_id,))
            existing = self.db.one("SELECT id FROM delivery_notes WHERE sale_id = ? ORDER BY id DESC LIMIT 1", (sale_id,))
            if existing and not messagebox.askyesno(APP_NAME, "Un bon existe deja pour cette vente. Creer un autre bon ?"):
                return
            now = datetime.now().isoformat(timespec="seconds")
            delivery_no = datetime.now().strftime("BL%Y%m%d%H%M%S")
            cur = self.db.execute(
                """
                INSERT INTO delivery_notes(delivery_no, sale_id, customer_id, customer_name, status, note, delivered_at, user_id, created_at)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (delivery_no, sale_id, sale["customer_id"], sale["customer_name"], status_var.get(), note_var.get().strip(), delivered_at_var.get().strip(), self.current_user["id"] if self.current_user else None, now),
            )
            delivery_id = cur.lastrowid
            items = self.db.all(
                """
                SELECT si.*, p.unit
                FROM sale_items si
                LEFT JOIN products p ON p.id = si.product_id
                WHERE si.sale_id = ?
                ORDER BY si.id
                """,
                (sale_id,),
            )
            for item in items:
                self.db.execute(
                    "INSERT INTO delivery_note_items(delivery_id, product_id, product_name, quantity, unit) VALUES (?, ?, ?, ?, ?)",
                    (delivery_id, item["product_id"], item["product_name"], item["quantity"], item["unit"] or "piece"),
                )
            messagebox.showinfo(APP_NAME, f"Bon de livraison cree: {delivery_no}")
            self.show_delivery_notes()

        ttk.Button(form, text="Creer bon de livraison", style="Accent.TButton", command=save_delivery).grid(row=6, column=0, columnspan=2, sticky="ew", pady=(12, 0))
        listing = self.card(body)
        listing.pack(side="left", fill="both", expand=True)
        actions = ttk.Frame(listing, style="Card.TFrame")
        actions.pack(fill="x", pady=(0, 8))
        tree = self.table(
            listing,
            ("ID", "Numero", "Vente", "Client", "Statut", "Livraison", "Date"),
            [
                (d["id"], d["delivery_no"], d["sale_id"] or "", d["customer_name"], d["status"], d["delivered_at"] or "", d["created_at"][:16])
                for d in self.db.all("SELECT * FROM delivery_notes ORDER BY id DESC LIMIT 150")
            ],
        )

        def selected_delivery():
            selected = tree.selection()
            if not selected:
                messagebox.showwarning(APP_NAME, "Selectionnez un bon de livraison.")
                return None
            return int(tree.item(selected[0])["values"][0])

        def preview_delivery():
            delivery_id = selected_delivery()
            if delivery_id:
                path = self.create_delivery_pdf(delivery_id)
                if path:
                    self.open_document(path)

        def print_delivery():
            delivery_id = selected_delivery()
            if delivery_id:
                path = self.create_delivery_document(delivery_id)
                if path:
                    self.print_text_file(path)

        def update_status():
            delivery_id = selected_delivery()
            if not delivery_id:
                return
            new_status = simpledialog.askstring(APP_NAME, "Nouveau statut: Prepare, En livraison, Livre, Annule", parent=self)
            if new_status not in {"Prepare", "En livraison", "Livre", "Annule"}:
                messagebox.showwarning(APP_NAME, "Statut invalide.")
                return
            delivered_at = datetime.now().isoformat(timespec="seconds") if new_status == "Livre" else ""
            self.db.execute("UPDATE delivery_notes SET status = ?, delivered_at = COALESCE(NULLIF(?, ''), delivered_at) WHERE id = ?", (new_status, delivered_at, delivery_id))
            self.show_delivery_notes()

        ttk.Button(actions, text="BL PDF", command=preview_delivery).pack(side="left", padx=(0, 8))
        ttk.Button(actions, text="Imprimer BL", command=print_delivery).pack(side="left", padx=(0, 8))
        ttk.Button(actions, text="Changer statut", style="Accent.TButton", command=update_status).pack(side="left")

    def show_invoices(self):
        self.clear("Factures/FNE", "CRUD factures, statut FNE, export officiel et documents professionnels.")
        body = ttk.Frame(self.main)
        body.pack(fill="both", expand=True, padx=24, pady=8)
        filters = self.card(body)
        filters.pack(fill="x")
        start_var = tk.StringVar(value=date.today().replace(day=1).isoformat())
        end_var = tk.StringVar(value=date.today().isoformat())
        fne_var = tk.StringVar(value="Tous")
        invoice_status_var = tk.StringVar(value="Tous")
        search_var = tk.StringVar()
        ttk.Label(filters, text="Debut", background=self.colors["surface"]).grid(row=0, column=0, sticky="w", padx=(0, 6))
        ttk.Entry(filters, textvariable=start_var, width=12).grid(row=0, column=1, sticky="w", padx=(0, 10))
        ttk.Label(filters, text="Fin", background=self.colors["surface"]).grid(row=0, column=2, sticky="w", padx=(0, 6))
        ttk.Entry(filters, textvariable=end_var, width=12).grid(row=0, column=3, sticky="w", padx=(0, 10))
        ttk.Label(filters, text="FNE", background=self.colors["surface"]).grid(row=0, column=4, sticky="w", padx=(0, 6))
        ttk.Combobox(filters, textvariable=fne_var, values=("Tous", "A exporter", "Exportee", "Certifiee", "Rejetee"), state="readonly", width=13).grid(row=0, column=5, sticky="w", padx=(0, 10))
        ttk.Label(filters, text="Facture", background=self.colors["surface"]).grid(row=0, column=6, sticky="w", padx=(0, 6))
        ttk.Combobox(filters, textvariable=invoice_status_var, values=("Tous", "Active", "Annulee"), state="readonly", width=12).grid(row=0, column=7, sticky="w", padx=(0, 10))
        ttk.Label(filters, text="Recherche", background=self.colors["surface"]).grid(row=1, column=0, sticky="w", padx=(0, 6), pady=(8, 0))
        ttk.Entry(filters, textvariable=search_var).grid(row=1, column=1, columnspan=5, sticky="ew", padx=(0, 10), pady=(8, 0))
        filters.columnconfigure(5, weight=1)

        summary = self.card(body)
        summary.pack(fill="x", pady=(12, 0))
        summary_text = tk.StringVar()
        ttk.Label(summary, textvariable=summary_text, style="CardTitle.TLabel").pack(anchor="w")

        actions = self.card(body)
        actions.pack(fill="x", pady=12)
        listing = self.card(body)
        listing.pack(fill="both", expand=True)
        holder = ttk.Frame(listing, style="Card.TFrame")
        holder.pack(fill="both", expand=True)
        tree_ref = {"tree": None}

        def filtered_invoices():
            clauses = ["date(created_at) BETWEEN ? AND ?"]
            params = [start_var.get().strip() or "1900-01-01", end_var.get().strip() or "2999-12-31"]
            if fne_var.get() != "Tous":
                clauses.append("fne_status = ?")
                params.append(fne_var.get())
            if invoice_status_var.get() != "Tous":
                clauses.append("invoice_status = ?")
                params.append(invoice_status_var.get())
            rows = self.db.all(f"SELECT * FROM sales WHERE {' AND '.join(clauses)} ORDER BY created_at DESC", params)
            query = search_var.get().strip().lower()
            if query:
                rows = [
                    row
                    for row in rows
                    if query in str(row["sale_no"]).lower()
                    or query in str(row["invoice_no"] or "").lower()
                    or query in str(row["customer_name"]).lower()
                    or query in str(row["fne_reference"] or "").lower()
                ]
            return rows

        def refresh(*_):
            for child in holder.winfo_children():
                child.destroy()
            rows = filtered_invoices()
            active_rows = [row for row in rows if row["invoice_status"] != "Annulee"]
            total = sum(float(row["total"] or 0) for row in active_rows)
            paid = sum(float(row["paid"] or 0) for row in active_rows)
            certified = sum(1 for row in active_rows if row["fne_status"] == "Certifiee")
            summary_text.set(f"{len(rows)} facture(s) | Total actif {self.money(total)} | Paye {self.money(paid)} | Certifiees FNE {certified}")
            tree_ref["tree"] = self.table(
                holder,
                ("ID", "Facture", "Vente", "Client", "Total", "Paye", "Paiement", "Facture", "FNE", "Reference FNE", "Date"),
                [
                    (
                        row["id"],
                        row["invoice_no"] or "",
                        row["sale_no"],
                        row["customer_name"],
                        self.money(row["total"]),
                        self.money(row["paid"]),
                        row["payment_method"],
                        row["invoice_status"] or "Active",
                        row["fne_status"] or "A exporter",
                        row["fne_reference"] or "",
                        row["created_at"][:16],
                    )
                    for row in rows
                ],
            )

        def selected_sale_id():
            tree = tree_ref["tree"]
            if not tree:
                return None
            selected = tree.selection()
            if not selected:
                messagebox.showwarning(APP_NAME, "Selectionnez une facture.")
                return None
            return int(tree.item(selected[0])["values"][0])

        def preview_invoice():
            sale_id = selected_sale_id()
            if not sale_id:
                return
            path = self.create_professional_invoice(sale_id)
            if path:
                self.show_receipt_window(path)

        def generate_pdf():
            sale_id = selected_sale_id()
            if not sale_id:
                return
            path = self.create_professional_invoice_pdf(sale_id)
            if path:
                messagebox.showinfo(APP_NAME, f"Facture PDF generee:\n{path}")
                self.open_document(path)

        def export_fne_one():
            sale_id = selected_sale_id()
            if not sale_id:
                return
            path = self.export_sale_fne(sale_id)
            if path:
                refresh()
                messagebox.showinfo(APP_NAME, f"Export FNE prepare:\n{path}")

        def export_fne_filtered():
            rows = [row for row in filtered_invoices() if row["invoice_status"] != "Annulee"]
            path = self.export_fne_batch(rows)
            if path:
                refresh()
                messagebox.showinfo(APP_NAME, f"Export FNE par lot prepare:\n{path}")

        def update_fne():
            sale_id = selected_sale_id()
            if not sale_id:
                return
            sale = self.db.one("SELECT * FROM sales WHERE id = ?", (sale_id,))
            window = tk.Toplevel(self)
            window.title("Suivi FNE")
            window.geometry("420x310")
            window.transient(self)
            status = tk.StringVar(value=sale["fne_status"] or "A exporter")
            reference = tk.StringVar(value=sale["fne_reference"] or "")
            note = tk.StringVar(value=sale["note"] or "")
            frame = self.card(window)
            frame.pack(fill="both", expand=True, padx=14, pady=14)
            ttk.Label(frame, text=f"Facture {sale['invoice_no'] or sale['sale_no']}", style="CardTitle.TLabel").grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 10))
            ttk.Label(frame, text="Statut FNE", background=self.colors["surface"]).grid(row=1, column=0, sticky="w", pady=6)
            ttk.Combobox(frame, textvariable=status, values=("A exporter", "Exportee", "Certifiee", "Rejetee"), state="readonly", width=26).grid(row=1, column=1, sticky="ew", padx=(10, 0), pady=6)
            self.form_row(frame, "Reference FNE", 2, reference, 28)
            self.form_row(frame, "Note", 3, note, 28)

            def save():
                if status.get() == "Certifiee" and not reference.get().strip():
                    messagebox.showwarning(APP_NAME, "La reference FNE est obligatoire pour certifier.")
                    return
                certified_at = datetime.now().isoformat(timespec="seconds") if status.get() == "Certifiee" else sale["fne_certified_at"]
                self.db.execute(
                    "UPDATE sales SET fne_status = ?, fne_reference = ?, fne_certified_at = ?, note = ? WHERE id = ?",
                    (status.get(), reference.get().strip(), certified_at, note.get().strip(), sale_id),
                )
                window.destroy()
                refresh()

            ttk.Button(frame, text="Enregistrer", style="Accent.TButton", command=save).grid(row=5, column=0, columnspan=2, sticky="ew", pady=(14, 0))
            frame.columnconfigure(1, weight=1)

        def toggle_cancel():
            sale_id = selected_sale_id()
            if not sale_id:
                return
            sale = self.db.one("SELECT * FROM sales WHERE id = ?", (sale_id,))
            new_status = "Active" if sale["invoice_status"] == "Annulee" else "Annulee"
            if new_status == "Annulee" and not self.require_admin("Annuler une facture"):
                return
            reason = sale["note"] or ""
            if new_status == "Annulee":
                reason = simpledialog.askstring(APP_NAME, "Motif d'annulation:", parent=self) or ""
            self.db.execute("UPDATE sales SET invoice_status = ?, note = ? WHERE id = ?", (new_status, reason, sale_id))
            refresh()

        ttk.Button(actions, text="Filtrer", style="Accent.TButton", command=refresh).pack(side="left", padx=(0, 8))
        ttk.Button(actions, text="Apercu facture", command=preview_invoice).pack(side="left", padx=(0, 8))
        ttk.Button(actions, text="Facture PDF", command=generate_pdf).pack(side="left", padx=(0, 8))
        ttk.Button(actions, text="Export FNE", command=export_fne_one).pack(side="left", padx=(0, 8))
        ttk.Button(actions, text="Export lot FNE", command=export_fne_filtered).pack(side="left", padx=(0, 8))
        ttk.Button(actions, text="Modifier FNE", command=update_fne).pack(side="left", padx=(0, 8))
        ttk.Button(actions, text="Annuler/Restaurer", style="Danger.TButton", command=toggle_cancel).pack(side="left")
        for variable in (start_var, end_var, fne_var, invoice_status_var, search_var):
            variable.trace_add("write", refresh)
        refresh()

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
        logo_var = tk.StringVar(value=self.setting("shop_logo_path"))
        logo_row = len(keys) + 1
        ttk.Label(body, text="Logo marchand", background="#ffffff").grid(row=logo_row, column=0, sticky="w", pady=5)
        logo_box = ttk.Frame(body, style="Card.TFrame")
        logo_box.grid(row=logo_row, column=1, sticky="ew", pady=5, padx=(10, 0))
        logo_box.columnconfigure(0, weight=1)
        ttk.Label(logo_box, textvariable=logo_var, style="MutedCard.TLabel", wraplength=360).grid(row=0, column=0, sticky="ew", padx=(0, 8))
        logo_preview = self.load_logo_image(max_width=120, max_height=70)
        if logo_preview:
            preview_label = ttk.Label(logo_box, image=logo_preview, background=self.colors["surface"])
            preview_label.image = logo_preview
            preview_label.grid(row=1, column=0, sticky="w", pady=(6, 0))

        def upload_logo():
            path = filedialog.askopenfilename(
                title="Choisir le logo du marchand",
                filetypes=[
                    ("Images", "*.png *.jpg *.jpeg *.webp *.gif"),
                    ("Tous les fichiers", "*.*"),
                ],
            )
            if not path:
                return
            source = Path(path)
            LOGO_UPLOADS_DIR.mkdir(parents=True, exist_ok=True)
            suffix = source.suffix.lower() or ".png"
            destination = LOGO_UPLOADS_DIR / f"logo-marchand-{datetime.now().strftime('%Y%m%d%H%M%S')}{suffix}"
            try:
                shutil.copy2(source, destination)
            except OSError as exc:
                messagebox.showerror(APP_NAME, f"Upload logo impossible: {exc}")
                return
            relative = str(destination.relative_to(BASE_DIR))
            logo_var.set(relative)
            self.db.execute("UPDATE settings SET value = ? WHERE key = 'shop_logo_path'", (relative,))
            self.logo_image = self.load_logo_image(max_width=180, max_height=120)
            self.splash_logo_image = self.load_project_logo_image(max_width=180, max_height=120)
            messagebox.showinfo(APP_NAME, "Logo marchand mis a jour.")
            if hasattr(self, "sidebar"):
                self.sidebar.destroy()
            if hasattr(self, "main"):
                self.main.destroy()
            self.build_shell()
            self.show_settings()

        def clear_logo():
            self.db.execute("UPDATE settings SET value = '' WHERE key = 'shop_logo_path'")
            logo_var.set("")
            self.logo_image = self.load_logo_image(max_width=180, max_height=120)
            self.splash_logo_image = self.load_project_logo_image(max_width=180, max_height=120)
            messagebox.showinfo(APP_NAME, "Logo marchand retire. Le logo par defaut sera utilise.")
            if hasattr(self, "sidebar"):
                self.sidebar.destroy()
            if hasattr(self, "main"):
                self.main.destroy()
            self.build_shell()
            self.show_settings()

        ttk.Button(logo_box, text="Uploader logo", style="Accent.TButton", command=upload_logo).grid(row=0, column=1, sticky="e", padx=(8, 0))
        ttk.Button(logo_box, text="Retirer", command=clear_logo).grid(row=1, column=1, sticky="e", padx=(8, 0), pady=(6, 0))

        def save():
            for key, var in vars_.items():
                self.db.execute("UPDATE settings SET value = ? WHERE key = ?", (var.get().strip(), key))
            self.db.execute("UPDATE settings SET value = ? WHERE key = ?", (theme_var.get().strip(), "theme"))
            self.db.execute("UPDATE settings SET value = ? WHERE key = ?", (logo_var.get().strip(), "shop_logo_path"))
            self.theme = theme_var.get().strip().lower() or "light"
            self.build_style()
            self.logo_image = self.load_logo_image(max_width=180, max_height=120)
            self.splash_logo_image = self.load_project_logo_image(max_width=180, max_height=120)
            messagebox.showinfo(APP_NAME, "Parametres enregistres.")
            self.show_settings()

        ttk.Button(body, text="Enregistrer", style="Accent.TButton", command=save).grid(row=len(keys) + 3, column=0, columnspan=2, sticky="ew", pady=(12, 0))

        if self.is_admin():
            users_box = self.card(self.main)
            users_box.pack(fill="both", expand=True, padx=24, pady=(0, 12))
            ttk.Label(users_box, text="Utilisateurs et PIN", style="CardTitle.TLabel").pack(anchor="w", pady=(0, 8))
            form = ttk.Frame(users_box, style="Card.TFrame")
            form.pack(fill="x", pady=(0, 10))
            user_name = tk.StringVar()
            user_role = tk.StringVar(value="Vendeur")
            user_pin = tk.StringVar()
            ttk.Label(form, text="Nom", background="#ffffff").grid(row=0, column=0, sticky="w", padx=(0, 6))
            ttk.Entry(form, textvariable=user_name, width=24).grid(row=0, column=1, sticky="ew", padx=(0, 10))
            ttk.Label(form, text="Role", background="#ffffff").grid(row=0, column=2, sticky="w", padx=(0, 6))
            ttk.Combobox(form, textvariable=user_role, values=("Admin", "Vendeur", "Magasinier", "Comptable"), state="readonly", width=16).grid(row=0, column=3, sticky="ew", padx=(0, 10))
            ttk.Label(form, text="PIN", background="#ffffff").grid(row=0, column=4, sticky="w", padx=(0, 6))
            ttk.Entry(form, textvariable=user_pin, show="*", width=12).grid(row=0, column=5, sticky="ew", padx=(0, 10))

            def add_user():
                name = user_name.get().strip()
                pin = user_pin.get().strip()
                if not name or len(pin) < 4:
                    messagebox.showwarning(APP_NAME, "Nom obligatoire et PIN de 4 caracteres minimum.")
                    return
                self.db.execute(
                    """
                    INSERT INTO users(name, role, pin_hash, active, created_at)
                    VALUES (?, ?, ?, 1, ?)
                    ON CONFLICT(name) DO UPDATE SET role = excluded.role, pin_hash = excluded.pin_hash, active = 1
                    """,
                    (name, user_role.get(), pin_hash(pin), datetime.now().isoformat(timespec="seconds")),
                )
                self.show_settings()

            ttk.Button(form, text="Ajouter / Modifier", style="Accent.TButton", command=add_user).grid(row=0, column=6, sticky="ew")
            user_tree = self.table(
                users_box,
                ("ID", "Nom", "Role", "Actif", "Date"),
                [(u["id"], u["name"], u["role"], "Oui" if u["active"] else "Non", u["created_at"][:10]) for u in self.db.all("SELECT * FROM users ORDER BY active DESC, role, name")],
            )

            def disable_user():
                selected = user_tree.selection()
                if not selected:
                    return
                user_id = int(user_tree.item(selected[0])["values"][0])
                if self.current_user and user_id == self.current_user["id"]:
                    messagebox.showwarning(APP_NAME, "Vous ne pouvez pas desactiver votre propre compte.")
                    return
                self.db.execute("UPDATE users SET active = 0 WHERE id = ?", (user_id,))
                self.show_settings()

            ttk.Button(users_box, text="Desactiver utilisateur selectionne", style="Danger.TButton", command=disable_user).pack(fill="x", pady=(8, 0))


if __name__ == "__main__":
    app = MagazinApp()
    app.mainloop()
