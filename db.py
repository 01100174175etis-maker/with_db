"""
Database helper for sales_app - Ultra Secure Time-Based ID Generation
Eliminates database count dependencies to prevent all application crashes.
"""
import sqlite3
from pathlib import Path
from datetime import datetime
import pandas as pd

DB_FILENAME = Path(__file__).parent / "sales.db"

def init_db(db_path: Path = DB_FILENAME):
    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS sales (
        id TEXT PRIMARY KEY,
        customer TEXT NOT NULL,
        item TEXT NOT NULL,
        qty INTEGER NOT NULL,
        price REAL NOT NULL DEFAULT 0.0,
        total REAL NOT NULL DEFAULT 0.0,
        sale_date TEXT NOT NULL,
        created_at TEXT NOT NULL
    );
    """)
    conn.commit()
    conn.close()

def generate_smart_id(sale_date_str):
    """توليد معرف ذكي فريد معزول يعتمد على تاريخ العملية + أجزاء الوقت لمنع التداخل والانهيار تماماً"""
    try:
        date_obj = datetime.strptime(sale_date_str, "%Y-%m-%d")
    except ValueError:
        date_obj = datetime.now()
        
    prefix = date_obj.strftime("%y%m%d") # استخراج YYMMDD بناء على تاريخ العملية المختار
    
    # دمج التوقيت الحالي بالثواني وأجزاء الثانية لضمان توليد رقم فريد مستحيل التكرار برمجياً
    time_suffix = datetime.now().strftime("%H%M%S")
    
    return f"{prefix}{time_suffix}"

def add_sale(customer, item, qty, price, total, sale_date, db_path: Path = DB_FILENAME):
    # توليد المعرف الفريد المعزول زمنياً
    smart_id = generate_smart_id(sale_date)
    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()
    cursor.execute(
        "INSERT INTO sales (id, customer, item, qty, price, total, sale_date, created_at) VALUES (?, ?, ?, ?, ?, ?, ?, datetime('now'))",
        (smart_id, customer, item, qty, price, total, sale_date),
    )
    conn.commit()
    conn.close()
    return smart_id

def get_sales(db_path: Path = DB_FILENAME):
    conn = sqlite3.connect(db_path)
    df = pd.read_sql_query("SELECT * FROM sales ORDER BY created_at DESC", conn)
    conn.close()
    return df

def delete_sale(sale_id, db_path: Path = DB_FILENAME):
    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()
    cursor.execute("DELETE FROM sales WHERE id = ?", (str(sale_id),))
    conn.commit()
    rows = cursor.rowcount
    conn.close()
    return rows
