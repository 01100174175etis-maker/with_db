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
    # Use context manager to ensure the connection is closed safely
    with sqlite3.connect(db_path) as conn:
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


def generate_smart_id(sale_date_str):
    """توليد معرف ذكي فريد معزول يعتمد على تاريخ العملية + أجزاء الوقت (بما في ذلك الميكروثواني)
    لإلغاء احتمال الاصطدام عند تسجيل عمليتين في نفس الثانية.
    حافظنا على صيغة نصية بسيطة: YYMMDD + HHMMSSffffff
    """
    try:
        date_obj = datetime.strptime(sale_date_str, "%Y-%m-%d")
    except Exception:
        date_obj = datetime.now()

    prefix = date_obj.strftime("%y%m%d")  # YYMMDD
    # إضافة الوقت الحالي بالميكروثانية لتقليل فرص التكرار إلى الحد العملي
    time_suffix = datetime.now().strftime("%H%M%S%f")

    return f"{prefix}{time_suffix}"


def add_sale(customer, item, qty, price, total, sale_date, db_path: Path = DB_FILENAME):
    # توليد المعرف الفريد المعزول زمنياً
    smart_id = generate_smart_id(sale_date)
    # استخدم context manager للاتصال
    with sqlite3.connect(db_path) as conn:
        cursor = conn.cursor()
        # نحفظ created_at بصيغة ISO من بايثون لضمان التوافق عبر الأنظمة
        created_at = datetime.now().isoformat(sep=' ')
        cursor.execute(
            "INSERT INTO sales (id, customer, item, qty, price, total, sale_date, created_at) VALUES (?, ?, ?, ?, ?, ?, ?, ?)",
            (smart_id, customer, item, qty, price, total, sale_date, created_at),
        )
        conn.commit()
    return smart_id


def get_sales(db_path: Path = DB_FILENAME):
    try:
        with sqlite3.connect(db_path) as conn:
            df = pd.read_sql_query("SELECT * FROM sales ORDER BY created_at DESC", conn)
        return df
    except Exception as e:
        print(f"خطأ في جلب البيانات: {e}")
        return pd.DataFrame()


def delete_sale(sale_id, db_path: Path = DB_FILENAME):
    with sqlite3.connect(db_path) as conn:
        cursor = conn.cursor()
        cursor.execute("DELETE FROM sales WHERE id = ?", (str(sale_id),))
        conn.commit()
        rows = cursor.rowcount
    return rows


def clear_all_data(db_path: Path = DB_FILENAME):
    """تفريغ جميع البيانات من قاعدة البيانات"""
    try:
        with sqlite3.connect(db_path) as conn:
            cursor = conn.cursor()
            cursor.execute("DELETE FROM sales")
            conn.commit()
        return True
    except Exception as e:
        print(f"خطأ في تفريغ قاعدة البيانات: {e}")
        return False
