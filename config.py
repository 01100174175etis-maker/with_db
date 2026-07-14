"""
ملف الإعدادات المركزي للبرنامج
Central configuration file for the application
"""
from pathlib import Path
from datetime import datetime

# =============== مسارات المشروع ===============
PROJECT_ROOT = Path(__file__).resolve().parent
DB_PATH = PROJECT_ROOT / "sales.db"
INVOICES_DIR = PROJECT_ROOT / "invoices"

# إنشاء مجلد الفواتير تلقائياً إن لم يكن موجوداً
INVOICES_DIR.mkdir(exist_ok=True)

# =============== إعدادات قاعدة البيانات ===============
DB_NAME = "sales.db"
DB_TIMEOUT = 5  # ثواني

# =============== إعدادات الواجهة الرسومية ===============
APP_TITLE = "نظام إدارة الحسابات والمبيعات المتطور - PyQt5 Enterprise"
APP_VERSION = "1.0.0"

WINDOW_WIDTH = 1250
WINDOW_HEIGHT = 720

# ألوان الواجهة
COLORS = {
    "primary": "#4CAF50",      # أخضر (الإضافة)
    "danger": "#f44336",       # أحمر (الحذف)
    "warning": "#ff9800",      # برتقالي (التحذير)
    "info": "#2196F3",         # أزرق (المعلومات)
    "secondary": "#00bcd4",    # سماوي (التعديل)
    "success": "#009688",      # أخضر داكن
    "gray": "#9e9e9e",         # رمادي
}

# =============== إعدادات الخطوط ===============
FONTS = {
    "header": ("Helvetica", 11, True),      # عنوان القسم
    "label": ("Helvetica", 12, True),       # تسميات الحقول
    "entry": ("Helvetica", 14, True),       # حقول الإدخال
    "button": ("Helvetica", 11, True),      # الأزرار
    "table": ("Helvetica", 11),             # الجدول
    "status": ("Helvetica", 10),            # شريط الحالة
}

# =============== إعدادات صيغ التاريخ والوقت ===============
DATE_FORMAT = "%Y-%m-%d"
DATETIME_FORMAT = "%Y-%m-%d %H:%M:%S"

# =============== إعدادات الأعمدة ===============
TABLE_HEADERS = [
    "المعرف الذكي",
    "اسم العميل",
    "المنتج",
    "الكمية",
    "السعر",
    "الإجمالي",
    "تاريخ البيع",
    "وقت الإدخال"
]

HEADERS_MAPPING = {
    "id": "المعرف الذكي",
    "customer": "اسم العميل",
    "item": "المنتج",
    "qty": "الكمية",
    "price": "السعر",
    "total": "الإجمالي",
    "sale_date": "تاريخ البيع",
    "created_at": "وقت الإدخال"
}

# =============== إعدادات تنسيق Excel ===============
EXCEL_ENCODING = "utf-8"
EXCEL_HEADER_FONT_SIZE = 14
EXCEL_DATA_FONT_SIZE = 12
EXCEL_DEFAULT_COLUMN_WIDTH = 18

# =============== إعدادات السلوك والميزات ===============
HIGHLIGHT_LARGE_TRANSACTIONS = True      # تمييز العمليات الكبيرة
LARGE_TRANSACTION_THRESHOLD = 10000      # حد العملية الكبيرة
HIGHLIGHT_COLOR = "#ffeb3b"              # لون التمييز (أصفر فاتح)

# =============== رسائل ومؤشرات النظام ===============
MESSAGES = {
    "ready": "النظام جاهز للعمل",
    "data_loaded": "تم تحميل البيانات بنجاح",
    "sale_added": "تم إضافة المبيعة بنجاح",
    "sale_edited": "تم تعديل المبيعة بنجاح",
    "sale_deleted": "تم حذف المبيعة بنجاح",
    "export_success": "تم التصدير بنجاح",
    "invalid_data": "بيانات غير صحيحة",
    "missing_data": "يوجد حقول مفقودة",
    "file_saved": "تم حفظ الملف بنجاح",
}

# =============== إعدادات الأداء ===============
BATCH_SIZE = 100              # حجم المعالجة الدفعية
CACHE_SIZE = 500              # حجم الذاكرة المؤقتة
REFRESH_INTERVAL = 1000       # فترة التحديث (ميلي ثانية)

# =============== إعدادات Windows-specific ===============
WINDOWS_ENCODING = "utf-8"
SUPPORTS_STARTFILE = True      # دعم os.startfile على Windows

# =============== إعدادات السجلات والتتبع ===============
ENABLE_LOGGING = True
LOG_LEVEL = "INFO"  # DEBUG, INFO, WARNING, ERROR, CRITICAL
LOG_FILE = PROJECT_ROOT / "app.log"
