"""
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ Fardan Apex --- Serializer ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
This application is responsible for serializing order items in the Fardan Apex system.

Author: Behnam Rabieyan
Company: Garma Gostar Fardan
Created: 2025
"""

# Standard library imports
import getpass
import json
import os
import re
import sys
from datetime import datetime

# Third-party library imports
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from cryptography.fernet import Fernet

# PyQt5 imports
from PyQt5.QtCore import Qt, QTimer, pyqtSignal
from PyQt5.QtGui import (
    QColor, QFont, QFontDatabase, QIcon, QIntValidator, QPixmap, QTextOption
)
from PyQt5.QtWidgets import (
    QApplication, QComboBox, QDialog, QFileDialog, QFormLayout, QGraphicsDropShadowEffect,
    QGroupBox, QHeaderView, QHBoxLayout, QLabel, QLineEdit, QMessageBox, QPushButton,
    QProgressDialog, QMainWindow, QSizePolicy, QTabWidget, QTableWidget, QTableWidgetItem,
    QTextEdit, QVBoxLayout, QWidget, QListWidget, QInputDialog, QListWidgetItem,
    QSplashScreen, QProgressBar
)



# ---------- تنظیمات ----------
ADMIN_USER = "BenRabin"       # نام کاربری ادمین در ویندوز
ADMIN_PASSWORD = "123.0"     # رمز مدیریت کاربران

# کلید ثابت
SECRET_KEY = b"SnZFJqzdj1xx6rxksdPL5P_-UKijvx4DRlR0a5-s1lQ="  
cipher = Fernet(SECRET_KEY)

EXCEL_FILE = r"\\fileserver\Mohandesi\سفارش ها\order.xlsx"
SHEET_NAME = "order"
TABLE_NAME = "ordertable"
ALLOWED_USERS = []


# ------------------ Helpers for resource paths ------------------
def resource_path(relative_path: str) -> str:
    """
    Return absolute path to resource, works for dev and for PyInstaller onefile.
    Use this for images, fonts, icons that are bundled with --add-data.
    """
    try:
        base_path = sys._MEIPASS  # PyInstaller extracted temp dir
    except Exception:
        base_path = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base_path, relative_path)


def app_dir_path(relative_path: str) -> str:
    """
    Return path next to the executable (where settings should live).
    Use this for writable files (settings.json) so the program reads/writes next to exe.
    """
    if getattr(sys, "frozen", False):
        base = os.path.dirname(sys.executable)
    else:
        base = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base, relative_path)


SETTINGS_FILE = app_dir_path("settings.json")


HEADERS = [
    "ردیف",
    "تاریخ",
    "شماره سفارش",
    "نوع محصول",
    "کد محصول",
    "تعداد",
    "ردیف آیتم",
    "سریال سفارش",
    "توضیحات",
    "کاربر ثبت",
    "تاریخ ثبت",
]

PRODUCT_MAP = {
    "MF": "F",
    "MR": "R",
    "MU": "U",
    "نفراست": "ن",
    "فویلی": "ف",
    "هیتر سیمی": "س",
    "لوله ای دیفراست": "د",
    "ترموسوییچ": "TS",
    "ترموفیوز": "TF"
}

GROUP_M = {
    "MF",
    "MR",
    "MU"
}


    # ---------- بررسی محدوده جدول در اکسل ----------
def update_excel_table_range(ws, table_name):
    """
    به‌روزرسانی محدوده جدول اکسل بعد از اضافه کردن داده جدید
    ws: نام شیت در اکسل
    table_name: نام جدول در اکسل
    """
    try:
        table = ws.tables[table_name]
        start_cell, end_cell = table.ref.split(':')
        min_col = ws[start_cell].col_idx
        min_row = ws[start_cell].row
        max_col = ws[end_cell].col_idx
        max_row = ws.max_row
        table.ref = f"{get_column_letter(min_col)}{min_row}:{get_column_letter(max_col)}{max_row}"
    except KeyError:
        QMessageBox.warning(
            None,
            "هشدار",
            f"جدول '{table_name}' پیدا نشد. داده‌ها ذخیره شدند ولی جدول آپدیت نشد."
        )


# ---------- بررسی فایل سفارش ----------
def ensure_excel(show_message=True):
    if not os.path.exists(EXCEL_FILE):
        if show_message:
            QMessageBox.warning(
                None,
                "هشدار",
                f"فایل اکسل یافت نشد:\n{EXCEL_FILE}\nلطفاً مسیر درست را در تنظیمات وارد کنید."
            )
        return False
    
    return True


# ---------- نرمال‌سازی حروف فارسی (عربی -> فارسی و trim) ----------
def normalize_farsi(text: str) -> str:
    if not text:
        return ""
    
    replacements = {
        "ي": "ی",
        "ك": "ک",
        "ة": "ه",
        "ۀ": "ه"
    }

    for src, dst in replacements.items():
        text = text.replace(src, dst)

    return re.sub(r"\s+", " ", text).strip()


# ---------- بارگذاری تنظیمات از settings.json ----------
def load_settings():
    try:
        if os.path.exists(SETTINGS_FILE):
            with open(SETTINGS_FILE, "rb") as f:
                encrypted = f.read()
                if not encrypted:
                    return {}
                raw = cipher.decrypt(encrypted)   # رمزگشایی
                data = json.loads(raw.decode("utf-8"))
                if isinstance(data, dict):
                    return data
    except Exception as e:
        print("خطا در بارگزاری تنظیمات:", e)

    return {}

# ---------- ذخیره دیکشنری در settings.json ----------
def save_settings(data: dict):
    try:
        raw = json.dumps(data, ensure_ascii=False, indent=2).encode("utf-8")
        encrypted = cipher.encrypt(raw)   # رمزنگاری
        with open(SETTINGS_FILE, "wb") as f:
            f.write(encrypted)
        return True
    except Exception as e:
        print("خطا در ذخیره تنظیمات:", e)
        return False


# ---------- پیدا کردن بیشترین ها (برای ردیف آیتم و شماره ردیف) ----------
def compute_maxes(ws):
    max_groupA = 0
    max_groupB = 0
    max_rowid = 0
    for row in ws.iter_rows(min_row=2, values_only=True):
        try:
            rowid = int(row[0]) if row[0] is not None and str(row[0]).strip() != "" else 0
        except:
            rowid = 0
        if rowid > max_rowid:
            max_rowid = rowid

        ptype = normalize_farsi(row[3]) if row[3] is not None else ""
        try:
            item_idx = int(row[6]) if row[6] is not None and str(row[6]).strip() != "" else 0
        except:
            item_idx = 0

        # تشخیص گروه میله‌ای: اگر متن نوع محصول MF/MR/MU باشه
        if str(ptype).upper() in GROUP_M:
            if item_idx > max_groupA:
                max_groupA = item_idx
        else:
            if item_idx > max_groupB:
                max_groupB = item_idx

    return max_groupA, max_groupB, max_rowid


# ---------- تابع ساخت سریال سفارش و ردیف آیتم بر اساس الگوریتم ----------
def next_item_and_serial(ws, date_text, product_type, max_groupA, max_groupB):
    p = normalize_farsi(product_type)
    key = p
    if re.match(r"^[A-Za-z]{1,4}$", p):
        key = p.upper()
    abbrev = PRODUCT_MAP.get(key, PRODUCT_MAP.get(key.lower(), None))
    if not abbrev:
        if key.upper() in GROUP_M:
            abbrev = key.upper()[0]
        else:
            abbrev = "0"

    yyyy = "0000"
    if date_text:
        m = re.search(r"\d{4}", date_text)
        if m:
            yyyy = m.group(0)
        else:
            yyyy = date_text[:4] if len(date_text) >= 4 else "0000"

    in_groupA = (str(key).upper() in GROUP_M)

    if in_groupA:
        max_groupA += 1
        item_idx = max_groupA
    else:
        max_groupB += 1
        item_idx = max_groupB

    serial = f"{item_idx}-{yyyy}-{abbrev}"

    return item_idx, serial, max_groupA, max_groupB


# ---------- حذف تمام ردیف‌های یک سفارش ----------
def delete_order_rows(ws, order_no):
    to_delete = []

    for idx, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
        if str(row[2]) == str(order_no):
            to_delete.append(idx)

    for r in reversed(to_delete):
        ws.delete_rows(r, 1)


# ---------- استایل ----------
APP_STYLESHEET = """
QWidget {
    background: #f5f7fb;
    font-family: 'Segoe UI', Tahoma, Arial;
}

QLineEdit, QTextEdit {
    background: white;
    border: 1px solid #d0d7df;
    border-radius: 6px;
    padding: 6px;
}

QTextEdit {
    font-family: Consolas, 'Courier New', monospace;
}

QPushButton {
    background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
        stop:0 #5aa9ff, stop:1 #2e7dff);
    color: white;
    border: none;
    padding: 8px 12px;
    border-radius: 8px;
}

QPushButton:hover {
    background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
        stop:0 #6bb8ff, stop:1 #3b8bff);
}

QPushButton#secondary {
    background: #eef4ff;
    color: #1a3b6e;
    border: 1px solid #d0dbff;
}

QTableWidget {
    background: white;
    border: 1px solid #e0e6ef;
    gridline-color: #f1f5fb;
}

QHeaderView::section {
    background: #eef4ff;
    padding: 6px;
    border: none;
}

QComboBox {
    background: white;
    border: 1px solid #d0d7df;
    border-radius: 6px;
    padding: 4px;
}

QTabBar::tab {
    background: transparent;
    padding: 8px 16px;
}

QTabWidget::pane {
    border: none;
}
"""


# ---------- Dialog افزودن/ویرایش محصول ----------
class ProductDialog(QDialog):
    product_added = pyqtSignal(tuple)  # (ptype, code, qty)

    def __init__(self, parent=None, preset=None):
        super().__init__(parent)
        self.setWindowTitle("افزودن/ویرایش محصول")
        self.setFixedSize(460, 220)

        shadow = QGraphicsDropShadowEffect(self)
        shadow.setBlurRadius(18)
        shadow.setXOffset(0)
        shadow.setYOffset(6)
        shadow.setColor(QColor(0, 0, 0, 60))
        self.setGraphicsEffect(shadow)

        font = QFont()
        font.setPointSize(10)
        self.setFont(font)

        form = QFormLayout()
        form.setLabelAlignment(Qt.AlignRight)
        self.setLayout(form)

        self.cb_type = QComboBox()

        # لیست نوع محصول در پنجره افزودن محصولات
        self.cb_type.addItems(
            [
            '',
            'فویلی',
            'هیتر سیمی',
            'نفراست',
            'لوله ای دیفراست',
            'ترموفیوز',
            'ترموسوییچ',
            'لوله استیل قطر 7 (60میل)',
            'MF',
            'MR',
            'MU'
            ]
        )
        self.cb_type.setEditable(True)
        self.cb_type.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        form.addRow("نوع محصول ", self.cb_type)

        self.e_code = QLineEdit()
        self.e_code.setAlignment(Qt.AlignRight)
        form.addRow("کد محصول ", self.e_code)

        self.e_qty = QLineEdit()
        self.e_qty.setValidator(QIntValidator(1, 10000000, self))
        form.addRow("تعداد ", self.e_qty)

        if preset:
            self.cb_type.setCurrentText(preset[0])
            self.e_code.setText(preset[1])
            try:
                self.e_qty.setText(str(int(preset[2])))
            except:
                self.e_qty.setText("")
        else:
            self.e_qty.setText("")

        btn_layout = QHBoxLayout()
        btn_register = QPushButton("ثبت")
        btn_register.setFixedWidth(120)
        btn_close = QPushButton("بستن")
        btn_close.setObjectName("secondary")
        btn_close.setFixedWidth(120)
        btn_layout.addStretch()
        btn_layout.addWidget(btn_register)
        btn_layout.addWidget(btn_close)
        form.addRow(btn_layout)

        btn_register.clicked.connect(self.on_register)
        btn_close.clicked.connect(self.reject)


    def on_register(self):
        ptype = normalize_farsi(self.cb_type.currentText())
        code = normalize_farsi(self.e_code.text())

        try:
            qty = int(self.e_qty.text())
            if qty <= 0:
                raise ValueError
        except:
            QMessageBox.critical(self, "خطا", "تعداد نامعتبر است. لطفا عددی بزرگتر از صفر وارد کنید.")
            return

        if not ptype or not code:
            QMessageBox.critical(self, "خطا", "همه فیلدها الزامی هستند.")
            return

        # ارسال داده به بیرون بدون بستن دیالوگ
        self.product_added.emit((ptype, code, qty))

        # ریست کردن فیلدها
        self.cb_type.setCurrentIndex(0)
        self.e_code.clear()
        self.e_qty.setText("")
        self.cb_type.setFocus()


# ---------- تابع Preloader ----------
def show_splash():
    app = QApplication(sys.argv)

    # تصویر PNG برای Splash Screen
    splash_pix = QPixmap(resource_path("SerializerFardanApex.png"))
    splash = QSplashScreen(splash_pix, Qt.WindowStaysOnTopHint)
    splash.setMask(splash_pix.mask())

    # نوار پیشرفت زیر تصویر
    progress = QProgressBar(splash)
    progress.setGeometry(90, splash_pix.height() - 100, splash_pix.width() - 180, 20)
    progress.setMaximum(100)
    progress.setValue(0)
    progress.setStyleSheet("""
        QProgressBar {
            border: 1px solid grey;
            border-radius: 5px;
            text-align: center;
        }
        QProgressBar::chunk {
            background-color: #2e7dff;
            width: 1px;
        }
    """)

    splash.show()

    window = App()

    counter = 0
    def update_progress():
        nonlocal counter
        counter += 1
        progress.setValue(counter)
        if counter >= 100:
            timer.stop()
            splash.close()
            window.show()

    timer = QTimer()
    timer.timeout.connect(update_progress)
    timer.start(30)

    sys.exit(app.exec_())


# ---------- کلاس اصلی ----------
class App(QMainWindow):
    def __init__(self):
        super().__init__()

        settings = load_settings()

        global EXCEL_FILE, SHEET_NAME, TABLE_NAME, ALLOWED_USERS
        EXCEL_FILE = settings.get("excel_file", EXCEL_FILE)
        SHEET_NAME = settings.get("sheet_name", SHEET_NAME)
        TABLE_NAME = settings.get("table_name", TABLE_NAME)
        ALLOWED_USERS = settings.get("allowed_users", ALLOWED_USERS)

        ensure_excel(show_message=False)
        
        self.setWindowTitle("تولید سریال سفارش - FardanApex")
        self.resize(900, 500)
        self.setStyleSheet(APP_STYLESHEET)

        central = QWidget(); self.setCentralWidget(central)
        shadow = QGraphicsDropShadowEffect(self)
        shadow.setBlurRadius(24); shadow.setXOffset(0); shadow.setYOffset(8); shadow.setColor(QColor(0,0,0,40))
        central.setGraphicsEffect(shadow)

        main_layout = QVBoxLayout(central)

        self.tabs = QTabWidget(); self.tabs.setDocumentMode(True); self.tabs.setTabPosition(QTabWidget.North)
        # وضعیت نمایشی تب انتخاب شده
        self.tabs.setStyleSheet("""
        QTabBar::tab {
            background: #f1f5f9;
            padding: 7px 14px;
            border: none;
            margin-right: 0px;
            font-size: 12px;
            color: #6b7280;
        }
        QTabBar::tab:selected {
            background: #ffffff;
            color: #111827;
            font-weight: bold;
            border-bottom: 3px solid #2563eb;
        }
        QTabBar::tab:!selected {
            background: #e5e7eb;
        }
        QTabWidget::pane {
            border-top: 2px solid #d1d5db;
            background: #ffffff;
        }
        """)
        main_layout.addWidget(self.tabs)

        self.tab_new = QWidget(); self.tab_search = QWidget(); self.tab_option = QWidget()
        self.tabs.addTab(self.tab_new, "ثبت سفارش جدید")
        self.tabs.addTab(self.tab_search, "جستجو و ویرایش")
        self.tabs.addTab(self.tab_option, "ویژگی‌ها")

        self.build_tab_new()
        self.build_tab_search()
        self.build_tab_option()


    # ---------- پنجره سفارش جدید ----------
    def build_tab_new(self):
        # لایه افقی: چپ فرم و جدول | راست پنل سریال‌ها
        main_hbox = QHBoxLayout(); self.tab_new.setLayout(main_hbox)
        left_layout = QVBoxLayout(); right_layout = QVBoxLayout()
        main_hbox.addLayout(left_layout, 3)
        main_hbox.addLayout(right_layout, 1)

        # بالای فرم
        top_layout = QHBoxLayout()
        lbl_order = QLabel("شماره سفارش "); top_layout.addWidget(lbl_order, 0, Qt.AlignRight)
        self.new_order_no = QLineEdit(); self.new_order_no.setAlignment(Qt.AlignRight); self.new_order_no.setFixedWidth(220)
        top_layout.addWidget(self.new_order_no)
        top_layout.addSpacing(12)
        lbl_date = QLabel("تاریخ "); top_layout.addWidget(lbl_date, 0, Qt.AlignRight)
        self.new_date = QLineEdit(); self.new_date.setAlignment(Qt.AlignRight); self.new_date.setFixedWidth(160)
        top_layout.addWidget(self.new_date)
        # دکمه ثبت سفارش جدید
        btn_new_top = QPushButton("ثبت سفارش جدید")
        btn_new_top.setObjectName("secondary")
        btn_new_top.clicked.connect(self.reset_new_order_form)
        top_layout.addWidget(btn_new_top)
        top_layout.addStretch()
        left_layout.addLayout(top_layout)

        # توضیحات
        desc_layout = QHBoxLayout()
        desc_layout.addWidget(QLabel("توضیحات "), 0, Qt.AlignRight)
        self.new_desc = QLineEdit(); self.new_desc.setAlignment(Qt.AlignRight)
        desc_layout.addWidget(self.new_desc)
        left_layout.addLayout(desc_layout)

        # جدول محصولات
        self.table_new = QTableWidget(0, 3)
        self.table_new.setHorizontalHeaderLabels(["نوع محصول", "کد محصول", "تعداد"])
        header = self.table_new.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.Stretch)
        header.setSectionResizeMode(1, QHeaderView.Stretch)
        header.setSectionResizeMode(2, QHeaderView.Stretch)
        self.table_new.verticalHeader().setVisible(False)
        left_layout.addWidget(self.table_new)

        # دکمه‌ها (افزودن/ویرایش/حذف/ذخیره)
        btn_layout = QHBoxLayout()
        btn_add = QPushButton("افزودن محصول")
        btn_edit = QPushButton("ویرایش محصول")
        btn_del = QPushButton("حذف محصول")
        btn_save = QPushButton("ذخیره سفارش")
        btn_add.clicked.connect(self.add_product_new)
        btn_edit.clicked.connect(self.edit_product_new)
        btn_del.clicked.connect(lambda: self.delete_selected(self.table_new))
        btn_save.clicked.connect(self.save_order_new_with_progress)
        btn_layout.addWidget(btn_add)
        btn_layout.addWidget(btn_edit)
        btn_layout.addWidget(btn_del)
        btn_layout.addStretch()
        btn_layout.addWidget(btn_save)
        left_layout.addLayout(btn_layout)

        # ----- پنل سریال‌ها -----
        right_layout.addWidget(QLabel("سریال‌های این سفارش:"))
        self.serial_box = QTextEdit()
        self.serial_box.setReadOnly(True)
        self.serial_box.setWordWrapMode(QTextOption.NoWrap)
        # بسیار مهم: چپ به راست برای کپی راحت‌تر
        self.serial_box.setLayoutDirection(Qt.LeftToRight)
        self.serial_box.setFont(QFont("Consolas", 10))
        right_layout.addWidget(self.serial_box)
        btn_copy = QPushButton("کپی سریال‌ها")
        btn_copy.setFixedWidth(130)
        btn_copy.clicked.connect(self.copy_serials)
        right_layout.addWidget(btn_copy)
        right_layout.addStretch()


    # ---------- پنجره جستجو ----------
    def build_tab_search(self):
        main_hbox = QHBoxLayout(); self.tab_search.setLayout(main_hbox)
        left_layout = QVBoxLayout(); right_layout = QVBoxLayout()
        main_hbox.addLayout(left_layout, 3)
        main_hbox.addLayout(right_layout, 1)

    # بالای فرم
        top_layout = QHBoxLayout()
        top_layout.addWidget(QLabel("شماره سفارش "), 0, Qt.AlignRight)
        self.search_order_no = QLineEdit(); self.search_order_no.setAlignment(Qt.AlignRight); self.search_order_no.setFixedWidth(220)
        top_layout.addWidget(self.search_order_no)
        btn_search = QPushButton("جستجو"); btn_search.clicked.connect(self.search_order_with_progress); top_layout.addWidget(btn_search)
        top_layout.addSpacing(12)
        top_layout.addWidget(QLabel("تاریخ "), 0, Qt.AlignRight)
        self.search_date = QLineEdit(); self.search_date.setAlignment(Qt.AlignRight); self.search_date.setFixedWidth(160)
        top_layout.addWidget(self.search_date)
        top_layout.addStretch(); left_layout.addLayout(top_layout)

    # توضیحات
        desc_layout = QHBoxLayout(); desc_layout.addWidget(QLabel("توضیحات "), 0, Qt.AlignRight)
        self.search_desc = QLineEdit(); self.search_desc.setAlignment(Qt.AlignRight); desc_layout.addWidget(self.search_desc)
        left_layout.addLayout(desc_layout)

    # جدول محصولات
        self.table_search = QTableWidget(0, 3)
        self.table_search.setHorizontalHeaderLabels(["نوع محصول", "کد محصول", "تعداد"])
        hdr2 = self.table_search.horizontalHeader()
        hdr2.setSectionResizeMode(0, QHeaderView.Stretch)
        hdr2.setSectionResizeMode(1, QHeaderView.Stretch)
        hdr2.setSectionResizeMode(2, QHeaderView.Stretch)
        self.table_search.verticalHeader().setVisible(False)
        left_layout.addWidget(self.table_search)

    # دکمه‌ها
        btn_layout = QHBoxLayout()
        btn_add = QPushButton("افزودن محصول")
        btn_edit = QPushButton("ویرایش محصول")
        btn_save = QPushButton("ذخیره تغییرات")
        btn_add.clicked.connect(self.add_product_search)
        btn_edit.clicked.connect(self.edit_product_search)
        btn_save.clicked.connect(self.save_changes_search_with_progress)
        btn_layout.addWidget(btn_add)
        btn_layout.addWidget(btn_edit)
        btn_layout.addStretch()
        btn_layout.addWidget(btn_save)
        left_layout.addLayout(btn_layout)

    # پنل سمت راست: سریال‌ها و دکمه کپی
        right_layout.addWidget(QLabel("سریال‌های این سفارش:"))
        self.serial_box_search = QTextEdit()
        self.serial_box_search.setReadOnly(True)
        self.serial_box_search.setWordWrapMode(QTextOption.NoWrap)
        self.serial_box_search.setLayoutDirection(Qt.LeftToRight)
        self.serial_box_search.setFont(QFont("Consolas", 10))
        right_layout.addWidget(self.serial_box_search)
        btn_copy = QPushButton("کپی سریال‌ها")
        btn_copy.setFixedWidth(130)
        btn_copy.clicked.connect(self.copy_serials_search)
        right_layout.addWidget(btn_copy)
        right_layout.addStretch()


    # ---------- پنجره آپشن ----------
    def build_tab_option(self):
        # main vertical container for the tab
        main_layout = QVBoxLayout()
        self.tab_option.setLayout(main_layout)

        # horizontal main: left = users (narrow), right = settings (expandable)
        h_main = QHBoxLayout()
        main_layout.addLayout(h_main)

        # ---------------- قاب مدیریت کاربران (سمت چپ، کوچک) ----------------
        self.user_group = QGroupBox("کاربران مجاز")
        self.user_group.setFixedWidth(280)
        user_group_layout = QVBoxLayout()
        user_group_layout.setAlignment(Qt.AlignTop)
        self.user_group.setAlignment(Qt.AlignRight)
        self.user_group.setLayout(user_group_layout)

        # بالای قاب: دکمه افزودن (سمت چپ) و فیلد نام کاربر (سمت راست)
        user_top_row = QHBoxLayout()
        user_top_row.setAlignment(Qt.AlignTop)
        btn_add_user = QPushButton("افزودن")
        btn_add_user.setFixedWidth(80)
        btn_add_user.clicked.connect(self.add_user)
        self.user_input = QLineEdit()
        self.user_input.setPlaceholderText("نام کاربری جدید...")
        self.user_input.setFixedWidth(170)
        # جای دکمه و فیلد عوض شد: دکمه اول، فیلد بعدی
        user_top_row.addWidget(btn_add_user)
        user_top_row.addSpacing(8)
        user_top_row.addWidget(self.user_input)
        user_top_row.addStretch()
        user_group_layout.addLayout(user_top_row)

        user_group_layout.addSpacing(6)

        # لیست کاربران — راست‌به‌چپ محیط برنامه حفظ می‌شود ولی خود لیست چپ-به-راست خواهد بود
        self.user_list = QListWidget()
        self.user_list.setMinimumHeight(260)
        self.user_list.setLayoutDirection(Qt.LeftToRight)
        user_group_layout.addWidget(self.user_list)

        # اضافه کردن قاب کاربران به ستون چپ، aligned top تا جمع بمونه
        h_main.addWidget(self.user_group, 0, Qt.AlignTop)

        # ---------------- قاب تنظیمات (سمت راست، بالا تراز) ----------------
        group_settings = QGroupBox("تنظیمات اکسل")
        settings_layout = QVBoxLayout()
        settings_layout.setAlignment(Qt.AlignTop)
        group_settings.setAlignment(Qt.AlignRight)
        group_settings.setLayout(settings_layout)

        # آدرس فایل اکسل (عنوان راست‌چین)
        lbl_file_layout = QHBoxLayout()
        lbl_file_layout.addStretch()
        lbl_file = QLabel("آدرس فایل اکسل")
        lbl_file_layout.addWidget(lbl_file)
        settings_layout.addLayout(lbl_file_layout)

        # ردیف: دکمه انتخاب فایل (چپ) + فیلد مسیر (راست)
        file_row = QHBoxLayout()
        btn_browse = QPushButton("انتخاب فایل")
        btn_browse.setFixedWidth(120)
        btn_browse.clicked.connect(self.browse_excel_file)
        self.e_excel_path = QLineEdit(self)
        self.e_excel_path.setText(EXCEL_FILE)
        file_row.addWidget(btn_browse)
        file_row.addWidget(self.e_excel_path)
        settings_layout.addLayout(file_row)

        settings_layout.addSpacing(12)

        # ردیف: نام برگه و نام جدول کنار هم (هر کدام با label بالای خودش)
        row2 = QHBoxLayout()
        # نام برگه
        sheet_v = QVBoxLayout()
        sheet_label_row = QHBoxLayout()
        sheet_label_row.addStretch()
        sheet_label_row.addWidget(QLabel("نام برگه"))
        sheet_v.addLayout(sheet_label_row)
        self.e_sheet_name = QLineEdit(self)
        self.e_sheet_name.setText(SHEET_NAME)
        sheet_v.addWidget(self.e_sheet_name)
        # نام جدول
        table_v = QVBoxLayout()
        table_label_row = QHBoxLayout()
        table_label_row.addStretch()
        table_label_row.addWidget(QLabel("نام جدول"))
        table_v.addLayout(table_label_row)
        self.e_table_name = QLineEdit(self)
        self.e_table_name.setText(TABLE_NAME)
        table_v.addWidget(self.e_table_name)

        row2.addLayout(sheet_v)
        row2.addSpacing(20)
        row2.addLayout(table_v)
        settings_layout.addLayout(row2)

        # اضافه کردن قاب تنظیمات به ستون راست و تراز بالا
        h_main.addWidget(group_settings, 1, Qt.AlignTop)

        # ---------------- دکمه‌های پایین تب ----------------
        bottom_row = QHBoxLayout()
        
        # دکمه درباره برنامه سمت چپ
        btn_about = QPushButton("درباره برنامه")
        btn_about.setFixedWidth(120)
        btn_about.clicked.connect(self.show_about)
        bottom_row.addWidget(btn_about, 0, Qt.AlignLeft)
        
        bottom_row.addStretch()
        
        # دکمه ذخیره سمت راست
        btn_save = QPushButton("ذخیره")
        btn_save.setFixedWidth(120)
        btn_save.clicked.connect(self.save_options)
        bottom_row.addWidget(btn_save, 0, Qt.AlignRight)
        
        main_layout.addLayout(bottom_row)

        # ---------- آماده‌سازی و populate اولیه لیست کاربران ----------
        # ساختار نگهداری آیتم‌ها: username -> {"item", "widget", "button", "label"}
        self.user_widgets = {}

        # populate initial users from ALLOWED_USERS using add_user_item (که پایین اضافه میشه)
        for u in ALLOWED_USERS:
            try:
                self.add_user_item(u)
            except Exception:
                pass


    # تابع انتخاب فایل
    def browse_excel_file(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "انتخاب فایل اکسل", "", "Excel Files (*.xlsx)")
        if file_path:
            self.e_excel_path.setText(file_path)


    # ---------- درباره برنامه ----------
    def show_about(self):
        dlg = QDialog(self)
        dlg.setWindowTitle("About")
        dlg.setFixedSize(500, 450)

        main_layout = QVBoxLayout(dlg)
        main_layout.setContentsMargins(20, 20, 20, 20)
        main_layout.setSpacing(15)

        intro_layout = QVBoxLayout()
        intro_layout.setSpacing(0)
        intro_layout.setContentsMargins(0, 0, 0, 0)
        lbl_intro = QLabel(
            "<h3><b>Fardan Apex — Serializer</b></h3>"
            "<h4>Production Serial Generator Application</h4><br>"
            "This application is designed to generate production series after order confirmation by the engineering unit.<br><br>"
            "Version: 2.1.7 — © 2025 All Rights Reserved<br>"
            "Developed exclusively for:<br>"
            "Garma Gostar Fardan Co."
        )

        lbl_intro.setWordWrap(True)
        lbl_intro.setAlignment(Qt.AlignLeft)
        intro_layout.addWidget(lbl_intro)

        logo = QLabel()
        logo_pix = QPixmap(resource_path("FardanLogo.jpg"))
        if logo_pix.isNull():
            logo_pix = QPixmap(resource_path("FardanLogoEN.png"))
        if not logo_pix.isNull():
            logo.setPixmap(logo_pix.scaledToWidth(175, Qt.SmoothTransformation))
        else:
            logo.setText("Fardan Apex")
        logo.setAlignment(Qt.AlignLeft | Qt.AlignTop)
        logo.setContentsMargins(35, 10, 0, 0)
        intro_layout.addWidget(logo)

        main_layout.addLayout(intro_layout)

        dev_layout = QVBoxLayout()
        dev_layout.setSpacing(0)
        dev_layout.setContentsMargins(5, 0, 0, 5)
        font_id = QFontDatabase.addApplicationFont(resource_path("BrittanySignature.ttf"))
        if font_id != -1:
            font_family = QFontDatabase.applicationFontFamilies(font_id)[0]
        else:
            font_family = "Sans Serif"

        lbl_dev = QLabel(
            f"<b>Design & Development:</b><br>"
            f"<span style='font-family:\"{font_family}\"; font-size:20pt; color:#4169E1;'>&nbsp;&nbsp;&nbsp;&nbsp;Behnam Rabieyan</span><br>"
            "website: behnamrabieyan.ir | E-mail: info@behnamrabieyan.ir"
        )

        lbl_dev.setWordWrap(True)
        lbl_dev.setAlignment(Qt.AlignLeft)
        dev_layout.addWidget(lbl_dev)

        main_layout.addLayout(dev_layout)

        dlg.exec_()


    # ---------- مدیریت آیتم‌های محصول ----------
    def add_product_new(self):
        dlg = ProductDialog(self)

        def add_row(data):
            row = self.table_new.rowCount()
            self.table_new.insertRow(row)
            self.table_new.setItem(row, 0, QTableWidgetItem(str(data[0])))
            self.table_new.setItem(row, 1, QTableWidgetItem(str(data[1])))
            self.table_new.setItem(row, 2, QTableWidgetItem(str(data[2])))

        dlg.product_added.connect(add_row)
        dlg.exec_()  # مودال می‌ماند ولی بعد از ثبت بسته نمی‌شود


    def edit_product_new(self):
        row = self.table_new.currentRow()
        if row == -1:
            QMessageBox.warning(self, "توجه", "ابتدا یک محصول را انتخاب کنید.")
            return

        preset = [self.table_new.item(row, i).text() if self.table_new.item(row, i) else "" for i in range(3)]
        dlg = ProductDialog(self, preset)


        def update_row(data):
            for col, val in enumerate(data):
                self.table_new.setItem(row, col, QTableWidgetItem(str(val)))

        dlg.product_added.connect(update_row)
        dlg.exec_()


    def add_product_search(self):
        dlg = ProductDialog(self)


        def add_row(data):
            row = self.table_search.rowCount()
            self.table_search.insertRow(row)
            self.table_search.setItem(row, 0, QTableWidgetItem(str(data[0])))
            self.table_search.setItem(row, 1, QTableWidgetItem(str(data[1])))
            self.table_search.setItem(row, 2, QTableWidgetItem(str(data[2])))

        dlg.product_added.connect(add_row)
        dlg.exec_()


    def edit_product_search(self):
        row = self.table_search.currentRow()
        if row == -1:
            QMessageBox.warning(self, "توجه", "ابتدا یک محصول را انتخاب کنید.")
            return

        preset = [self.table_search.item(row, i).text() if self.table_search.item(row, i) else "" for i in range(3)]
        dlg = ProductDialog(self, preset)


        def update_row(data):
            for col, val in enumerate(data):
                self.table_search.setItem(row, col, QTableWidgetItem(str(val)))

        dlg.product_added.connect(update_row)
        dlg.exec_()


    def delete_selected(self, table):
        indexes = table.selectionModel().selectedRows()
        for idx in sorted([r.row() for r in indexes], reverse=True):
            table.removeRow(idx)


    # ---------- ذخیره سفارش جدید ----------
    def save_order_new(self):
        if not ensure_excel(show_message=True):
           return
        date_text = normalize_farsi(self.new_date.text())
        order_no = normalize_farsi(self.new_order_no.text())
        desc = normalize_farsi(self.new_desc.text())

        if not date_text or not order_no:
            QMessageBox.critical(self, "خطا", "تاریخ و شماره سفارش الزامی هستند.")
            return

        items = []
        for row in range(self.table_new.rowCount()):
            ptype = normalize_farsi(self.table_new.item(row, 0).text() if self.table_new.item(row, 0) else "")
            code = normalize_farsi(self.table_new.item(row, 1).text() if self.table_new.item(row, 1) else "")
            try:
                qty = int(self.table_new.item(row, 2).text())
                if qty <= 0:
                    raise ValueError
            except:
                QMessageBox.critical(self, "خطا", f"تعداد نامعتبر برای آیتم: {ptype} - {code}")
                return
            items.append((ptype, code, qty))

        if not items:
            QMessageBox.critical(self, "خطا", "حداقل یک محصول باید اضافه شود.")
            return

        try:
            wb = load_workbook(EXCEL_FILE)
            ws = wb[SHEET_NAME]
            username = getpass.getuser()
            now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        except PermissionError:
            QMessageBox.critical(self, "خطا", f"فایل {EXCEL_FILE} باز است. لطفاً آن را ببندید.")
            return
        
        except Exception as e:
            QMessageBox.critical(self, "خطا", f"خطا در باز کردن فایل: {e}")
            return
        
        maxA, maxB, max_rowid = compute_maxes(ws)
        serial_lines = []

        for ptype, code, qty in items:
            item_idx, serial, maxA, maxB = next_item_and_serial(ws, date_text, ptype, maxA, maxB)
            max_rowid += 1
            ws.append(
                [
                    max_rowid,
                    date_text,
                    order_no,
                    ptype,
                    code,
                    qty,
                    item_idx,
                    serial,
                    desc,
                    username,
                    now_str,
                ]
            )
            serial_lines.append('\u200E' + serial)

    # آپدیت محدوده جدول
        update_excel_table_range(ws, TABLE_NAME)

        try:
            wb.save(EXCEL_FILE)
            QMessageBox.information(self, "موفق", "سفارش با موفقیت ثبت شد. سریال‌ها در پنل سمت راست درج شدند.")
            self.serial_box.clear()
            self.serial_box.setPlainText("\n".join(serial_lines))
        except PermissionError:
            QMessageBox.critical(self, "خطا", f"فایل {EXCEL_FILE} باز است. لطفاً آن را ببندید.")
        except Exception as e:
            QMessageBox.critical(self, "خطا", f"خطا در ذخیره‌سازی: {e}")


# ---------- پروسه ذخیره سفارش جدید ----------
    def save_order_new_with_progress(self):
        progress = QProgressDialog("در حال ذخیره سفارش، لطفا صبر کنید...", "لغو", 0, 0, self)
        progress.setWindowTitle("در حال پردازش")
        progress.setWindowModality(Qt.ApplicationModal)
        progress.setMinimumDuration(0)
        progress.show()
        QApplication.processEvents()

        try:
            self.save_order_new()
        finally:
            progress.close()


    # ---------- دکمه ثبت سفارش جدید: پاکسازی فرم و پنل ----------
    def reset_new_order_form(self):
        if hasattr(self, "new_order_no"):
            self.new_order_no.clear()
        if hasattr(self, "new_date"):
            self.new_date.setText("")
        if hasattr(self, "new_desc"):
            self.new_desc.clear()
        if hasattr(self, "table_new"):
            self.table_new.setRowCount(0)
        if hasattr(self, "serial_box"):
            self.serial_box.clear()


    # ---------- جستجو سفارش ----------
    def search_order(self):
        if not ensure_excel(show_message=True):
            return
        
        order_no = normalize_farsi(self.search_order_no.text())

        if not order_no:
            QMessageBox.critical(self, "خطا", "شماره سفارش را وارد کنید.")
            return

        try:
            wb = load_workbook(EXCEL_FILE)
            ws = wb[SHEET_NAME]
        except PermissionError:
            QMessageBox.critical(self, "خطا", f"فایل {EXCEL_FILE} باز است.")
            return
        
        rows = list(ws.iter_rows(min_row=2, values_only=True))
        found_rows = [r for r in rows if str(r[2]) == str(order_no)]

        if not found_rows:
            QMessageBox.information(self, "یافت نشد", "سفارشی با این شماره پیدا نشد.")
            self.search_date.clear(); self.search_desc.clear(); self.table_search.setRowCount(0); self.serial_box_search.clear()
            return
        
        first = found_rows[0]
        self.search_date.setText(first[1] if first[1] is not None else "")
        self.search_desc.setText(first[8] if first[8] is not None else "")
        self.table_search.setRowCount(0)
        serial_lines = []
        for r in found_rows:
            row_idx = self.table_search.rowCount()
            self.table_search.insertRow(row_idx)
            self.table_search.setItem(row_idx, 0, QTableWidgetItem(str(r[3] if r[3] is not None else "")))
            self.table_search.setItem(row_idx, 1, QTableWidgetItem(str(r[4] if r[4] is not None else "")))
            self.table_search.setItem(row_idx, 2, QTableWidgetItem(str(r[5] if r[5] is not None else "")))
            serial = str(r[7] if r[7] is not None else "")
            serial_lines.append('\u200E' + serial)
        self.serial_box_search.setPlainText("\n".join(serial_lines))


    # ---------- پروسه جستجو سفارش ----------
    def search_order_with_progress(self):
        progress = QProgressDialog("در حال جستجوی سفارش، لطفا صبر کنید...", "لغو", 0, 0, self)
        progress.setWindowTitle("در حال پردازش")
        progress.setWindowModality(Qt.ApplicationModal)
        progress.setMinimumDuration(0)
        progress.show()
        QApplication.processEvents()

        try:
            self.search_order()
        finally:
            progress.close()


    # ---------- ذخیر تغییرات در تب جستجو ----------
    def save_changes_search(self):
        if not ensure_excel(show_message=True):
            return
        
        order_no = normalize_farsi(self.search_order_no.text())
        date_text = normalize_farsi(self.search_date.text())
        desc = normalize_farsi(self.search_desc.text())
        if not order_no or not date_text:
            QMessageBox.critical(self, "خطا", "شماره سفارش و تاریخ الزامی هستند.")
            return

        try:
            wb = load_workbook(EXCEL_FILE)
            ws = wb[SHEET_NAME]
            username = getpass.getuser()
            now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        except PermissionError:
            QMessageBox.critical(self, "خطا", f"فایل {EXCEL_FILE} باز است. لطفا ببندید و دوباره تلاش کنید.")
            return

        existing_rows = []
        for idx, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
            if str(row[2]) == str(order_no):
                existing_rows.append(
                    {
                    "ws_idx": idx,
                    "rowid": row[0],
                    "date": row[1],
                    "ptype": row[3],
                    "code": row[4],
                    "qty": row[5],
                    "item_idx": row[6],
                    "serial": row[7],
                    "desc": row[8],
                    "username": row[9],
                    "now_str": row[10]
                    }
                )

        maxA, maxB, max_rowid = compute_maxes(ws)
        serial_lines = []

        for row_idx in range(self.table_search.rowCount()):
            ptype_new = normalize_farsi(self.table_search.item(row_idx, 0).text() if self.table_search.item(row_idx, 0) else "")
            code_new = normalize_farsi(self.table_search.item(row_idx, 1).text() if self.table_search.item(row_idx, 1) else "")
            try:
                qty_new = int(self.table_search.item(row_idx, 2).text())
            except:
                QMessageBox.critical(self, "خطا", f"تعداد نامعتبر در ردیف {row_idx+1}: {ptype_new} - {code_new}")
                return

            if row_idx < len(existing_rows):
                row_data = existing_rows[row_idx]
                ws_idx = row_data["ws_idx"]

                if row_data["date"] != date_text:
                    ws.cell(row=ws_idx, column=2, value=date_text)
                if row_data["desc"] != desc:
                    ws.cell(row=ws_idx, column=9, value=desc)

                if row_data["code"] != code_new:
                    ws.cell(row=ws_idx, column=5, value=code_new)
                if row_data["qty"] != qty_new:
                    ws.cell(row=ws_idx, column=6, value=qty_new)

                if row_data["ptype"] != ptype_new:
                    ws.cell(row=ws_idx, column=4, value=ptype_new)
                    parts = str(row_data["serial"]).split("-")
                    if len(parts) == 3:
                        item_idx = parts[0]
                        year_part = parts[1]
                        key = ptype_new
                        if re.match(r"^[A-Za-z]{1,4}$", key):
                            key = key.upper()
                        abbrev = PRODUCT_MAP.get(key, PRODUCT_MAP.get(key.lower(), None))
                        if not abbrev:
                            abbrev = "0"
                        new_serial = f"{item_idx}-{year_part}-{abbrev}"
                        ws.cell(row=ws_idx, column=8, value=new_serial)
                        serial_lines.append('\u200E' + new_serial)
                    else:
                        serial_lines.append('\u200E' + row_data["serial"])
                else:
                    serial_lines.append('\u200E' + row_data["serial"])
                # write audit columns (col 10 -> user, col 11 -> timestamp)
                ws.cell(row=ws_idx, column=10, value=username)
                ws.cell(row=ws_idx, column=11, value=now_str)

            else:
                max_rowid += 1
                item_idx, serial, maxA, maxB = next_item_and_serial(ws, date_text, ptype_new, maxA, maxB)
                ws.append(
                    [
                    max_rowid,
                    date_text,
                    order_no,
                    ptype_new,
                    code_new,
                    qty_new,
                    item_idx,
                    serial,
                    desc,
                    username,
                    now_str,
                    ]
                )

                serial_lines.append('\u200E' + serial)

    # آپدیت محدوده جدول
        update_excel_table_range(ws, TABLE_NAME)

        try:
            wb.save(EXCEL_FILE)
            QMessageBox.information(self, "موفق", "تغییرات با موفقیت ذخیره شد.")
            self.serial_box_search.setPlainText("\n".join(serial_lines))
        except PermissionError:
            QMessageBox.critical(self, "خطا", f"فایل {EXCEL_FILE} باز است. لطفا آن را ببندید.")
        except Exception as e:
            QMessageBox.critical(self, "خطا", f"خطا در ذخیره‌سازی: {e}")


# ---------- پروسه ذخیره تغییرات سفارش ----------
    def save_changes_search_with_progress(self):
        progress = QProgressDialog("در حال ذخیره تغییرات، لطفا صبر کنید...", "لغو", 0, 0, self)
        progress.setWindowTitle("در حال پردازش")
        progress.setWindowModality(Qt.ApplicationModal)
        progress.setMinimumDuration(0)
        progress.show()
        QApplication.processEvents()

        try:
            self.save_changes_search()  # اجرای تابع اصلی
        finally:
            progress.close()


    # ---------- کپی کردن سریال در ثبت سفارش ----------
    def copy_serials(self):
            text = self.serial_box.toPlainText()
            if not text.strip():
                QMessageBox.information(self, "هشدار", "هیچ سریالی برای کپی وجود ندارد.")
                return
            
            QApplication.clipboard().setText(text)
            QMessageBox.information(self, "کپی شد", "سریال‌ها به کلیپ‌بورد کپی شدند.")


    # ---------- ذخیر کپی کردن سریال در ویرایش ----------
    def copy_serials_search(self):
            text = self.serial_box_search.toPlainText()
            if not text.strip():
                QMessageBox.information(self, "هشدار", "هیچ سریالی برای کپی وجود ندارد.")
                return
            
            QApplication.clipboard().setText(text)
            QMessageBox.information(self, "کپی شد", "سریال‌ها به کلیپ‌بورد کپی شدند.")


    # ---------- افزودن و حذف کاربر ----------
    def ask_admin_password(self):
        """Ask for admin password (admin password is hardcoded in code)."""
        pwd, ok = QInputDialog.getText(
            self,
            "رمز مدیر",
            "رمز عبور مدیر را وارد کنید:",
            QLineEdit.Password
        )

        if not ok:
            return False
            
        # ADMIN_PASSWORD is expected to be defined in code (hardcoded)
        return pwd == ADMIN_PASSWORD


    def add_user(self):
        """Add user to UI list only (do not save to JSON here)."""
        if not self.ask_admin_password():
            QMessageBox.warning(self, "خطا", "رمز عبور نادرست یا عملیات لغو شد.")
            return

        new_user = normalize_farsi(self.user_input.text().strip())
        if not new_user:
            QMessageBox.information(self, "توجه", "نام کاربر خالی است.")
            return

        if new_user == ADMIN_USER:
            QMessageBox.information(self, "توجه", "این حساب متعلق به ادمین است.")
            return

        if new_user in getattr(self, "user_widgets", {}):
            QMessageBox.information(self, "توجه", "این کاربر قبلاً اضافه شده است.")
            return

        # اضافه کردن آیتم سفارشی
        self.add_user_item(new_user)
        self.user_input.clear()
        QMessageBox.information(self, "انجام شد", f"کاربر '{new_user}' به لیست اضافه شد.")


    def remove_user(self):
        """Remove selected user from UI list only (asks admin password)."""
        if not self.ask_admin_password():
            QMessageBox.warning(self, "خطا", "رمز عبور نادرست یا عملیات لغو شد.")
            return

        item = self.user_list.currentItem()
        if not item:
            QMessageBox.information(self, "توجه", "ابتدا یک کاربر را انتخاب کنید.")
            return

        username = item.data(Qt.UserRole)
        if username == ADMIN_USER:
            QMessageBox.warning(self, "خطا", "نمی‌توان ادمین را حذف کرد.")
            return

        row = self.user_list.row(item)
        self.user_list.takeItem(row)
        if username in getattr(self, "user_widgets", {}):
            del self.user_widgets[username]

        QMessageBox.information(self, "انجام شد", f"کاربر '{username}' از لیست حذف شد.")


    # ---------- ذخیر تنظیمات ----------
    def add_user_item(self, username):
        """
        یک آیتم سفارشی به QListWidget اضافه می‌کند که شامل دکمه حذف (✘) و نام کاربر است.
        دکمه حذف همیشه فضای خودش را رزرو می‌کند ولی پیش‌فرض غیرفعال/نامرئی است.
        """
        username = str(username)
        if username in getattr(self, "user_widgets", {}):
            return

        item = QListWidgetItem()
        item.setData(Qt.UserRole, username)

        widget = QWidget()
        hl = QHBoxLayout()
        hl.setContentsMargins(6, 4, 6, 4)

        # دکمه حذف (فقط فضای خودش را اشغال می‌کند، پیش‌فرض غیرقابل کلیک و شفاف)
        btn_del = QPushButton("✘")
        btn_del.setStyleSheet("""
            QPushButton {
                background: transparent;
                color: red;
                border: none;
                font-weight: bold;
                padding-top: 4px;
                padding-bottom: 4px;
            }
        """)
        btn_del.clicked.connect(lambda _, u=username: self.on_delete_button_clicked(u))

        lbl = QLabel(username)
        lbl.setAlignment(Qt.AlignRight | Qt.AlignVCenter)

        # ترتیب: دکمه (چپ) - لیبل - فاصله
        hl.addWidget(btn_del)
        hl.addSpacing(6)
        hl.addWidget(lbl)
        hl.addStretch()
        widget.setLayout(hl)

        item.setSizeHint(widget.sizeHint())
        self.user_list.addItem(item)
        self.user_list.setItemWidget(item, widget)

        if not hasattr(self, "user_widgets"):
            self.user_widgets = {}
        self.user_widgets[username] = {"item": item, "widget": widget, "button": btn_del, "label": lbl}


    # ---------- کلید حذف کاربران ----------
    def on_delete_button_clicked(self, username):
        """
        وقتی دکمهٔ ✘ کنار نام کلیک شد: رمز مدیر خواسته میشه و پس از معتبر بودن، کاربر حذف می‌شود.
        """
        if not self.ask_admin_password():
            QMessageBox.warning(self, "خطا", "رمز عبور نادرست یا عملیات لغو شد.")
            return

        entry = self.user_widgets.get(username)
        if not entry:
            QMessageBox.information(self, "توجه", "کاربر مورد نظر پیدا نشد.")
            return

        item = entry["item"]
        row = self.user_list.row(item)
        self.user_list.takeItem(row)
        try:
            del self.user_widgets[username]
        except KeyError:
            pass

        QMessageBox.information(self, "انجام شد", f"کاربر '{username}' حذف شد.")


    # ---------- ذخیر تنظیمات ----------
    def save_options(self):
        global EXCEL_FILE, SHEET_NAME, TABLE_NAME, ALLOWED_USERS

        EXCEL_FILE = self.e_excel_path.text().strip() or EXCEL_FILE
        SHEET_NAME = self.e_sheet_name.text().strip() or SHEET_NAME
        TABLE_NAME = self.e_table_name.text().strip() or TABLE_NAME

        # collect users from UI (custom widgets)
        if hasattr(self, "user_widgets"):
            ALLOWED_USERS = list(self.user_widgets.keys())
        else:
            ALLOWED_USERS = [self.user_list.item(i).text() for i in range(self.user_list.count())]

        settings = {
            "excel_file": EXCEL_FILE,
            "sheet_name": SHEET_NAME,
            "table_name": TABLE_NAME,
            "allowed_users": ALLOWED_USERS,
        }

        ok = save_settings(settings)
        if ok:
            QMessageBox.information(self, "ذخیره شد", "تنظیمات با موفقیت ذخیره شد.")
        else:
            QMessageBox.warning(self, "خطا", "خطا در ذخیره تنظیمات (لطفاً دسترسی نوشتن را بررسی کنید).")


# ---------- اجرای برنامه ----------
if __name__ == "__main__":
    app = QApplication(sys.argv)
    # اضافه کردن فونت IRAN
    QFontDatabase.addApplicationFont(resource_path("IRAN.ttf"))
    app.setFont(QFont("IRAN", 10))
    # اضافه کردن آیکون
    app.setWindowIcon(QIcon(resource_path("icon.ico")))
    
    show_splash()
