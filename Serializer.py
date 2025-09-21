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
from functools import wraps


# Third-party library imports
from cryptography.fernet import Fernet
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter


# PyQt5 imports
from PyQt5.QtCore import Qt, QTimer, pyqtSignal
from PyQt5.QtGui import (
    QColor, QFont, QFontDatabase, QIcon, QIntValidator, QPixmap, QTextOption
)
from PyQt5.QtWidgets import (
    QApplication, QComboBox, QDialog, QFileDialog, QFormLayout,
    QGraphicsDropShadowEffect, QGroupBox, QHeaderView, QHBoxLayout, QLabel,
    QLineEdit, QMessageBox, QPushButton, QProgressDialog, QMainWindow,
    QSizePolicy, QTabWidget, QTableWidget, QTableWidgetItem, QTextEdit,
    QVBoxLayout, QWidget, QListWidget, QInputDialog, QListWidgetItem,
    QSplashScreen, QProgressBar
)


# ---------- تنظیمات ----------
ADMIN_USER = "BenRabin"
""" نام کاربری ادمین در ویندوز """
ADMIN_PASSWORD = "123.0"
""" رمز مدیریت کاربران """

SECRET_KEY = b"SnZFJqzdj1xx6rxksdPL5P_-UKijvx4DRlR0a5-s1lQ="
""" کلید امنیتی """
cipher = Fernet(SECRET_KEY)

EXCEL_FILE = r"\\fileserver\Mohandesi\سفارش ها\orders.xlsx"
SHEET_NAME = "order"
TABLE_NAME = "ordertable"
ALLOWED_USERS = []


# ------------------ Helpers for resource paths ------------------
def resource_path(relative_path: str) -> str:
    """
    Return absolute path to resource, works for dev and for PyInstaller.
    """
    try:
        base_path = sys._MEIPASS  # PyInstaller extracted temp dir
    except AttributeError:
        base_path = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base_path, relative_path)


def app_dir_path(relative_path: str) -> str:
    """
    Return path next to the executable (where settings should live).
    """
    if getattr(sys, "frozen", False):
        base = os.path.dirname(sys.executable)
    else:
        base = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base, relative_path)


SETTINGS_FILE = app_dir_path("settings.json")

HEADERS = [
    "ردیف", "تاریخ", "شماره سفارش", "نوع محصول", "کد محصول", "تعداد",
    "ردیف آیتم", "سریال سفارش", "توضیحات", "کاربر ثبت", "تاریخ ثبت",
]

PRODUCT_MAP = {
    "MF": "F", "MR": "R", "MU": "U", "نفراست": "ن", "فویلی": "ف",
    "هیتر سیمی": "س", "لوله ای دیفراست": "د", "ترموسوییچ": "TS",
    "ترموفیوز": "TF"
}

GROUP_M = {"MF", "MR", "MU"}


# ---------- توابع کاربردی (Helpers) ----------
def update_excel_table_range(ws, table_name):
    """به‌روزرسانی محدوده جدول اکسل بعد از اضافه کردن داده جدید."""
    try:
        table = ws.tables[table_name]
        start_cell, end_cell = table.ref.split(':')
        min_col = ws[start_cell].col_idx
        min_row = ws[start_cell].row
        max_col = ws[end_cell].col_idx
        max_row = ws.max_row
        new_ref = (
            f"{get_column_letter(min_col)}{min_row}:"
            f"{get_column_letter(max_col)}{max_row}"
        )
        table.ref = new_ref
    except KeyError:
        QMessageBox.warning(
            None, "هشدار",
            f"جدول '{table_name}' پیدا نشد. "
            f"داده‌ها ذخیره شدند ولی جدول آپدیت نشد."
        )


def ensure_excel(show_message=True):
    """بررسی وجود فایل اکسل سفارشات."""
    if not os.path.exists(EXCEL_FILE):
        if show_message:
            QMessageBox.warning(
                None, "هشدار",
                f"فایل اکسل یافت نشد:\n{EXCEL_FILE}\n"
                f"لطفاً مسیر درست را در تنظیمات وارد کنید."
            )
        return False
    return True


def normalize_farsi(text: str) -> str:
    """نرمال‌سازی حروف فارسی (عربی -> فارسی و trim)."""
    if not text:
        return ""
    replacements = {"ي": "ی", "ك": "ک", "ة": "ه", "ۀ": "ه"}
    for src, dst in replacements.items():
        text = text.replace(src, dst)
    return re.sub(r"\s+", " ", text).strip()


def load_settings():
    """بارگذاری تنظیمات از فایل settings.json."""
    try:
        if os.path.exists(SETTINGS_FILE):
            with open(SETTINGS_FILE, "rb") as f:
                encrypted = f.read()
                if not encrypted:
                    return {}
                raw = cipher.decrypt(encrypted)
                data = json.loads(raw.decode("utf-8"))
                if isinstance(data, dict):
                    return data
    except Exception as e:
        print("خطا در بارگزاری تنظیمات:", e)
    return {}


def save_settings(data: dict):
    """ذخیره دیکشنری تنظیمات در فایل settings.json."""
    try:
        raw = json.dumps(data, ensure_ascii=False, indent=2).encode("utf-8")
        encrypted = cipher.encrypt(raw)
        with open(SETTINGS_FILE, "wb") as f:
            f.write(encrypted)
        return True
    except Exception as e:
        print("خطا در ذخیره تنظیمات:", e)
        return False


def compute_maxes(ws):
    """پیدا کردن بیشترین مقادیر برای ردیف آیتم و شماره ردیف."""
    max_groupA = 0
    max_groupB = 0
    max_rowid = 0
    for row in ws.iter_rows(min_row=2, values_only=True):
        try:
            rowid = int(row[0]) if row[0] is not None else 0
        except (ValueError, TypeError):
            rowid = 0
        if rowid > max_rowid:
            max_rowid = rowid

        ptype = normalize_farsi(row[3]) if row[3] is not None else ""
        try:
            item_idx = int(row[6]) if row[6] is not None else 0
        except (ValueError, TypeError):
            item_idx = 0

        if str(ptype).upper() in GROUP_M:
            if item_idx > max_groupA:
                max_groupA = item_idx
        else:
            if item_idx > max_groupB:
                max_groupB = item_idx
    return max_groupA, max_groupB, max_rowid


def next_item_and_serial(ws, date_text, product_type, max_groupA, max_groupB):
    """ساخت سریال سفارش و ردیف آیتم بر اساس الگوریتم."""
    p = normalize_farsi(product_type)
    key = p.upper() if re.match(r"^[A-Za-z]{1,4}$", p) else p
    abbrev = PRODUCT_MAP.get(key, "0")

    yyyy = "0000"
    if date_text:
        m = re.search(r"\d{4}", date_text)
        yyyy = m.group(0) if m else date_text[:4]

    in_groupA = (str(key).upper() in GROUP_M)
    if in_groupA:
        max_groupA += 1
        item_idx = max_groupA
    else:
        max_groupB += 1
        item_idx = max_groupB

    serial = f"{item_idx}-{yyyy}-{abbrev}"
    return item_idx, serial, max_groupA, max_groupB


def delete_order_rows(ws, order_no):
    """حذف تمام ردیف‌های یک سفارش."""
    to_delete = [
        idx for idx, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2)
        if str(row[2]) == str(order_no)
    ]
    for r in reversed(to_delete):
        ws.delete_rows(r, 1)


# ---------- استایل برنامه ----------
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
QTextEdit { font-family: Consolas, 'Courier New', monospace; }
QPushButton {
    background: qlineargradient(x1:0, y1:0, x2:0, y2:1, stop:0 #5aa9ff, stop:1 #2e7dff);
    color: white;
    border: none;
    padding: 8px 12px;
    border-radius: 8px;
}
QPushButton:hover {
    background: qlineargradient(x1:0, y1:0, x2:0, y2:1, stop:0 #6bb8ff, stop:1 #3b8bff);
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
QTabBar::tab { background: transparent; padding: 8px 16px; }
QTabWidget::pane { border: none; }
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
        self.cb_type.addItems([
            '', 'فویلی', 'هیتر سیمی', 'نفراست', 'لوله ای دیفراست',
            'ترموفیوز', 'ترموسوییچ', 'لوله استیل قطر 7 (60میل)',
            'MF', 'MR', 'MU'
        ])
        self.cb_type.setEditable(True)
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
            self.e_qty.setText(str(preset[2]))
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
        qty_text = self.e_qty.text()

        if not ptype or not code or not qty_text:
            QMessageBox.critical(self, "خطا", "همه فیلدها الزامی هستند.")
            return

        try:
            qty = int(qty_text)
            if qty <= 0:
                raise ValueError
        except ValueError:
            QMessageBox.critical(
                self, "خطا",
                "تعداد نامعتبر است. لطفا عددی بزرگتر از صفر وارد کنید."
            )
            return

        self.product_added.emit((ptype, code, qty))

        # Reset fields for next entry
        self.cb_type.setCurrentIndex(0)
        self.e_code.clear()
        self.e_qty.clear()
        self.cb_type.setFocus()


# ---------- تابع Preloader ----------
def show_splash():
    app = QApplication.instance()
    splash_pix = QPixmap(resource_path("SerializerFardanApex.png"))
    splash = QSplashScreen(splash_pix, Qt.WindowStaysOnTopHint)
    splash.setMask(splash_pix.mask())

    progress = QProgressBar(splash)
    progress.setGeometry(
        90, splash_pix.height() - 100, splash_pix.width() - 180, 20
    )
    progress.setMaximum(100)
    progress.setValue(0)
    progress.setStyleSheet("""
        QProgressBar { border: 1px solid grey; border-radius: 5px; text-align: center; }
        QProgressBar::chunk { background-color: #2e7dff; width: 1px; }
    """)
    splash.show()

    main_window = App()

    def update_progress(val):
        progress.setValue(val)
        if val >= 100:
            timer.stop()
            splash.close()
            main_window.show()

    timer = QTimer()
    # Simple animation logic
    values = list(range(1, 101))
    current_step = 0

    def next_step():
        nonlocal current_step
        if current_step < len(values):
            update_progress(values[current_step])
            current_step += 1

    timer.timeout.connect(next_step)
    timer.start(25)  # Update every 25ms

    app.exec_()

# ---------- Decorator for Progress Dialog ----------
def with_progress_dialog(title, label):
    """Decorator to show a progress dialog while a function is running."""
    def decorator(func):
        @wraps(func)
        def wrapper(self, *args, **kwargs):
            progress = QProgressDialog(label, "لغو", 0, 0, self)
            progress.setWindowTitle(title)
            progress.setWindowModality(Qt.ApplicationModal)
            progress.setMinimumDuration(0)
            progress.show()
            QApplication.processEvents()
            try:
                # The fix is here: Call the original function without extra args
                return func(self)
            finally:
                progress.close()
        return wrapper
    return decorator


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

        central = QWidget()
        self.setCentralWidget(central)
        shadow = QGraphicsDropShadowEffect(self)
        shadow.setBlurRadius(24)
        shadow.setXOffset(0)
        shadow.setYOffset(8)
        shadow.setColor(QColor(0, 0, 0, 40))
        central.setGraphicsEffect(shadow)

        main_layout = QVBoxLayout(central)
        self.tabs = QTabWidget()
        self.tabs.setDocumentMode(True)
        self.tabs.setStyleSheet("""
            QTabBar::tab:selected {
                background: #ffffff; color: #111827; font-weight: bold;
                border-bottom: 3px solid #2563eb;
            }
            QTabBar::tab:!selected { background: #e5e7eb; }
            QTabWidget::pane { border-top: 2px solid #d1d5db; background: #ffffff; }
        """)
        main_layout.addWidget(self.tabs)

        self.tab_new = QWidget()
        self.tab_search = QWidget()
        self.tab_option = QWidget()
        self.tabs.addTab(self.tab_new, "ثبت سفارش جدید")
        self.tabs.addTab(self.tab_search, "جستجو و ویرایش")
        self.tabs.addTab(self.tab_option, "ویژگی‌ها")

        self.build_tab_new()
        self.build_tab_search()
        self.build_tab_option()

    # ---------- متدهای Refactor شده ----------
    def _create_order_tab_widgets(self, is_search_tab=False):
        """A factory method to build common UI widgets for order tabs."""
        top_layout = QHBoxLayout()
        order_no = QLineEdit()
        order_no.setAlignment(Qt.AlignRight)
        order_no.setFixedWidth(220)
        date = QLineEdit()
        date.setAlignment(Qt.AlignRight)
        date.setFixedWidth(160)
        top_layout.addWidget(QLabel("شماره سفارش "), 0, Qt.AlignRight)
        top_layout.addWidget(order_no)
        top_layout.addSpacing(12)
        top_layout.addWidget(QLabel("تاریخ "), 0, Qt.AlignRight)
        top_layout.addWidget(date)

        desc_layout = QHBoxLayout()
        description = QLineEdit()
        description.setAlignment(Qt.AlignRight)
        desc_layout.addWidget(QLabel("توضیحات "), 0, Qt.AlignRight)
        desc_layout.addWidget(description)

        if is_search_tab:
            table = QTableWidget(0, 4)
            table.setHorizontalHeaderLabels(["حذف", "نوع محصول", "کد محصول", "تعداد"])
            header = table.horizontalHeader()
            header.setSectionResizeMode(0, QHeaderView.ResizeToContents)
            header.setSectionResizeMode(1, QHeaderView.Stretch)
            header.setSectionResizeMode(2, QHeaderView.Stretch)
            header.setSectionResizeMode(3, QHeaderView.Stretch)
        else:
            table = QTableWidget(0, 3)
            table.setHorizontalHeaderLabels(["نوع محصول", "کد محصول", "تعداد"])
            header = table.horizontalHeader()
            for i in range(3):
                header.setSectionResizeMode(i, QHeaderView.Stretch)

        table.verticalHeader().setVisible(False)

        serial_box = QTextEdit()
        serial_box.setReadOnly(True)
        serial_box.setWordWrapMode(QTextOption.NoWrap)
        serial_box.setLayoutDirection(Qt.LeftToRight)
        serial_box.setFont(QFont("Consolas", 10))

        return {
            "top_layout": top_layout,
            "order_no": order_no,
            "date": date,
            "desc_layout": desc_layout,
            "description": description,
            "table": table,
            "serial_box": serial_box
        }

    def _add_product_to_table(self, table: QTableWidget):
        """Adds a new product to the specified table."""
        dlg = ProductDialog(self)


        def add_row(data):
            row_count = table.rowCount()
            table.insertRow(row_count)
            
            # If it's the search table, add a delete button
            if table.columnCount() == 4:
                btn_del = QPushButton("×")
                btn_del.setStyleSheet("color: red; font-weight: bold;")
                btn_del.clicked.connect(self.toggle_row_for_deletion)
                table.setCellWidget(row_count, 0, btn_del)
                # Data starts from column 1
                table.setItem(row_count, 1, QTableWidgetItem(str(data[0])))
                table.setItem(row_count, 2, QTableWidgetItem(str(data[1])))
                table.setItem(row_count, 3, QTableWidgetItem(str(data[2])))
            else:
                 # Data starts from column 0
                table.setItem(row_count, 0, QTableWidgetItem(str(data[0])))
                table.setItem(row_count, 1, QTableWidgetItem(str(data[1])))
                table.setItem(row_count, 2, QTableWidgetItem(str(data[2])))

        dlg.product_added.connect(add_row)
        dlg.exec_()

    def _edit_product_in_table(self, table: QTableWidget):
        """Edits the selected product in the specified table."""
        row = table.currentRow()
        if row == -1:
            QMessageBox.warning(self, "توجه", "ابتدا یک محصول را انتخاب کنید.")
            return
        
        # Adjust column indices for search table
        start_col = 1 if table.columnCount() == 4 else 0
        end_col = table.columnCount()
        
        preset_data = []
        for i in range(start_col, end_col):
            item = table.item(row, i)
            preset_data.append(item.text() if item else "")

        # The preset for ProductDialog expects (type, code, qty)
        dlg = ProductDialog(self, preset=preset_data)

        def update_row(data):
            for i, val in enumerate(data):
                table.setItem(row, i + start_col, QTableWidgetItem(str(val)))
            dlg.accept()
        dlg.product_added.connect(update_row)
        dlg.exec_()

    def _copy_serials_to_clipboard(self, serial_box: QTextEdit):
        """Copies content from a QTextEdit to the clipboard."""
        text = serial_box.toPlainText()
        if not text.strip():
            QMessageBox.information(self, "هشدار", "هیچ سریالی برای کپی وجود ندارد.")
            return
        QApplication.clipboard().setText(text)
        QMessageBox.information(self, "کپی شد", "سریال‌ها به کلیپ‌بورد کپی شدند.")

    # ---------- ساخت تب‌ها ----------
    def build_tab_new(self):
        """Builds the 'New Order' tab."""
        main_hbox = QHBoxLayout()
        self.tab_new.setLayout(main_hbox)
        left_layout, right_layout = QVBoxLayout(), QVBoxLayout()
        main_hbox.addLayout(left_layout, 3)
        main_hbox.addLayout(right_layout, 1)

        widgets = self._create_order_tab_widgets()
        self.new_order_no = widgets["order_no"]
        self.new_date = widgets["date"]
        self.new_desc = widgets["description"]
        self.table_new = widgets["table"]
        self.serial_box = widgets["serial_box"]

        btn_new_top = QPushButton("ثبت سفارش جدید")
        btn_new_top.setObjectName("secondary")
        btn_new_top.clicked.connect(self.reset_new_order_form)
        widgets["top_layout"].addWidget(btn_new_top)
        widgets["top_layout"].addStretch()

        left_layout.addLayout(widgets["top_layout"])
        left_layout.addLayout(widgets["desc_layout"])
        left_layout.addWidget(self.table_new)

        btn_layout = QHBoxLayout()
        btn_add = QPushButton("افزودن محصول")
        btn_edit = QPushButton("ویرایش محصول")
        btn_del = QPushButton("حذف محصول")
        btn_save = QPushButton("ذخیره سفارش")
        btn_add.clicked.connect(lambda: self._add_product_to_table(self.table_new))
        btn_edit.clicked.connect(lambda: self._edit_product_in_table(self.table_new))
        btn_del.clicked.connect(lambda: self.delete_selected(self.table_new))
        btn_save.clicked.connect(self.save_order_new)
        btn_layout.addWidget(btn_add)
        btn_layout.addWidget(btn_edit)
        btn_layout.addWidget(btn_del)
        btn_layout.addStretch()
        btn_layout.addWidget(btn_save)
        left_layout.addLayout(btn_layout)

        right_layout.addWidget(QLabel("سریال‌های این سفارش:"))
        right_layout.addWidget(self.serial_box)
        btn_copy = QPushButton("کپی سریال‌ها")
        btn_copy.setFixedWidth(130)
        btn_copy.clicked.connect(lambda: self._copy_serials_to_clipboard(self.serial_box))
        right_layout.addWidget(btn_copy)

    def build_tab_search(self):
        """Builds the 'Search and Edit' tab."""
        main_hbox = QHBoxLayout()
        self.tab_search.setLayout(main_hbox)
        left_layout, right_layout = QVBoxLayout(), QVBoxLayout()
        main_hbox.addLayout(left_layout, 3)
        main_hbox.addLayout(right_layout, 1)

        widgets = self._create_order_tab_widgets(is_search_tab=True)
        self.search_order_no = widgets["order_no"]
        self.search_date = widgets["date"]
        self.search_desc = widgets["description"]
        self.table_search = widgets["table"]
        self.serial_box_search = widgets["serial_box"]

        btn_search = QPushButton("جستجو")
        btn_search.clicked.connect(self.search_order)
        widgets["top_layout"].insertWidget(2, btn_search)
        widgets["top_layout"].addStretch()

        left_layout.addLayout(widgets["top_layout"])
        left_layout.addLayout(widgets["desc_layout"])
        left_layout.addWidget(self.table_search)

        btn_layout = QHBoxLayout()
        btn_add = QPushButton("افزودن محصول")
        btn_edit = QPushButton("ویرایش محصول")
        btn_save = QPushButton("ذخیره تغییرات")
        btn_add.clicked.connect(lambda: self._add_product_to_table(self.table_search))
        btn_edit.clicked.connect(lambda: self._edit_product_in_table(self.table_search))
        btn_save.clicked.connect(self.save_changes_search)
        btn_layout.addWidget(btn_add)
        btn_layout.addWidget(btn_edit)
        btn_layout.addStretch()
        btn_layout.addWidget(btn_save)
        left_layout.addLayout(btn_layout)

        right_layout.addWidget(QLabel("سریال‌های این سفارش:"))
        right_layout.addWidget(self.serial_box_search)
        btn_copy = QPushButton("کپی سریال‌ها")
        btn_copy.setFixedWidth(130)
        btn_copy.clicked.connect(
            lambda: self._copy_serials_to_clipboard(self.serial_box_search)
        )
        right_layout.addWidget(btn_copy)

    def build_tab_option(self):
        main_layout = QVBoxLayout(self.tab_option)
        h_main = QHBoxLayout()
        main_layout.addLayout(h_main)

        # User management panel
        self.user_group = QGroupBox("کاربران مجاز")
        self.user_group.setFixedWidth(280)
        user_group_layout = QVBoxLayout(self.user_group)
        h_main.addWidget(self.user_group, 0, Qt.AlignTop)

        user_top_row = QHBoxLayout()
        btn_add_user = QPushButton("افزودن")
        btn_add_user.setFixedWidth(80)
        btn_add_user.clicked.connect(self.add_user)
        self.user_input = QLineEdit(placeholderText="نام کاربری جدید...")
        user_top_row.addWidget(self.user_input)
        user_top_row.addWidget(btn_add_user)
        user_group_layout.addLayout(user_top_row)

        self.user_list = QListWidget()
        self.user_list.setLayoutDirection(Qt.LeftToRight)
        user_group_layout.addWidget(self.user_list)

        btn_remove_user = QPushButton("حذف کاربر انتخاب شده")
        btn_remove_user.setObjectName("secondary")
        btn_remove_user.clicked.connect(self.remove_selected_user)
        user_group_layout.addWidget(btn_remove_user)

        # Excel settings panel
        group_settings = QGroupBox("تنظیمات اکسل")
        settings_layout = QFormLayout(group_settings)
        h_main.addWidget(group_settings, 1, Qt.AlignTop)

        file_row = QHBoxLayout()
        btn_browse = QPushButton("انتخاب فایل")
        btn_browse.setFixedWidth(120)
        btn_browse.clicked.connect(self.browse_excel_file)
        self.e_excel_path = QLineEdit(EXCEL_FILE)
        file_row.addWidget(self.e_excel_path)
        file_row.addWidget(btn_browse)
        settings_layout.addRow("آدرس فایل اکسل:", file_row)

        self.e_sheet_name = QLineEdit(SHEET_NAME)
        settings_layout.addRow("نام برگه:", self.e_sheet_name)
        self.e_table_name = QLineEdit(TABLE_NAME)
        settings_layout.addRow("نام جدول:", self.e_table_name)

        bottom_row = QHBoxLayout()
        btn_about = QPushButton("درباره برنامه")
        btn_about.clicked.connect(self.show_about)
        btn_save = QPushButton("ذخیره")
        btn_save.clicked.connect(self.save_options)
        bottom_row.addWidget(btn_about)
        bottom_row.addStretch()
        bottom_row.addWidget(btn_save)
        main_layout.addLayout(bottom_row)

        # Initialize user list
        for u in ALLOWED_USERS:
            self.add_user_item(u)


    # ---------- متدهای عملکردی ----------
    @with_progress_dialog("در حال پردازش", "در حال ذخیره سفارش، لطفا صبر کنید...")
    def save_order_new(self):
        if not ensure_excel():
            return

        date_text = normalize_farsi(self.new_date.text())
        order_no = normalize_farsi(self.new_order_no.text())
        if not date_text or not order_no:
            QMessageBox.critical(self, "خطا", "تاریخ و شماره سفارش الزامی هستند.")
            return

        items = []
        for row in range(self.table_new.rowCount()):
            try:
                ptype = self.table_new.item(row, 0).text()
                code = self.table_new.item(row, 1).text()
                qty = int(self.table_new.item(row, 2).text())
                if qty <= 0:
                    raise ValueError
                items.append((ptype, code, qty))
            except (AttributeError, ValueError):
                QMessageBox.critical(self, "خطا", f"داده نامعتبر در ردیف {row + 1}")
                return

        if not items:
            QMessageBox.critical(self, "خطا", "حداقل یک محصول باید اضافه شود.")
            return

        try:
            wb = load_workbook(EXCEL_FILE)
            ws = wb[SHEET_NAME]
        except PermissionError:
            QMessageBox.critical(self, "خطا", f"فایل {EXCEL_FILE} باز است.")
            return
        except Exception as e:
            QMessageBox.critical(self, "خطا", f"خطا در باز کردن فایل: {e}")
            return

        maxA, maxB, max_rowid = compute_maxes(ws)
        serial_lines = []
        username = getpass.getuser()
        now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        for ptype, code, qty in items:
            item_idx, serial, maxA, maxB = next_item_and_serial(
                ws, date_text, ptype, maxA, maxB
            )
            max_rowid += 1
            ws.append([
                max_rowid, date_text, order_no, ptype, code, qty,
                item_idx, serial, normalize_farsi(self.new_desc.text()),
                username, now_str
            ])
            serial_lines.append('\u200E' + serial)

        update_excel_table_range(ws, TABLE_NAME)

        try:
            wb.save(EXCEL_FILE)
            QMessageBox.information(
                self, "موفق",
                "سفارش ثبت شد. سریال‌ها در پنل راست نمایش داده شدند."
            )
            self.serial_box.setPlainText("\n".join(serial_lines))
        except PermissionError:
            QMessageBox.critical(self, "خطا", f"فایل {EXCEL_FILE} باز است.")
        except Exception as e:
            QMessageBox.critical(self, "خطا", f"خطا در ذخیره‌سازی: {e}")

    @with_progress_dialog("در حال پردازش", "در حال جستجوی سفارش، لطفا صبر کنید...")
    def search_order(self):
        if not ensure_excel(): return
        order_no = normalize_farsi(self.search_order_no.text())
        if not order_no:
            QMessageBox.critical(self, "خطا", "شماره سفارش را وارد کنید.")
            return

        try:
            wb = load_workbook(EXCEL_FILE, read_only=True)
            ws = wb[SHEET_NAME]
            found_rows = [
                r for r in ws.iter_rows(min_row=2, values_only=True)
                if str(r[2]) == str(order_no)
            ]
            wb.close()
        except Exception as e:
            QMessageBox.critical(self, "خطا", f"خطا در خواندن فایل: {e}")
            return

        self.table_search.setRowCount(0)
        self.serial_box_search.clear()
        if not found_rows:
            QMessageBox.information(self, "یافت نشد", "سفارشی با این شماره پیدا نشد.")
            self.search_date.clear()
            self.search_desc.clear()
            return

        first = found_rows[0]
        self.search_date.setText(str(first[1] or ""))
        self.search_desc.setText(str(first[8] or ""))
        serial_lines = []
        for r in found_rows:
            row_idx = self.table_search.rowCount()
            self.table_search.insertRow(row_idx)

            btn_del = QPushButton("×")
            btn_del.setStyleSheet("color: red; font-weight: bold;")
            btn_del.clicked.connect(self.toggle_row_for_deletion)
            self.table_search.setCellWidget(row_idx, 0, btn_del)

            self.table_search.setItem(row_idx, 1, QTableWidgetItem(str(r[3] or "")))
            self.table_search.setItem(row_idx, 2, QTableWidgetItem(str(r[4] or "")))
            self.table_search.setItem(row_idx, 3, QTableWidgetItem(str(r[5] or "")))

            serial_lines.append('\u200E' + str(r[7] or ""))
        self.serial_box_search.setPlainText("\n".join(serial_lines))
    
    
    def toggle_row_for_deletion(self):
        """Toggles the strikethrough state of a row based on the button clicked."""
        button = self.sender()
        if not button:
            return

        index = self.table_search.indexAt(button.pos())
        if not index.isValid():
            return
        
        row = index.row()
        try:
            font = self.table_search.item(row, 1).font()
            is_struck_out = font.strikeOut()

            font.setStrikeOut(not is_struck_out)
            
            if not is_struck_out:
                button.setText("⟲") # Undo symbol
                button.setStyleSheet("color: green; font-weight: bold;")
            else:
                button.setText("×")
                button.setStyleSheet("color: red; font-weight: bold;")

            for col in range(1, self.table_search.columnCount()):
                item = self.table_search.item(row, col)
                if item:
                    item.setFont(font)
        except AttributeError:
            pass


    @with_progress_dialog("در حال پردازش", "در حال ذخیره تغییرات، لطفا صبر کنید...")
    def save_changes_search(self):
        if not ensure_excel(): return
        order_no = normalize_farsi(self.search_order_no.text())
        date_text = normalize_farsi(self.search_date.text())
        if not order_no or not date_text:
            QMessageBox.critical(self, "خطا", "شماره سفارش و تاریخ الزامی هستند.")
            return

        items_to_keep = []
        for row in range(self.table_search.rowCount()):
            item_font = self.table_search.item(row, 1).font()
            if not item_font.strikeOut():
                try:
                    ptype = self.table_search.item(row, 1).text()
                    code = self.table_search.item(row, 2).text()
                    qty = int(self.table_search.item(row, 3).text())
                    items_to_keep.append((ptype, code, qty))
                except (AttributeError, ValueError):
                    QMessageBox.critical(self, "خطا", f"داده نامعتبر در ردیف {row + 1}")
                    return

        try:
            wb = load_workbook(EXCEL_FILE)
            ws = wb[SHEET_NAME]
        except Exception as e:
            QMessageBox.critical(self, "خطا", f"خطا در باز کردن فایل اکسل: {e}")
            return

        delete_order_rows(ws, order_no)

        maxA, maxB, max_rowid = compute_maxes(ws)
        serial_lines = []
        username = getpass.getuser()
        now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        for ptype, code, qty in items_to_keep:
            item_idx, serial, maxA, maxB = next_item_and_serial(
                ws, date_text, ptype, maxA, maxB
            )
            max_rowid += 1
            ws.append([
                max_rowid, date_text, order_no, ptype, code, qty,
                item_idx, serial, normalize_farsi(self.search_desc.text()),
                username, now_str
            ])
            serial_lines.append('\u200E' + serial)

        update_excel_table_range(ws, TABLE_NAME)
        try:
            wb.save(EXCEL_FILE)
            self.serial_box_search.setPlainText("\n".join(serial_lines))
            
            for row in reversed(range(self.table_search.rowCount())):
                font = self.table_search.item(row, 1).font()
                if font.strikeOut():
                    self.table_search.removeRow(row)

            QMessageBox.information(self, "موفق", "تغییرات با موفقیت ذخیره شد.")

        except Exception as e:
            QMessageBox.critical(self, "خطا", f"خطا در ذخیره‌سازی: {e}")


    def reset_new_order_form(self):
        self.new_order_no.clear()
        self.new_date.setText("")
        self.new_desc.clear()
        self.table_new.setRowCount(0)
        self.serial_box.clear()


    def delete_selected(self, table):
        indexes = table.selectionModel().selectedRows()
        for idx in sorted([r.row() for r in indexes], reverse=True):
            table.removeRow(idx)


    # ---------- متدهای تب آپشن ----------
    def browse_excel_file(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "انتخاب فایل اکسل", "", "Excel Files (*.xlsx)"
        )
        if path:
            self.e_excel_path.setText(path)


    def save_options(self):
        global EXCEL_FILE, SHEET_NAME, TABLE_NAME, ALLOWED_USERS
        EXCEL_FILE = self.e_excel_path.text().strip()
        SHEET_NAME = self.e_sheet_name.text().strip()
        TABLE_NAME = self.e_table_name.text().strip()
        
        ALLOWED_USERS = [self.user_list.item(i).text() for i in range(self.user_list.count())]

        settings = {
            "excel_file": EXCEL_FILE, "sheet_name": SHEET_NAME,
            "table_name": TABLE_NAME, "allowed_users": ALLOWED_USERS,
        }
        if save_settings(settings):
            QMessageBox.information(self, "ذخیره شد", "تنظیمات با موفقیت ذخیره شد.")
        else:
            QMessageBox.warning(self, "خطا", "خطا در ذخیره تنظیمات.")


    def ask_admin_password(self):
        pwd, ok = QInputDialog.getText(
            self, "رمز مدیر", "رمز عبور مدیر را وارد کنید:", QLineEdit.Password
        )
        return ok and pwd == ADMIN_PASSWORD


    def add_user(self):
        if not self.ask_admin_password():
            QMessageBox.warning(self, "خطا", "رمز عبور نادرست است.")
            return
        new_user = normalize_farsi(self.user_input.text())
        if not new_user:
            return
        
        current_users = [self.user_list.item(i).text() for i in range(self.user_list.count())]
        if new_user in current_users:
            QMessageBox.information(self, "توجه", "این کاربر قبلاً اضافه شده است.")
            return
            
        self.add_user_item(new_user)
        self.user_input.clear()

    def add_user_item(self, username):
        item = QListWidgetItem(username)
        self.user_list.addItem(item)

    def remove_selected_user(self):
        selected_items = self.user_list.selectedItems()
        if not selected_items:
            QMessageBox.warning(self, "توجه", "لطفا ابتدا یک کاربر را برای حذف انتخاب کنید.")
            return

        if not self.ask_admin_password():
            QMessageBox.warning(self, "خطا", "رمز عبور نادرست است.")
            return

        for item in selected_items:
            row = self.user_list.row(item)
            self.user_list.takeItem(row)


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
            "Version: 1.4.0 — © 2025 All Rights Reserved<br>"
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


# ---------- اجرای برنامه ----------
if __name__ == "__main__":
    app = QApplication(sys.argv)
    QFontDatabase.addApplicationFont(resource_path("IRAN.ttf"))
    app.setFont(QFont("IRAN", 10))
    app.setWindowIcon(QIcon(resource_path("icon.ico")))
    show_splash()


