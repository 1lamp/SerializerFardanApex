"""
Fardan Apex --- Serializer
@2025
Author: Behnam Rabieyan
Company: Garma Gostar Fardan
"""

from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QTabWidget, QVBoxLayout, QHBoxLayout, QFileDialog,
    QLabel, QLineEdit, QPushButton, QTableWidget, QTableWidgetItem, QComboBox, QMessageBox, QDialog,
    QFormLayout, QHeaderView, QSizePolicy, QTextEdit, QProgressDialog, QGraphicsDropShadowEffect, QGroupBox
)
from PyQt5.QtGui import QFont, QIcon, QColor, QTextOption, QFontDatabase, QIntValidator
from PyQt5.QtCore import Qt, pyqtSignal
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
import os, re, sys

# ---------- تنظیمات ----------
EXCEL_FILE = r"D:\MyWork\G.G.Fardan\order.xlsx"
SHEET_NAME = "order"
HEADERS = ["ردیف", "تاریخ", "شماره سفارش", "نوع محصول", "کد محصول",
           "تعداد", "ردیف آیتم", "سریال سفارش", "توضیحات"]
PRODUCT_MAP = {
    "MF": "F", "MR": "R", "MU": "U",
    "نفراست": "ن", "فویلی": "ف", "فویل": "ف",
    "ترموسوییچ": "TS", "ترموسوئیچ": "TS",
    "هیتر سیمی": "س", "لوله ای دیفراست": "د", "لوله‌ای دیفراست": "د",
    "ترموفیوز": "TF"
}
GROUP_M = {"MF", "MR", "MU"}

    # ---------- بررسی محدوده جدول در اکسل ----------
def update_excel_table_range(ws, table_name):
    """
    به‌روزرسانی محدوده جدول اکسل بعد از اضافه کردن داده جدید
    ws: شیء Worksheet
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
        QMessageBox.warning(None, "هشدار", f"جدول '{table_name}' پیدا نشد. داده‌ها ذخیره شدند ولی جدول آپدیت نشد.")

# ---------- بررسی فایل سفارش ----------
def ensure_excel():
    if not os.path.exists(EXCEL_FILE):
        QMessageBox.critical(None, "خطا", f"فایل اکسل سفارشات یافت نشد:\n{EXCEL_FILE}")
        raise FileNotFoundError(f"Excel file not found: {EXCEL_FILE}")

# ---------- نرمال‌سازی حروف فارسی (عربی -> فارسی و trim) ----------
def normalize_farsi(s):
    if s is None:
        return ""
    s = str(s)
    mapping = {"ي": "ی", "ك": "ک", "ة": "ه", "ۀ": "ه"}
    for a, b in mapping.items():
        s = s.replace(a, b)
    s = re.sub(r"\s+", " ", s).strip()
    return s

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
QWidget { background: #f5f7fb; font-family: 'Segoe UI', Tahoma, Arial; }
QLineEdit, QTextEdit { background: white; border: 1px solid #d0d7df; border-radius: 6px; padding: 6px; }
QTextEdit { font-family: Consolas, 'Courier New', monospace; }
QPushButton { background: qlineargradient(x1:0,y1:0,x2:0,y2:1, stop:0 #5aa9ff, stop:1 #2e7dff); color: white; border: none; padding: 8px 12px; border-radius: 8px; }
QPushButton:hover { background: qlineargradient(x1:0,y1:0,x2:0,y2:1, stop:0 #6bb8ff, stop:1 #3b8bff); }
QPushButton#secondary { background: #eef4ff; color: #1a3b6e; border: 1px solid #d0dbff; }
QTableWidget { background: white; border: 1px solid #e0e6ef; gridline-color: #f1f5fb; }
QHeaderView::section { background: #eef4ff; padding: 6px; border: none; }
QComboBox { background: white; border: 1px solid #d0d7df; border-radius: 6px; padding: 4px; }
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
        self.cb_type.addItems(['', 'فویلی', 'هیتر سیمی', 'نفراست', 'لوله ای دیفراست', 'ترموفیوز', 'ترموسوییچ', 'MF', 'MR', 'MU'])
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
                self.e_qty.setText("1")
        else:
            self.e_qty.setText("1")

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
            QMessageBox.critical(self, "خطا", "تعداد نامعتبر است. لطفاً عددی بزرگتر از صفر وارد کنید.")
            return

        if not ptype or not code:
            QMessageBox.critical(self, "خطا", "همه فیلدها الزامی هستند.")
            return

        # ارسال داده به بیرون بدون بستن دیالوگ
        self.product_added.emit((ptype, code, qty))

        # ریست کردن فیلدها
        self.cb_type.setCurrentIndex(0)
        self.e_code.clear()
        self.e_qty.setText("1")
        self.cb_type.setFocus()

# ---------- کلاس اصلی ----------
class App(QMainWindow):
    def __init__(self):
        super().__init__()
        ensure_excel()
        self.setWindowTitle("تولید سریال سفارش - FardanApex")
        self.resize(900, 500)
        self.setStyleSheet(APP_STYLESHEET)

        # Drop shadow for main window (subtle on central container)
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
        layout = QVBoxLayout()
        self.tab_option.setLayout(layout)

        group_settings = QGroupBox("تنظیمات")
        group_settings.setMaximumHeight(900)
        settings_layout = QVBoxLayout()
        group_settings.setLayout(settings_layout)

        # آدرس فایل اکسل
        lbl_file = QLabel("     آدرس فایل اکسل")
        lbl_file_layout = QHBoxLayout()
        lbl_file_layout.addStretch()
        lbl_file_layout.addWidget(lbl_file)

        self.e_excel_path = QLineEdit(self)
        self.e_excel_path.setText(EXCEL_FILE)

        btn_browse = QPushButton("انتخاب فایل")
        btn_browse.clicked.connect(self.browse_excel_file)

        row1 = QVBoxLayout()
        row1.addLayout(lbl_file_layout)

        file_row = QHBoxLayout()
        file_row.addWidget(btn_browse)
        file_row.addWidget(self.e_excel_path)
        row1.addLayout(file_row)

        settings_layout.addLayout(row1)

        # فاصله بین بخش‌ها
        settings_layout.addSpacing(20)

        # نام برگه
        lbl_sheet = QLabel("    نام برگه")
        sheet_label_layout = QHBoxLayout()
        sheet_label_layout.addStretch()
        sheet_label_layout.addWidget(lbl_sheet)

        self.e_sheet_name = QLineEdit(self)
        self.e_sheet_name.setText(SHEET_NAME)

        sheet_layout = QVBoxLayout()
        sheet_layout.addLayout(sheet_label_layout)
        sheet_layout.addWidget(self.e_sheet_name)

        # نام جدول
        lbl_table = QLabel("    نام جدول")
        table_label_layout = QHBoxLayout()
        table_label_layout.addStretch()
        table_label_layout.addWidget(lbl_table)

        self.e_table_name = QLineEdit(self)
        self.e_table_name.setText("ordertable")

        table_layout = QVBoxLayout()
        table_layout.addLayout(table_label_layout)
        table_layout.addWidget(self.e_table_name)

        # ترکیب دو بخش در یک ردیف
        row2 = QHBoxLayout()
        row2.addLayout(sheet_layout)
        row2.addSpacing(20)
        row2.addLayout(table_layout)

        settings_layout.addLayout(row2)

        # دکمه ذخیره
        btn_layout = QHBoxLayout()
        btn_save = QPushButton("ذخیره")
        btn_save.clicked.connect(self.save_options)
        btn_layout.addStretch()
        btn_layout.addWidget(btn_save)
        settings_layout.addLayout(btn_layout)

        layout.addWidget(group_settings, alignment=Qt.AlignTop)

        # دکمه درباره برنامه
        btn_about_layout = QHBoxLayout()
        btn_about = QPushButton("درباره برنامه")
        btn_about.clicked.connect(self.show_about)
        btn_about_layout.addWidget(btn_about, alignment=Qt.AlignLeft)
        layout.addLayout(btn_about_layout)

    # تابع انتخاب فایل
    def browse_excel_file(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "انتخاب فایل اکسل",
            "",
            "Excel Files (*.xlsx)"
        )
        if file_path:
            self.e_excel_path.setText(file_path)


    # ---------- درباره برنامه ----------
    def show_about(self):
        msg = QMessageBox(self)
        msg.setWindowTitle("Fardan Apex --- Serializer")
        msg.setIcon(QMessageBox.Information)
        msg.setText("Fardan Apex — Serializer\n\n"
                    "Serial Generator\n\n"
                    "This application is designed to generate production series after order confirmation by the engineering unit. All calculations and data entries are handled automatically and saved to the designated Excel file.\n\n"
                    "Developed exclusively for:\n"
                    "Garma Gostar Fardan Co.\n\n"
                    "Version: 2.1.7\n"
                    "© 2025 All Rights Reserved\n\n"
                    "Design & Development: Behnam Rabieyan\n"
                    "Email: behnamrabieyan@live.com\n"
                    "Web: behnamrabieyan.ir\n")
        msg.setStandardButtons(QMessageBox.Ok)
        msg.exec_()

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
        ensure_excel()
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
            ws.append([max_rowid, date_text, order_no, ptype, code, qty, item_idx, serial, desc])
            serial_lines.append('\u200E' + serial)

    # آپدیت محدوده جدول
        update_excel_table_range(ws, getattr(self, "table_name", "ordertable"))

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
        ensure_excel()
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
        ensure_excel()
        order_no = normalize_farsi(self.search_order_no.text())
        date_text = normalize_farsi(self.search_date.text())
        desc = normalize_farsi(self.search_desc.text())
        if not order_no or not date_text:
            QMessageBox.critical(self, "خطا", "شماره سفارش و تاریخ الزامی هستند.")
            return

        try:
            wb = load_workbook(EXCEL_FILE)
            ws = wb[SHEET_NAME]
        except PermissionError:
            QMessageBox.critical(self, "خطا", f"فایل {EXCEL_FILE} باز است. لطفا ببندید و دوباره تلاش کنید.")
            return

        existing_rows = []
        for idx, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
            if str(row[2]) == str(order_no):
                existing_rows.append({
                    "ws_idx": idx,
                    "rowid": row[0],
                    "date": row[1],
                    "ptype": row[3],
                    "code": row[4],
                    "qty": row[5],
                    "item_idx": row[6],
                    "serial": row[7],
                    "desc": row[8]
                })

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

            else:
                max_rowid += 1
                item_idx, serial, maxA, maxB = next_item_and_serial(ws, date_text, ptype_new, maxA, maxB)
                ws.append([max_rowid, date_text, order_no, ptype_new, code_new, qty_new, item_idx, serial, desc])
                serial_lines.append('\u200E' + serial)

    # آپدیت محدوده جدول
        update_excel_table_range(ws, getattr(self, "table_name", "ordertable"))

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

    # ---------- ذخیر کپی کردن سریال ----------
    def copy_serials(self):
            text = self.serial_box.toPlainText()
            if not text.strip():
                QMessageBox.information(self, "هشدار", "هیچ سریالی برای کپی وجود ندارد.")
                return
            QApplication.clipboard().setText(text)
            QMessageBox.information(self, "کپی شد", "سریال‌ها به کلیپ‌بورد کپی شدند.")

    # ---------- ذخیر کپی کردن سریال ----------
    def copy_serials_search(self):
            text = self.serial_box_search.toPlainText()
            if not text.strip():
                QMessageBox.information(self, "هشدار", "هیچ سریالی برای کپی وجود ندارد.")
                return
            QApplication.clipboard().setText(text)
            QMessageBox.information(self, "کپی شد", "سریال‌ها به کلیپ‌بورد کپی شدند.")

    # ---------- ذخیر تنظیمات ----------
    def save_options(self):
        global EXCEL_FILE, SHEET_NAME
        EXCEL_FILE = self.e_excel_path.text().strip()
        SHEET_NAME = self.e_sheet_name.text().strip()
        self.table_name = self.e_table_name.text().strip() or "ordertable"
        QMessageBox.information(self, "ذخیره شد", "تنظیمات با موفقیت ذخیره شد.")

# ---------- اجرای برنامه ----------
if __name__ == "__main__":
    app = QApplication(sys.argv)
    # اضافه کردن فونت IRAN
    QFontDatabase.addApplicationFont("IRAN.ttf")
    app.setFont(QFont("IRAN", 10))

# اضافه کردن آیکون
    app.setWindowIcon(QIcon("icon.ico"))
    window = App()
    window.show()
    sys.exit(app.exec_())
