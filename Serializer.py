"""
Fardan Apex --- Serializer
@2025
Author: Behnam Rabieyan
Company: Garma Gostar Fardan
"""

import tkinter as tk
from tkinter import ttk, messagebox
from openpyxl import Workbook, load_workbook
import os
import re
from datetime import datetime

# ---------- تنظیمات ----------
EXCEL_FILE = r"D:\MyWork\G.G.Fardan\order.xlsx"
SHEET_NAME = "order"

# سربرگ‌ها در شیت
HEADERS = ["ردیف", "تاریخ", "شماره سفارش", "نوع محصول", "کد محصول",
           "تعداد", "ردیف آیتم", "سریال سفارش", "توضیحات"]

# کاوش نوع محصول جهت اختصار سریال
PRODUCT_MAP = {
    "MF": "F", "MR": "R", "MU": "U",
    "نفراست": "ن", "فویلی": "ف", "فویل": "ف",
    "ترموسوییچ": "TS", "ترموسوئیچ": "TS",
    "هیتر سیمی": "س", "لوله ای دیفراست": "د", "لوله‌ای دیفراست": "د",
    "ترموفیوز": "TF"
}
GROUP_A_LATIN = {"MF", "MR", "MU"}  # گروه میله ای
# -------------------------------------------------


# ---------- کمکی: ساخت فایل اگر وجود نداشت ----------
def ensure_excel():
    if not os.path.exists(EXCEL_FILE):
        wb = Workbook()
        ws = wb.active
        ws.title = SHEET_NAME
        ws.append(HEADERS)
        wb.save(EXCEL_FILE)


# ---------- نرمال‌سازی حروف فارسی (عربی -> فارسی و trim) ----------
def normalize_farsi(s):
    if s is None:
        return ""
    s = str(s)
    mapping = {
        "ي": "ی",
        "ك": "ک",
        "ة": "ه",
        "ۀ": "ه",
    }
    for a, b in mapping.items():
        s = s.replace(a, b)
    # حذف فاصله‌های اضافی و نرمال‌سازی فاصله
    s = re.sub(r"\s+", " ", s).strip()
    return s


# ---------- پیدا کردن بیشترین ها (برای ردیف آیتم و شماره ردیف) ----------
def compute_maxes(ws):
    """بررسی شیت و محاسبهٔ حداکثر ردیف آیتم برای گروه میله ای و غیر میله ای و حداکثر شماره ردیف (ستون A)."""
    max_groupA = 0
    max_groupB = 0
    max_rowid = 0
    for row in ws.iter_rows(min_row=2, values_only=True):
        # row indices: 0=ردیف,1=تاریخ,2=شماره سفارش,3=نوع محصول,4=کد,5=تعداد,6=ردیف آیتم,7=سریال,8=توضیحات
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

        # تشخیص گروه میله ای: اگر متنِ نوع محصول MF/MR/MU باشه
        if str(ptype).upper() in GROUP_A_LATIN:
            if item_idx > max_groupA:
                max_groupA = item_idx
        else:
            # برای بقیه
            if item_idx > max_groupB:
                max_groupB = item_idx

    return max_groupA, max_groupB, max_rowid


# ---------- تابع ساخت سریال سفارش و ردیف آیتم بر اساس الگوریتم ----------
def next_item_and_serial(ws, date_text, product_type, max_groupA, max_groupB):
    """برای یک آیتم جدید، بر اساس مقدار فعلی max_groupA/max_groupB مقدار ردیف آیتم و سریال بساز و return کن."""
    # نرمال‌سازی
    p = normalize_farsi(product_type)
    # تشخیص اختصار
    key = p
    # اگر لاتین وارد شده (مانند "mf" یا "MF")، آن را به uppercase تبدیل کن
    if re.match(r"^[A-Za-z]{1,4}$", p):
        key = p.upper()
    # انتخاب اختصار
    abbrev = PRODUCT_MAP.get(key, PRODUCT_MAP.get(key.lower(), None))
    if not abbrev:
        # اگر توی دیکشنری نبود، اما key لاتین از MF/MR/MU بود، نگاشت مستقیم:
        if key.upper() in GROUP_A_LATIN:
            abbrev = key.upper()[0]
        else:
            abbrev = "0"

    # استخراج YYYY از date_text (چهار رقم اول عددی)
    yyyy = "0000"
    if date_text:
        m = re.search(r"\d{4}", date_text)
        if m:
            yyyy = m.group(0)
        else:
            yyyy = date_text[:4] if len(date_text) >= 4 else "0000"

    # تشخیص گروه
    in_groupA = (str(key).upper() in GROUP_A_LATIN)

    if in_groupA:
        max_groupA += 1
        item_idx = max_groupA
    else:
        max_groupB += 1
        item_idx = max_groupB

    serial = f"{item_idx}-{yyyy}-{abbrev}"
    return item_idx, serial, max_groupA, max_groupB


# ---------- حذف تمام ردیف‌های یک سفارش (برای بازنویسی هنگام ویرایش) ----------
def delete_order_rows(ws, order_no):
    to_delete = []
    for idx, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
        if str(row[2]) == str(order_no):
            to_delete.append(idx)
    for r in reversed(to_delete):
        ws.delete_rows(r, 1)


# ---------- UI ----------
class App:
    def __init__(self, root):
        ensure_excel()
        self.root = root
        root.title("ورود سفارش به اکسل")
        root.geometry("900x620")

        # Notebook (دو تب ثبت و ویرایش)
        self.nb = ttk.Notebook(root)
        self.tab_new = ttk.Frame(self.nb)
        self.tab_search = ttk.Frame(self.nb)
        self.nb.add(self.tab_new, text="ثبت سفارش جدید")
        self.nb.add(self.tab_search, text="جستجو و ویرایش")
        self.nb.pack(fill="both", expand=True)

        # ---------- Tab: ثبت سفارش جدید ----------
        self.build_tab_new()

        # ---------- Tab: جستجو و ویرایش ----------
        self.build_tab_search()

    def build_tab_new(self):
        frm = self.tab_new
        # بالای فرم: شماره سفارش - تاریخ - توضیحات (با راست‌چین)
        tk.Label(frm, text="شماره سفارش:").place(x=20, y=10)
        self.new_order_no = tk.Entry(frm, justify="right", width=18)
        self.new_order_no.place(x=110, y=10)

        tk.Label(frm, text="تاریخ:").place(x=320, y=10)
        self.new_date = tk.Entry(frm, justify="right", width=18)
        self.new_date.place(x=360, y=10)
        self.new_date.insert(0, datetime.today().strftime("%Y/%m/%d"))

        tk.Label(frm, text="توضیحات:").place(x=20, y=40)
        self.new_desc = tk.Entry(frm, justify="right", width=70)
        self.new_desc.place(x=110, y=40)

        # Treeview محصولات (وسط)
        cols = ("نوع محصول", "کد محصول", "تعداد")
        self.tree_new = ttk.Treeview(frm, columns=cols, show="headings", height=18)
        self.tree_new.heading("نوع محصول", text="نوع محصول")
        self.tree_new.heading("کد محصول", text="کد محصول")
        self.tree_new.heading("تعداد", text="تعداد")
        self.tree_new.column("نوع محصول", width=420, anchor="center")
        self.tree_new.column("کد محصول", width=200, anchor="center")
        self.tree_new.column("تعداد", width=100, anchor="center")
        self.tree_new.place(x=20, y=90, width=840, height=400)

        # دکمه‌های پایین (افزودن، ویرایش، حذف) سمت چپ و ذخیره سمت راست
        btn_x = 20
        btn_y = 510
        tk.Button(frm, text="افزودن محصول", width=16, command=self.popup_add_product_new).place(x=btn_x, y=btn_y)
        tk.Button(frm, text="ویرایش محصول", width=16, command=self.edit_product_new).place(x=btn_x+140, y=btn_y)
        tk.Button(frm, text="حذف محصول", width=16, command=lambda: self.delete_selected(self.tree_new)).place(x=btn_x+280, y=btn_y)

        tk.Button(frm, text="ذخیره سفارش", width=14, command=self.save_order_new).place(x=760, y=550)

    def build_tab_search(self):
        frm = self.tab_search
        # بالای فرم: شماره سفارش - تاریخ - توضیحات
        tk.Label(frm, text="شماره سفارش:").place(x=20, y=10)
        self.search_order_no = tk.Entry(frm, justify="right", width=18)
        self.search_order_no.place(x=110, y=10)

        tk.Button(frm, text="جستجو", command=self.search_order).place(x=290, y=8)

        tk.Label(frm, text="تاریخ:").place(x=360, y=10)
        self.search_date = tk.Entry(frm, justify="right", width=18)
        self.search_date.place(x=400, y=10)

        tk.Label(frm, text="توضیحات:").place(x=20, y=40)
        self.search_desc = tk.Entry(frm, justify="right", width=70)
        self.search_desc.place(x=110, y=40)

        # Treeview محصولات تب جستجو (شامل ردیف آیتم و سریال جهت مشاهده)
        cols = ("نوع محصول", "کد محصول", "تعداد", "ردیف آیتم", "سریال سفارش")
        self.tree_search = ttk.Treeview(frm, columns=cols, show="headings", height=18)
        for c in cols:
            self.tree_search.heading(c, text=c)
            # تنظیم عرض
        self.tree_search.column("نوع محصول", width=360, anchor="center")
        self.tree_search.column("کد محصول", width=220, anchor="center")
        self.tree_search.column("تعداد", width=100, anchor="center")
        self.tree_search.column("ردیف آیتم", width=90, anchor="center")
        self.tree_search.column("سریال سفارش", width=160, anchor="center")
        self.tree_search.place(x=20, y=90, width=840, height=400)

        # دکمه‌ها پایین تب جستجو
        btn_x = 20
        btn_y = 510
        tk.Button(frm, text="افزودن محصول", width=16, command=self.popup_add_product_search).place(x=btn_x, y=btn_y)
        tk.Button(frm, text="ویرایش محصول", width=16, command=self.edit_product_search).place(x=btn_x+140, y=btn_y)
        tk.Button(frm, text="حذف محصول", width=16, command=lambda: self.delete_selected(self.tree_search)).place(x=btn_x+280, y=btn_y)

        tk.Button(frm, text="ذخیره تغییرات", width=14, command=self.save_changes_search).place(x=760, y=550)

    # ---------- پنجره افزودن / ویرایش محصول برای تب new ----------
    def popup_add_product_new(self, preset=None, edit_iid=None):
        # preset = (ptype, code, qty) or None
        popup = tk.Toplevel(self.root)
        popup.title("افزودن/ویرایش محصول")
        popup.geometry("420x200")
        popup.transient(self.root)
        # لیست محصولات (قابل تایپ)
        product_options = ['نفراست', 'فویلی', 'هیتر سیمی', 'لوله ای دیفراست', 'ترموفیوز',
                           'ترموسوییچ', 'ترموسوئیچ', 'MF', 'MR', 'MU']
        tk.Label(popup, text="نوع محصول:").place(x=300, y=10)
        cb_type = ttk.Combobox(popup, values=product_options, width=36)
        cb_type.place(x=20, y=10)
        cb_type.configure(state="normal")

        tk.Label(popup, text="کد محصول:").place(x=300, y=50)
        e_code = tk.Entry(popup, width=52, justify="right")
        e_code.place(x=20, y=50)

        tk.Label(popup, text="تعداد:").place(x=300, y=90)
        e_qty = tk.Entry(popup, width=20, justify="right")
        e_qty.place(x=20, y=90)

        if preset:
            cb_type.set(preset[0])
            e_code.insert(0, preset[1])
            e_qty.insert(0, preset[2])

        def on_register():
            ptype = normalize_farsi(cb_type.get())
            code = normalize_farsi(e_code.get())
            qty = e_qty.get().strip()
            if not ptype or not code or not qty:
                messagebox.showerror("خطا", "همهٔ فیلدها الزامی هستند.", parent=popup)
                return
            try:
                q = int(qty)
                if q <= 0:
                    raise ValueError
            except:
                messagebox.showerror("خطا", "تعداد باید عدد صحیح مثبت باشد.", parent=popup)
                return

            # اگر حالت ویرایش (edit_iid مشخص است) => آپدیت ردیف انتخاب شده
            if edit_iid:
                self.tree_new.item(edit_iid, values=(ptype, code, q))
                popup.destroy()
            else:
                # افزودن به جدول در تب جدید و خالی کردن فیلدها (پنجره باز بماند)
                self.tree_new.insert("", "end", values=(ptype, code, q))
                cb_type.set(product_options[0])
                e_code.delete(0, tk.END)
                e_qty.delete(0, tk.END)
                cb_type.focus_set()

        tk.Button(popup, text="ثبت", width=12, command=on_register).place(x=40, y=140)
        tk.Button(popup, text="بستن", width=12, command=popup.destroy).place(x=180, y=140)

    # ---------- popup افزودن برای تب جستجو (اضافه کردن محصول جدید هنگام ویرایش سفارش) ----------
    def popup_add_product_search(self, preset=None, edit_iid=None):
        popup = tk.Toplevel(self.root)
        popup.title("افزودن/ویرایش محصول")
        popup.geometry("420x200")
        popup.transient(self.root)

        product_options = ['نفراست', 'فویلی', 'هیتر سیمی', 'لوله ای دیفراست', 'ترموفیوز',
                           'ترموسوییچ', 'ترموسوئیچ', 'MF', 'MR', 'MU']
        tk.Label(popup, text="نوع محصول:").place(x=300, y=10)
        cb_type = ttk.Combobox(popup, values=product_options, width=36)
        cb_type.place(x=20, y=10)
        cb_type.configure(state="normal")

        tk.Label(popup, text="کد محصول:").place(x=300, y=50)
        e_code = tk.Entry(popup, width=52, justify="right")
        e_code.place(x=20, y=50)

        tk.Label(popup, text="تعداد:").place(x=300, y=90)
        e_qty = tk.Entry(popup, width=20, justify="right")
        e_qty.place(x=20, y=90)

        if preset:
            cb_type.set(preset[0])
            e_code.insert(0, preset[1])
            e_qty.insert(0, preset[2])

        def on_register():
            ptype = normalize_farsi(cb_type.get())
            code = normalize_farsi(e_code.get())
            qty = e_qty.get().strip()
            if not ptype or not code or not qty:
                messagebox.showerror("خطا", "همهٔ فیلدها الزامی هستند.", parent=popup)
                return
            try:
                q = int(qty)
                if q <= 0:
                    raise ValueError
            except:
                messagebox.showerror("خطا", "تعداد باید عدد صحیح مثبت باشد.", parent=popup)
                return

            if edit_iid:
                # ویرایش سطر انتخاب‌شده در tree_search
                self.tree_search.item(edit_iid, values=(ptype, code, q,
                                                        self.tree_search.set(edit_iid, "ردیف آیتم"),
                                                        self.tree_search.set(edit_iid, "سریال سفارش")))
                popup.destroy()
            else:
                # افزودن یک آیتم جدید در تب جستجو (هنگام ویرایش سفارش)
                self.tree_search.insert("", "end", values=(ptype, code, q, "", ""))
                cb_type.set(product_options[0])
                e_code.delete(0, tk.END)
                e_qty.delete(0, tk.END)
                cb_type.focus_set()

        tk.Button(popup, text="ثبت", width=12, command=on_register).place(x=40, y=140)
        tk.Button(popup, text="بستن", width=12, command=popup.destroy).place(x=180, y=140)

    # ---------- ویرایش: در تب new ----------
    def edit_product_new(self):
        sel = self.tree_new.selection()
        if not sel:
            messagebox.showwarning("توجه", "ابتدا یک محصول را انتخاب کنید.")
            return
        iid = sel[0]
        vals = self.tree_new.item(iid, "values")
        self.popup_add_product_new(preset=vals, edit_iid=iid)

    # ---------- ویرایش: در تب search ----------
    def edit_product_search(self):
        sel = self.tree_search.selection()
        if not sel:
            messagebox.showwarning("توجه", "ابتدا یک محصول را انتخاب کنید.")
            return
        iid = sel[0]
        vals = self.tree_search.item(iid, "values")
        # preset: (ptype, code, qty)
        preset = (vals[0], vals[1], vals[2])
        self.popup_add_product_search(preset=preset, edit_iid=iid)

    # ---------- حذف سطر انتخاب شده از یک tree ----------
    def delete_selected(self, tree):
        sel = tree.selection()
        if not sel:
            messagebox.showwarning("توجه", "هیچ موردی انتخاب نشده.")
            return
        for s in sel:
            tree.delete(s)

    # ---------- ذخیره سفارش جدید (Tab New) ----------
    def save_order_new(self):
        ensure_excel()
        date_text = normalize_farsi(self.new_date.get())
        order_no = normalize_farsi(self.new_order_no.get())
        desc = normalize_farsi(self.new_desc.get())

        if not date_text or not order_no:
            messagebox.showerror("خطا", "تاریخ و شماره سفارش الزامی هستند.")
            return
        items = []
        for child in self.tree_new.get_children():
            ptype, code, qty = self.tree_new.item(child, "values")
            ptype = normalize_farsi(ptype)
            code = normalize_farsi(code)
            try:
                qty = int(qty)
                if qty <= 0:
                    raise ValueError
            except:
                messagebox.showerror("خطا", f"تعداد نامعتبر برای آیتم: {ptype} - {code}")
                return
            items.append((ptype, code, qty))

        if not items:
            messagebox.showerror("خطا", "حداقل یک محصول باید اضافه شود.")
            return

        # باز کردن اکسل و محاسبهٔ maxها
        try:
            wb = load_workbook(EXCEL_FILE)
            ws = wb[SHEET_NAME]
        except PermissionError:
            messagebox.showerror("خطا", f"فایل {EXCEL_FILE} باز است. لطفاً آن را ببندید و مجدداً تلاش کنید.")
            return
        except Exception as e:
            messagebox.showerror("خطا", f"خطا در باز کردن فایل: {e}")
            return

        maxA, maxB, max_rowid = compute_maxes(ws)
        # برای آیتم‌های جدید یکی یکی تولید کن (تا آیتم بعدی از آیتم اضافه‌شده استفاده کنه)
        for ptype, code, qty in items:
            item_idx, serial, maxA, maxB = next_item_and_serial(ws, date_text, ptype, maxA, maxB)
            max_rowid += 1
            # ردیف اکسل را خود برنامه می‌نویسد (ستون A)
            ws.append([max_rowid, date_text, order_no, ptype, code, qty, item_idx, serial, desc])

        try:
            wb.save(EXCEL_FILE)
            messagebox.showinfo("موفق", "سفارش با موفقیت ثبت شد.")
            # پاکسازی فرم و درخت
            self.new_order_no.delete(0, tk.END)
            self.new_desc.delete(0, tk.END)
            self.new_date.delete(0, tk.END)
            self.new_date.insert(0, datetime.today().strftime("%Y/%m/%d"))
            for ch in self.tree_new.get_children():
                self.tree_new.delete(ch)
        except PermissionError:
            messagebox.showerror("خطا", f"فایل {EXCEL_FILE} باز است. لطفاً آن را ببندید و دوباره تلاش کنید.")
        except Exception as e:
            messagebox.showerror("خطا", f"خطا در ذخیره‌سازی: {e}")

    # ---------- جستجو سفارش (Tab Search) ----------
    def search_order(self):
        ensure_excel()
        order_no = normalize_farsi(self.search_order_no.get())
        if not order_no:
            messagebox.showerror("خطا", "شماره سفارش را وارد کنید.")
            return
        try:
            wb = load_workbook(EXCEL_FILE)
            ws = wb[SHEET_NAME]
        except PermissionError:
            messagebox.showerror("خطا", f"فایل {EXCEL_FILE} باز است. لطفاً آن را ببندید.")
            return

        rows = list(ws.iter_rows(min_row=2, values_only=True))
        found_rows = [r for r in rows if str(r[2]) == str(order_no)]
        if not found_rows:
            messagebox.showinfo("یافت نشد", "سفارشی با این شماره پیدا نشد.")
            # پاکسازی نمایش
            self.search_date.delete(0, tk.END)
            self.search_desc.delete(0, tk.END)
            for ch in self.tree_search.get_children():
                self.tree_search.delete(ch)
            return

        # پر کردن تاریخ و توضیحات (از اولین ردیف)
        first = found_rows[0]
        self.search_date.delete(0, tk.END)
        self.search_date.insert(0, first[1] if first[1] is not None else "")
        self.search_desc.delete(0, tk.END)
        self.search_desc.insert(0, first[8] if first[8] is not None else "")

        # پر کردن جدول با جزئیات آیتم‌ها (نوع،کد،تعداد,ردیف آیتم،سریال)
        for ch in self.tree_search.get_children():
            self.tree_search.delete(ch)
        for r in found_rows:
            ptype = r[3] if r[3] is not None else ""
            code = r[4] if r[4] is not None else ""
            qty = r[5] if r[5] is not None else ""
            itemidx = r[6] if r[6] is not None else ""
            serial = r[7] if r[7] is not None else ""
            self.tree_search.insert("", "end", values=(ptype, code, qty, itemidx, serial))

    # ---------- ذخیرهٔ تغییرات در تب جستجو (بازنویسی ردیف‌های آن سفارش) ----------
    def save_changes_search(self):
        order_no = normalize_farsi(self.search_order_no.get())
        date_text = normalize_farsi(self.search_date.get())
        desc = normalize_farsi(self.search_desc.get())

        if not order_no or not date_text:
            messagebox.showerror("خطا", "شماره سفارش و تاریخ الزامی هستند.")
            return

        # ساخت لیست آیتم‌ها از tree_search
        items = []
        for ch in self.tree_search.get_children():
            vals = self.tree_search.item(ch, "values")
            ptype = normalize_farsi(vals[0])
            code = normalize_farsi(vals[1])
            try:
                qty = int(vals[2])
            except:
                messagebox.showerror("خطا", f"تعداد نامعتبر در آیتم: {ptype} - {code}")
                return
            items.append((ptype, code, qty))

        if not items:
            messagebox.showerror("خطا", "حداقل یک محصول لازم است.")
            return

        try:
            wb = load_workbook(EXCEL_FILE)
            ws = wb[SHEET_NAME]
        except PermissionError:
            messagebox.showerror("خطا", f"فایل {EXCEL_FILE} باز است. لطفاً ببندید و دوباره تلاش کنید.")
            return

        # حذف ردیف‌های قبلی این سفارش
        delete_order_rows(ws, order_no)
        # محاسبهٔ maxها دوباره
        maxA, maxB, max_rowid = compute_maxes(ws)

        # اضافه‌کردن ردیف‌های جدید (مثل save_order_new)
        for ptype, code, qty in items:
            item_idx, serial, maxA, maxB = next_item_and_serial(ws, date_text, ptype, maxA, maxB)
            max_rowid += 1
            ws.append([max_rowid, date_text, order_no, ptype, code, qty, item_idx, serial, desc])

        try:
            wb.save(EXCEL_FILE)
            messagebox.showinfo("موفق", "تغییرات با موفقیت ذخیره شد.")
        except PermissionError:
            messagebox.showerror("خطا", f"فایل {EXCEL_FILE} باز است. لطفاً آن را ببندید و دوباره تلاش کنید.")
        except Exception as e:
            messagebox.showerror("خطا", f"خطا در ذخیره: {e}")


if __name__ == "__main__":
    ensure_excel()
    root = tk.Tk()
    app = App(root)
    root.mainloop()

