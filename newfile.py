import os
import shutil
import sqlite3
from datetime import datetime

from kivy.app import App
from kivy.core.text import LabelBase
from kivy.core.window import Window
from kivy.metrics import dp, sp
from kivy.graphics import Color, Rectangle
from kivy.uix.boxlayout import BoxLayout
from kivy.uix.gridlayout import GridLayout
from kivy.uix.scrollview import ScrollView
from kivy.uix.label import Label
from kivy.uix.button import Button
from kivy.uix.textinput import TextInput
from kivy.uix.popup import Popup
from kivy.uix.spinner import Spinner, SpinnerOption

import arabic_reshaper
from bidi.algorithm import get_display

try:
    from openpyxl import Workbook
    from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
    OPENPYXL_AVAILABLE = True
except ImportError:
    OPENPYXL_AVAILABLE = False


# =========================================================
# إعدادات التصميم والخط
# =========================================================
FONT_PATH = "/storage/emulated/0/Download/NotoNaskhArabic[wght].ttf"
LabelBase.register(name="Arabic", fn_regular=FONT_PATH)

NAVY = (0.055, 0.105, 0.19, 1)
BLUE = (0.12, 0.22, 0.36, 1)
SILVER = (0.78, 0.82, 0.88, 1)
LIGHT = (0.94, 0.96, 0.98, 1)
WHITE = (1, 1, 1, 1)
DARK = (0.12, 0.16, 0.23, 1)
GREEN = (0.12, 0.48, 0.34, 1)
RED = (0.68, 0.20, 0.20, 1)
GRAY = (0.42, 0.47, 0.54, 1)
ORANGE = (0.85, 0.47, 0.15, 1)

CURRENCIES = ["USD", "SYP", "EUR"]
KINDS = {"debt": "إضافة دين", "receipt": "سند قبض", "payment": "سند دفع"}
KIND_KEYS = {v: k for k, v in KINDS.items()}

BACKUP_DIR = "/storage/emulated/0/Download/AccountingBackups"
EXPORT_DIR = "/storage/emulated/0/Download"


class ArabicSpinnerOption(SpinnerOption):
    def __init__(self, **kwargs):
        kwargs.setdefault("font_name", "Arabic")
        kwargs.setdefault("font_size", sp(16))
        super().__init__(**kwargs)


def ar(value):
    return get_display(arabic_reshaper.reshape(str(value)))


def now_text():
    return datetime.now().strftime("%Y-%m-%d %H:%M")


class AccountingApp(App):

    def build(self):
        Window.clearcolor = LIGHT
        os.makedirs(self.user_data_dir, exist_ok=True)
        os.makedirs(BACKUP_DIR, exist_ok=True)

        self.db_path = os.path.join(self.user_data_dir, "accounting.db")
        self.init_database()

        self.root_box = BoxLayout(
            orientation="vertical", spacing=dp(6), padding=dp(8)
        )
        self.set_background(self.root_box, LIGHT)
        self.home()
        return self.root_box

    # -----------------------------------------------------
    # قاعدة البيانات
    # -----------------------------------------------------
    def connect(self):
        db = sqlite3.connect(self.db_path)
        db.execute("PRAGMA foreign_keys=ON")
        return db

    def init_database(self):
        db = self.connect()
        cur = db.cursor()

        cur.execute("""
            CREATE TABLE IF NOT EXISTS customers (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT NOT NULL,
                phone TEXT
            )
        """)
        cur.execute("""
            CREATE TABLE IF NOT EXISTS transactions (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                customer_id INTEGER NOT NULL,
                kind TEXT NOT NULL,
                amount REAL NOT NULL,
                currency TEXT NOT NULL,
                note TEXT,
                created_at TEXT NOT NULL,
                FOREIGN KEY(customer_id) REFERENCES customers(id)
            )
        """)
        cur.execute("""
            CREATE TABLE IF NOT EXISTS transfers (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                customer1 INTEGER NOT NULL,
                customer2 INTEGER NOT NULL,
                amount REAL NOT NULL,
                currency TEXT NOT NULL,
                fee1 REAL DEFAULT 0,
                fee1_direction TEXT DEFAULT 'عليه',
                fee2 REAL DEFAULT 0,
                fee2_direction TEXT DEFAULT 'عليه',
                note TEXT,
                created_at TEXT NOT NULL,
                FOREIGN KEY(customer1) REFERENCES customers(id),
                FOREIGN KEY(customer2) REFERENCES customers(id)
            )
        """)
        cur.execute("""
            CREATE TABLE IF NOT EXISTS extra_debts (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                person TEXT NOT NULL,
                amount REAL NOT NULL,
                currency TEXT NOT NULL,
                direction TEXT NOT NULL,
                note TEXT,
                created_at TEXT NOT NULL
            )
        """)
        db.commit()
        db.close()

    # -----------------------------------------------------
    # عناصر الواجهة
    # -----------------------------------------------------
    def set_background(self, widget, color):
        with widget.canvas.before:
            Color(*color)
            rect = Rectangle(pos=widget.pos, size=widget.size)
        widget.bind(
            pos=lambda obj, value: setattr(rect, "pos", value),
            size=lambda obj, value: setattr(rect, "size", value)
        )

    def label(self, text, size=17, color=DARK, height=dp(42)):
        lbl = Label(
            text=ar(text), font_name="Arabic", font_size=sp(size),
            color=color, size_hint_y=None, height=height,
            halign="center", valign="middle"
        )
        lbl.bind(size=lambda obj, value: setattr(obj, "text_size", value))
        return lbl

    def button(self, text, callback, color=SILVER,
               text_color=NAVY, height=dp(56)):
        btn = Button(
            text=ar(text), font_name="Arabic", font_size=sp(16),
            color=text_color, background_normal="", background_down="",
            background_color=color, size_hint_y=None, height=height
        )
        btn.bind(on_press=callback)
        return btn

    def input_field(self, hint, number=False):
        return TextInput(
            hint_text=ar(hint), multiline=False,
            input_filter="float" if number else None,
            font_name="Arabic", font_size=sp(17),
            size_hint_y=None, height=dp(50),
            foreground_color=NAVY, background_normal="",
            background_active="", background_color=WHITE,
            cursor_color=NAVY, hint_text_color=GRAY,
            padding=[dp(12), dp(12)], halign="right"
        )

    def clear(self):
        self.root_box.clear_widgets()

    def header(self, title, back=True):
        bar = BoxLayout(
            size_hint_y=None, height=dp(58),
            spacing=dp(8), padding=[dp(8), dp(6)]
        )
        self.set_background(bar, NAVY)
        bar.add_widget(self.label(title, 20, WHITE, dp(46)))
        if back:
            b = self.button("الرئيسية", lambda x: self.home(),
                            SILVER, NAVY, dp(44))
            b.size_hint_x = None
            b.width = dp(105)
            bar.add_widget(b)
        self.root_box.add_widget(bar)

    def page(self, title):
        self.clear()
        self.header(title)

    def make_scroll(self):
        scroll = ScrollView(do_scroll_x=False)
        items = GridLayout(
            cols=1, spacing=dp(6), padding=dp(2),
            size_hint_y=None
        )
        items.bind(minimum_height=items.setter("height"))
        scroll.add_widget(items)
        return scroll, items

    def popup(self, content, height=0.7):
        p = Popup(
            title="", content=content, size_hint=(0.94, height),
            background_color=WHITE, separator_color=SILVER,
            auto_dismiss=True
        )
        p.open()
        return p

    def message(self, text):
        box = BoxLayout(orientation="vertical", padding=dp(14), spacing=dp(10))
        box.add_widget(self.label(text, 17, NAVY))
        p = self.popup(box, 0.32)
        box.add_widget(self.button("حسناً", lambda x: p.dismiss(),
                                   NAVY, WHITE, dp(48)))

    def confirm(self, text, yes_callback):
        box = BoxLayout(orientation="vertical", padding=dp(14), spacing=dp(10))
        box.add_widget(self.label(text, 17, RED))
        p = self.popup(box, 0.36)
        row = BoxLayout(spacing=dp(8), size_hint_y=None, height=dp(50))
        row.add_widget(self.button(
            "نعم", lambda x: (p.dismiss(), yes_callback()), RED, WHITE, dp(48)
        ))
        row.add_widget(self.button("إلغاء", lambda x: p.dismiss(), height=dp(48)))
        box.add_widget(row)

    def currency_spinner(self, value="USD"):
        return Spinner(
            text=value, values=CURRENCIES, option_cls=ArabicSpinnerOption,
            size_hint_y=None, height=dp(46), font_name="Arabic",
            font_size=sp(16), color=NAVY
        )

    # -----------------------------------------------------
    # الصفحة الرئيسية
    # -----------------------------------------------------
    def home(self, instance=None):
        self.clear()

        top = BoxLayout(
            orientation="vertical", size_hint_y=None,
            height=dp(100), padding=dp(8), spacing=dp(2)
        )
        self.set_background(top, NAVY)
        top.add_widget(self.label("نظام المحاسبة", 26, WHITE, dp(50)))
        top.add_widget(self.label("إدارة الحسابات والعملاء", 14, SILVER, dp(28)))
        self.root_box.add_widget(top)

        self.root_box.add_widget(
            self.label("لوحة التحكم", 18, NAVY, dp(40))
        )

        scroll = ScrollView(do_scroll_x=False)
        grid = GridLayout(
            cols=2, spacing=dp(9), padding=dp(4), size_hint_y=None
        )
        grid.bind(minimum_height=grid.setter("height"))

        items = [
            ("الصندوق", self.cashbox, WHITE),
            ("العملاء", self.customers_page, WHITE),
            ("إضافة عميل", self.add_customer, GREEN),
            ("سند قبض", lambda x: self.choose_customer("receipt"), WHITE),
            ("سند دفع", lambda x: self.choose_customer("payment"), WHITE),
            ("تحويل بين العملاء", self.transfer_form, WHITE),
            ("الديون المتفرقة", self.extra_debts_page, WHITE),
            ("التقارير", self.reports, WHITE),
            ("تصدير Excel", self.export_excel, ORANGE),
            ("نسخ احتياطي", self.backup_database, BLUE),
            ("استعادة نسخة", self.restore_page, RED),
        ]

        for title, action, color in items:
            text_color = WHITE if color in (GREEN, BLUE, RED, ORANGE, NAVY) else NAVY
            grid.add_widget(self.button(title, action, color, text_color, dp(58)))

        scroll.add_widget(grid)
        self.root_box.add_widget(scroll)

    # -----------------------------------------------------
    # نسخ احتياطي واستعادة
    # -----------------------------------------------------
    def backup_database(self, instance=None):
        try:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            dest = os.path.join(BACKUP_DIR, f"accounting_{timestamp}.db")
            shutil.copy2(self.db_path, dest)
            self.message("تم إنشاء نسخة احتياطية بنجاح\n" + dest)
        except Exception as e:
            self.message("فشل النسخ الاحتياطي:\n" + str(e))

    def restore_page(self, instance=None):
        self.page("استعادة نسخة احتياطية")
        self.root_box.add_widget(self.label(
            "اختر نسخة لاستعادتها (سيتم استبدال البيانات الحالية)",
            15, RED, dp(50)
        ))

        scroll, items = self.make_scroll()
        try:
            files = [f for f in os.listdir(BACKUP_DIR) if f.endswith(".db")]
            files.sort(reverse=True)
        except Exception:
            files = []

        if not files:
            items.add_widget(self.label("لا توجد نسخ احتياطية", 16))
        else:
            for f in files[:20]:
                path = os.path.join(BACKUP_DIR, f)
                items.add_widget(self.button(
                    f,
                    lambda x, p=path: self.confirm_restore(p),
                    WHITE, NAVY, dp(54)
                ))
        self.root_box.add_widget(scroll)

    def confirm_restore(self, path):
        def do_restore():
            try:
                shutil.copy2(path, self.db_path)
                self.message("تمت الاستعادة بنجاح")
                self.home()
            except Exception as e:
                self.message("فشلت الاستعادة:\n" + str(e))
        self.confirm(
            "هل أنت متأكد من استعادة هذه النسخة؟\nسيتم استبدال جميع البيانات الحالية.",
            do_restore
        )

    # -----------------------------------------------------
    # تصدير إلى Excel
    # -----------------------------------------------------
    def export_excel(self, instance=None):
        if not OPENPYXL_AVAILABLE:
            self.message(
                "مكتبة openpyxl غير مثبتة.\n"
                "افتح Pip في Pydroid وثبت: openpyxl"
            )
            return

        try:
            wb = Workbook()

            header_fill = PatternFill("solid", fgColor="0E1A30")
            header_font = Font(bold=True, color="FFFFFF", size=12)
            thin = Border(
                left=Side(style='thin'),
                right=Side(style='thin'),
                top=Side(style='thin'),
                bottom=Side(style='thin')
            )

            def style_header(ws, cols):
                for col in range(1, cols + 1):
                    cell = ws.cell(1, col)
                    cell.fill = header_fill
                    cell.font = header_font
                    cell.alignment = Alignment(horizontal="center")
                    cell.border = thin

            db = self.connect()

            # ورقة العملاء
            ws = wb.active
            ws.title = "العملاء"
            headers = ["الرقم", "الاسم", "الهاتف", "USD", "SYP", "EUR"]
            ws.append(headers)
            style_header(ws, len(headers))

            customers = db.execute(
                "SELECT id, name, phone FROM customers ORDER BY name"
            ).fetchall()
            for cid, name, phone in customers:
                row = [
                    cid, name, phone or "",
                    round(self.balance(cid, "USD"), 2),
                    round(self.balance(cid, "SYP"), 2),
                    round(self.balance(cid, "EUR"), 2)
                ]
                ws.append(row)

            for col in ws.columns:
                max_len = max(len(str(cell.value or "")) for cell in col)
                ws.column_dimensions[col[0].column_letter].width = min(max_len + 4, 30)

            # ورقة الحركات
            ws2 = wb.create_sheet("الحركات")
            headers2 = ["الرقم", "العميل", "النوع", "المبلغ", "العملة", "البيان", "التاريخ"]
            ws2.append(headers2)
            style_header(ws2, len(headers2))

            rows = db.execute("""
                SELECT t.id, c.name, t.kind, t.amount, t.currency, t.note, t.created_at
                FROM transactions t
                LEFT JOIN customers c ON c.id = t.customer_id
                ORDER BY t.id DESC
            """).fetchall()
            for r in rows:
                ws2.append([
                    r[0], r[1] or "", KINDS.get(r[2], r[2]),
                    r[3], r[4], r[5] or "", r[6]
                ])

            # ورقة التحويلات
            ws3 = wb.create_sheet("التحويلات")
            headers3 = ["الرقم", "من", "إلى", "المبلغ", "العملة",
                        "أجور المحول", "اتجاه1", "أجور المستفيد", "اتجاه2",
                        "ملاحظات", "التاريخ"]
            ws3.append(headers3)
            style_header(ws3, len(headers3))

            rows = db.execute("""
                SELECT t.id, c1.name, c2.name, t.amount, t.currency,
                       t.fee1, t.fee1_direction, t.fee2, t.fee2_direction,
                       t.note, t.created_at
                FROM transfers t
                LEFT JOIN customers c1 ON c1.id = t.customer1
                LEFT JOIN customers c2 ON c2.id = t.customer2
                ORDER BY t.id DESC
            """).fetchall()
            for r in rows:
                ws3.append(list(r))

            # ورقة الديون المتفرقة
            ws4 = wb.create_sheet("الديون المتفرقة")
            headers4 = ["الرقم", "الشخص", "المبلغ", "العملة", "الاتجاه", "ملاحظات", "التاريخ"]
            ws4.append(headers4)
            style_header(ws4, len(headers4))

            rows = db.execute(
                "SELECT id, person, amount, currency, direction, note, created_at "
                "FROM extra_debts ORDER BY id DESC"
            ).fetchall()
            for r in rows:
                ws4.append(list(r))

            # ورقة الصندوق
            ws5 = wb.create_sheet("الصندوق")
            ws5.append(["العملة", "المقبوض", "المدفوع", "رصيد الصندوق"])
            style_header(ws5, 4)

            for currency in CURRENCIES:
                rows = db.execute("""
                    SELECT kind, COALESCE(SUM(amount),0)
                    FROM transactions WHERE currency=?
                    AND kind IN ('receipt','payment') GROUP BY kind
                """, (currency,)).fetchall()
                totals = dict(rows)
                received = totals.get("receipt", 0) or 0
                paid = totals.get("payment", 0) or 0
                ws5.append([currency, received, paid, received - paid])

            db.close()

            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            path = os.path.join(EXPORT_DIR, f"تقرير_محاسبة_{timestamp}.xlsx")
            wb.save(path)
            self.message("تم تصدير التقرير بنجاح\n" + path)

        except Exception as e:
            self.message("فشل التصدير:\n" + str(e))

    # -----------------------------------------------------
    # العملاء
    # -----------------------------------------------------
    def customers_page(self, instance=None):
        self.page("قائمة العملاء")
        self.root_box.add_widget(self.button(
            "إضافة عميل جديد", self.add_customer, GREEN, WHITE, dp(52)
        ))
        search = self.input_field("ابحث عن اسم العميل")
        self.root_box.add_widget(search)
        scroll, items = self.make_scroll()
        self.root_box.add_widget(scroll)

        def refresh(term=""):
            items.clear_widgets()
            db = self.connect()
            rows = db.execute("""
                SELECT id,name,phone FROM customers
                WHERE name LIKE ? ORDER BY name
            """, ("%" + term + "%",)).fetchall()
            db.close()
            if not rows:
                items.add_widget(self.label("لا يوجد عملاء", 16))
            for cid, name, phone in rows:
                caption = name + ((" - " + phone) if phone else "")
                items.add_widget(self.button(
                    caption,
                    lambda x, customer_id=cid: self.customer_details(customer_id),
                    WHITE, NAVY, dp(56)
                ))
        search.bind(text=lambda obj, value: refresh(value))
        refresh()

    def add_customer(self, instance=None, customer_id=None):
        editing = customer_id is not None
        box = BoxLayout(orientation="vertical", padding=dp(12), spacing=dp(8))
        box.add_widget(self.label(
            "تعديل بيانات العميل" if editing else "إضافة عميل جديد", 20, NAVY
        ))
        name = self.input_field("اسم العميل")
        phone = self.input_field("رقم الهاتف")
        if editing:
            db = self.connect()
            row = db.execute(
                "SELECT name,phone FROM customers WHERE id=?", (customer_id,)
            ).fetchone()
            db.close()
            if row:
                name.text, phone.text = row[0], row[1] or ""
        box.add_widget(name)
        box.add_widget(phone)
        p = self.popup(box, 0.48)

        def save(btn):
            if not name.text.strip():
                self.message("يرجى إدخال اسم العميل")
                return
            db = self.connect()
            if editing:
                db.execute(
                    "UPDATE customers SET name=?,phone=? WHERE id=?",
                    (name.text.strip(), phone.text.strip(), customer_id)
                )
            else:
                db.execute(
                    "INSERT INTO customers(name,phone) VALUES(?,?)",
                    (name.text.strip(), phone.text.strip())
                )
            db.commit()
            db.close()
            p.dismiss()
            self.customers_page()
            self.message("تم حفظ بيانات العميل")
        box.add_widget(self.button("حفظ", save, GREEN, WHITE, dp(50)))

    # -----------------------------------------------------
    # حساب الرصيد
    # -----------------------------------------------------
    def balance(self, cid, currency):
        db = self.connect()
        rows = db.execute("""
            SELECT kind, COALESCE(SUM(amount),0)
            FROM transactions
            WHERE customer_id=? AND currency=? GROUP BY kind
        """, (cid, currency)).fetchall()
        totals = dict(rows)
        result = (
            (totals.get("debt", 0) or 0)
            - (totals.get("receipt", 0) or 0)
            - (totals.get("payment", 0) or 0)
        )

        sent = db.execute("""
            SELECT COALESCE(SUM(amount),0), COALESCE(SUM(fee1),0)
            FROM transfers WHERE customer1=? AND currency=?
        """, (cid, currency)).fetchone()
        received = db.execute("""
            SELECT COALESCE(SUM(amount),0), COALESCE(SUM(fee2),0)
            FROM transfers WHERE customer2=? AND currency=?
        """, (cid, currency)).fetchone()

        result += (sent[0] or 0) - (received[0] or 0)

        fee_rows = db.execute("""
            SELECT fee1,fee1_direction FROM transfers
            WHERE customer1=? AND currency=?
            UNION ALL
            SELECT fee2,fee2_direction FROM transfers
            WHERE customer2=? AND currency=?
        """, (cid, currency, cid, currency)).fetchall()
        for fee, direction in fee_rows:
            fee = fee or 0
            result += fee if direction == "عليه" else -fee

        db.close()
        return result

    def customer_details(self, cid):
        self.page("حساب العميل")
        db = self.connect()
        customer = db.execute(
            "SELECT name,phone FROM customers WHERE id=?", (cid,)
        ).fetchone()
        transactions = db.execute("""
            SELECT id,kind,amount,currency,note,created_at
            FROM transactions WHERE customer_id=? ORDER BY id DESC
        """, (cid,)).fetchall()
        transfers = db.execute("""
            SELECT id,customer1,customer2,amount,currency,
                   fee1,fee1_direction,fee2,fee2_direction,note,created_at
            FROM transfers
            WHERE customer1=? OR customer2=? ORDER BY id DESC
        """, (cid, cid)).fetchall()
        db.close()

        if not customer:
            self.message("العميل غير موجود")
            self.customers_page()
            return

        self.root_box.add_widget(self.label(customer[0], 22, NAVY))
        if customer[1]:
            self.root_box.add_widget(self.label("الهاتف: " + customer[1], 14))
        self.root_box.add_widget(self.button(
            "تعديل بيانات العميل",
            lambda x: self.add_customer(customer_id=cid), height=dp(48)
        ))

        self.root_box.add_widget(self.label("الأرصدة حسب العملة", 17, NAVY))
        for currency in CURRENCIES:
            amount = self.balance(cid, currency)
            self.root_box.add_widget(self.button(
                currency + "     " + format(amount, ".2f"),
                lambda x, c=currency: self.currency_details(cid, c),
                WHITE, NAVY, dp(46)
            ))

        self.root_box.add_widget(self.label("سجل الحركات", 17, NAVY))
        scroll, items = self.make_scroll()
        for tid, kind, amount, currency, note, date in transactions:
            text = (KINDS.get(kind, kind) + " | " + format(amount, ".2f")
                    + " " + currency + "\n" + date)
            if note:
                text += "\n" + note
            items.add_widget(self.button(
                text,
                lambda x, transaction_id=tid:
                    self.transaction_details(transaction_id, cid),
                WHITE, DARK, dp(74)
            ))
        if not transactions:
            items.add_widget(self.label("لا توجد حركات عادية", 14))
        self.root_box.add_widget(scroll)

        self.root_box.add_widget(self.label("التحويلات المرتبطة", 17, NAVY))
        tscroll, titems = self.make_scroll()
        for row in transfers:
            (tid, c1, c2, amount, currency, fee1, dir1,
             fee2, dir2, note, date) = row
            db2 = self.connect()
            names = db2.execute(
                "SELECT id,name FROM customers WHERE id IN (?,?)", (c1, c2)
            ).fetchall()
            db2.close()
            name_map = dict(names)
            direction_text = (
                ("من: " + name_map.get(c1, "?") + " ← إلى: " + name_map.get(c2, "?"))
            )
            text = (direction_text + "\n" + format(amount, ".2f") + " " +
                    currency + " | " + date)
            if fee1 or fee2:
                text += "\nأجور المحوّل: " + str(fee1) + " (" + dir1 + ")"
                text += " | أجور المستفيد: " + str(fee2) + " (" + dir2 + ")"
            if note:
                text += "\n" + note
            titems.add_widget(self.button(
                text,
                lambda x, transfer_id=tid, owner=cid:
                    self.transfer_details(transfer_id, owner),
                SILVER, NAVY, dp(88)
            ))
        if not transfers:
            titems.add_widget(self.label("لا توجد تحويلات", 14))
        self.root_box.add_widget(tscroll)

        actions = GridLayout(cols=3, spacing=dp(6),
                             size_hint_y=None, height=dp(52))
        for key, title in [("debt", "إضافة دين"), ("receipt", "قبض"), ("payment", "دفع")]:
            actions.add_widget(self.button(
                title, lambda x, k=key: self.transaction_form(cid, k),
                SILVER, NAVY, dp(48)
            ))
        self.root_box.add_widget(actions)

        self.root_box.add_widget(self.button(
            "حذف العميل",
            lambda x: self.delete_customer(cid),
            RED, WHITE, dp(48)
        ))

    def delete_customer(self, cid):
        db = self.connect()
        linked = db.execute("""
            SELECT COUNT(*) FROM transfers
            WHERE customer1=? OR customer2=?
        """, (cid, cid)).fetchone()[0]
        db.close()
        if linked:
            self.message(
                "لا يمكن حذف هذا العميل لأن لديه تحويلات مرتبطة.\n"
                "سيبقى حسابه وسجل تحويلاته محفوظاً."
            )
            return

        balances = [self.balance(cid, c) for c in CURRENCIES]
        if any(abs(v) > 0.000001 for v in balances):
            self.message(
                "لا يمكن حذف العميل لأن رصيده ليس صفراً في جميع العملات."
            )
            return

        def do_delete():
            db2 = self.connect()
            db2.execute("DELETE FROM transactions WHERE customer_id=?", (cid,))
            db2.execute("DELETE FROM customers WHERE id=?", (cid,))
            db2.commit()
            db2.close()
            self.customers_page()
            self.message("تم حذف العميل")
        self.confirm("رصيد العميل صفر ولا توجد تحويلات مرتبطة. هل تريد حذفه؟",
                     do_delete)

    # -----------------------------------------------------
    # كشف العملة
    # -----------------------------------------------------
    def currency_details(self, cid, currency):
        self.page("كشف حساب العملة")
        db = self.connect()
        customer = db.execute(
            "SELECT name FROM customers WHERE id=?", (cid,)
        ).fetchone()
        rows = db.execute("""
            SELECT kind,amount,note,created_at FROM transactions
            WHERE customer_id=? AND currency=? ORDER BY id DESC
        """, (cid, currency)).fetchall()
        db.close()
        self.root_box.add_widget(self.label(customer[0] if customer else "", 19, NAVY))
        self.root_box.add_widget(self.label(
            currency + " : " + format(self.balance(cid, currency), ".2f"),
            20, BLUE
        ))
        scroll, items = self.make_scroll()
        for kind, amount, note, date in rows:
            text = KINDS.get(kind, kind) + " | " + format(amount, ".2f") + "\n" + date
            if note:
                text += "\n" + note
            items.add_widget(self.label(text, 15, DARK, dp(68)))
        self.root_box.add_widget(scroll)
        self.root_box.add_widget(self.button(
            "رجوع إلى حساب العميل", lambda x: self.customer_details(cid), height=dp(48)
        ))

    # -----------------------------------------------------
    # اختيار عميل لسند
    # -----------------------------------------------------
    def choose_customer(self, kind):
        db = self.connect()
        rows = db.execute("SELECT id,name FROM customers ORDER BY name").fetchall()
        db.close()
        if not rows:
            self.message("أضف عميلاً أولاً")
            return
        box = BoxLayout(orientation="vertical", padding=dp(10), spacing=dp(7))
        box.add_widget(self.label("اختر العميل", 19, NAVY))
        scroll, items = self.make_scroll()
        box.add_widget(scroll)
        p = self.popup(box, 0.75)
        for cid, name in rows:
            items.add_widget(self.button(
                name,
                lambda x, customer_id=cid:
                    (p.dismiss(), self.transaction_form(customer_id, kind)),
                height=dp(50)
            ))

    # -----------------------------------------------------
    # سندات
    # -----------------------------------------------------
    def transaction_form(self, cid, kind, tid=None):
        editing = tid is not None
        old = None
        if editing:
            db = self.connect()
            old = db.execute("""
                SELECT kind,amount,currency,note FROM transactions WHERE id=?
            """, (tid,)).fetchone()
            db.close()
            if not old:
                self.message("الحركة غير موجودة")
                return

        box = BoxLayout(orientation="vertical", padding=dp(12), spacing=dp(7))
        box.add_widget(self.label("تعديل الحركة" if editing else KINDS[kind], 19, NAVY))
        kind_spinner = Spinner(
            text=KINDS[old[0]] if editing else KINDS[kind],
            values=list(KINDS.values()), size_hint_y=None,
            height=dp(46), font_name="Arabic", font_size=sp(15),
            color=NAVY, disabled=not editing
        )
        amount = self.input_field("المبلغ بالأرقام الإنجليزية", True)
        currency = self.currency_spinner(old[2] if editing else "USD")
        note = self.input_field("البيان والملاحظات")
        if editing:
            amount.text = str(old[1])
            note.text = old[3] or ""
        for widget in [self.label("نوع الحركة"), kind_spinner,
                       self.label("المبلغ"), amount,
                       self.label("العملة"), currency,
                       self.label("البيان"), note]:
            box.add_widget(widget)
        p = self.popup(box, 0.88)

        def save(btn):
            try:
                value = float(amount.text)
                if value <= 0:
                    raise ValueError
            except ValueError:
                self.message("أدخل مبلغاً صحيحاً أكبر من صفر")
                return
            selected = KIND_KEYS.get(kind_spinner.text)
            if not selected:
                self.message("نوع الحركة غير صحيح")
                return
            db = self.connect()
            if editing:
                db.execute("""
                    UPDATE transactions SET kind=?,amount=?,currency=?,note=?
                    WHERE id=?
                """, (selected, value, currency.text, note.text.strip(), tid))
            else:
                db.execute("""
                    INSERT INTO transactions
                    (customer_id,kind,amount,currency,note,created_at)
                    VALUES(?,?,?,?,?,?)
                """, (cid, selected, value, currency.text,
                      note.text.strip(), now_text()))
            db.commit()
            db.close()
            p.dismiss()
            self.customer_details(cid)
        box.add_widget(self.button("حفظ الحركة", save, GREEN, WHITE, dp(50)))

    def transaction_details(self, tid, cid):
        db = self.connect()
        row = db.execute("""
            SELECT kind,amount,currency,note,created_at
            FROM transactions WHERE id=?
        """, (tid,)).fetchone()
        db.close()
        if not row:
            self.message("الحركة غير موجودة")
            return
        kind, amount, currency, note, date = row
        box = BoxLayout(orientation="vertical", padding=dp(12), spacing=dp(8))
        box.add_widget(self.label(KINDS.get(kind, kind), 19, NAVY))
        box.add_widget(self.label(format(amount, ".2f") + " " + currency, 18))
        box.add_widget(self.label(date, 14))
        if note:
            box.add_widget(self.label(note, 15))
        p = self.popup(box, 0.56)
        box.add_widget(self.button(
            "تعديل", lambda x: (p.dismiss(), self.transaction_form(cid, kind, tid)),
            height=dp(48)
        ))

        def remove():
            db2 = self.connect()
            db2.execute("DELETE FROM transactions WHERE id=?", (tid,))
            db2.commit()
            db2.close()
            self.customer_details(cid)
        box.add_widget(self.button(
            "حذف الحركة",
            lambda x: (p.dismiss(), self.confirm("هل تريد حذف الحركة؟", remove)),
            RED, WHITE, dp(48)
        ))

    # -----------------------------------------------------
    # التحويل بين العملاء
    # -----------------------------------------------------
    def get_customers(self):
        db = self.connect()
        rows = db.execute("SELECT id,name FROM customers ORDER BY name").fetchall()
        db.close()
        return rows

    def transfer_form(self, instance=None, transfer_id=None, owner_id=None):
        editing = transfer_id is not None
        old = None
        if editing:
            db = self.connect()
            old = db.execute("""
                SELECT customer1,customer2,amount,currency,fee1,
                       fee1_direction,fee2,fee2_direction,note
                FROM transfers WHERE id=?
            """, (transfer_id,)).fetchone()
            db.close()
            if not old:
                self.message("التحويل غير موجود")
                return

        customers = self.get_customers()
        if len(customers) < 2:
            self.message("يجب إضافة عميلين على الأقل لإجراء تحويل")
            return
        names = {cid: name for cid, name in customers}
        ids = [cid for cid, name in customers]
        display_to_id = {ar(name): cid for cid, name in customers}
        values = [ar(name) for cid, name in customers]
        direction_values = [ar("عليه"), ar("له")]
        direction_to_raw = {ar("عليه"): "عليه", ar("له"): "له"}

        box = BoxLayout(orientation="vertical", padding=dp(10), spacing=dp(6))
        box.add_widget(self.label("تحويل بين العملاء", 20, NAVY))

        def spinner_for(cid=None):
            raw_selected = names.get(cid, customers[0][1])
            selected = ar(raw_selected)
            return Spinner(
                text=selected, values=values,
                option_cls=ArabicSpinnerOption, size_hint_y=None,
                height=dp(44), font_name="Arabic", font_size=sp(15),
                color=NAVY
            )

        c1 = spinner_for(old[0] if editing else None)
        c2 = spinner_for(old[1] if editing else ids[1])
        amount = self.input_field("مبلغ التحويل", True)
        currency = self.currency_spinner(old[3] if editing else "USD")
        fee1 = self.input_field("أجور العميل المحوّل (اختياري)", True)
        dir1 = Spinner(
            text=ar(old[5] if editing else "عليه"),
            values=direction_values, option_cls=ArabicSpinnerOption,
            size_hint_y=None, height=dp(42), font_name="Arabic", color=NAVY
        )
        fee2 = self.input_field("أجور العميل المستفيد (اختياري)", True)
        dir2 = Spinner(
            text=ar(old[7] if editing else "عليه"),
            values=direction_values, option_cls=ArabicSpinnerOption,
            size_hint_y=None, height=dp(42), font_name="Arabic", color=NAVY
        )
        note = self.input_field("ملاحظات التحويل")
        if editing:
            amount.text = str(old[2])
            fee1.text = str(old[4] or 0)
            fee2.text = str(old[6] or 0)
            note.text = old[8] or ""

        fields = [
            self.label("العميل المحوّل (المبلغ عليه)"),
            c1,
            self.label("العميل المستفيد (المبلغ له)"),
            c2,
            self.label("المبلغ والعملة"),
            amount, currency,
            self.label("أجور المحوّل ثم اتجاهها"),
            fee1, dir1,
            self.label("أجور المستفيد ثم اتجاهها"),
            fee2, dir2,
            self.label("ملاحظات"),
            note
        ]
        scroll = ScrollView(do_scroll_x=False)
        inner = GridLayout(cols=1, spacing=dp(5), padding=dp(2), size_hint_y=None)
        inner.bind(minimum_height=inner.setter("height"))
        for w in fields:
            inner.add_widget(w)
        scroll.add_widget(inner)
        box.add_widget(scroll)
        p = self.popup(box, 0.94)

        def save(btn):
            id1 = display_to_id.get(c1.text)
            id2 = display_to_id.get(c2.text)
            if id1 is None or id2 is None or id1 == id2:
                self.message("اختر عميلين مختلفين")
                return
            try:
                val = float(amount.text)
                f1 = float(fee1.text or 0)
                f2 = float(fee2.text or 0)
                if val <= 0 or f1 < 0 or f2 < 0:
                    raise ValueError
            except ValueError:
                self.message("تحقق من المبلغ والأجور، يجب أن تكون القيم صحيحة")
                return
            db2 = self.connect()
            data = (id1, id2, val, currency.text, f1,
                    direction_to_raw.get(dir1.text, "عليه"),
                    f2, direction_to_raw.get(dir2.text, "عليه"),
                    note.text.strip())
            if editing:
                db2.execute("""
                    UPDATE transfers SET customer1=?,customer2=?,amount=?,
                    currency=?,fee1=?,fee1_direction=?,fee2=?,
                    fee2_direction=?,note=? WHERE id=?
                """, data + (transfer_id,))
            else:
                db2.execute("""
                    INSERT INTO transfers
                    (customer1,customer2,amount,currency,fee1,fee1_direction,
                     fee2,fee2_direction,note,created_at)
                    VALUES(?,?,?,?,?,?,?,?,?,?)
                """, data + (now_text(),))
            db2.commit()
            db2.close()
            p.dismiss()
            if owner_id is not None:
                self.customer_details(owner_id)
            else:
                self.transfer_list()
            self.message("تم حفظ التحويل وتحديث حسابي العميلين")

        box.add_widget(self.button("حفظ التحويل", save, GREEN, WHITE, dp(50)))

    def transfer_list(self):
        self.page("سجل التحويلات")
        scroll, items = self.make_scroll()
        db = self.connect()
        rows = db.execute("""
            SELECT t.id,c1.name,c2.name,t.amount,t.currency,t.created_at
            FROM transfers t
            LEFT JOIN customers c1 ON c1.id=t.customer1
            LEFT JOIN customers c2 ON c2.id=t.customer2
            ORDER BY t.id DESC
        """).fetchall()
        db.close()
        for tid, n1, n2, amount, currency, date in rows:
            items.add_widget(self.button(
                (n1 or "؟") + " ← " + (n2 or "؟") + "\n" +
                format(amount, ".2f") + " " + currency + " | " + date,
                lambda x, transfer_id=tid: self.transfer_details(transfer_id),
                WHITE, NAVY, dp(70)
            ))
        if not rows:
            items.add_widget(self.label("لا توجد تحويلات", 16))
        self.root_box.add_widget(scroll)
        self.root_box.add_widget(self.button(
            "إضافة تحويل جديد", self.transfer_form, GREEN, WHITE, dp(50)
        ))

    def transfer_details(self, transfer_id, owner_id=None):
        db = self.connect()
        row = db.execute("""
            SELECT t.customer1,t.customer2,c1.name,c2.name,t.amount,t.currency,
                   t.fee1,t.fee1_direction,t.fee2,t.fee2_direction,t.note,t.created_at
            FROM transfers t
            LEFT JOIN customers c1 ON c1.id=t.customer1
            LEFT JOIN customers c2 ON c2.id=t.customer2
            WHERE t.id=?
        """, (transfer_id,)).fetchone()
        db.close()
        if not row:
            self.message("التحويل غير موجود")
            return
        c1, c2, n1, n2, amount, currency, f1, d1, f2, d2, note, date = row
        box = BoxLayout(orientation="vertical", padding=dp(12), spacing=dp(7))
        box.add_widget(self.label("تفاصيل التحويل", 20, NAVY))
        box.add_widget(self.label("المحوّل: " + str(n1), 15))
        box.add_widget(self.label("المستفيد: " + str(n2), 15))
        box.add_widget(self.label("المبلغ: " + format(amount, ".2f") + " " + currency, 17))
        box.add_widget(self.label("أجور المحوّل: " + str(f1) + " - " + d1, 14))
        box.add_widget(self.label("أجور المستفيد: " + str(f2) + " - " + d2, 14))
        box.add_widget(self.label("التاريخ: " + date, 13))
        if note:
            box.add_widget(self.label(note, 14))
        p = self.popup(box, 0.76)

        def after_change():
            if owner_id is not None:
                self.customer_details(owner_id)
            else:
                self.transfer_list()

        box.add_widget(self.button(
            "تعديل التحويل",
            lambda x: (p.dismiss(), self.transfer_form(
                transfer_id=transfer_id, owner_id=owner_id
            )), height=dp(48)
        ))

        def remove():
            db2 = self.connect()
            db2.execute("DELETE FROM transfers WHERE id=?", (transfer_id,))
            db2.commit()
            db2.close()
            after_change()
            self.message("تم حذف التحويل وتحديث حسابي العميلين")

        box.add_widget(self.button(
            "حذف التحويل",
            lambda x: (p.dismiss(), self.confirm("حذف التحويل من حسابي العميلين؟", remove)),
            RED, WHITE, dp(48)
        ))

    # -----------------------------------------------------
    # الديون المتفرقة
    # -----------------------------------------------------
    def extra_debts_page(self, instance=None):
        self.page("الديون المتفرقة")
        self.root_box.add_widget(self.button(
            "إضافة دين متفرق", self.extra_debt_form, GREEN, WHITE, dp(50)
        ))
        scroll, items = self.make_scroll()
        db = self.connect()
        rows = db.execute("""
            SELECT id,person,amount,currency,direction,note,created_at
            FROM extra_debts ORDER BY id DESC
        """).fetchall()
        db.close()
        for row in rows:
            did, person, amount, currency, direction, note, date = row
            text = person + " | " + format(amount, ".2f") + " " + currency
            text += " | " + direction + "\n" + date
            if note:
                text += "\n" + note
            items.add_widget(self.button(
                text, lambda x, debt_id=did: self.extra_debt_details(debt_id),
                WHITE, NAVY, dp(76)
            ))
        if not rows:
            items.add_widget(self.label("لا توجد ديون متفرقة", 16))
        self.root_box.add_widget(scroll)

    def extra_debt_form(self, instance=None, debt_id=None):
        editing = debt_id is not None
        old = None
        if editing:
            db = self.connect()
            old = db.execute("""
                SELECT person,amount,currency,direction,note
                FROM extra_debts WHERE id=?
            """, (debt_id,)).fetchone()
            db.close()
            if not old:
                self.message("الدين غير موجود")
                return
        box = BoxLayout(orientation="vertical", padding=dp(12), spacing=dp(7))
        person = self.input_field("اسم الشخص أو الجهة")
        amount = self.input_field("المبلغ", True)
        currency = self.currency_spinner(old[2] if editing else "USD")
        direction = Spinner(
            text=old[3] if editing else "له",
            values=["له", "عليه"], size_hint_y=None,
            height=dp(46), color=NAVY
        )
        note = self.input_field("ملاحظات")
        if editing:
            person.text, amount.text, note.text = old[0], str(old[1]), old[4] or ""
        for w in [self.label("اسم الشخص"), person, self.label("المبلغ"),
                  amount, self.label("العملة"), currency,
                  self.label("اتجاه الدين"), direction,
                  self.label("ملاحظات"), note]:
            box.add_widget(w)
        p = self.popup(box, 0.88)

        def save(btn):
            if not person.text.strip():
                self.message("أدخل اسم الشخص")
                return
            try:
                val = float(amount.text)
                if val <= 0:
                    raise ValueError
            except ValueError:
                self.message("أدخل مبلغاً صحيحاً")
                return
            data = (person.text.strip(), val, currency.text,
                    direction.text, note.text.strip())
            db2 = self.connect()
            if editing:
                db2.execute("""
                    UPDATE extra_debts SET person=?,amount=?,currency=?,
                    direction=?,note=? WHERE id=?
                """, data + (debt_id,))
            else:
                db2.execute("""
                    INSERT INTO extra_debts
                    (person,amount,currency,direction,note,created_at)
                    VALUES(?,?,?,?,?,?)
                """, data + (now_text(),))
            db2.commit()
            db2.close()
            p.dismiss()
            self.extra_debts_page()
        box.add_widget(self.button("حفظ", save, GREEN, WHITE, dp(50)))

    def extra_debt_details(self, debt_id):
        db = self.connect()
        row = db.execute("""
            SELECT person,amount,currency,direction,note
            FROM extra_debts WHERE id=?
        """, (debt_id,)).fetchone()
        db.close()
        if not row:
            self.message("الدين غير موجود")
            return
        person, amount, currency, direction, note = row
        box = BoxLayout(orientation="vertical", padding=dp(12), spacing=dp(8))
        box.add_widget(self.label(person, 19, NAVY))
        box.add_widget(self.label(format(amount, ".2f") + " " + currency, 17))
        box.add_widget(self.label("الاتجاه: " + direction, 15))
        if note:
            box.add_widget(self.label(note, 14))
        p = self.popup(box, 0.54)
        box.add_widget(self.button(
            "تعديل", lambda x: (p.dismiss(), self.extra_debt_form(debt_id=debt_id)),
            height=dp(48)
        ))

        def remove():
            db2 = self.connect()
            db2.execute("DELETE FROM extra_debts WHERE id=?", (debt_id,))
            db2.commit()
            db2.close()
            self.extra_debts_page()
        box.add_widget(self.button(
            "حذف",
            lambda x: (p.dismiss(), self.confirm("هل تريد حذف هذا الدين؟", remove)),
            RED, WHITE, dp(48)
        ))

    # -----------------------------------------------------
    # الصندوق والتقارير
    # -----------------------------------------------------
    def cashbox(self, instance=None):
        self.page("الصندوق")
        self.root_box.add_widget(self.label("أرصدة الصندوق حسب العملة", 18, NAVY))
        db = self.connect()
        for currency in CURRENCIES:
            rows = db.execute("""
                SELECT kind,COALESCE(SUM(amount),0)
                FROM transactions WHERE currency=?
                AND kind IN ('receipt','payment') GROUP BY kind
            """, (currency,)).fetchall()
            totals = dict(rows)
            received = totals.get("receipt", 0) or 0
            paid = totals.get("payment", 0) or 0
            self.root_box.add_widget(self.button(
                currency + "\nالمقبوض: " + format(received, ".2f") +
                " | المدفوع: " + format(paid, ".2f") +
                "\nرصيد الصندوق: " + format(received-paid, ".2f"),
                lambda x: None, WHITE, NAVY, dp(80)
            ))
        db.close()

    def reports(self, instance=None):
        self.page("التقارير")
        db = self.connect()
        count = db.execute("SELECT COUNT(*) FROM customers").fetchone()[0]
        db.close()
        self.root_box.add_widget(self.label("عدد العملاء: " + str(count), 17, NAVY))
        self.root_box.add_widget(self.label("ملخص أرصدة العملاء", 18, NAVY))
        for currency in CURRENCIES:
            db2 = self.connect()
            ids = db2.execute("SELECT id FROM customers").fetchall()
            db2.close()
            total = sum(self.balance(row[0], currency) for row in ids)
            self.root_box.add_widget(self.button(
                currency + " | صافي أرصدة العملاء: " + format(total, ".2f"),
                lambda x: None, WHITE, NAVY, dp(54)
            ))
        self.root_box.add_widget(self.button(
            "عرض سجل التحويلات", lambda x: self.transfer_list(), height=dp(50)
        ))
        self.root_box.add_widget(self.button(
            "تصدير إلى Excel", self.export_excel, ORANGE, WHITE, dp(50)
        ))


if __name__ == "__main__":
    AccountingApp().run()
