import os
import sys
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
import subprocess
import platform

# محاولة استيراد مكتبة السحب والإسقاط المدعومة إن وجدت، مع وضع بديل آمن
try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
    HAS_DND = True
except ImportError:
    HAS_DND = False

class RoundedButton(tk.Canvas):
    def __init__(self, parent, text, command, bg_color, hover_color, fg_color="white", width=160, height=48, radius=18, font=("Arial", 12, "bold")):
        super().__init__(parent, width=width, height=height, bg=parent["bg"], highlightthickness=0)
        self.command = command
        self.bg_color = bg_color
        self.hover_color = hover_color
        self.fg_color = fg_color
        self.radius = radius
        self.width = width
        self.height = height
        self.font = font
        self.font_text = text
        
        self._draw(self.bg_color)
        self.bind("<Enter>", self._on_enter)
        self.bind("<Leave>", self._on_leave)
        self.bind("<Button-1>", self._on_click)

    def _draw(self, color):
        self.delete("all")
        r, w, h = self.radius, self.width, self.height
        
        self.create_arc((0, 0, 2*r, 2*r), start=90, extent=90, fill=color, outline=color)
        self.create_arc((w-2*r, 0, w, 2*r), start=0, extent=90, fill=color, outline=color)
        self.create_arc((0, h-2*r, 2*r, h), start=180, extent=90, fill=color, outline=color)
        self.create_arc((w-2*r, h-2*r, w, h), start=270, extent=90, fill=color, outline=color)
        
        self.create_rectangle((r, 0, w-r, h), fill=color, outline=color)
        self.create_rectangle((0, r, w, h-r), fill=color, outline=color)
        
        self.create_text(w/2, h/2, text=self.font_text, fill=self.fg_color, font=self.font)

    def _on_enter(self, event): self._draw(self.hover_color)
    def _on_leave(self, event): self._draw(self.bg_color)
    def _on_click(self, event): 
        if self.command: self.command()

# --- استخراج المواد الفريدة لطلبيات المبيع ---
def get_unique_items(input_path):
    wb_src = openpyxl.load_workbook(input_path, data_only=True)
    ws_src = wb_src.active
    
    header_row = 1
    for r in range(1, min(10, ws_src.max_row + 1)):
        row_vals = [str(ws_src.cell(row=r, column=c).value or '').strip() for c in range(1, ws_src.max_column + 1)]
        if any(k in row_vals for k in ['المادة', 'تجهيز', 'اسم العميل']):
            header_row = r
            break
            
    col_item = 3
    for c in range(1, ws_src.max_column + 1):
        val = str(ws_src.cell(row=header_row, column=c).value or '').strip()
        if val in ['المادة', 'اسم المادة']:
            col_item = c
            break

    items = set()
    for r in range(header_row + 1, ws_src.max_row + 1):
        val = ws_src.cell(row=r, column=col_item).value
        if val is not None and str(val).strip() != '':
            items.add(str(val).strip())
            
    return sorted(list(items)), header_row

# --- معالجة الملف الأول: محلاية + خمس مواد ---
def format_sales_orders(input_path, output_path, selected_items, remove_empty=True):
    wb_src = openpyxl.load_workbook(input_path, data_only=True)
    ws_src = wb_src.active
    
    header_row = 1
    for r in range(1, min(10, ws_src.max_row + 1)):
        row_vals = [str(ws_src.cell(row=r, column=c).value or '').strip() for c in range(1, ws_src.max_column + 1)]
        if any(k in row_vals for k in ['المادة', 'تجهيز', 'اسم العميل']):
            header_row = r
            break

    col_item, col_client, col_qty = 3, 5, 6
    for c in range(1, ws_src.max_column + 1):
        val = str(ws_src.cell(row=header_row, column=c).value or '').strip()
        if val in ['المادة', 'اسم المادة']: col_item = c
        elif val in ['اسم العميل', 'العميل', 'المحل']: col_client = c
        elif val in ['تجهيز', 'الكمية', 'كمية']: col_qty = c

    wb_out = openpyxl.Workbook()
    ws_out = wb_out.active
    ws_out.title = "محلاية و خمس مواد"
    ws_out.views.sheetView[0].rightToLeft = True

    ws_out.append(["المادة", "اسم العميل", "الكمية"])

    totals_per_item = {}

    for r in range(header_row + 1, ws_src.max_row + 1):
        qty_val = ws_src.cell(row=r, column=col_qty).value
        
        if remove_empty and (qty_val is None or str(qty_val).strip() in ['', '0', 'None', '0.0']):
            continue

        item_val = str(ws_src.cell(row=r, column=col_item).value or '').strip()
        client_val = ws_src.cell(row=r, column=col_client).value

        if selected_items and item_val not in selected_items:
            continue

        try:
            num_qty = float(qty_val)
        except (ValueError, TypeError):
            num_qty = 0.0

        if isinstance(qty_val, float) and qty_val.is_integer():
            qty_val = int(qty_val)
        elif isinstance(qty_val, (int, float)):
            pass
        else:
            try:
                qty_val = int(num_qty) if num_qty.is_integer() else num_qty
            except:
                pass

        totals_per_item[item_val] = totals_per_item.get(item_val, 0) + num_qty

        ws_out.append([item_val, client_val, qty_val])

    ws_out.append([])
    
    total_rows_start = ws_out.max_row + 1
    for item, total in totals_per_item.items():
        total_formatted = int(total) if total.is_integer() else round(total, 2)
        ws_out.append([f"مجموع {item}", "الإجمالي", total_formatted])

    font_header = Font(name='Arial', size=13, bold=True, color='FFFFFF')
    font_body = Font(name='Arial', size=12, bold=True, color='000000')
    font_total = Font(name='Arial', size=13, bold=True, color='000000')

    fill_header = PatternFill(start_color='000000', end_color='000000', fill_type='solid')
    fill_zebra = PatternFill(start_color='F2F2F2', end_color='F2F2F2', fill_type='solid')
    fill_total = PatternFill(start_color='E2E2E2', end_color='E2E2E2', fill_type='solid')

    thin_side = Side(style='thin', color='000000')
    double_side = Side(style='double', color='000000')
    
    border_all = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)
    border_total = Border(left=thin_side, right=thin_side, top=thin_side, bottom=double_side)

    center_alignment = Alignment(horizontal='center', vertical='center')

    ws_out.row_dimensions[1].height = 28

    for r in range(1, ws_out.max_row + 1):
        if r > 1:
            ws_out.row_dimensions[r].height = 25
        is_header = (r == 1)
        is_total_row = (r >= total_rows_start)
        
        if not is_header and ws_out.cell(row=r, column=1).value is None:
            ws_out.row_dimensions[r].height = 10
            continue

        for c in range(1, 4):
            cell = ws_out.cell(row=r, column=c)
            cell.alignment = center_alignment
            
            if is_header:
                cell.font = font_header
                cell.border = border_all
                cell.fill = fill_header
            elif is_total_row:
                cell.font = font_total
                cell.border = border_total
                cell.fill = fill_total
            else:
                cell.font = font_body
                cell.border = border_all
                if r % 2 == 0:
                    cell.fill = fill_zebra

    ws_out.column_dimensions['A'].width = 20
    ws_out.column_dimensions['B'].width = 12
    ws_out.column_dimensions['C'].width = 6

    ws_out.page_margins.left = ws_out.page_margins.right = 0.01
    ws_out.page_margins.top = ws_out.page_margins.bottom = 0.01

    ws_out.page_setup.orientation = ws_out.ORIENTATION_PORTRAIT
    ws_out.sheet_properties.pageSetUpPr.fitToPage = True
    ws_out.page_setup.fitToWidth = 1
    ws_out.page_setup.fitToHeight = 0

    wb_out.save(output_path)

# --- معالجة الملفات القياسية (ورقة الفرن، هرايس بانواعها، مستودع الجاهز) ---
def format_standard_two_column_sheet(input_path, output_path, report_title):
    wb_src = openpyxl.load_workbook(input_path, data_only=True)
    ws_src = wb_src.active

    data = []
    for row in ws_src.iter_rows(values_only=True):
        if any(row):
            valid_vals = [v for v in row if v is not None and str(v).strip() != '']
            if len(valid_vals) >= 2:
                item_name = valid_vals[0]
                qty = valid_vals[-1]
                data.append([item_name, qty])

    wb_out = openpyxl.Workbook()
    ws_out = wb_out.active
    ws_out.title = report_title
    ws_out.views.sheetView[0].rightToLeft = True

    ws_out.merge_cells('A1:B1')
    ws_out['A1'] = report_title
    ws_out['A2'] = "اسم المادة"
    ws_out['B2'] = "المطلوب"

    for r in data:
        item_name, qty = r[0], r[1]
        if item_name in ['اسم المادة', None] or str(item_name).strip() == '':
            continue
        try:
            val_float = float(qty)
            formatted_qty = int(val_float) if val_float.is_integer() else round(val_float, 2)
        except (ValueError, TypeError):
            formatted_qty = qty
        ws_out.append([item_name, formatted_qty])

    font_title = Font(name='Arial', size=24, bold=True)
    font_header = Font(name='Arial', size=19, bold=True, color='FFFFFF')
    font_body = Font(name='Arial', size=18, bold=True)

    fill_header = PatternFill(start_color='000000', end_color='000000', fill_type='solid')
    fill_zebra = PatternFill(start_color='F2F2F2', end_color='F2F2F2', fill_type='solid')

    thin_side = Side(style='thin', color='000000')
    border_all = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)
    center_align = Alignment(horizontal='center', vertical='center', wrap_text=True)

    ws_out.row_dimensions[1].height = 48
    ws_out['A1'].font = font_title
    ws_out['A1'].alignment = center_align

    ws_out.row_dimensions[2].height = 38
    for c in range(1, 3):
        cell = ws_out.cell(row=2, column=c)
        cell.font = font_header
        cell.fill = fill_header
        cell.alignment = center_align
        cell.border = border_all

    for r in range(3, ws_out.max_row + 1):
        ws_out.row_dimensions[r].height = 34
        for c in range(1, 3):
            cell = ws_out.cell(row=r, column=c)
            cell.font = font_body
            cell.alignment = center_align
            cell.border = border_all
            if r % 2 == 1:
                cell.fill = fill_zebra

    ws_out.column_dimensions['A'].width = 24
    ws_out.column_dimensions['B'].width = 11

    ws_out.page_margins.left = ws_out.page_margins.right = 0.01
    ws_out.page_margins.top = ws_out.page_margins.bottom = 0.01
    ws_out.page_setup.orientation = ws_out.ORIENTATION_PORTRAIT
    ws_out.sheet_properties.pageSetUpPr.fitToPage = True
    ws_out.page_setup.fitToWidth = 1
    ws_out.page_setup.fitToHeight = 0

    wb_out.save(output_path)

# --- دالة الطباعة المباشرة ---
def print_excel_file(file_path):
    try:
        if platform.system() == 'Windows':
            os.startfile(file_path, "print")
        else:
            subprocess.run(['lpr', file_path], check=True)
        return True
    except Exception as e:
        return str(e)

# --- نافذة تحديد المواد لطلبيات المبيع ---
class ItemSelectorWindow(tk.Toplevel):
    def __init__(self, parent, items, callback):
        super().__init__(parent)
        self.title("تحديد المواد المطلوبة للطباعة")
        self.geometry("520x620")
        self.resizable(False, False)
        self.callback = callback
        self.item_vars = {}

        lbl = tk.Label(self, text="اختر المواد المراد إدراجها في التقرير:", font=("Arial", 11, "bold"))
        lbl.pack(pady=8)

        frame_quick = tk.LabelFrame(self, text="اختيار سريع مخصص", font=("Arial", 10, "bold"))
        frame_quick.pack(fill="x", padx=15, pady=5)

        btn_both = RoundedButton(
            frame_quick, text="محلاية + غريبة", command=self.select_both, 
            bg_color="#6f42c1", hover_color="#522f92", width=165, height=46, radius=16, font=("Arial", 12, "bold")
        )
        btn_both.pack(side="right", padx=10, pady=10)

        btn_five = RoundedButton(
            frame_quick, text="خمس مواد", command=self.select_five_items, 
            bg_color="#17a2b8", hover_color="#117a8b", width=140, height=46, radius=16, font=("Arial", 12, "bold")
        )
        btn_five.pack(side="right", padx=10, pady=10)

        frame_btns = tk.Frame(self)
        frame_btns.pack(fill="x", padx=15, pady=5)

        btn_all = tk.Button(frame_btns, text="تحديد الكل", font=("Arial", 10), command=self.select_all)
        btn_all.pack(side="right", padx=5)

        btn_none = tk.Button(frame_btns, text="إلغاء الكل", font=("Arial", 10), command=self.deselect_all)
        btn_none.pack(side="right", padx=5)

        container = tk.Frame(self)
        container.pack(fill="both", expand=True, padx=15, pady=5)

        self.canvas = tk.Canvas(container)
        scrollbar = ttk.Scrollbar(container, orient="vertical", command=self.canvas.yview)
        scrollable_frame = tk.Frame(self.canvas)

        scrollable_frame.bind("<Configure>", lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))
        self.canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        self.canvas.configure(yscrollcommand=scrollbar.set)

        self.canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        for item in items:
            var = tk.BooleanVar(value=True)
            chk = tk.Checkbutton(scrollable_frame, text=item, variable=var, font=("Arial", 10), anchor="w")
            chk.pack(fill="x", pady=2, padx=5)
            self.item_vars[item] = var

        btn_confirm = tk.Button(self, text="استخراج الملف وجدولة الطباعة", font=("Arial", 11, "bold"), bg="#28a745", fg="white", padx=12, pady=8, command=self.confirm_selection)
        btn_confirm.pack(pady=10)

    def select_all(self):
        for var in self.item_vars.values(): var.set(True)

    def deselect_all(self):
        for var in self.item_vars.values(): var.set(False)

    def select_both(self):
        for item, var in self.item_vars.items():
            name = item.strip()
            var.set(name == 'محلاية' or name == 'غريبة بالقشطة')

    def select_five_items(self):
        target_items = {"غاز سائل كبير", "عش البلبل فستق نية", "عش لحمة نية", "كريمة", "كنافة ناعمة"}
        for item, var in self.item_vars.items():
            var.set(item.strip() in target_items)

    def confirm_selection(self):
        selected = [item for item, var in self.item_vars.items() if var.get()]
        if not selected:
            messagebox.showwarning("تنبيه", "يرجى تحديد مادة واحدة على الأقل!")
            return
        self.callback(selected)
        self.destroy()

# --- الشاشة الرئيسية للبرنامج (مع دعم السحب والإسقاط والطباعة المباشرة والمعالجة الجماعية) ---
class App:
    def __init__(self, root):
        self.root = root
        self.root.title("منسق طلبات الأمين الحراري (8سم) - الإصدار الشامل الذكي")
        self.root.geometry("700x660")
        self.root.resizable(False, False)

        self.current_files = [] # قائمة الملفات (للملف الفردي أو الدفعة)

        lbl_title = tk.Label(root, text="منسق ملفات الأمين للطابعة الحرارية (8سم)", font=("Arial", 14, "bold"))
        lbl_title.pack(pady=10)

        # إطار الخيارات والتحكم
        frame_controls = tk.LabelFrame(root, text="إعدادات ومعالجة التقارير", font=("Arial", 10, "bold"))
        frame_controls.pack(fill="x", padx=15, pady=5)

        frame_type = tk.Frame(frame_controls)
        frame_type.pack(pady=8, fill="x", padx=10)

        tk.Label(frame_type, text="نوع التقرير:", font=("Arial", 11, "bold")).pack(side="right", padx=5)
        
        self.file_type_var = tk.StringVar(value="تعرّف تلقائي")
        self.combo_type = ttk.Combobox(
            frame_type, textvariable=self.file_type_var, 
            values=["تعرّف تلقائي", "محلاية + خمس مواد", "ورقة الفرن", "هرايس بانواعها", "مستودع الجاهز"],
            state="readonly", font=("Arial", 10, "bold"), width=22
        )
        self.combo_type.pack(side="right", padx=5)

        self.chk_var = tk.BooleanVar(value=True)
        chk = tk.Checkbutton(frame_controls, text="حذف الصفوف فارغة/صفرية الكمية تلقائياً (لمحلاية + خمس مواد)", variable=self.chk_var, font=("Arial", 10))
        chk.pack(anchor="e", padx=15, pady=3)

        self.direct_print_var = tk.BooleanVar(value=False)
        chk_print = tk.Checkbutton(frame_controls, text="إرسال للطابعة الحرارية مباشرة بعد المعالجة (بدون حفظ يدوي)", variable=self.direct_print_var, font=("Arial", 10, "bold"), fg="#d9534f")
        chk_print.pack(anchor="e", padx=15, pady=3)

        # أزرار اختيار الملفات
        frame_btns = tk.Frame(root)
        frame_btns.pack(pady=8)

        btn_select = tk.Button(frame_btns, text="اختر ملف أو عدة ملفات Excel", font=("Arial", 11, "bold"), bg="#007bff", fg="white", padx=15, pady=6, command=self.load_files_dialog)
        btn_select.pack(side="left", padx=5)

        # منطقة السحب والإسقاط (أو المعاينة)
        frame_preview = tk.LabelFrame(root, text="معاينة الملفات المسحوبة أو المختارة (اسحب الملفات وأفلتها هنا)", font=("Arial", 10, "bold"))
        frame_preview.pack(fill="both", expand=True, padx=15, pady=5)

        scroll_x = ttk.Scrollbar(frame_preview, orient="horizontal")
        scroll_y = ttk.Scrollbar(frame_preview, orient="vertical")

        self.tree = ttk.Treeview(
            frame_preview, show="headings", height=7, 
            xscrollcommand=scroll_x.set, yscrollcommand=scroll_y.set
        )
        
        scroll_x.config(command=self.tree.xview)
        scroll_y.config(command=self.tree.yview)

        scroll_x.pack(side="bottom", fill="x")
        scroll_y.pack(side="left", fill="y")
        self.tree.pack(fill="both", expand=True, padx=5, pady=5)

        # تفعيل خاصية السحب والإسقاط (Drag and Drop)
        if HAS_DND:
            try:
                root.drop_target_register(DND_FILES)
                root.dnd_bind('<<Drop>>', self.handle_drop)
                frame_preview.config(text="معاينة الملفات (اسحب ملفات الأكسل وأفلتها هنا مباشرة 📂)")
            except Exception:
                pass

        self.btn_process = tk.Button(root, text="معالجة واستخراج (أو طباعة) الملفات دفعة واحدة", font=("Arial", 12, "bold"), bg="#28a745", fg="white", padx=20, pady=10, state="disabled", command=self.process_files)
        self.btn_process.pack(pady=10)

    def load_files_dialog(self):
        file_paths = filedialog.askopenfilenames(filetypes=[("Excel Files", "*.xlsx *.xls")])
        if not file_paths:
            return
        self.handle_loaded_files(list(file_paths))

    def handle_drop(self, event):
        raw_data = event.data
        if not raw_data:
            return
        import re
        if platform.system() == 'Windows':
            files = re.findall(r'\{([^}]+)\}|(\S+)', raw_data)
            file_paths = [f[0] or f[1] for f in files if f[0] or f[1]]
        else:
            file_paths = raw_data.split()

        valid_files = [f for f in file_paths if f.lower().endswith(('.xlsx', '.xls'))]
        if valid_files:
            self.handle_loaded_files(valid_files)
        else:
            messagebox.showwarning("تنبيه", "يرجى سحب وإسقاط ملفات إكسل صالحة (.xlsx أو .xls)!")

    def handle_loaded_files(self, file_paths):
        self.current_files = file_paths
        self.preview_files(file_paths)
        if len(file_paths) == 1:
            self.file_type_var.set("تعرّف تلقائي")
            self.auto_detect_type(file_paths[0])
        else:
            self.file_type_var.set("تعرّف تلقائي")
        self.btn_process.config(state="normal")

    def auto_detect_type(self, file_path):
        try:
            wb = openpyxl.load_workbook(file_path, data_only=True)
            if 'تأثير الطلبيات على المخزون' in wb.sheetnames:
                if 'رأسية التقرير' in wb.sheetnames:
                    ws_head = wb['رأسية التقرير']
                    val = str(ws_head.cell(row=1, column=2).value or '')
                    if 'هرايس' in val:
                        return "هرايس بانواعها"
                    elif 'مستودع' in val:
                        return "مستودع الجاهز"
                return "ورقة الفرن"
            else:
                return "محلاية + خمس مواد"
        except Exception:
            return "ورقة الفرن"

    def preview_files(self, file_paths):
        for item in self.tree.get_children():
            self.tree.delete(item)

        try:
            first_file = file_paths[0]
            wb = openpyxl.load_workbook(first_file, data_only=True)
            ws = wb.active
            rows = list(ws.iter_rows(values_only=True))
            if not rows:
                return

            header_row = rows[0]
            cols = [f"col_{i}" for i in range(len(header_row))]
            self.tree["columns"] = cols

            for i, col_name in enumerate(header_row):
                header_text = str(col_name) if col_name is not None else f"عمود {i+1}"
                self.tree.heading(f"col_{i}", text=header_text)
                self.tree.column(f"col_{i}", width=120, anchor="center")

            for row in rows[1:8]:
                self.tree.insert("", "end", values=[str(v) if v is not None else "" for v in row])
        except Exception as e:
            messagebox.showerror("خطأ", f"تعذر معاينة الملف:\n{str(e)}")

    def process_files(self):
        if not self.current_files:
            return

        success_count = 0
        direct_print = self.direct_print_var.get()

        for file_path in self.current_files:
            detected_type = self.file_type_var.get()
            if detected_type == "تعرّف تلقائي":
                detected_type = self.auto_detect_type(file_path)

            if detected_type == "محلاية + خمس مواد":
                try:
                    items, _ = get_unique_items(file_path)
                    if not items:
                        continue
                    
                    def process_with_selection(selected_items):
                        base_name = os.path.basename(file_path)
                        save_path = os.path.join(os.path.dirname(file_path), "جاهز_للطباعة_" + base_name)
                        format_sales_orders(file_path, save_path, selected_items, remove_empty=self.chk_var.get())
                        if direct_print:
                            print_excel_file(save_path)
                        nonlocal success_count
                        success_count += 1

                    if len(self.current_files) == 1:
                        ItemSelectorWindow(self.root, items, process_with_selection)
                        return
                    else:
                        base_name = os.path.basename(file_path)
                        save_path = os.path.join(os.path.dirname(file_path), "جاهز_للطباعة_" + base_name)
                        format_sales_orders(file_path, save_path, items, remove_empty=self.chk_var.get())
                        if direct_print:
                            print_excel_file(save_path)
                        success_count += 1
                except Exception:
                    pass

            elif detected_type in ["ورقة الفرن", "هرايس بانواعها", "مستودع الجاهز"]:
                try:
                    base_name = os.path.basename(file_path)
                    save_path = os.path.join(os.path.dirname(file_path), "جاهز_للطباعة_" + base_name)
                    format_standard_two_column_sheet(file_path, save_path, detected_type)
                    if direct_print:
                        print_excel_file(save_path)
                    success_count += 1
                except Exception:
                    pass

        if success_count > 0:
            msg = f"تمت معالجة وإخراج {success_count} ملف بنجاح!"
            if direct_print:
                msg += "\nوتم إرسالها للطباعة المباشرة على الطابعة الحرارية."
            messagebox.showinfo("نجاح تام", msg)

if __name__ == "__main__":
    if HAS_DND:
        root = TkinterDnD.Tk()
    else:
        root = tk.Tk()
    app = App(root)
    root.mainloop()
