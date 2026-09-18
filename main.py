import os
import sys
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
import subprocess
import platform

# --- معالجة الملفات ذات الثلاثة أعمدة (محلاية + غريبة أو الخمس مواد) ---
def format_sales_orders_custom(input_path, output_path, selected_items, report_title, remove_empty=True):
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
    ws_out.title = report_title
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

    font_header = Font(name='Arial', size=12, bold=True, color='FFFFFF')
    font_body = Font(name='Arial', size=11, bold=True, color='000000')
    font_total = Font(name='Arial', size=12, bold=True, color='000000')

    fill_header = PatternFill(start_color='000000', end_color='000000', fill_type='solid')
    fill_zebra = PatternFill(start_color='F2F2F2', end_color='F2F2F2', fill_type='solid')
    fill_total = PatternFill(start_color='E2E2E2', end_color='E2E2E2', fill_type='solid')

    thin_side = Side(style='thin', color='000000')
    double_side = Side(style='double', color='000000')
    
    border_all = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)
    border_total = Border(left=thin_side, right=thin_side, top=thin_side, bottom=double_side)

    center_alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)

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

    # ضبط عروض الأعمدة الثلاثة لتملأ الصفحة الحرارية تماماً دون فراغات
    ws_out.column_dimensions['A'].width = 18
    ws_out.column_dimensions['B'].width = 18
    ws_out.column_dimensions['C'].width = 8

    ws_out.page_margins.left = ws_out.page_margins.right = 0.0
    ws_out.page_margins.top = ws_out.page_margins.bottom = 0.01

    ws_out.page_setup.orientation = ws_out.ORIENTATION_PORTRAIT
    ws_out.sheet_properties.pageSetUpPr.fitToPage = True
    ws_out.page_setup.fitToWidth = 1
    ws_out.page_setup.fitToHeight = 1

    wb_out.save(output_path)

# --- معالجة الملفات القياسية ذات العمودين (ورقة الفرن، هرايس بانواعها، مستودع الجاهز) ---
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

    font_title = Font(name='Arial', size=16, bold=True)
    font_header = Font(name='Arial', size=14, bold=True, color='FFFFFF')
    font_body = Font(name='Arial', size=13, bold=True)

    fill_header = PatternFill(start_color='000000', end_color='000000', fill_type='solid')
    fill_zebra = PatternFill(start_color='F2F2F2', end_color='F2F2F2', fill_type='solid')

    thin_side = Side(style='thin', color='000000')
    border_all = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)
    center_align = Alignment(horizontal='center', vertical='center', wrap_text=True)

    ws_out.row_dimensions[1].height = 32
    ws_out['A1'].font = font_title
    ws_out['A1'].alignment = center_align

    ws_out.row_dimensions[2].height = 28
    for c in range(1, 3):
        cell = ws_out.cell(row=2, column=c)
        cell.font = font_header
        cell.fill = fill_header
        cell.alignment = center_align
        cell.border = border_all

    for r in range(3, ws_out.max_row + 1):
        ws_out.row_dimensions[r].height = 26
        for c in range(1, 3):
            cell = ws_out.cell(row=r, column=c)
            cell.font = font_body
            cell.alignment = center_align
            cell.border = border_all
            if r % 2 == 1:
                cell.fill = fill_zebra

    ws_out.column_dimensions['A'].width = 28
    ws_out.column_dimensions['B'].width = 15

    ws_out.page_margins.left = ws_out.page_margins.right = 0.0
    ws_out.page_margins.top = ws_out.page_margins.bottom = 0.01
    ws_out.page_setup.orientation = ws_out.ORIENTATION_PORTRAIT
    ws_out.sheet_properties.pageSetUpPr.fitToPage = True
    ws_out.page_setup.fitToWidth = 1
    ws_out.page_setup.fitToHeight = 1

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

# --- الشاشة الرئيسية للبرنامج ---
class App:
    def __init__(self, root):
        self.root = root
        self.root.title("منسق طلبات الأمين الحراري (8سم) - الإصدار الشامل الذكي")
        self.root.geometry("700x660")
        self.root.resizable(False, False)

        self.current_files = []

        lbl_title = tk.Label(root, text="منسق ملفات الأمين للطابعة الحرارية (8سم)", font=("Arial", 14, "bold"))
        lbl_title.pack(pady=10)

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
        chk = tk.Checkbutton(frame_controls, text="حذف الصفوف فارغة/صفرية الكمية تلقائياً", variable=self.chk_var, font=("Arial", 10))
        chk.pack(anchor="e", padx=15, pady=3)

        self.direct_print_var = tk.BooleanVar(value=False)
        chk_print = tk.Checkbutton(frame_controls, text="إرسال الملفات للطابعة الحرارية مباشرة بعد إنتاجها", variable=self.direct_print_var, font=("Arial", 10, "bold"), fg="#d9534f")
        chk_print.pack(anchor="e", padx=15, pady=3)

        frame_btns = tk.Frame(root)
        frame_btns.pack(pady=8)

        btn_select = tk.Button(frame_btns, text="📂 اختر ملف أو عدة ملفات Excel معاً", font=("Arial", 11, "bold"), bg="#007bff", fg="white", padx=15, pady=6, command=self.load_files_dialog)
        btn_select.pack(side="left", padx=5)

        frame_preview = tk.LabelFrame(root, text="معاينة الملفات المختارة", font=("Arial", 10, "bold"))
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

        self.btn_process = tk.Button(root, text="⚡ معالجة واستخراج جميع الملفات إلى مجلد سطح المكتب", font=("Arial", 12, "bold"), bg="#28a745", fg="white", padx=20, pady=10, state="disabled", command=self.process_files)
        self.btn_process.pack(pady=10)

    def load_files_dialog(self):
        file_paths = filedialog.askopenfilenames(filetypes=[("Excel Files", "*.xlsx *.xls")])
        if not file_paths:
            return
        self.handle_loaded_files(list(file_paths))

    def handle_loaded_files(self, file_paths):
        self.current_files = file_paths
        self.preview_files(file_paths)
        
        if len(file_paths) == 1:
            detected = self.auto_detect_type(file_paths[0])
            self.file_type_var.set(detected)
        else:
            self.file_type_var.set("تعرّف تلقائي")
            
        self.btn_process.config(state="normal")

    def auto_detect_type(self, file_path):
        try:
            filename = os.path.basename(file_path).lower()
            
            if 'هرايس' in filename or 'harees' in filename:
                return "هرايس بانواعها"
            elif 'مستودع' in filename or 'جاهز' in filename or 'warehouse' in filename:
                return "مستودع الجاهز"
            elif 'فرن' in filename or 'oven' in filename:
                return "ورقة الفرن"
            elif 'مبيع' in filename or 'طلبات' in filename or 'محلاية' in filename:
                return "محلاية + خمس مواد"

            wb = openpyxl.load_workbook(file_path, data_only=True)
            if 'تأثير الطلبيات على المخزون' in wb.sheetnames:
                if 'رأسية التقرير' in wb.sheetnames:
                    ws_head = wb['رأسية التقرير']
                    for r in range(1, 4):
                        val = str(ws_head.cell(row=r, column=2).value or '').lower()
                        if 'هرايس' in val:
                            return "هرايس بانواعها"
                        elif 'مستودع' in val:
                            return "مستودع الجاهز"
                return "ورقة الفرن"
            else:
                return "محلاية + خمس مواد"
        except Exception:
            return "محلاية + خمس مواد"

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

        desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")
        output_dir = os.path.join(desktop_path, "جاهز للطباعه")
        os.makedirs(output_dir, exist_ok=True)

        direct_print = self.direct_print_var.get()
        processed_count = 0
        manual_type = self.file_type_var.get()

        for file_path in self.current_files:
            if manual_type != "تعرّف تلقائي":
                detected_type = manual_type
            else:
                detected_type = self.auto_detect_type(file_path)

            if detected_type in ["ورقة الفرن", "هرايس بانواعها", "مستودع الجاهز"]:
                try:
                    if detected_type == "هرايس بانواعها":
                        out_name = "جاهز_للطباعة_هرايس.xlsx"
                    elif detected_type == "مستودع الجاهز":
                        out_name = "جاهز_للطباعة_مستودع_الجاهز.xlsx"
                    else:
                        out_name = "جاهز_للطباعة_ورقة_الفرن.xlsx"

                    save_path = os.path.join(output_dir, out_name)
                    format_standard_two_column_sheet(file_path, save_path, detected_type)
                    if direct_print:
                        print_excel_file(save_path)
                    processed_count += 1
                except Exception:
                    pass
            elif detected_type == "محلاية + خمس مواد":
                try:
                    path_both = os.path.join(output_dir, "جاهز_للطباعة_محلاية_وغريبة.xlsx")
                    format_sales_orders_custom(
                        file_path, path_both, 
                        selected_items=['محلاية', 'غريبة بالقشطة'], 
                        report_title="محلاية + غريبة", 
                        remove_empty=self.chk_var.get()
                    )
                    if direct_print:
                        print_excel_file(path_both)

                    five_items_set = {"غاز سائل كبير", "عش البلبل فستق نية", "عش لحمة نية", "كريمة", "كنافة ناعمة"}
                    path_five = os.path.join(output_dir, "جاهز_للطباعة_خمس_مواد.xlsx")
                    format_sales_orders_custom(
                        file_path, path_five, 
                        selected_items=five_items_set, 
                        report_title="خمس مواد", 
                        remove_empty=self.chk_var.get()
                    )
                    if direct_print:
                        print_excel_file(path_five)

                    processed_count += 2
                except Exception:
                    pass

        msg = f"تمت معالجة وإخراج جميع الملفات بنجاح تام ({processed_count} ملف)!\nموجودة الآن على سطح المكتب داخل مجلد (جاهز للطباعه)."
        if direct_print:
            msg += "\nوتم إرسالها للطباعة المباشرة على الطابعة الحرارية."
        messagebox.showinfo("نجاح تام", msg)

if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop()
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
    ws_out.page_setup.fitToHeight = 1  # فرض الطباعة على صفحة واحدة طولياً

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

    # تم تصغير أحجام الخطوط قليلاً لتناسب ورقة واحدة بالطباعة الحرارية دون امتداد مزعج
    font_title = Font(name='Arial', size=16, bold=True)
    font_header = Font(name='Arial', size=13, bold=True, color='FFFFFF')
    font_body = Font(name='Arial', size=12, bold=True)

    fill_header = PatternFill(start_color='000000', end_color='000000', fill_type='solid')
    fill_zebra = PatternFill(start_color='F2F2F2', end_color='F2F2F2', fill_type='solid')

    thin_side = Side(style='thin', color='000000')
    border_all = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)
    center_align = Alignment(horizontal='center', vertical='center', wrap_text=True)

    ws_out.row_dimensions[1].height = 32
    ws_out['A1'].font = font_title
    ws_out['A1'].alignment = center_align

    ws_out.row_dimensions[2].height = 28
    for c in range(1, 3):
        cell = ws_out.cell(row=2, column=c)
        cell.font = font_header
        cell.fill = fill_header
        cell.alignment = center_align
        cell.border = border_all

    for r in range(3, ws_out.max_row + 1):
        ws_out.row_dimensions[r].height = 26
        for c in range(1, 3):
            cell = ws_out.cell(row=r, column=c)
            cell.font = font_body
            cell.alignment = center_align
            cell.border = border_all
            if r % 2 == 1:
                cell.fill = fill_zebra

    ws_out.column_dimensions['A'].width = 22
    ws_out.column_dimensions['B'].width = 10

    ws_out.page_margins.left = ws_out.page_margins.right = 0.01
    ws_out.page_margins.top = ws_out.page_margins.bottom = 0.01
    ws_out.page_setup.orientation = ws_out.ORIENTATION_PORTRAIT
    ws_out.sheet_properties.pageSetUpPr.fitToPage = True
    ws_out.page_setup.fitToWidth = 1
    ws_out.page_setup.fitToHeight = 1  # فرض الطباعة على صفحة واحدة طولياً تماماً

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

# --- الشاشة الرئيسية للبرنامج ---
class App:
    def __init__(self, root):
        self.root = root
        self.root.title("منسق طلبات الأمين الحراري (8سم) - الإصدار الشامل الذكي")
        self.root.geometry("700x660")
        self.root.resizable(False, False)

        self.current_files = []

        lbl_title = tk.Label(root, text="منسق ملفات الأمين للطابعة الحرارية (8سم)", font=("Arial", 14, "bold"))
        lbl_title.pack(pady=10)

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
        chk = tk.Checkbutton(frame_controls, text="حذف الصفوف فارغة/صفرية الكمية تلقائياً", variable=self.chk_var, font=("Arial", 10))
        chk.pack(anchor="e", padx=15, pady=3)

        self.direct_print_var = tk.BooleanVar(value=False)
        chk_print = tk.Checkbutton(frame_controls, text="إرسال الملفات للطابعة الحرارية مباشرة بعد إنتاجها", variable=self.direct_print_var, font=("Arial", 10, "bold"), fg="#d9534f")
        chk_print.pack(anchor="e", padx=15, pady=3)

        frame_btns = tk.Frame(root)
        frame_btns.pack(pady=8)

        btn_select = tk.Button(frame_btns, text="📂 اختر ملف أو عدة ملفات Excel معاً", font=("Arial", 11, "bold"), bg="#007bff", fg="white", padx=15, pady=6, command=self.load_files_dialog)
        btn_select.pack(side="left", padx=5)

        frame_preview = tk.LabelFrame(root, text="معاينة الملفات المختارة", font=("Arial", 10, "bold"))
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

        self.btn_process = tk.Button(root, text="⚡ معالجة واستخراج جميع الملفات إلى مجلد سطح المكتب", font=("Arial", 12, "bold"), bg="#28a745", fg="white", padx=20, pady=10, state="disabled", command=self.process_files)
        self.btn_process.pack(pady=10)

    def load_files_dialog(self):
        file_paths = filedialog.askopenfilenames(filetypes=[("Excel Files", "*.xlsx *.xls")])
        if not file_paths:
            return
        self.handle_loaded_files(list(file_paths))

    def handle_loaded_files(self, file_paths):
        self.current_files = file_paths
        self.preview_files(file_paths)
        
        if len(file_paths) == 1:
            detected = self.auto_detect_type(file_paths[0])
            self.file_type_var.set(detected)
        else:
            self.file_type_var.set("تعرّف تلقائي")
            
        self.btn_process.config(state="normal")

    def auto_detect_type(self, file_path):
        try:
            filename = os.path.basename(file_path).lower()
            
            if 'هرايس' in filename or 'harees' in filename:
                return "هرايس بانواعها"
            elif 'مستودع' in filename or 'جاهز' in filename or 'warehouse' in filename:
                return "مستودع الجاهز"
            elif 'فرن' in filename or 'oven' in filename:
                return "ورقة الفرن"
            elif 'مبيع' in filename or 'طلبات' in filename or 'محلاية' in filename:
                return "محلاية + خمس مواد"

            wb = openpyxl.load_workbook(file_path, data_only=True)
            if 'تأثير الطلبيات على المخزون' in wb.sheetnames:
                if 'رأسية التقرير' in wb.sheetnames:
                    ws_head = wb['رأسية التقرير']
                    for r in range(1, 4):
                        val = str(ws_head.cell(row=r, column=2).value or '').lower()
                        if 'هرايس' in val:
                            return "هرايس بانواعها"
                        elif 'مستودع' in val:
                            return "مستودع الجاهز"
                return "ورقة الفرن"
            else:
                return "محلاية + خمس مواد"
        except Exception:
            return "محلاية + خمس مواد"

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

        desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")
        output_dir = os.path.join(desktop_path, "جاهز للطباعه")
        os.makedirs(output_dir, exist_ok=True)

        direct_print = self.direct_print_var.get()
        processed_count = 0
        manual_type = self.file_type_var.get()

        for file_path in self.current_files:
            if manual_type != "تعرّف تلقائي":
                detected_type = manual_type
            else:
                detected_type = self.auto_detect_type(file_path)

            if detected_type in ["ورقة الفرن", "هرايس بانواعها", "مستودع الجاهز"]:
                try:
                    if detected_type == "هرايس بانواعها":
                        out_name = "جاهز_للطباعة_هرايس.xlsx"
                    elif detected_type == "مستودع الجاهز":
                        out_name = "جاهز_للطباعة_مستودع_الجاهز.xlsx"
                    else:
                        out_name = "جاهز_للطباعة_ورقة_الفرن.xlsx"

                    save_path = os.path.join(output_dir, out_name)
                    format_standard_two_column_sheet(file_path, save_path, detected_type)
                    if direct_print:
                        print_excel_file(save_path)
                    processed_count += 1
                except Exception:
                    pass
            elif detected_type == "محلاية + خمس مواد":
                try:
                    path_both = os.path.join(output_dir, "جاهز_للطباعة_محلاية_وغريبة.xlsx")
                    format_sales_orders_custom(
                        file_path, path_both, 
                        selected_items=['محلاية', 'غريبة بالقشطة'], 
                        report_title="محلاية + غريبة", 
                        remove_empty=self.chk_var.get()
                    )
                    if direct_print:
                        print_excel_file(path_both)

                    five_items_set = {"غاز سائل كبير", "عش البلبل فستق نية", "عش لحمة نية", "كريمة", "كنافة ناعمة"}
                    path_five = os.path.join(output_dir, "جاهز_للطباعة_خمس_مواد.xlsx")
                    format_sales_orders_custom(
                        file_path, path_five, 
                        selected_items=five_items_set, 
                        report_title="خمس مواد", 
                        remove_empty=self.chk_var.get()
                    )
                    if direct_print:
                        print_excel_file(path_five)

                    processed_count += 2
                except Exception:
                    pass

        msg = f"تمت معالجة وإخراج جميع الملفات بنجاح تام ({processed_count} ملف)!\nموجودة الآن على سطح المكتب داخل مجلد (جاهز للطباعه)."
        if direct_print:
            msg += "\nوتم إرسالها للطباعة المباشرة على الطابعة الحرارية."
        messagebox.showinfo("نجاح تام", msg)

if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop()
