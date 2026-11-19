import os
import sys
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
import subprocess
import platform


def format_sales_orders_custom(input_path, output_path, selected_items, report_title, remove_empty=True):
    wb_src = openpyxl.load_workbook(input_path, data_only=True)
    ws_src = wb_src.active
    header_row = 1
    for r in range(1, min(10, ws_src.max_row + 1)):
        row_vals = [str(ws_src.cell(row=r, column=c).value or '').strip() for c in range(1, ws_src.max_column + 1)]
        if any((k in row_vals for k in ['المادة', 'تجهيز', 'اسم العميل'])):
            header_row = r
            break
    col_item, col_client, col_qty = (3, 5, 6)
    for c in range(1, ws_src.max_column + 1):
        val = str(ws_src.cell(row=header_row, column=c).value or '').strip()
        if val in ['المادة', 'اسم المادة']:
            col_item = c
        elif val in ['اسم العميل', 'العميل', 'المحل']:
            col_client = c
        elif val in ['تجهيز', 'الكمية', 'كمية']:
            col_qty = c

    wb_out = openpyxl.Workbook()
    ws_out = wb_out.active
    ws_out.title = report_title[:31]
    ws_out.views.sheetView[0].rightToLeft = True

    # الصف 1: عنوان الصفحة مدمج بخلفية سوداء وخط أبيض
    ws_out.merge_cells('A1:C1')
    ws_out['A1'] = report_title

    # الصف 2: فاصل واضح بين العنوان الرئيسي وعناوين الأعمدة
    ws_out.append([])
    title_sep_row = ws_out.max_row

    # الصف 3: عناوين الأعمدة
    ws_out.append(['المادة', 'اسم العميل', 'الكمية'])
    col_header_row = ws_out.max_row

    totals_per_item = {}
    data_rows = []

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
        elif not isinstance(qty_val, (int, float)):
            try:
                qty_val = int(num_qty) if num_qty.is_integer() else num_qty
            except Exception:
                pass
        totals_per_item[item_val] = totals_per_item.get(item_val, 0) + num_qty
        ws_out.append([item_val, client_val, qty_val])
        data_rows.append(ws_out.max_row)

    ws_out.append([])
    sep_row = ws_out.max_row

    total_rows = []
    for item, total in totals_per_item.items():
        total_formatted = int(total) if float(total).is_integer() else round(total, 2)
        ws_out.append([f'مجموع {item}', 'الإجمالي', total_formatted])
        total_rows.append(ws_out.max_row)

    font_title = Font(name='Arial', size=14, bold=True, color='FFFFFF')
    font_header = Font(name='Arial', size=12, bold=True, color='FFFFFF')
    font_body = Font(name='Arial', size=11, bold=True, color='000000')
    font_total = Font(name='Arial', size=12, bold=True, color='000000')

    fill_black = PatternFill(start_color='000000', end_color='000000', fill_type='solid')
    fill_zebra = PatternFill(start_color='F2F2F2', end_color='F2F2F2', fill_type='solid')
    fill_total = PatternFill(start_color='E2E2E2', end_color='E2E2E2', fill_type='solid')

    thin_side = Side(style='thin', color='000000')
    double_side = Side(style='double', color='000000')
    border_all = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)
    border_total = Border(left=thin_side, right=thin_side, top=thin_side, bottom=double_side)
    center_alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)

    ws_out.row_dimensions[1].height = 34
    for c in range(1, 4):
        cell = ws_out.cell(row=1, column=c)
        cell.font = font_title
        cell.fill = fill_black
        cell.alignment = center_alignment
        cell.border = border_all

    ws_out.row_dimensions[title_sep_row].height = 10

    ws_out.row_dimensions[col_header_row].height = 28
    for c in range(1, 4):
        cell = ws_out.cell(row=col_header_row, column=c)
        cell.font = font_header
        cell.fill = fill_black
        cell.alignment = center_alignment
        cell.border = border_all

    ws_out.row_dimensions[sep_row].height = 10

    for idx, r in enumerate(data_rows):
        ws_out.row_dimensions[r].height = 25
        for c in range(1, 4):
            cell = ws_out.cell(row=r, column=c)
            cell.alignment = center_alignment
            cell.font = font_body
            cell.border = border_all
            if idx % 2 == 1:
                cell.fill = fill_zebra

    for r in total_rows:
        ws_out.row_dimensions[r].height = 25
        for c in range(1, 4):
            cell = ws_out.cell(row=r, column=c)
            cell.alignment = center_alignment
            cell.font = font_total
            cell.fill = fill_total
            cell.border = border_total

    ws_out.column_dimensions['A'].width = 18
    ws_out.column_dimensions['B'].width = 18
    ws_out.column_dimensions['C'].width = 8
    ws_out.page_margins.left = ws_out.page_margins.right = 0.0
    ws_out.page_margins.top = ws_out.page_margins.bottom = 0.01
    ws_out.page_setup.orientation = ws_out.ORIENTATION_PORTRAIT
    ws_out.sheet_properties.pageSetUpPr.fitToPage = True
    ws_out.page_setup.fitToWidth = 1
    ws_out.page_setup.fitToHeight = 0
    wb_out.save(output_path)


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
    ws_out.title = report_title[:31]
    ws_out.views.sheetView[0].rightToLeft = True

    # الصف 1: عنوان الصفحة مدمج بخلفية سوداء وخط أبيض
    ws_out.merge_cells('A1:B1')
    ws_out['A1'] = report_title

    # الصف 2: فاصل واضح بين العنوان الرئيسي وعناوين الأعمدة
    ws_out.append([])
    title_sep_row = ws_out.max_row

    # الصف 3: عناوين الأعمدة
    ws_out.append(['اسم المادة', 'المطلوب'])
    col_header_row = ws_out.max_row

    data_rows = []
    for r in data:
        item_name, qty = (r[0], r[1])
        if item_name in ('اسم المادة', None) or str(item_name).strip() == '':
            continue
        else:
            try:
                val_float = float(qty)
                formatted_qty = int(val_float) if val_float.is_integer() else round(val_float, 2)
            except (ValueError, TypeError):
                formatted_qty = qty
            ws_out.append([item_name, formatted_qty])
            data_rows.append(ws_out.max_row)

    font_title = Font(name='Arial', size=14, bold=True, color='FFFFFF')
    font_header = Font(name='Arial', size=13, bold=True, color='FFFFFF')
    font_body = Font(name='Arial', size=12, bold=True, color='000000')
    fill_black = PatternFill(start_color='000000', end_color='000000', fill_type='solid')
    fill_zebra = PatternFill(start_color='F2F2F2', end_color='F2F2F2', fill_type='solid')
    thin_side = Side(style='thin', color='000000')
    border_all = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)
    center_align = Alignment(horizontal='center', vertical='center', wrap_text=True)

    ws_out.row_dimensions[1].height = 34
    for c in range(1, 3):
        cell = ws_out.cell(row=1, column=c)
        cell.font = font_title
        cell.fill = fill_black
        cell.alignment = center_align
        cell.border = border_all

    ws_out.row_dimensions[title_sep_row].height = 10

    ws_out.row_dimensions[col_header_row].height = 28
    for c in range(1, 3):
        cell = ws_out.cell(row=col_header_row, column=c)
        cell.font = font_header
        cell.fill = fill_black
        cell.alignment = center_align
        cell.border = border_all

    for idx, r in enumerate(data_rows):
        ws_out.row_dimensions[r].height = 26
        for c in range(1, 3):
            cell = ws_out.cell(row=r, column=c)
            cell.font = font_body
            cell.alignment = center_align
            cell.border = border_all
            if idx % 2 == 1:
                cell.fill = fill_zebra

    ws_out.column_dimensions['A'].width = 28
    ws_out.column_dimensions['B'].width = 15
    ws_out.page_margins.left = ws_out.page_margins.right = 0.0
    ws_out.page_margins.top = ws_out.page_margins.bottom = 0.01
    ws_out.page_setup.orientation = ws_out.ORIENTATION_PORTRAIT
    ws_out.sheet_properties.pageSetUpPr.fitToPage = True
    ws_out.page_setup.fitToWidth = 1
    ws_out.page_setup.fitToHeight = 0
    wb_out.save(output_path)


def format_paired_items_by_client(input_path, output_path, items_pair, report_title, remove_empty=True):
    """ورقة طباعة حرارية (8سم) لصنفين مفروزة حسب المحل مع إجمالي كل صنف في الأسفل."""
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
        if val in ['المادة', 'اسم المادة']:
            col_item = c
        elif val in ['اسم العميل', 'العميل', 'المحل']:
            col_client = c
        elif val in ['تجهيز', 'الكمية', 'كمية']:
            col_qty = c

    item1, item2 = items_pair[0], items_pair[1]
    target_set = {item1, item2}

    client_data = {}
    client_order = []

    for r in range(header_row + 1, ws_src.max_row + 1):
        item_val = str(ws_src.cell(row=r, column=col_item).value or '').strip()
        if item_val not in target_set:
            continue
        client_val = str(ws_src.cell(row=r, column=col_client).value or '').strip()
        qty_val = ws_src.cell(row=r, column=col_qty).value

        if remove_empty and (qty_val is None or str(qty_val).strip() in ['', '0', 'None', '0.0']):
            continue

        try:
            num_qty = float(qty_val)
        except (ValueError, TypeError):
            num_qty = 0.0

        if remove_empty and num_qty == 0:
            continue

        if client_val not in client_data:
            client_data[client_val] = {item1: 0.0, item2: 0.0}
            client_order.append(client_val)

        client_data[client_val][item_val] += num_qty

    wb_out = openpyxl.Workbook()
    ws_out = wb_out.active
    ws_out.title = report_title[:31]
    ws_out.views.sheetView[0].rightToLeft = True

    # الصف 1: عنوان الصفحة مدمج بخلفية سوداء وخط أبيض
    ws_out.merge_cells('A1:C1')
    ws_out['A1'] = report_title

    # الصف 2: فاصل واضح بين العنوان الرئيسي وعناوين الأعمدة
    ws_out.append([])
    title_sep_row = ws_out.max_row

    # الصف 3: عناوين الأعمدة
    ws_out.append(['المادة', 'اسم المحل', 'الكمية'])
    col_header_row = ws_out.max_row

    total_item1 = 0.0
    total_item2 = 0.0
    data_rows = []

    for client_name in client_order:
        for itm in [item1, item2]:
            q = client_data[client_name].get(itm, 0.0)
            if remove_empty and q == 0:
                continue
            if itm == item1:
                total_item1 += q
            else:
                total_item2 += q
            fmt_q = int(q) if float(q).is_integer() else round(q, 4)
            ws_out.append([itm, client_name, fmt_q])
            data_rows.append(ws_out.max_row)

    ws_out.append([])
    sep_row = ws_out.max_row

    fmt_tot1 = int(total_item1) if float(total_item1).is_integer() else round(total_item1, 4)
    fmt_tot2 = int(total_item2) if float(total_item2).is_integer() else round(total_item2, 4)

    ws_out.append([f'مجموع {item1}', 'الإجمالي', fmt_tot1])
    tot1_row = ws_out.max_row

    ws_out.append([f'مجموع {item2}', 'الإجمالي', fmt_tot2])
    tot2_row = ws_out.max_row

    font_title = Font(name='Arial', size=14, bold=True, color='FFFFFF')
    font_header = Font(name='Arial', size=12, bold=True, color='FFFFFF')
    font_body = Font(name='Arial', size=11, bold=True, color='000000')
    font_total = Font(name='Arial', size=12, bold=True, color='000000')

    fill_black = PatternFill(start_color='000000', end_color='000000', fill_type='solid')
    fill_zebra = PatternFill(start_color='F2F2F2', end_color='F2F2F2', fill_type='solid')
    fill_total = PatternFill(start_color='E2E2E2', end_color='E2E2E2', fill_type='solid')

    thin_side = Side(style='thin', color='000000')
    double_side = Side(style='double', color='000000')
    border_all = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)
    border_total = Border(left=thin_side, right=thin_side, top=thin_side, bottom=double_side)

    center_alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)

    ws_out.row_dimensions[1].height = 34
    for c in range(1, 4):
        cell = ws_out.cell(row=1, column=c)
        cell.font = font_title
        cell.fill = fill_black
        cell.alignment = center_alignment
        cell.border = border_all

    ws_out.row_dimensions[title_sep_row].height = 10

    ws_out.row_dimensions[col_header_row].height = 28
    for c in range(1, 4):
        cell = ws_out.cell(row=col_header_row, column=c)
        cell.font = font_header
        cell.fill = fill_black
        cell.alignment = center_alignment
        cell.border = border_all

    ws_out.row_dimensions[sep_row].height = 10

    for idx, r in enumerate(data_rows):
        ws_out.row_dimensions[r].height = 25
        for c in range(1, 4):
            cell = ws_out.cell(row=r, column=c)
            cell.alignment = center_alignment
            cell.font = font_body
            cell.border = border_all
            if idx % 2 == 1:
                cell.fill = fill_zebra

    for r in (tot1_row, tot2_row):
        ws_out.row_dimensions[r].height = 25
        for c in range(1, 4):
            cell = ws_out.cell(row=r, column=c)
            cell.alignment = center_alignment
            cell.font = font_total
            cell.fill = fill_total
            cell.border = border_total

    ws_out.column_dimensions['A'].width = 18
    ws_out.column_dimensions['B'].width = 18
    ws_out.column_dimensions['C'].width = 8

    ws_out.page_margins.left = ws_out.page_margins.right = 0.0
    ws_out.page_margins.top = ws_out.page_margins.bottom = 0.01
    ws_out.page_setup.orientation = ws_out.ORIENTATION_PORTRAIT
    ws_out.sheet_properties.pageSetUpPr.fitToPage = True
    ws_out.page_setup.fitToWidth = 1
    ws_out.page_setup.fitToHeight = 0

    wb_out.save(output_path)


def print_excel_file(file_path):
    try:
        if platform.system() == 'Windows':
            os.startfile(file_path, 'print')
            return True
        else:
            subprocess.run(['lpr', file_path], check=True)
            return True
    except Exception as e:
        return str(e)


class App:
    def __init__(self, root):
        self.root = root
        self.root.title('منسق طلبات الأمين الحراري الشامل (مع محول TXT)')
        self.root.geometry('720x740')
        self.root.resizable(False, False)

        self.current_files = []
        self.txt_converter_files = []

        # --- قسم منسق طلبات الأمين (البرنامج الرئيسي) ---
        lbl_title = tk.Label(root, text='منسق ملفات الأمين للطابعة الحرارية (8سم)', font=('Tahoma', 13, 'bold'))
        lbl_title.pack(pady=5)

        frame_controls = tk.LabelFrame(root, text='إعدادات ومعالجة التقارير الحرارية', font=('Tahoma', 10, 'bold'))
        frame_controls.pack(fill='x', padx=15, pady=3)

        frame_type = tk.Frame(frame_controls)
        frame_type.pack(pady=5, fill='x', padx=10)
        tk.Label(frame_type, text='نوع التقرير:', font=('Tahoma', 10, 'bold')).pack(side='right', padx=5)
        self.file_type_var = tk.StringVar(value='تعرّف تلقائي')
        self.combo_type = ttk.Combobox(
            frame_type,
            textvariable=self.file_type_var,
            values=[
                'تعرّف تلقائي',
                'محلاية + خمس مواد',
                'ورقة الفرن',
                'هرايس بانواعها',
                'مستودع الجاهز',
                'ورقة الشعبيات',
                'جبنة حلوم وحليب مطبوخ'
            ],
            state='readonly',
            font=('Tahoma', 10, 'bold'),
            width=22
        )
        self.combo_type.pack(side='right', padx=5)

        self.chk_var = tk.BooleanVar(value=True)
        chk = tk.Checkbutton(frame_controls, text='حذف الصفوف فارغة/صفرية الكمية تلقائياً', variable=self.chk_var, font=('Tahoma', 9))
        chk.pack(anchor='e', padx=15, pady=2)

        self.direct_print_var = tk.BooleanVar(value=False)
        chk_print = tk.Checkbutton(frame_controls, text='إرسال الملفات للطابعة الحرارية مباشرة بعد إنتاجها', variable=self.direct_print_var, font=('Tahoma', 9, 'bold'), fg='#d9534f')
        chk_print.pack(anchor='e', padx=15, pady=2)

        frame_btns = tk.Frame(root)
        frame_btns.pack(pady=5)
        btn_select = tk.Button(
            frame_btns,
            text='📂 اختر ملف أو عدة ملفات Excel',
            font=('Tahoma', 10, 'bold'),
            bg='#007bff',
            fg='#ffffff',
            activebackground='#0056b3',
            activeforeground='#ffffff',
            disabledforeground='#ffffff',
            padx=10,
            pady=5,
            command=self.load_files_dialog
        )
        btn_select.pack(side='left', padx=5)

        frame_preview = tk.LabelFrame(root, text='معاينة الملفات المختارة', font=('Tahoma', 9, 'bold'))
        frame_preview.pack(fill='both', expand=True, padx=15, pady=3)

        scroll_x = ttk.Scrollbar(frame_preview, orient='horizontal')
        scroll_y = ttk.Scrollbar(frame_preview, orient='vertical')
        self.tree = ttk.Treeview(frame_preview, show='headings', height=5, xscrollcommand=scroll_x.set, yscrollcommand=scroll_y.set)
        scroll_x.config(command=self.tree.xview)
        scroll_y.config(command=self.tree.yview)
        scroll_x.pack(side='bottom', fill='x')
        scroll_y.pack(side='left', fill='y')
        self.tree.pack(fill='both', expand=True, padx=5, pady=5)

        self.btn_process = tk.Button(
            root,
            text='⚡ معالجة واستخراج جميع الملفات إلى مجلد سطح المكتب',
            font=('Tahoma', 11, 'bold'),
            bg='#28a745',
            fg='#ffffff',
            activebackground='#1e7e34',
            activeforeground='#ffffff',
            disabledforeground='#ffffff',
            padx=15,
            pady=7,
            state='disabled',
            command=self.process_files
        )
        self.btn_process.pack(pady=5)

        # --- قسم أداة تحويل Excel إلى TXT (Unicode) بالزرين المنفصلين ---
        frame_txt_tool = tk.LabelFrame(root, text='أداة محول ملفات Excel إلى TXT (Unicode)', font=('Tahoma', 10, 'bold'), fg='#0056b3')
        frame_txt_tool.pack(fill='x', padx=15, pady=5)

        self.lbl_txt_status = tk.Label(frame_txt_tool, text='لم يتم اختيار أي ملف للتحويل', font=('Tahoma', 9), fg='gray')
        self.lbl_txt_status.pack(pady=2)

        frame_txt_btns = tk.Frame(frame_txt_tool)
        frame_txt_btns.pack(pady=5)

        btn_txt_select = tk.Button(
            frame_txt_btns,
            text='📂 (واحدة أو أكثر) Excel اختر ملفات',
            font=('Tahoma', 10, 'bold'),
            bg='#28a745',
            fg='#ffffff',
            activebackground='#1e7e34',
            activeforeground='#ffffff',
            disabledforeground='#ffffff',
            padx=10,
            pady=5,
            command=self.load_txt_files_dialog
        )
        btn_txt_select.pack(side='left', padx=5)

        self.btn_txt_convert = tk.Button(
            frame_txt_btns,
            text='🔄 TXT (Unicode) تحويل الملفات إلى',
            font=('Tahoma', 11, 'bold'),
            bg='#007bff',
            fg='#ffffff',
            activebackground='#0056b3',
            activeforeground='#ffffff',
            disabledforeground='#ffffff',
            padx=12,
            pady=6,
            state='disabled',
            command=self.convert_to_txt_files
        )
        self.btn_txt_convert.pack(side='left', padx=5)

    def load_files_dialog(self):
        file_paths = filedialog.askopenfilenames(filetypes=[('Excel Files', '*.xlsx *.xls')])
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
            self.file_type_var.set('تعرّف تلقائي')
        self.btn_process.config(state='normal')

    def auto_detect_type(self, file_path):
        try:
            filename = os.path.basename(file_path).lower()
            if 'شعبيات' in filename:
                return 'ورقة الشعبيات'
            elif 'حلوم' in filename or 'حليب' in filename:
                return 'جبنة حلوم وحليب مطبوخ'
            elif 'هرايس' in filename or 'harees' in filename:
                return 'هرايس بانواعها'
            elif 'مستودع' in filename or 'جاهز' in filename or 'warehouse' in filename:
                return 'مستودع الجاهز'
            elif 'فرن' in filename or 'oven' in filename:
                return 'ورقة الفرن'
            elif 'مبيع' in filename or 'طلبات' in filename or 'محلاية' in filename:
                return 'محلاية + خمس مواد'
            else:
                wb = openpyxl.load_workbook(file_path, data_only=True)
                if 'تأثير الطلبيات على المخزون' in wb.sheetnames:
                    if 'رأسية التقرير' in wb.sheetnames:
                        ws_head = wb['رأسية التقرير']
                        for r in range(1, 4):
                            val = str(ws_head.cell(row=r, column=2).value or '').lower()
                            if 'هرايس' in val:
                                return 'هرايس بانواعها'
                            elif 'مستودع' in val:
                                return 'مستودع الجاهز'
                    return 'ورقة الفرن'
                else:
                    return 'محلاية + خمس مواد'
        except Exception:
            return 'محلاية + خمس مواد'

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
            cols = [f'col_{i}' for i in range(len(header_row))]
            self.tree['columns'] = cols
            for i, col_name in enumerate(header_row):
                header_text = str(col_name) if col_name is not None else f'عمود {i + 1}'
                self.tree.heading(f'col_{i}', text=header_text)
                self.tree.column(f'col_{i}', width=100, anchor='center')
            for row in rows[1:6]:
                self.tree.insert('', 'end', values=[str(v) if v is not None else '' for v in row])
        except Exception as e:
            messagebox.showerror('خطأ', f'تعذر معاينة الملف:\n{str(e)}')

    def process_files(self):
        if not self.current_files:
            return
        desktop_path = os.path.join(os.path.expanduser('~'), 'Desktop')
        output_dir = os.path.join(desktop_path, 'جاهز للطباعه')
        os.makedirs(output_dir, exist_ok=True)
        direct_print = self.direct_print_var.get()
        processed_count = 0
        manual_type = self.file_type_var.get()

        for file_path in self.current_files:
            if manual_type != 'تعرّف تلقائي':
                detected_type = manual_type
            else:
                detected_type = self.auto_detect_type(file_path)

            if detected_type == 'ورقة الشعبيات':
                try:
                    save_path = os.path.join(output_dir, 'جاهز_للطباعة_ورقة_الشعبيات.xlsx')
                    format_paired_items_by_client(
                        file_path,
                        save_path,
                        items_pair=['شعيبيات بالقشطة', 'شعيبيات جوز'],
                        report_title='ورقة الشعبيات',
                        remove_empty=self.chk_var.get()
                    )
                    if direct_print:
                        print_excel_file(save_path)
                    processed_count += 1
                except Exception:
                    continue
            elif detected_type == 'جبنة حلوم وحليب مطبوخ':
                try:
                    save_path = os.path.join(output_dir, 'جاهز_للطباعة_جبنة_حلوم_وحليب_مطبوخ.xlsx')
                    format_paired_items_by_client(
                        file_path,
                        save_path,
                        items_pair=['جبنة حلوم', 'حليب مطبوخ'],
                        report_title='جبنة حلوم وحليب مطبوخ',
                        remove_empty=self.chk_var.get()
                    )
                    if direct_print:
                        print_excel_file(save_path)
                    processed_count += 1
                except Exception:
                    continue
            elif detected_type in ['ورقة الفرن', 'هرايس بانواعها', 'مستودع الجاهز']:
                try:
                    if detected_type == 'هرايس بانواعها':
                        out_name = 'جاهز_للطباعة_هرايس.xlsx'
                    elif detected_type == 'مستودع الجاهز':
                        out_name = 'جاهز_للطباعة_مستودع_الجاهز.xlsx'
                    else:
                        out_name = 'جاهز_للطباعة_ورقة_الفرن.xlsx'

                    save_path = os.path.join(output_dir, out_name)
                    format_standard_two_column_sheet(file_path, save_path, detected_type)
                    if direct_print:
                        print_excel_file(save_path)
                    processed_count += 1
                except Exception:
                    continue
            else:
                if detected_type == 'محلاية + خمس مواد':
                    try:
                        path_both = os.path.join(output_dir, 'جاهز_للطباعة_محلاية_وغريبة.xlsx')
                        format_paired_items_by_client(
                            file_path,
                            path_both,
                            items_pair=['محلاية', 'غريبة بالقشطة'],
                            report_title='محلاية + غريبة',
                            remove_empty=self.chk_var.get()
                        )
                        if direct_print:
                            print_excel_file(path_both)

                        five_items_set = {'كنافة ناعمة', 'غاز سائل كبير', 'كريمة', 'عش لحمة نية', 'عش البلبل فستق نية'}
                        path_five = os.path.join(output_dir, 'جاهز_للطباعة_خمس_مواد.xlsx')
                        format_sales_orders_custom(file_path, path_five, selected_items=five_items_set, report_title='خمس مواد', remove_empty=self.chk_var.get())
                        if direct_print:
                            print_excel_file(path_five)

                        # استخراج الورقتين الجديدتين تلقائياً مع ملف حركة المواد
                        path_shaabiyat = os.path.join(output_dir, 'جاهز_للطباعة_ورقة_الشعبيات.xlsx')
                        format_paired_items_by_client(
                            file_path,
                            path_shaabiyat,
                            items_pair=['شعيبيات بالقشطة', 'شعيبيات جوز'],
                            report_title='ورقة الشعبيات',
                            remove_empty=self.chk_var.get()
                        )
                        if direct_print:
                            print_excel_file(path_shaabiyat)

                        path_cheese_milk = os.path.join(output_dir, 'جاهز_للطباعة_جبنة_حلوم_وحليب_مطبوخ.xlsx')
                        format_paired_items_by_client(
                            file_path,
                            path_cheese_milk,
                            items_pair=['جبنة حلوم', 'حليب مطبوخ'],
                            report_title='جبنة حلوم وحليب مطبوخ',
                            remove_empty=self.chk_var.get()
                        )
                        if direct_print:
                            print_excel_file(path_cheese_milk)

                        processed_count += 4
                    except Exception:
                        pass

        msg = f'تمت معالجة وإخراج جميع الملفات بنجاح تام ({processed_count} ملف)!\nموجودة الآن على سطح المكتب داخل مجلد (جاهز للطباعه).'
        if direct_print:
            msg += '\nوتم إرسالها للطباعة المباشرة على الطابعة الحرارية.'
        messagebox.showinfo('نجاح تام', msg)

    # --- دوال أداة تحويل الـ TXT باستخدام openpyxl فقط ---
    def load_txt_files_dialog(self):
        file_paths = filedialog.askopenfilenames(filetypes=[('Excel Files', '*.xlsx *.xls')])
        if not file_paths:
            return
        self.txt_converter_files = list(file_paths)
        self.lbl_txt_status.config(text=f'تم اختيار {len(self.txt_converter_files)} ملف للتحويل', fg='green')
        self.btn_txt_convert.config(state='normal')

    def convert_to_txt_files(self):
        if not self.txt_converter_files:
            return
        desktop_path = os.path.join(os.path.expanduser('~'), 'Desktop')
        output_dir = os.path.join(desktop_path, 'ملفات_TXT_الأمين')
        os.makedirs(output_dir, exist_ok=True)

        success_count = 0
        try:
            for file_path in self.txt_converter_files:
                wb = openpyxl.load_workbook(file_path, data_only=True)
                ws = wb.active

                base_name = os.path.splitext(os.path.basename(file_path))[0]
                out_txt_path = os.path.join(output_dir, f'{base_name}.txt')

                with open(out_txt_path, 'w', encoding='utf-16') as f:
                    for row in ws.iter_rows(values_only=True):
                        if any(row):
                            row_vals = [str(v) if v is not None else '' for v in row]
                            f.write('\t'.join(row_vals) + '\n')

                success_count += 1

            messagebox.showinfo('نجاح التحويل', f'تم تحويل {success_count} ملف بنجاح!\nموجودة الآن في مجلد (ملفات_TXT_الأمين) على سطح المكتب.')
            self.lbl_txt_status.config(text='لم يتم اختيار أي ملف للتحويل', fg='gray')
            self.txt_converter_files = []
            self.btn_txt_convert.config(state='disabled')
        except Exception as e:
            messagebox.showerror('خطأ في التحويل', f'حدث خطأ أثناء معالجة الملفات:\n{str(e)}')


if __name__ == '__main__':
    root = tk.Tk()
    app = App(root)
    root.mainloop()
