import os
import sys
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill

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

# --- دالة استخراج المواد الفريدة من ملف طلبيات المبيع ---
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

# --- معالجة الملف الثاني: ورقة الفرن ---
def format_oven_sheet(input_path, output_path):
    wb_src = openpyxl.load_workbook(input_path, data_only=True)
    ws_src = wb_src.active

    data = []
    for row in ws_src.iter_rows(values_only=True):
        if any(row):
            data.append([row[0], row[1]])

    wb_out = openpyxl.Workbook()
    ws_out = wb_out.active
    ws_out.title = "ورقة الفرن"
    ws_out.views.sheetView[0].rightToLeft = True

    ws_out.merge_cells('A1:B1')
    ws_out['A1'] = "ورقة الفرن"
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

    font_title = Font(name='Arial', size=26, bold=True)
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
            frame_quick, 
            text="محلاية + غريبة", 
            command=self.select_both, 
            bg_color="#6f42c1", 
            hover_color="#522f92", 
            width=165, 
            height=46, 
            radius=16, 
            font=("Arial", 12, "bold")
        )
        btn_both.pack(side="right", padx=10, pady=10)

        btn_five = RoundedButton(
            frame_quick, 
            text="خمس مواد", 
            command=self.select_five_items, 
            bg_color="#17a2b8", 
            hover_color="#117a8b", 
            width=140, 
            height=46, 
            radius=16, 
            font=("Arial", 12, "bold")
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

        scrollable_frame.bind(
            "<Configure>",
            lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        )

        self.canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        self.canvas.configure(yscrollcommand=scrollbar.set)

        self.canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        self.bind_mouse_wheel(self)
        self.bind_mouse_wheel(self.canvas)
        self.bind_mouse_wheel(scrollable_frame)

        for item in items:
            var = tk.BooleanVar(value=True)
            chk = tk.Checkbutton(scrollable_frame, text=item, variable=var, font=("Arial", 10), anchor="w")
            chk.pack(fill="x", pady=2, padx=5)
            self.bind_mouse_wheel(chk)
            self.item_vars[item] = var

        btn_confirm = tk.Button(self, text="استخراج ملف Excel الجاهز", font=("Arial", 11, "bold"), bg="#28a745", fg="white", padx=12, pady=8, command=self.confirm_selection)
        btn_confirm.pack(pady=10)

    def bind_mouse_wheel(self, widget):
        widget.bind("<MouseWheel>", self._on_mousewheel)
        widget.bind("<Button-4>", self._on_mousewheel)
        widget.bind("<Button-5>", self._on_mousewheel)

    def _on_mousewheel(self, event):
        if event.num == 4:
            self.canvas.yview_scroll(-1, "units")
        elif event.num == 5:
            self.canvas.yview_scroll(1, "units")
        else:
            self.canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

    def select_all(self):
        for var in self.item_vars.values(): var.set(True)

    def deselect_all(self):
        for var in self.item_vars.values(): var.set(False)

    def select_both(self):
        for item, var in self.item_vars.items():
            name = item.strip()
            var.set(name == 'محلاية' or name == 'غريبة بالقشطة')

    def select_five_items(self):
        target_items = {
            "غاز سائل كبير",
            "عش البلبل فستق نية",
            "عش لحمة نية",
            "كريمة",
            "كنافة ناعمة"
        }
        for item, var in self.item_vars.items():
            name = item.strip()
            var.set(name in target_items)

    def confirm_selection(self):
        selected = [item for item, var in self.item_vars.items() if var.get()]
        if not selected:
            messagebox.showwarning("تنبيه", "يرجى تحديد مادة واحدة على الأقل!")
            return
        self.callback(selected)
        self.destroy()

# --- الشاشة الرئيسية الشاملة للبرنامج ---
class App:
    def __init__(self, root):
        self.root = root
        self.root.title("منسق طلبات الأمين الحراري (8سم)")
        self.root.geometry("680x620")
        self.root.resizable(False, False)

        self.current_file_path = None

        lbl_title = tk.Label(root, text="منسق ملفات الأمين للطابعة الحرارية (8سم)", font=("Arial", 14, "bold"))
        lbl_title.pack(pady=10)

        frame_type = tk.Frame(root)
        frame_type.pack(pady=5)

        tk.Label(frame_type, text="نوع التقرير:", font=("Arial", 11, "bold")).pack(side="right", padx=5)
        
        self.file_type_var = tk.StringVar(value="تعرّف تلقائي")
        self.combo_type = ttk.Combobox(
            frame_type, 
            textvariable=self.file_type_var, 
            values=["تعرّف تلقائي", "محلاية + خمس مواد", "ورقة الفرن", "الملف الثالث (قريباً)"],
            state="readonly", 
            font=("Arial", 10, "bold"),
            width=22
        )
        self.combo_type.pack(side="right", padx=5)

        self.chk_var = tk.BooleanVar(value=True)
        chk = tk.Checkbutton(root, text="حذف الصفوف فارغة/صفرية الكمية تلقائياً (لمحلاية + خمس مواد)", variable=self.chk_var, font=("Arial", 10))
        chk.pack(pady=3)

        btn_select = tk.Button(root, text="اختر ملف Excel لاستعراضه وتنسيقه", font=("Arial", 11, "bold"), bg="#007bff", fg="white", padx=15, pady=6, command=self.load_file)
        btn_select.pack(pady=8)

        frame_preview = tk.LabelFrame(root, text="معاينة سريعة للملف", font=("Arial", 10, "bold"))
        frame_preview.pack(fill="both", expand=True, padx=15, pady=5)

        scroll_x = ttk.Scrollbar(frame_preview, orient="horizontal")
        scroll_y = ttk.Scrollbar(frame_preview, orient="vertical")

        self.tree = ttk.Treeview(
            frame_preview, 
            show="headings", 
            height=7, 
            xscrollcommand=scroll_x.set, 
            yscrollcommand=scroll_y.set
        )
        
        scroll_x.config(command=self.tree.xview)
        scroll_y.config(command=self.tree.yview)

        scroll_x.pack(side="bottom", fill="x")
        scroll_y.pack(side="left", fill="y")
        self.tree.pack(fill="both", expand=True, padx=5, pady=5)

        self.btn_process = tk.Button(root, text="استخراج الملف الجاهز للطباعة", font=("Arial", 12, "bold"), bg="#28a745", fg="white", padx=20, pady=8, state="disabled", command=self.process_file)
        self.btn_process.pack(pady=12)

    def load_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx *.xls")])
        if not file_path:
            return

        self.current_file_path = file_path
        self.preview_excel(file_path)
        
        # إعادة ضبط نوع التقرير إلى التعرّف التلقائي ليتم فحص الملف الجديد بدقة وتحديث القائمة
        self.file_type_var.set("تعرّف تلقائي")
        self.auto_detect_type(file_path)
        
        self.btn_process.config(state="normal")

    def auto_detect_type(self, file_path):
        if self.file_type_var.get() != "تعرّف تلقائي":
            return
        try:
            wb = openpyxl.load_workbook(file_path, data_only=True)
            if 'تأثير الطلبيات على المخزون' in wb.sheetnames:
                self.combo_type.set("ورقة الفرن")
            else:
                self.combo_type.set("محلاية + خمس مواد")
        except Exception:
            pass

    def preview_excel(self, file_path):
        for item in self.tree.get_children():
            self.tree.delete(item)

        try:
            wb = openpyxl.load_workbook(file_path, data_only=True)
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

            for row in rows[1:10]:
                self.tree.insert("", "end", values=[str(v) if v is not None else "" for v in row])

        except Exception as e:
            messagebox.showerror("خطأ في المعاينة", f"تعذر قراءة الملف للمعاينة:\n{str(e)}")

    def process_file(self):
        if not self.current_file_path:
            return

        selected_type = self.file_type_var.get()
        
        if selected_type == "تعرّف تلقائي":
            self.auto_detect_type(self.current_file_path)
            selected_type = self.file_type_var.get()

        if selected_type == "محلاية + خمس مواد":
            try:
                items, _ = get_unique_items(self.current_file_path)
                if not items:
                    messagebox.showerror("خطأ", "لم يتم العثور على أي مواد داخل الملف!")
                    return

                def on_items_selected(selected_items):
                    base_name = os.path.basename(self.current_file_path)
                    save_path = filedialog.asksaveasfilename(defaultextension=".xlsx", initialfile="جاهز_للطباعة_" + base_name, filetypes=[("Excel Files", "*.xlsx")])
                    if not save_path:
                        return
                    try:
                        format_sales_orders(self.current_file_path, save_path, selected_items, remove_empty=self.chk_var.get())
                        messagebox.showinfo("نجاح", f"تم استخراج الملف بنجاح:\n{save_path}")
                    except Exception as e:
                        messagebox.showerror("خطأ", f"حدث خطأ أثناء حفظ الملف:\n{str(e)}")

                ItemSelectorWindow(self.root, items, on_items_selected)

            except Exception as e:
                messagebox.showerror("خطأ", f"حدث خطأ أثناء قراءة المواد:\n{str(e)}")

        elif selected_type == "ورقة الفرن":
            base_name = os.path.basename(self.current_file_path)
            save_path = filedialog.asksaveasfilename(defaultextension=".xlsx", initialfile="جاهز_للطباعة_" + base_name, filetypes=[("Excel Files", "*.xlsx")])
            if not save_path:
                return
            try:
                format_oven_sheet(self.current_file_path, save_path)
                messagebox.showinfo("نجاح", f"تم استخراج ملف ورقة الفرن بنجاح:\n{save_path}")
            except Exception as e:
                messagebox.showerror("خطأ", f"حدث خطأ أثناء معالجة الملف:\n{str(e)}")
        else:
            messagebox.showwarning("تنبيه", "يرجى تحديد نوع ملف صالح للمعالجة.")

if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop()
