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

# --- معالجة الملف الأول: ورقة الفرن ---
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

# --- الواجهة الرئيسية ---
class App:
    def __init__(self, root):
        self.root = root
        self.root.title("منسق طلبات الأمين الحراري (8سم)")
        self.root.geometry("600x550")
        self.root.resizable(False, False)

        self.current_file_path = None

        lbl_title = tk.Label(root, text="منسق ملفات الأمين للطابعة الحرارية (8سم)", font=("Arial", 14, "bold"))
        lbl_title.pack(pady=10)

        # إطار اختيار نوع التقرير
        frame_type = tk.Frame(root)
        frame_type.pack(pady=5)

        tk.Label(frame_type, text="نوع التقرير:", font=("Arial", 11, "bold")).pack(side="right", padx=5)
        
        self.file_type_var = tk.StringVar(value="تعرّف تلقائي")
        self.combo_type = ttk.Combobox(
            frame_type, 
            textvariable=self.file_type_var, 
            values=["تعرّف تلقائي", "ورقة الفرن", "الملف الثاني (قريباً)", "الملف الثالث (قريباً)"],
            state="readonly", 
            font=("Arial", 10, "bold"),
            width=22
        )
        self.combo_type.pack(side="right", padx=5)

        # زر أخذ الملف
        btn_select = tk.Button(root, text="اختر ملف Excel لاستعراضه وتنسيقه", font=("Arial", 11, "bold"), bg="#007bff", fg="white", padx=15, pady=6, command=self.load_file)
        btn_select.pack(pady=10)

        # إطار المعاينة
        frame_preview = tk.LabelFrame(root, text="معاينة سريعة للملف", font=("Arial", 10, "bold"))
        frame_preview.pack(fill="both", expand=True, padx=15, pady=5)

        self.tree = ttk.Treeview(frame_preview, show="headings", height=8)
        self.tree.pack(fill="both", expand=True, padx=5, pady=5)

        # زر المعالجة والاستخراج
        self.btn_process = tk.Button(root, text="استخراج الملف الجاهز للطباعة", font=("Arial", 12, "bold"), bg="#28a745", fg="white", padx=20, pady=8, state="disabled", command=self.process_file)
        self.btn_process.pack(pady=15)

    def load_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx *.xls")])
        if not file_path:
            return

        self.current_file_path = file_path
        self.preview_excel(file_path)
        self.auto_detect_type(file_path)
        self.btn_process.config(state="normal")

    def auto_detect_type(self, file_path):
        if self.file_type_var.get() != "تعرّف تلقائي":
            return
        try:
            wb = openpyxl.load_workbook(file_path, data_only=True)
            if 'تأثير الطلبيات على المخزون' in wb.sheetnames:
                self.combo_type.set("ورقة الفرن")
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

            cols = [f"Col {i+1}" for i in range(len(rows[0]))]
            self.tree["columns"] = cols

            for col in cols:
                self.tree.heading(col, text=col)
                self.tree.column(col, width=100, anchor="center")

            for row in rows[:8]:
                self.tree.insert("", "end", values=[str(v) if v is not None else "" for v in row])

        except Exception as e:
            messagebox.showerror("خطأ في المعاينة", f"تعذر قراءة الملف للمعاينة:\n{str(e)}")

    def process_file(self):
        if not self.current_file_path:
            return

        selected_type = self.file_type_var.get()
        base_name = os.path.basename(self.current_file_path)
        save_path = filedialog.asksaveasfilename(defaultextension=".xlsx", initialfile="جاهز_للطباعة_" + base_name, filetypes=[("Excel Files", "*.xlsx")])

        if not save_path:
            return

        try:
            if selected_type == "ورقة الفرن":
                format_oven_sheet(self.current_file_path, save_path)
                messagebox.showinfo("نجاح", f"تم استخراج ملف ورقة الفرن بنجاح:\n{save_path}")
            else:
                messagebox.showwarning("تنبيه", "يرجى تحديد نوع ملف صالح للمعالجة.")
        except Exception as e:
            messagebox.showerror("خطأ", f"حدث خطأ أثناء معالجة الملف:\n{str(e)}")

if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop()
