import pandas as pd
import sys
import os

def process_excel(file_path):
    # قراءة ملف الإكسل
    df = pd.read_excel(file_path)
    
    # تنظيف وتعديل الأعمدة حسب الحاجة لطباعة حرارية 8 سم
    # يمكنك تعديل المعالجة بناءً على هيكل ملفات برنامج الأمين لديك
    
    is_header = True
    if is_header:
        # معالجة رأس الفاتورة أو الطلب
        print("Processing header...")
        
    # مثال على معالجة البيانات
    processed_df = df.dropna(how='all')
    
    return processed_df

if __name__ == "__main__":
    if len(sys.argv) > 1:
        input_file = sys.argv[1]
        process_excel(input_file)
    else:
        print("Please provide the Excel file path.")
