import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
import os

def convert_to_excel(csv_path, separator, output_path):
    encodings = ['utf-8', 'latin1', 'cp1252']
    for encoding in encodings:
        try:
            df = pd.read_csv(csv_path, sep=separator, encoding=encoding, on_bad_lines='skip')
            df.to_excel(output_path, index=False)
            messagebox.showinfo("Success", f"CSV has been converted to Excel using encoding {encoding}: {output_path}")
            return
        except Exception as e:
            last_exception = e
    messagebox.showerror("Error", f"Failed to convert CSV to Excel. Last error: {str(last_exception)}")

def select_file():
    file_path = filedialog.askopenfilename(title="Select CSV file", filetypes=[("CSV files", "*.csv")])
    if file_path:
        default_output_path = os.path.splitext(file_path)[0] + '.xlsx'
        output_path = filedialog.asksaveasfilename(title="Save as Excel file", defaultextension=".xlsx", initialfile=default_output_path, filetypes=[("Excel files", "*.xlsx")])
        if output_path:
            separator = separator_entry.get()
            convert_to_excel(file_path, separator, output_path)

# 创建Tkinter窗口
root = tk.Tk()
root.title("CSV to Excel Converter")
root.geometry("300x150")  # 固定窗口大小

# 标签和输入框
tk.Label(root, text="分隔符:").pack(pady=10)
separator_entry = tk.Entry(root, width=5)
separator_entry.pack(pady=5)
separator_entry.insert(0, ",")

# 转换按钮
convert_button = tk.Button(root, text="转换", command=select_file)
convert_button.pack(pady=20)
convert_button.config(width=20, height=2)

# 运行主循环
root.mainloop()
