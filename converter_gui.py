import tkinter as tk
from tkinter import ttk
import re

def convert_to_comma_separated(input_str, remove_duplicates=False):
    # 使用正则表达式匹配所有非字母数字字符作为分隔符
    numbers = re.split(r'\W+', input_str.strip())
    # 过滤掉空字符串
    numbers = [num for num in numbers if num]
    if remove_duplicates:
        numbers = list(dict.fromkeys(numbers))
    output_str = ','.join(numbers)
    # 统计逗号分隔的字符串数量
    num_count = len(numbers)
    return output_str, num_count

def convert_and_display():
    input_str = input_text.get("1.0", tk.END)
    remove_duplicates = remove_duplicates_var.get()
    output_str, num_count = convert_to_comma_separated(input_str, remove_duplicates)
    output_text.delete("1.0", tk.END)
    output_text.insert(tk.END, output_str)
    char_count_label.config(text=f"逗号分隔的字符串数量: {num_count}")

# 创建主窗口
root = tk.Tk()
root.title("任意分隔符转换为逗号")

# 设置行列权重，使得组件随窗口大小调整
root.grid_rowconfigure(1, weight=1)
root.grid_rowconfigure(3, weight=1)
root.grid_columnconfigure(0, weight=1)

# 创建输入标签和文本框
input_label = ttk.Label(root, text="输入 (任意分隔符):")
input_label.grid(row=0, column=0, padx=10, pady=5, sticky='w')
input_text = tk.Text(root, height=10)
input_text.grid(row=1, column=0, padx=10, pady=5, sticky='nsew')

# 创建输出标签和文本框
output_label = ttk.Label(root, text="输出 (以逗号分隔):")
output_label.grid(row=2, column=0, padx=10, pady=5, sticky='w')
output_text = tk.Text(root, height=10)
output_text.grid(row=3, column=0, padx=10, pady=5, sticky='nsew')

# 创建字符串数量标签
char_count_label = ttk.Label(root, text="逗号分隔的字符串数量: 0")
char_count_label.grid(row=4, column=0, padx=10, pady=5, sticky='w')

# 创建去除重复字符串复选框
remove_duplicates_var = tk.BooleanVar()
remove_duplicates_checkbox = ttk.Checkbutton(
    root, text="去除重复的字符串", variable=remove_duplicates_var)
remove_duplicates_checkbox.grid(row=5, column=0, padx=10, pady=5, sticky='w')

# 创建转换按钮
convert_button = ttk.Button(root, text="转换", command=convert_and_display)
convert_button.grid(row=6, column=0, padx=10, pady=10)

# 运行主循环
root.mainloop()
