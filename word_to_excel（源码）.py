import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog
from docx import Document
import pandas as pd

def extract_data_from_docx(docx_path, field_names):
    doc = Document(docx_path)
    data = []
    for para in doc.paragraphs:
        if para.text.strip():  # 忽略空行
            parts = para.text.split(',')
            if len(parts) == len(field_names):  # 假设文档中包含与用户定义的字段数相同的字段
                data.append({field: part.strip() for field, part in zip(field_names, parts)})
    return data

def save_data_to_excel(data, excel_path):
    df = pd.DataFrame(data)
    df.to_excel(excel_path, index=False)

def process_docx_to_excel():
    # 让用户输入他们想要的字段名
    field_names = simpledialog.askstring("Input", "请输入字段名，用逗号分隔（例如：姓名,性别,班级,学号,学分,学时）")
    if not field_names:
        messagebox.showerror("Error", "未输入字段名")
        return
    field_names = [name.strip() for name in field_names.split(',')]
    
    docx_path = filedialog.askopenfilename(filetypes=[("Word files", "*.docx")])
    if not docx_path:
        return
    data = extract_data_from_docx(docx_path, field_names)
    if data:
        excel_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if excel_path:
            save_data_to_excel(data, excel_path)
            messagebox.showinfo("Success", f"数据已成功导出到Excel文件：{excel_path}")
    else:
        messagebox.showerror("Error", "未找到有效数据或格式不正确")

# 创建主窗口
root = tk.Tk()
root.title("Word to Excel Converter")

# 创建按钮
button_convert = tk.Button(root, text="转换Word文档到Excel", command=process_docx_to_excel)
button_convert.pack(pady=20)

# 运行主循环
root.mainloop()