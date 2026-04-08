
import pandas as pd
import json
import re
import os
import sys
from tkinter import Tk, Label, Button, Entry, StringVar, filedialog, messagebox, ttk
from tkinter.scrolledtext import ScrolledText

def convert_excel_to_json(excel_path, json_path=None):
    """将Excel文件转换为JSON格式"""
    try:
        # 读取Excel文件
        df = pd.read_excel(excel_path, header=1)

        # 重新命名列
        df.columns = ['Type', 'Name', 'Address', 'DataType', 'InitValue', 'Comment', 'Attribute']

        # 删除可能的重复表头行
        df = df[df['Type'] != 'Type'].reset_index(drop=True)

        # 创建JSON列表
        json_list = []

        for _, row in df.iterrows():
            # 提取Name，去掉 DI_ 或 DO_ 前缀
            name = row['Name']
            if isinstance(name, str):
                name_clean = re.sub(r'^(DI_|DO_)', '', name)
            else:
                name_clean = str(name)

            # 提取Address，去掉 % 符号
            address = row['Address']
            if isinstance(address, str):
                address_clean = address.replace('%', '')
            else:
                address_clean = str(address)

            # 创建JSON对象
            json_obj = {
                "Name": name_clean,
                "Address": address_clean
            }
            json_list.append(json_obj)

        # 如果未指定输出路径，使用默认路径
        if json_path is None:
            base_name = os.path.splitext(excel_path)[0]
            json_path = base_name + '.json'

        # 保存JSON文件
        with open(json_path, 'w', encoding='utf-8') as f:
            json.dump(json_list, f, ensure_ascii=False, indent=2)

        return True, f"转换成功！\n共转换 {len(json_list)} 条记录\n保存至: {json_path}"

    except Exception as e:
        return False, f"转换失败: {str(e)}"

class ExcelToJsonConverter:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel转JSON转换器")
        self.root.geometry("600x450")
        self.root.resizable(False, False)

        # 设置样式
        self.style = ttk.Style()
        self.style.configure('TButton', font=('微软雅黑', 10))
        self.style.configure('TLabel', font=('微软雅黑', 10))

        # 文件选择区域
        ttk.Label(root, text="Excel文件路径:").grid(row=0, column=0, padx=10, pady=10, sticky='w')

        self.excel_path = StringVar()
        ttk.Entry(root, textvariable=self.excel_path, width=50).grid(row=0, column=1, padx=5, pady=10)
        ttk.Button(root, text="浏览...", command=self.browse_excel).grid(row=0, column=2, padx=5, pady=10)

        # 输出路径区域
        ttk.Label(root, text="JSON输出路径:").grid(row=1, column=0, padx=10, pady=10, sticky='w')

        self.json_path = StringVar()
        ttk.Entry(root, textvariable=self.json_path, width=50).grid(row=1, column=1, padx=5, pady=10)
        ttk.Button(root, text="浏览...", command=self.browse_json).grid(row=1, column=2, padx=5, pady=10)

        # 转换按钮
        ttk.Button(root, text="开始转换", command=self.convert, width=20).grid(row=2, column=1, pady=20)

        # 结果显示区域
        ttk.Label(root, text="转换结果:").grid(row=3, column=0, padx=10, pady=5, sticky='nw')

        self.result_text = ScrolledText(root, width=70, height=15, wrap='word')
        self.result_text.grid(row=4, column=0, columnspan=3, padx=10, pady=5)

        # 状态栏
        self.status_label = ttk.Label(root, text="就绪", foreground="gray")
        self.status_label.grid(row=5, column=0, columnspan=3, pady=5)

        # 说明文字
        ttk.Label(root, text="说明: 支持 .xls 和 .xlsx 格式的Excel文件", 
                 foreground="gray", font=('微软雅黑', 8)).grid(row=6, column=0, columnspan=3, pady=5)

    def browse_excel(self):
        filename = filedialog.askopenfilename(
            title="选择Excel文件",
            filetypes=[("Excel文件", "*.xls *.xlsx"), ("所有文件", "*.*")]
        )
        if filename:
            self.excel_path.set(filename)
            # 自动设置默认输出路径
            if not self.json_path.get():
                base_name = os.path.splitext(filename)[0]
                self.json_path.set(base_name + '.json')

    def browse_json(self):
        filename = filedialog.asksaveasfilename(
            title="保存JSON文件",
            defaultextension=".json",
            filetypes=[("JSON文件", "*.json"), ("所有文件", "*.*")]
        )
        if filename:
            self.json_path.set(filename)

    def convert(self):
        excel_file = self.excel_path.get().strip()
        json_file = self.json_path.get().strip()

        if not excel_file:
            messagebox.showwarning("警告", "请选择Excel文件！")
            return

        if not os.path.exists(excel_file):
            messagebox.showerror("错误", "Excel文件不存在！")
            return

        self.status_label.config(text="正在转换...", foreground="blue")
        self.root.update()

        success, message = convert_excel_to_json(excel_file, json_file if json_file else None)

        self.result_text.delete(1.0, 'end')
        self.result_text.insert('end', message)

        if success:
            self.status_label.config(text="转换完成", foreground="green")
            messagebox.showinfo("成功", "转换完成！")
        else:
            self.status_label.config(text="转换失败", foreground="red")
            messagebox.showerror("错误", message)

def main():
    # 支持命令行模式
    if len(sys.argv) > 1:
        # 命令行模式
        excel_path = sys.argv[1]
        json_path = sys.argv[2] if len(sys.argv) > 2 else None

        if not os.path.exists(excel_path):
            print(f"错误: 文件不存在 - {excel_path}")
            sys.exit(1)

        success, message = convert_excel_to_json(excel_path, json_path)
        print(message)
        sys.exit(0 if success else 1)
    else:
        # GUI模式
        root = Tk()

        # 设置窗口图标（如果有的话）
        try:
            root.iconbitmap('icon.ico')
        except:
            pass

        app = ExcelToJsonConverter(root)
        root.mainloop()

if __name__ == "__main__":
    main()
