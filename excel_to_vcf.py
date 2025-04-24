import os
import tkinter as tk
from tkinter import filedialog, ttk, messagebox
import pandas as pd

class ExcelToVcfConverter:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel联系人转VCF工具")
        self.root.geometry("800x600")
        
        self.excel_file_path = None
        self.df = None
        self.column_mappings = {}
        
        self.create_widgets()
    
    def create_widgets(self):
        # 顶部框架 - 按钮区域
        btn_frame = tk.Frame(self.root)
        btn_frame.pack(fill=tk.X, padx=10, pady=10)
        
        import_btn = tk.Button(btn_frame, text="导入Excel", command=self.import_excel)
        import_btn.pack(side=tk.LEFT, padx=5)
        
        export_btn = tk.Button(btn_frame, text="导出VCF", command=self.export_vcf)
        export_btn.pack(side=tk.LEFT, padx=5)
        
        # 中间框架 - 表格预览
        preview_frame = tk.Frame(self.root)
        preview_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        self.tree = ttk.Treeview(preview_frame)
        self.tree.pack(fill=tk.BOTH, expand=True, side=tk.LEFT)
        
        scrollbar = ttk.Scrollbar(preview_frame, orient="vertical", command=self.tree.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.tree.configure(yscrollcommand=scrollbar.set)
        
        # 底部框架 - 字段映射
        mapping_frame = tk.Frame(self.root)
        mapping_frame.pack(fill=tk.X, padx=10, pady=10)
        
        # VCF所需字段标签
        tk.Label(mapping_frame, text="字段映射:").grid(row=0, column=0, sticky=tk.W)
        
        # 字段映射区域
        self.mapping_subframe = tk.Frame(mapping_frame)
        self.mapping_subframe.grid(row=1, column=0, sticky=tk.W)
    
    def import_excel(self):
        file_path = filedialog.askopenfilename(
            title="选择Excel文件",
            filetypes=[("Excel文件", "*.xlsx *.xls")]
        )
        
        if not file_path:
            return
        
        try:
            self.excel_file_path = file_path
            self.df = pd.read_excel(file_path)
            self.display_data()
            self.setup_mapping_controls()
            messagebox.showinfo("成功", f"已成功导入Excel文件: {os.path.basename(file_path)}")
        except Exception as e:
            messagebox.showerror("错误", f"导入Excel文件时出错: {str(e)}")
    
    def display_data(self):
        # 清除现有数据
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        # 设置列
        self.tree["columns"] = list(self.df.columns)
        self.tree["show"] = "headings"
        
        # 设置列标题
        for col in self.df.columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=100)
        
        # 添加数据行
        for i, row in self.df.iterrows():
            if i < 100:  # 只显示前100行，避免过多数据影响性能
                values = list(row)
                self.tree.insert("", tk.END, values=values)
    
    def setup_mapping_controls(self):
        # 清除现有的映射控件
        for widget in self.mapping_subframe.winfo_children():
            widget.destroy()
        
        self.column_mappings = {}
        excel_columns = [""] + list(self.df.columns)
        
        # 创建必要的VCF字段映射
        vcf_fields = [
            ("姓名", "FN"),
            ("手机", "TEL"),
            ("电子邮件", "EMAIL"),
            ("地址", "ADR"),
            ("组织", "ORG"),
            ("职位", "TITLE")
        ]
        
        # 为每个VCF字段创建下拉菜单
        for i, (field_label, field_code) in enumerate(vcf_fields):
            tk.Label(self.mapping_subframe, text=f"{field_label}:").grid(row=i, column=0, sticky=tk.W, padx=5, pady=2)
            combobox = ttk.Combobox(self.mapping_subframe, values=excel_columns, width=20)
            combobox.grid(row=i, column=1, sticky=tk.W, padx=5, pady=2)
            
            # 尝试智能匹配
            for col in self.df.columns:
                if field_label.lower() in col.lower() or field_code.lower() in col.lower():
                    combobox.set(col)
                    self.column_mappings[field_code] = col
                    break
            
            # 绑定选择变更事件
            field_code_copy = field_code  # 创建副本以避免闭包问题
            combobox.bind("<<ComboboxSelected>>", lambda event, code=field_code_copy: self.update_mapping(event, code))
    
    def update_mapping(self, event, field_code):
        selected_column = event.widget.get()
        if selected_column:
            self.column_mappings[field_code] = selected_column
        elif field_code in self.column_mappings:
            del self.column_mappings[field_code]
    
    def export_vcf(self):
        if self.df is None:
            messagebox.showwarning("警告", "请先导入Excel文件")
            return
        
        if not self.column_mappings:
            messagebox.showwarning("警告", "请先设置字段映射")
            return
        
        # 选择导出位置
        save_path = filedialog.asksaveasfilename(
            title="保存VCF文件",
            filetypes=[("VCF文件", "*.vcf")],
            defaultextension=".vcf"
        )
        
        if not save_path:
            return
        
        try:
            with open(save_path, 'w', encoding='utf-8') as vcf_file:
                for i, row in self.df.iterrows():
                    # 开始一个新的vCard
                    vcf_file.write("BEGIN:VCARD\n")
                    vcf_file.write("VERSION:3.0\n")
                    
                    # 添加各个字段
                    for field_code, excel_col in self.column_mappings.items():
                        if excel_col and excel_col in row:
                            value = str(row[excel_col])
                            if value and value != 'nan':
                                if field_code == "FN":
                                    vcf_file.write(f"FN:{value}\n")
                                    # 添加N字段
                                    vcf_file.write(f"N:{value};;;;\n")
                                elif field_code == "TEL":
                                    vcf_file.write(f"TEL;TYPE=CELL:{value}\n")
                                elif field_code == "EMAIL":
                                    vcf_file.write(f"EMAIL;TYPE=WORK:{value}\n")
                                elif field_code == "ADR":
                                    vcf_file.write(f"ADR;TYPE=HOME:;;{value};;;;\n")
                                elif field_code == "ORG":
                                    vcf_file.write(f"ORG:{value}\n")
                                elif field_code == "TITLE":
                                    vcf_file.write(f"TITLE:{value}\n")
                    
                    # 结束vCard
                    vcf_file.write("END:VCARD\n\n")
            
            messagebox.showinfo("成功", f"已成功导出VCF文件: {os.path.basename(save_path)}")
        except Exception as e:
            messagebox.showerror("错误", f"导出VCF文件时出错: {str(e)}")


if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelToVcfConverter(root)
    root.mainloop() 