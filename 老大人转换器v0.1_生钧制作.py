import pandas as pd
import tkinter as tk
from tkinter import filedialog, scrolledtext, messagebox, ttk
import os

class App:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel 亲称生成工具 (庙用·最终版)")
        self.root.geometry("650x550")
        self.root.resizable(True, True)

        self.file_path = None
        self.df_columns = [] # 用于存储Excel的列名

        # --- 顶部框架：文件选择 ---
        top_frame = tk.Frame(root)
        top_frame.pack(fill=tk.X, pady=5, padx=10)
        
        self.select_btn = tk.Button(top_frame, text="选择 XLSX 文件", command=self.select_file)
        self.select_btn.pack(side=tk.LEFT)

        self.file_label = tk.Label(top_frame, text="未选择文件", fg="gray")
        self.file_label.pack(side=tk.LEFT, padx=10)

        # --- 中部框架：列映射 ---
        mapping_frame = ttk.LabelFrame(root, text="请指定 Excel 中对应的列")
        mapping_frame.pack(fill=tk.X, pady=10, padx=10)

        self.yang_shang_var = tk.StringVar()
        self.yang_xia_var = tk.StringVar()
        self.appellation_var = tk.StringVar()

        tk.Label(mapping_frame, text="阳上人姓名列：").grid(row=0, column=0, sticky='w', padx=5, pady=5)
        self.yang_shang_menu = ttk.Combobox(mapping_frame, textvariable=self.yang_shang_var, state="readonly")
        self.yang_shang_menu.grid(row=0, column=1, sticky='ew', padx=5, pady=5)

        tk.Label(mapping_frame, text="亡人姓名列：").grid(row=1, column=0, sticky='w', padx=5, pady=5)
        self.yang_xia_menu = ttk.Combobox(mapping_frame, textvariable=self.yang_xia_var, state="readonly")
        self.yang_xia_menu.grid(row=1, column=1, sticky='ew', padx=5, pady=5)

        tk.Label(mapping_frame, text="称谓列：").grid(row=2, column=0, sticky='w', padx=5, pady=5)
        self.appellation_menu = ttk.Combobox(mapping_frame, textvariable=self.appellation_var, state="readonly")
        self.appellation_menu.grid(row=2, column=1, sticky='ew', padx=5, pady=5)
        
        mapping_frame.columnconfigure(1, weight=1)


        # --- 中部框架：操作按钮 ---
        action_frame = tk.Frame(root)
        action_frame.pack(pady=10)
        
        self.process_btn = tk.Button(action_frame, text="一键转换", command=self.process_file, state=tk.DISABLED, font=("微软雅黑", 10, "bold"))
        self.process_btn.pack()

        # --- 底部框架：日志 ---
        log_frame = ttk.LabelFrame(root, text="运行日志")
        log_frame.pack(fill=tk.BOTH, expand=True, pady=5, padx=10)

        self.log_area = scrolledtext.ScrolledText(log_frame, height=15, wrap=tk.WORD)
        self.log_area.pack(fill=tk.BOTH, expand=True)

    def log(self, message):
        self.log_area.insert(tk.END, message + "\n")
        self.log_area.see(tk.END)
        self.root.update_idletasks()

    def select_file(self):
        self.file_path = filedialog.askopenfilename(
            title="选择 Excel 文件",
            filetypes=[("Excel files", "*.xlsx")]
        )
        if self.file_path:
            self.file_label.config(text=os.path.basename(self.file_path), fg="black")
            self.process_btn.config(state=tk.NORMAL)
            self.log(f"已选择文件: {self.file_path}")
            
            # 读取列名并更新下拉菜单
            try:
                df_sample = pd.read_excel(self.file_path, nrows=0)
                self.df_columns = list(df_sample.columns)
                self.yang_shang_menu['values'] = self.df_columns
                self.yang_xia_menu['values'] = self.df_columns
                self.appellation_menu['values'] = self.df_columns
                
                # 尝试智能预设
                cols_lower = [c.lower() for c in self.df_columns]
                if '阳上人' in cols_lower: self.yang_shang_var.set(self.df_columns[cols_lower.index('阳上人')])
                elif '斋主' in ''.join(cols_lower): self.yang_shang_var.set(self.df_columns[[c.lower().find('斋主') for c in self.df_columns].index(0)])
                
                if '阳下人' in cols_lower: self.yang_xia_var.set(self.df_columns[cols_lower.index('阳下人')])
                elif '亡人' in ''.join(cols_lower): self.yang_xia_var.set(self.df_columns[[c.lower().find('亡人') for c in self.df_columns].index(0)])
                
                if '称谓' in cols_lower: self.appellation_var.set(self.df_columns[cols_lower.index('称谓')])

                self.log("列名已加载，请在下拉框中确认或调整。")
            except Exception as e:
                self.log(f"读取文件列名失败: {e}")

    def get_name_parts(self, abf, cde):
        parts = {'A': '', 'C': '', 'DE': ''}
        if pd.isna(abf) or not isinstance(abf, str): abf = ""
        if pd.isna(cde) or not isinstance(cde, str): cde = ""

        abf_str = str(abf).strip()
        cde_str = str(cde).strip()

        parts['A'] = abf_str[0] if abf_str else ''
        parts['C'] = cde_str[0] if cde_str else ''
        parts['DE'] = cde_str[1:] if len(cde_str) > 1 else ''
        return parts

    def normalize_text(self, text):
        return ' '.join(str(text).split())

    def generate_title_logic(self, abf, cde, appellation):
        if pd.isna(appellation):
            return pd.Series(['', True])

        appellation = str(appellation).strip()
        parts = self.get_name_parts(abf, cde)

        if appellation in ['父亲', '父子', '父女', '公公']:
            title = f"显考 {parts['C']}公 {parts['DE']} 老大人"
        elif appellation in ['母亲', '母子', '母女', '婆婆']:
            title = f"显妣 {parts['A']}母 {parts['C']}氏 {parts['DE']} 老孺人"
        elif appellation in ['祖父', '爷爷'] and '外' not in appellation:
            title = f"祖考 {parts['C']}公 {parts['DE']} 老大人"
        elif appellation in ['祖母', '奶奶'] and '外' not in appellation:
            title = f"祖妣 {parts['A']}母 {parts['C']}氏 {parts['DE']} 老孺人"
        elif appellation in ['外祖父', '姥爷', '外公']:
            title = f"外祖考 {parts['C']}公 {parts['DE']} 老大人"
        elif appellation in ['外祖母', '姥姥', '外婆']:
            title = f"外祖妣 {parts['A']}母 {parts['C']}氏 {parts['DE']} 老孺人"
        else:
            title = f"未知称谓：{appellation}"
            return pd.Series([self.normalize_text(title), True])

        return pd.Series([self.normalize_text(title), False])

    def process_file(self):
        if not self.file_path or not self.yang_shang_var.get():
            messagebox.showerror("错误", "请先选择文件并指定所有列。")
            return

        try:
            self.log("开始处理文件...")
            ys_col = self.yang_shang_var.get()
            yx_col = self.yang_xia_var.get()
            app_col = self.appellation_var.get()

            xls = pd.ExcelFile(self.file_path)
            base_dir = os.path.dirname(self.file_path)
            new_file_path = os.path.join(base_dir, f"已生成称呼版_{os.path.basename(self.file_path)}")

            with pd.ExcelWriter(new_file_path, engine='xlsxwriter') as writer:
                for sheet_name in xls.sheet_names:
                    self.log(f"正在处理工作表: {sheet_name}")
                    df = pd.read_excel(xls, sheet_name=sheet_name)
                    
                    # 检查所选列是否存在
                    if not all(col in df.columns for col in [ys_col, yx_col, app_col]):
                        self.log(f"警告: 工作表 '{sheet_name}' 缺少指定的列，已跳过。")
                        continue

                    result = df.apply(lambda row: self.generate_title_logic(row[ys_col], row[yx_col], row[app_col]), axis=1)
                    
                    # 将结果追加到最后一列
                    df['完整亲称'] = result[0]

                    df.to_excel(writer, sheet_name=sheet_name, index=False)

                    workbook = writer.book
                    worksheet = writer.sheets[sheet_name]
                    red_format = workbook.add_format({'font_color': 'red'})

                    appellation_col_idx = df.columns.get_loc(app_col)
                    for row_idx, is_error in enumerate(result[1]):
                        if is_error:
                            worksheet.write(row_idx + 1, appellation_col_idx, df.iloc[row_idx][app_col], red_format)
            
            self.log(f"\n✅ 文件处理成功！已保存至: {new_file_path}")
            messagebox.showinfo("完成", f"文件已成功转换并保存至:\n{new_file_path}")

        except Exception as e:
            self.log(f"\n❌ 发生严重错误: {str(e)}")
            messagebox.showerror("错误", f"处理过程中发生错误:\n{str(e)}")

if __name__ == "__main__":
    root = tk.Tk()
    App(root)
    root.mainloop()