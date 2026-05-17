# -*- coding: utf-8 -*-
# 高DPI适配（必须放在最开头！）
import ctypes
try:
    ctypes.windll.shcore.SetProcessDpiAwarenessContext(-4)
except Exception:
    try:
        ctypes.windll.user32.SetProcessDPIAware()
    except Exception:
        pass

import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox, filedialog
import pandas as pd
import os

# ====================== 内置表格数据 ======================
year_data = {
    "甲子": {"欠债_num": 53000, "看经卷": 18, "纳库": 3, "曹官姓": "元"},
    "丙子": {"欠债_num": 73000, "看经卷": 25, "纳库": 9, "曹官姓": "王"},
    "戊子": {"欠债_num": 63000, "看经卷": 21, "纳库": 6, "曹官姓": "伍"},
    "庚子": {"欠债_num": 110000, "看经卷": 37, "纳库": 53, "曹官姓": "李"},
    "壬子": {"欠债_num": 70000, "看经卷": 24, "纳库": 38, "曹官姓": "孟"},
    "甲午": {"欠债_num": 40000, "看经卷": 14, "纳库": 21, "曹官姓": "牛"},
    "丙午": {"欠债_num": 53000, "看经卷": 18, "纳库": 60, "曹官姓": "蕭"},
    "戊午": {"欠债_num": 90000, "看经卷": 30, "纳库": 39, "曹官姓": "史"},
    "庚午": {"欠债_num": 62000, "看经卷": 21, "纳库": 22, "曹官姓": "陳"},
    "壬午": {"欠债_num": 70000, "看经卷": 24, "纳库": 44, "曹官姓": "孔"},
    "乙丑": {"欠债_num": 280000, "看经卷": 94, "纳库": 13, "曹官姓": "田"},
    "丁丑": {"欠债_num": 42000, "看经卷": 15, "纳库": 34, "曹官姓": "崔"},
    "己丑": {"欠债_num": 80000, "看经卷": 27, "纳库": 7, "曹官姓": "周"},
    "辛丑": {"欠债_num": 110000, "看经卷": 37, "纳库": 28, "曹官姓": "吉"},
    "癸丑": {"欠债_num": 27000, "看经卷": 9, "纳库": 8, "曹官姓": "習"},
    "乙未": {"欠债_num": 40000, "看经卷": 51, "纳库": 51, "曹官姓": "皇"},
    "丁未": {"欠债_num": 91000, "看经卷": 52, "纳库": 52, "曹官姓": "朱"},
    "己未": {"欠债_num": 43000, "看经卷": 15, "纳库": 5, "曹官姓": "卞"},
    "辛未": {"欠债_num": 101000, "看经卷": 59, "纳库": 59, "曹官姓": "常"},
    "癸未": {"欠债_num": 52000, "看经卷": 18, "纳库": 48, "曹官姓": "朱"},
    "甲寅": {"欠债_num": 33000, "看经卷": 11, "纳库": 31, "曹官姓": "杜"},
    "丙寅": {"欠债_num": 80000, "看经卷": 27, "纳库": 23, "曹官姓": "马"},
    "戊寅": {"欠债_num": 60000, "看经卷": 20, "纳库": 11, "曹官姓": "郭"},
    "庚寅": {"欠债_num": 51000, "看经卷": 17, "纳库": 25, "曹官姓": "毛"},
    "壬寅": {"欠债_num": 56000, "看经卷": 32, "纳库": 11, "曹官姓": "施"},
    "甲申": {"欠债_num": 70000, "看经卷": 24, "纳库": 56, "曹官姓": "吕"},
    "丙申": {"欠债_num": 33000, "看经卷": 11, "纳库": 57, "曹官姓": "紐"},
    "戊申": {"欠债_num": 80000, "看经卷": 27, "纳库": 58, "曹官姓": "柴"},
    "庚申": {"欠债_num": 61000, "看经卷": 21, "纳库": 45, "曹官姓": "胡"},
    "壬申": {"欠债_num": 42000, "看经卷": 14, "纳库": 49, "曹官姓": "王"},
    "乙卯": {"欠债_num": 80000, "看经卷": 27, "纳库": 48, "曹官姓": "柳"},
    "丁卯": {"欠债_num": 23000, "看经卷": 8, "纳库": 46, "曹官姓": "许"},
    "己卯": {"欠债_num": 80000, "看经卷": 27, "纳库": 26, "曹官姓": "宋"},
    "辛卯": {"欠债_num": 80000, "看经卷": 27, "纳库": 4, "曹官姓": "张"},
    "癸卯": {"欠债_num": 36000, "看经卷": 11, "纳库": 27, "曹官姓": "王"},
    "乙酉": {"欠债_num": 40000, "看经卷": 14, "纳库": 54, "曹官姓": "安"},
    "丁酉": {"欠债_num": 170000, "看经卷": 57, "纳库": 29, "曹官姓": "胡"},
    "己酉": {"欠债_num": 90000, "看经卷": 30, "纳库": 22, "曹官姓": "孫"},
    "辛酉": {"欠债_num": 27000, "看经卷": 9, "纳库": 15, "曹官姓": "丁"},
    "癸酉": {"欠债_num": 50000, "看经卷": 17, "纳库": 12, "曹官姓": "申"},
    "甲辰": {"欠债_num": 29000, "看经卷": 10, "纳库": 19, "曹官姓": "董"},
    "丙辰": {"欠债_num": 32000, "看经卷": 11, "纳库": 33, "曹官姓": "贾"},
    "戊辰": {"欠债_num": 18000, "看经卷": 18, "纳库": 14, "曹官姓": "冯"},
    "庚辰": {"欠债_num": 19000, "看经卷": 19, "纳库": 24, "曹官姓": "刘"},
    "壬辰": {"欠债_num": 15000, "看经卷": 15, "纳库": 1, "曹官姓": "赵"},
    "甲戌": {"欠债_num": 27000, "看经卷": 9, "纳库": 17, "曹官姓": "井"},
    "丙戌": {"欠债_num": 8000, "看经卷": 27, "纳库": 35, "曹官姓": "左"},
    "戊戌": {"欠债_num": 42000, "看经卷": 14, "纳库": 36, "曹官姓": "晋"},
    "庚戌": {"欠债_num": 11000, "看经卷": 37, "纳库": 2, "曹官姓": "辛"},
    "壬戌": {"欠债_num": 72000, "看经卷": 24, "纳库": 40, "曹官姓": "彭"},
    "乙巳": {"欠债_num": 90000, "看经卷": 30, "纳库": 41, "曹官姓": "杨"},
    "丁巳": {"欠债_num": 72000, "看经卷": 24, "纳库": 16, "曹官姓": "程"},
    "己巳": {"欠债_num": 72000, "看经卷": 24, "纳库": 30, "曹官姓": "曹"},
    "辛巳": {"欠债_num": 57000, "看经卷": 19, "纳库": 37, "曹官姓": "高"},
    "癸巳": {"欠债_num": 39000, "看经卷": 13, "纳库": 50, "曹官姓": "牛"},
    "乙亥": {"欠债_num": 48000, "看经卷": 16, "纳库": 42, "曹官姓": "成"},
    "丁亥": {"欠债_num": 39000, "看经卷": 13, "纳库": 47, "曹官姓": "吉"},
    "己亥": {"欠债_num": 72000, "看经卷": 24, "纳库": 55, "曹官姓": "卞"},
    "辛亥": {"欠债_num": 71000, "看经卷": 24, "纳库": 46, "曹官姓": "卞"},
    "癸亥": {"欠债_num": 75000, "看经卷": 25, "纳库": 43, "曹官姓": "仇"},
}

month_gan_data = {
    "甲": 10000, "乙": 20000, "丙": 30000, "丁": 40000, "戊": 50000,
    "己": 60000, "庚": 70000, "辛": 80000, "壬": 90000, "癸": 100000,
}

lunar_month_data = {i: i*1000 for i in range(1,13)}
lunar_day_data = {i: i*1000 for i in range(1,31)}

hour_gan_data = {
    "甲": 10000, "乙": 20000, "丙": 30000, "丁": 40000, "戊": 50000,
    "己": 6000, "庚": 7000, "辛": 8000, "壬": 9000, "癸": 12000,
}

hour_zhi_data = {
    "子":1000,"丑":2000,"寅":3000,"卯":4000,"辰":5000,"巳":6000,
    "午":7000,"未":8000,"申":9000,"酉":10000,"戌":11000,"亥":12000
}

HOUR_MAX_FIX = 61000

# ====================== 批量Excel处理功能 ======================
class ExcelProcessor:
    def __init__(self, parent):
        self.parent = parent
        self.frame = ttk.Frame(parent)
        self.frame.pack(fill="both", expand=True)
        
        self.file_path = ""
        
        # 标题
        ttk.Label(self.frame, text="📊 善信信息批量处理", font=("微软雅黑",12,"bold")).pack(pady=10)
        
        # 文件选择区域
        file_frame = ttk.LabelFrame(self.frame, text="文件操作", padding=10)
        file_frame.pack(pady=5, fill="x", padx=20)
        
        self.file_label = ttk.Label(file_frame, text="未选择文件", width=50)
        self.file_label.grid(row=0, column=0, padx=5, columnspan=2)
        
        ttk.Button(file_frame, text="选择Excel文件", command=self.select_file).grid(row=1, column=0, padx=5, pady=5)
        ttk.Button(file_frame, text="创建标准模板", command=self.create_template).grid(row=1, column=1, padx=5, pady=5)
        
        # 模板说明 - 自动换行
        tip1 = ttk.Label(self.frame, text="📋 输入模板表头：序号、姓名、性别、年干支、月干、阴历月份（阿拉伯数字）、阴历日（阿拉伯数字）、时干、时支", font=("微软雅黑",9), wraplength=850)
        tip1.pack(pady=2)
        tip2 = ttk.Label(self.frame, text="⚠️ 时辰未知请填：未知，程序会自动按最高值61000贯计算", font=("微软雅黑",9,"bold"))
        tip2.pack(pady=2)
        
        # 转换按钮
        btn_frame = ttk.Frame(self.frame)
        btn_frame.pack(pady=15)
        ttk.Button(btn_frame, text="✅ 开始转换并导出Excel", command=self.process_excel, width=25).pack()
        
        # 日志
        log_frame = ttk.LabelFrame(self.frame, text="处理日志", padding=10)
        log_frame.pack(pady=5, padx=15, fill="both", expand=True)
        self.log_text = scrolledtext.ScrolledText(log_frame, height=8, width=70, font=("微软雅黑",9))
        self.log_text.pack(fill="both", expand=True)
        
    def select_file(self):
        path = filedialog.askopenfilename(filetypes=[("Excel文件", "*.xlsx")])
        if path:
            self.file_path = path
            self.file_label.config(text=os.path.basename(path))
            self.log("已选择文件：" + os.path.basename(path))
            
    def create_template(self):
        """创建标准Excel模板"""
        headers = ["序号", "姓名", "性别", "年干支", "月干", "阴历月份（阿拉伯数字）", "阴历日（阿拉伯数字）", "时干", "时支"]
        template_df = pd.DataFrame(columns=headers)
        save_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel文件", "*.xlsx")],
            initialfile="受生债批量处理模板.xlsx"
        )
        if save_path:
            template_df.to_excel(save_path, index=False)
            self.log(f"✅ 模板已创建：{os.path.basename(save_path)}")
            messagebox.showinfo("创建成功", f"标准模板已创建到：\n{save_path}")
            
    def log(self, msg):
        self.log_text.insert(tk.END, msg + "\n")
        self.log_text.see(tk.END)
        
    def process_excel(self):
        if not self.file_path:
            messagebox.showwarning("提示","请先选择Excel文件！")
            return
            
        try:
            df = pd.read_excel(self.file_path)
            required = ["序号","姓名","性别","年干支","月干","阴历月份（阿拉伯数字）","阴历日（阿拉伯数字）","时干","时支"]
            for col in required:
                if col not in df.columns:
                    messagebox.showerror("错误",f"模板错误：缺少列【{col}】")
                    return
                    
            self.log("开始计算 " + str(len(df)) + " 条善信数据...")
            results = []
            
            for idx, row in df.iterrows():
                year_pillar = str(row["年干支"]).strip()
                month_gan = str(row["月干"]).strip()
                lunar_m = int(row["阴历月份（阿拉伯数字）"])
                lunar_d = int(row["阴历日（阿拉伯数字）"])
                h_gan = str(row["时干"]).strip()
                hour_val = str(row["时支"]).strip()
                
                # 基础计算
                a = year_data[year_pillar]["欠债_num"]
                b = month_gan_data.get(month_gan, 0)
                c = lunar_month_data.get(lunar_m, 0)
                d = lunar_day_data.get(lunar_d, 0)
                fixed = a + b + c + d
                
                # 时辰处理
                if "未知" in hour_val or hour_val.strip() == "":
                    e = 0
                    f = HOUR_MAX_FIX
                else:
                    e = hour_gan_data.get(h_gan, 0)
                    f = hour_zhi_data.get(hour_val, 0)
                    
                total = fixed + e + f
                caoguan = year_data[year_pillar]["曹官姓"]
                naku = year_data[year_pillar]["纳库"]
                jing = year_data[year_pillar]["看经卷"]
                
                results.append([total, caoguan, naku, jing])
                self.log(f"→ {row['姓名']} 计算完成")
                
            # 写入结果
            res_df = pd.DataFrame(results, columns=["受生债总额(贯)", "曹官", "库房号", "需看经卷数"])
            out_df = pd.concat([df, res_df], axis=1)
            
            save_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel文件", "*.xlsx")])
            if save_path:
                out_df.to_excel(save_path, index=False)
                self.log("✅ 全部完成！文件已保存")
                messagebox.showinfo("成功", f"处理完成！\n共 {len(df)} 人")
                
        except Exception as e:
            messagebox.showerror("处理失败", str(e))
            self.log("错误：" + str(e))

# ====================== 原手动计算功能 ======================
def show_data_table():
    table_win = tk.Toplevel(root)
    table_win.title("受生债完整对照表")
    table_win.geometry("900x700")
    txt = scrolledtext.ScrolledText(table_win, font=("微软雅黑", 10))
    txt.pack(fill="both", expand=True, padx=10, pady=10)
    
    txt.insert(tk.END, "="*60 + "\n【年柱 欠债/经卷/纳库/曹官】\n" + "="*60 + "\n")
    for y, d in year_data.items():
        txt.insert(tk.END, f"{y}：欠债{d['欠债_num']:,}贯 | 经卷{d['看经卷']}卷 | 纳库第{d['纳库']}库 | 曹官{d['曹官姓']}姓\n")
    
    txt.insert(tk.END, "\n【月干】【阴历月份】【日期】【时干】【时支】\n")
    txt.config(state=tk.DISABLED)

def calculate_manual():
    year_pillar = entry_year.get().strip()
    month_pillar = entry_month.get().strip()
    day_pillar = entry_day.get().strip()
    hour_pillar = entry_hour.get().strip()
    
    try:
        lunar_month = int(entry_lunar_month.get().strip())
        lunar_day = int(entry_lunar_day.get().strip())
    except:
        messagebox.showerror("错误","阴历月日必须是数字！")
        return

    if not all([year_pillar, month_pillar]):
        messagebox.showerror("错误","年柱、月柱不能为空！")
        return
    if year_pillar not in year_data:
        messagebox.showerror("错误",f"年柱【{year_pillar}】不在表格内！")
        return

    month_gan = month_pillar[0]
    a = year_data[year_pillar]["欠债_num"]
    b = month_gan_data.get(month_gan, 0)
    c = lunar_month_data.get(lunar_month, 0)
    d = lunar_day_data.get(lunar_day, 0)
    fixed_sum = a + b + c + d

    is_hour_empty = not hour_pillar or len(hour_pillar) < 2
    info = year_data[year_pillar]

    result_text.delete(1.0, tk.END)
    result_text.insert(tk.END, "="*50 + "\n           受生债计算明细\n" + "="*50 + "\n\n")
    result_text.insert(tk.END, f"① 年柱【{year_pillar}】：{a:,} 贯\n")
    result_text.insert(tk.END, f"② 月干【{month_gan}】：{b:,} 贯\n")
    result_text.insert(tk.END, f"③ 阴历{lunar_month}月：{c:,} 贯\n")
    result_text.insert(tk.END, f"④ 阴历{lunar_day}日：{d:,} 贯\n")

    if is_hour_empty:
        total = fixed_sum + HOUR_MAX_FIX
        result_text.insert(tk.END, "⑤ 时柱：【时辰不详，按61000贯核算】\n\n")
        result_text.insert(tk.END, "="*30 + f"\n✅ 总阴债：{total:,} 贯\n" + "="*30 + "\n\n")
    else:
        hg = hour_pillar[0]
        hz = hour_pillar[1]
        e = hour_gan_data.get(hg, 0)
        f = hour_zhi_data.get(hz, 0)
        total = fixed_sum + e + f
        result_text.insert(tk.END, f"⑤ 时干【{hg}】：{e:,} 贯\n⑥ 时支【{hz}】：{f:,} 贯\n\n")
        result_text.insert(tk.END, "="*30 + f"\n✅ 总阴债：{total:,} 贯\n" + "="*30 + "\n\n")

    result_text.insert(tk.END, f"📖 经卷：{info['看经卷']}卷\n🏦 库房：第{info['纳库']}库\n👤 曹官：{info['曹官姓']}姓\n")

# ====================== 主界面 ======================
root = tk.Tk()
root.title("受生债计算器 | 玉清宫专用版")
root.geometry("900x700")

# 选项卡
notebook = ttk.Notebook(root)
notebook.pack(fill="both", expand=True)

# 手动计算页
manual_frame = ttk.Frame(notebook)
notebook.add(manual_frame, text="🧮 手动单条计算")

# 批量处理页
excel_frame = ttk.Frame(notebook)
notebook.add(excel_frame, text="📊 批量Excel转换")
processor = ExcelProcessor(excel_frame)

# ---------- 手动界面：每行一个输入项 ----------
style = ttk.Style()
style.configure("TLabel", font=("微软雅黑", 11))

frame_input = ttk.LabelFrame(manual_frame, text="输入信息", padding=10)
frame_input.pack(fill="x", padx=15, pady=5)

# 每行一个输入项
ttk.Label(frame_input, text="年柱（干支）：").pack(anchor="w", pady=2)
entry_year = ttk.Entry(frame_input, width=30)
entry_year.pack(anchor="w", pady=2)

ttk.Label(frame_input, text="月柱（干支）：").pack(anchor="w", pady=2)
entry_month = ttk.Entry(frame_input, width=30)
entry_month.pack(anchor="w", pady=2)

ttk.Label(frame_input, text="日柱（干支）：").pack(anchor="w", pady=2)
entry_day = ttk.Entry(frame_input, width=30)
entry_day.pack(anchor="w", pady=2)

ttk.Label(frame_input, text="时柱（干支）：").pack(anchor="w", pady=2)
entry_hour = ttk.Entry(frame_input, width=30)
entry_hour.pack(anchor="w", pady=2)

ttk.Label(frame_input, text="阴历月份（阿拉伯数字）：").pack(anchor="w", pady=2)
entry_lunar_month = ttk.Entry(frame_input, width=30)
entry_lunar_month.pack(anchor="w", pady=2)

ttk.Label(frame_input, text="阴历日期（阿拉伯数字）：").pack(anchor="w", pady=2)
entry_lunar_day = ttk.Entry(frame_input, width=30)
entry_lunar_day.pack(anchor="w", pady=2)

# 按钮
btn_frame = ttk.Frame(manual_frame)
btn_frame.pack(pady=8)
ttk.Button(btn_frame, text="🧮 计算", command=calculate_manual).grid(row=0, column=0, padx=10)
ttk.Button(btn_frame, text="📋 查看对照表", command=show_data_table).grid(row=0, column=1, padx=10)

# 结果框
frame_result = ttk.LabelFrame(manual_frame, text="计算结果", padding=10)
frame_result.pack(fill="both", expand=True, padx=15, pady=5)
result_text = scrolledtext.ScrolledText(frame_result, font=("微软雅黑", 11))
result_text.pack(fill="both", expand=True)

# 提示
ttk.Label(manual_frame, text="时柱留空 = 自动按61000贯最高值计算", font=("微软雅黑",9,"bold")).pack(pady=2)

root.mainloop()