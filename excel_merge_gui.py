import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import os
import warnings
warnings.filterwarnings("ignore")

class ExcelMergeGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel合并工具（严格保留第一个文件格式 · 防科学计数法）")
        self.root.geometry("700x420")
        self.root.resizable(False, False)

        self.folder_path = tk.StringVar()
        self.output_path = tk.StringVar()

        main_frame = ttk.Frame(root, padding=25)
        main_frame.pack(fill=tk.BOTH, expand=True)

        # 标题说明
        ttk.Label(main_frame, text="📊 批量合并 Excel，格式 = 第一个文件格式", font=("微软雅黑", 12, "bold")).grid(row=0, column=0, columnspan=3, p=10)

        # 选择文件夹
        ttk.Label(main_frame, text="1. 选择待合并文件夹：", font=("微软雅黑", 10)).grid(row=1, column=0, sticky="w", pady=8)
        ttk.Entry(main_frame, textvariable=self.folder_path, width=58).grid(row=1, column=1, padx=5)
        ttk.Button(main_frame, text="选择", command=self.select_folder).grid(row=1, column=2)

        # 输出文件
        ttk.Label(main_frame, text="2. 选择输出文件：", font=("微软雅黑", 10)).grid(row=2, column=0, sticky="w", pady=8)
        ttk.Entry(main_frame, textvariable=self.output_path, width=58).grid(row=2, column=1, padx=5)
        ttk.Button(main_frame, text="保存", command=self.select_output).grid(row=2, column=2)

        # 提示
        ttk.Label(main_frame, text="✅ 以第一个文件格式为准 | ✅ 带'文本列保持文本 | ✅ 无科学计数法", foreground="green", font=("微软雅黑", 9)).grid(row=3, column=0, columnspan=3, sticky="w", pady=5)

        # 合并按钮
        self.merge_btn = ttk.Button(main_frame, text="✅ 开始合并", command=self.start_merge, width=25)
        self.merge_btn.grid(row=4, column=1, pady=15)

        # 日志
        ttk.Label(main_frame, text="执行日志：").grid(row=5, column=0, sticky="w", pady=5)
        self.log_text = tk.Text(main_frame, height=10, width=80)
        self.log_text.grid(row=6, column=0, columnspan=3, pady=5)

    def log(self, msg):
        self.log_text.insert(tk.END, msg + "\n")
        self.log_text.see(tk.END)
        self.root.update()

    def select_folder(self):
        path = filedialog.askdirectory()
        if path:
            self.folder_path.set(path)
            self.log(f"已选择：{path}")

    def select_output(self):
        path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel 文件", "*.xlsx")])
        if path:
            self.output_path.set(path)
            self.log(f"输出到：{path}")

    def start_merge(self):
        folder = self.folder_path.get()
        out = self.output_path.get()
        if not folder or not out:
            messagebox.showerror("错误", "请先选择文件夹和输出路径！")
            return

        self.merge_btn.config(state=tk.DISABLED, text="合并中...")
        try:
            self.merge_excel(folder, out)
            self.log("\n🎉 合并完成！")
            messagebox.showinfo("成功", f"合并完成！\n格式严格匹配第一个文件\n无科学计数法！")
        except Exception as e:
            self.log(f"❌ 失败：{str(e)}")
            messagebox.showerror("失败", str(e))
        finally:
            self.merge_btn.config(state=tk.NORMAL, text="✅ 开始合并")

    def merge_excel(self, folder, output):
        excel_ext = (".xlsx", ".xls")
        files = [f for f in os.listdir(folder) if f.endswith(excel_ext)]
        if not files:
            raise Exception("未找到 Excel 文件")

        merged = []
        self.log(f"📌 标准格式来源：{files[0]}")

        # 第一个文件：文本模式读取（保留'）
        first_path = os.path.join(folder, files[0])
        df_first = pd.read_excel(first_path, dtype=str)
        cols = df_first.columns.tolist()
        merged.append(df_first)
        self.log(f"✅ 标准列数：{len(cols)} 列")

        # 合并其余文件
        for fname in files[1:]:
            self.log(f"正在合并：{fname}")
            fp = os.path.join(folder, fname)
            df = pd.read_excel(fp, dtype=str)
            df = df.reindex(columns=cols)
            merged.append(df)

        final = pd.concat(merged, ignore_index=True)
        final = final.fillna("")

        # 保存（完全保持文本格式）
        with pd.ExcelWriter(output, engine="openpyxl") as w:
            final.to_excel(w, index=False)

if __name__ == "__main__":
    root = tk.Tk()
    ExcelMergeGUI(root)
    root.mainloop()
