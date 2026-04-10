import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog, Toplevel
import pandas as pd
import os
import warnings
warnings.filterwarnings("ignore")

class ExcelMergeGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel合并工具（保留原格式+可选排序）")
        self.root.geometry("720x450")
        self.root.resizable(False, False)

        self.folder_path = tk.StringVar()
        self.output_path = tk.StringVar()
        self.sort_column = None
        self.sort_ascending = None
        self.first_columns = []

        main_frame = ttk.Frame(root, padding=25)
        main_frame.pack(fill=tk.BOTH, expand=True)

        ttk.Label(main_frame, text="📊 Excel 合并工具（严格保留第一个文件格式｜可选排序）", font=("微软雅黑", 12, "bold")).grid(row=0, column=0, columnspan=3, pady=10)

        ttk.Label(main_frame, text="1. 选择待合并文件夹：", font=("微软雅黑", 10)).grid(row=1, column=0, sticky="w", pady=8)
        ttk.Entry(main_frame, textvariable=self.folder_path, width=60).grid(row=1, column=1, padx=5)
        ttk.Button(main_frame, text="选择", command=self.select_folder).grid(row=1, column=2)

        ttk.Label(main_frame, text="2. 选择输出文件：", font=("微软雅黑", 10)).grid(row=2, column=0, sticky="w", pady=8)
        ttk.Entry(main_frame, textvariable=self.output_path, width=60).grid(row=2, column=1, padx=5)
        ttk.Button(main_frame, text="保存", command=self.select_output).grid(row=2, column=2)

        ttk.Label(main_frame, text="✅ 以第一个文件格式为准｜✅ 防科学计数法｜✅ 可选排序", foreground="green", font=("微软雅黑", 9)).grid(row=3, column=0, columnspan=3, sticky="w", pady=5)

        self.merge_btn = ttk.Button(main_frame, text="✅ 开始合并", command=self.start_merge, width=25)
        self.merge_btn.grid(row=4, column=1, pady=15)

        ttk.Label(main_frame, text="执行日志：").grid(row=5, column=0, sticky="w", pady=5)
        self.log_text = tk.Text(main_frame, height=12, width=83)
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

    def get_first_file_columns(self):
        folder = self.folder_path.get()
        excel_ext = (".xlsx", ".xls")
        files = [f for f in os.listdir(folder) if f.endswith(excel_ext)]
        if not files:
            raise Exception("未找到Excel文件")

        for f in files:
            p = os.path.join(folder, f)
            try:
                df = pd.read_excel(p, nrows=1, dtype=str)
                if not df.empty:
                    self.first_columns = df.columns.tolist()
                    self.log(f"📌 第一个有效文件：{f}")
                    self.log(f"📋 识别到列名：{self.first_columns}")
                    return
            except:
                continue
        raise Exception("没有有效数据文件")

    def choose_sort_option(self):
        win = Toplevel(self.root)
        win.title("选择排序")
        win.geometry("450x220")
        win.resizable(False, False)
        win.transient(self.root)
        win.grab_set()

        ttk.Label(win, text="请选择排序方式", font=("微软雅黑",11,"bold")).pack(pady=10)

        frame = ttk.Frame(win)
        frame.pack(pady=5)

        sort_var = tk.StringVar()
        col_combo = ttk.Combobox(frame, textvariable=sort_var, values=self.first_columns, width=30, state="readonly")
        col_combo.pack(side=tk.LEFT, padx=5)
        col_combo.current(0)

        order_var = tk.StringVar(value="倒序")
        ttk.Radiobutton(frame, text="正序", variable=order_var, value="正序").pack(side=tk.LEFT, padx=3)
        ttk.Radiobutton(frame, text="倒序", variable=order_var, value="倒序").pack(side=tk.LEFT, padx=3)

        def confirm():
            self.sort_column = sort_var.get()
            self.sort_ascending = (order_var.get() == "正序")
            win.destroy()

        def skip():
            self.sort_column = None
            self.sort_ascending = None
            win.destroy()

        btn_frame = ttk.Frame(win)
        btn_frame.pack(pady=15)
        ttk.Button(btn_frame, text="使用此列排序", command=confirm).pack(side=tk.LEFT, pad=10)
        ttk.Button(btn_frame, text="跳过不排序", command=skip).pack(side=tk.LEFT, pad=10)

        self.root.wait_window(win)

    def start_merge(self):
        folder = self.folder_path.get()
        out = self.output_path.get()
        if not folder or not out:
            messagebox.showerror("错误", "请先选择文件夹和输出路径！")
            return

        self.merge_btn.config(state=tk.DISABLED, text="处理中...")
        try:
            self.log("\n==== 开始读取第一个文件列名 ====")
            self.get_first_file_columns()

            self.log("\n==== 请选择排序方式 ====")
            self.choose_sort_option()

            if self.sort_column:
                self.log(f"✅ 排序方式：按【{self.sort_column}】{'正序' if self.sort_ascending else '倒序'}")
            else:
                self.log("✅ 跳过排序，直接合并")

            self.log("\n==== 开始合并文件 ====")
            self.merge_excel(folder, out)

            self.log("\n🎉 合并完成！")
            messagebox.showinfo("成功", "合并完成！")
        except Exception as e:
            self.log(f"❌ 失败：{str(e)}")
            messagebox.showerror("失败", str(e))
        finally:
            self.merge_btn.config(state=tk.NORMAL, text="✅ 开始合并")

    def merge_excel(self, folder, output):
        excel_ext = (".xlsx", ".xls")
        files = [f for f in os.listdir(folder) if f.endswith(excel_ext)]
        merged = []
        target_cols = None

        for fname in files:
            fp = os.path.join(folder, fname)
            self.log(f"正在读取：{fname}")
            df = pd.read_excel(fp, dtype=str)
            if target_cols is None:
                target_cols = df.columns.tolist()
            df = df.reindex(columns=target_cols)
            merged.append(df)

        final = pd.concat(merged, ignore_index=True)
        final = final.fillna("")

        if self.sort_column and self.sort_column in final.columns:
            final = final.sort_values(by=self.sort_column, ascending=self.sort_ascending)

        with pd.ExcelWriter(output, engine="openpyxl") as w:
            final.to_excel(w, index=False)

if __name__ == "__main__":
    root = tk.Tk()
    ExcelMergeGUI(root)
    root.mainloop()
