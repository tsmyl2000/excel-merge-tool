import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import os
import warnings
warnings.filterwarnings("ignore")

class ExcelMergeGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel批量合并工具 - Windows专用")
        self.root.geometry("680x380")
        self.root.resizable(False, False)

        self.folder_path = tk.StringVar()
        self.output_path = tk.StringVar()

        main_frame = ttk.Frame(root, padding=20)
        main_frame.pack(fill=tk.BOTH, expand=True)

        ttk.Label(main_frame, text="1. 选择待合并Excel文件夹：", font=("微软雅黑", 10)).grid(row=0, column=0, sticky="w", pady=10)
        ttk.Entry(main_frame, textvariable=self.folder_path, width=55).grid(row=0, column=1, padx=5)
        ttk.Button(main_frame, text="选择文件夹", command=self.select_folder).grid(row=0, column=2)

        ttk.Label(main_frame, text="2. 选择输出保存位置：", font=("微软雅黑", 10)).grid(row=1, column=0, sticky="w", pady=10)
        ttk.Entry(main_frame, textvariable=self.output_path, width=55).grid(row=1, column=1, padx=5)
        ttk.Button(main_frame, text="保存", command=self.select_output).grid(row=1, column=2)

        ttk.Label(main_frame, text="✅ 以第一个文件格式为准 | ✅ 文本保持文本 | ✅ 数字保持数字 | ✅ 无科学计数法", foreground="green", font=("微软雅黑", 9)).grid(row=2, column=0, columnspan=3, sticky="w", pady=5)

        self.merge_btn = ttk.Button(main_frame, text="✅ 开始合并Excel", command=self.start_merge, width=25)
        self.merge_btn.grid(row=3, column=1, pady=15)

        ttk.Label(main_frame, text="执行日志：").grid(row=4, column=0, sticky="w", pady=5)
        self.log_text = tk.Text(main_frame, height=9, width=75)
        self.log_text.grid(row=5, column=0, columnspan=3, pady=5)

    def log(self, msg):
        self.log_text.insert(tk.END, msg + "\n")
        self.log_text.see(tk.END)
        self.root.update()

    def select_folder(self):
        path = filedialog.askdirectory()
        if path:
            self.folder_path.set(path)
            self.log(f"📂 已选择文件夹：{path}")

    def select_output(self):
        path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel文件", "*.xlsx")])
        if path:
            self.output_path.set(path)
            self.log(f"💾 输出文件：{output_file}")

    def start_merge(self):
        folder = self.folder_path.get()
        output = self.output_path.get()
        if not folder or not output:
            messagebox.showerror("错误", "请先选择文件夹和输出路径！")
            return

        self.merge_btn.config(state=tk.DISABLED, text="合并中...")
        self.log("="*60)
        self.log("开始合并Excel文件...")

        try:
            self.merge_excel(folder, output)
            self.log("✅ 合并完成！严格按照第一个文件格式输出")
            messagebox.showinfo("成功", "Excel合并完成！\n已按第一个文件格式输出！")
        except Exception as e:
            self.log(f"❌ 合并失败：{str(e)}")
            messagebox.showerror("失败", f"合并出错：{str(e)}")
        finally:
            self.merge_btn.config(state=tk.NORMAL, text="✅ 开始合并Excel")

    def merge_excel(self, folder_path, output_file):
        excel_ext = (".xlsx", ".xls")
        merged_data = []
        has_header = False

        # 第一个文件的结构（列名 + 列数据类型）
        first_dtypes = None
        first_columns = None

        for filename in os.listdir(folder_path):
            if not filename.endswith(excel_ext):
                continue
            file_path = os.path.join(folder_path, filename)

            try:
                self.log(f"正在读取：{filename}")

                # --------------------------
                # 第一个文件：读取格式 + 记录类型
                # --------------------------
                if first_dtypes is None:
                    df = pd.read_excel(file_path, dtype=str)
                    first_columns = df.columns.tolist()
                    first_dtypes = df.dtypes
                    self.log(f"📌 以第一个文件格式为准：{filename}")

                # --------------------------
                # 后续文件：按第一个文件的列对齐
                # --------------------------
                else:
                    df = pd.read_excel(file_path, dtype=str)
                    df = df.reindex(columns=first_columns)

                # 清空NaN，避免格式异常
                df = df.fillna("")

                # 添加到合并列表
                if not has_header:
                    merged_data.append(df)
                    has_header = True
                else:
                    merged_data.append(df.iloc[1:])

            except Exception as e:
                self.log(f"❌ 读取失败：{filename} | 错误：{str(e)}")
                continue

        if not merged_data:
            raise Exception("未找到任何有效Excel文件！")

        # 合并
        final_df = pd.concat(merged_data, ignore_index=True)
        self.log("🔗 数据合并完成")

        # --------------------------
        # 排序（交易日期倒序）
        # --------------------------
        try:
            if "交易日期" in final_df.columns:
                self.log("📅 按交易日期倒序排序...")
                original_date = final_df["交易日期"].copy()
                final_df["__sort__"] = pd.to_datetime(original_date, errors="coerce")
                final_df = final_df.sort_values(by="__sort__", ascending=False)
                final_df["交易日期"] = original_date
                final_df = final_df.drop(columns=["__sort__"])
        except Exception as e:
            self.log(f"⚠️ 排序异常：{str(e)}，跳过排序")

        # --------------------------
        # 核心：恢复第一个文件的格式
        # --------------------------
        for col in final_df.columns:
            try:
                final_df[col] = final_df[col].astype(first_dtypes[col])
            except:
                pass

        # --------------------------
        # 输出Excel
        # --------------------------
        with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
            final_df.to_excel(writer, index=False)

        self.log(f"✅ 文件已保存：{output_file}")

if __name__ == "__main__":
    root = tk.Tk()
    ExcelMergeGUI(root)
    root.mainloop()
