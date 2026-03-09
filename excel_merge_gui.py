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
        self.root.geometry("650x350")
        self.root.resizable(False, False)

        # 路径变量
        self.folder_path = tk.StringVar()
        self.output_path = tk.StringVar()

        # 界面布局
        main_frame = ttk.Frame(root, padding=20)
        main_frame.pack(fill=tk.BOTH, expand=True)

        # 选择文件夹
        ttk.Label(main_frame, text="1. 选择待合并Excel文件夹：", font=("微软雅黑", 10)).grid(row=0, column=0, sticky="w", pady=10)
        ttk.Entry(main_frame, textvariable=self.folder_path, width=50).grid(row=0, column=1, padx=5)
        ttk.Button(main_frame, text="选择文件夹", command=self.select_folder).grid(row=0, column=2)

        # 选择输出文件
        ttk.Label(main_frame, text="2. 选择输出保存位置：", font=("微软雅黑", 10)).grid(row=1, column=0, sticky="w", pady=10)
        ttk.Entry(main_frame, textvariable=self.output_path, width=50).grid(row=1, column=1, padx=5)
        ttk.Button(main_frame, text="保存", command=self.select_output).grid(row=1, column=2)

        # 功能说明
        ttk.Label(main_frame, text="✅ 第二列(商户编号)=文本格式 | ✅ 交易日期仅排序不修改", foreground="green", font=("微软雅黑", 9)).grid(row=2, column=0, columnspan=3, sticky="w", pady=5)
        
        # 合并按钮
        self.merge_btn = ttk.Button(main_frame, text="✅ 开始合并Excel", command=self.start_merge, width=25)
        self.merge_btn.grid(row=3, column=1, pady=15)

        # 日志框
        ttk.Label(main_frame, text="执行日志：").grid(row=4, column=0, sticky="w", pady=5)
        self.log_text = tk.Text(main_frame, height=8, width=70)
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
            self.log(f"💾 输出文件：{path}")

    def start_merge(self):
        folder = self.folder_path.get()
        output = self.output_path.get()
        if not folder or not output:
            messagebox.showerror("错误", "请先选择文件夹和输出路径！")
            return

        self.merge_btn.config(state=tk.DISABLED, text="合并中...")
        self.log("="*50)
        self.log("开始合并Excel文件...")

        try:
            self.merge_excel(folder, output)
            self.log("✅ 合并完成！")
            messagebox.showinfo("成功", "Excel合并完成！\n文件已保存到：" + output)
        except Exception as e:
            self.log(f"❌ 合并失败：{str(e)}")
            messagebox.showerror("失败", f"合并出错：{str(e)}")
        finally:
            self.merge_btn.config(state=tk.NORMAL, text="✅ 开始合并Excel")

    def merge_excel(self, folder_path, output_file):
        excel_ext = (".xlsx", ".xls")
        merged_data = []
        has_header = False

        # 遍历文件夹
        for filename in os.listdir(folder_path):
            if not filename.endswith(excel_ext):
                continue
            file_path = os.path.join(folder_path, filename)

            try:
                self.log(f"正在读取：{filename}")
                # 核心：第二列强制文本格式（商户编号）
                df = pd.read_excel(file_path, dtype={1: str})

                if df.empty:
                    self.log(f"⚠️ {filename} 是空文件，已跳过")
                    continue

                # 清理商户编号空格
                if df.shape[1] >= 2:
                    df.iloc[:, 1] = df.iloc[:, 1].astype(str).str.strip()

                # 合并数据
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

        # 合并所有数据
        final_df = pd.concat(merged_data, ignore_index=True)
        self.log("🔗 数据合并完成，开始排序...")

        # 排序：交易日期仅转换排序，不修改原始格式
        try:
            original_date = final_df["交易日期"].copy()
            final_df["排序日期"] = pd.to_datetime(final_df["交易日期"], errors="coerce")
            final_df = final_df.sort_values(by="排序日期", ascending=False)
            final_df["交易日期"] = original_date
            final_df = final_df.drop(columns=["排序日期"])
            self.log("📅 按交易日期倒序排序完成")
        except KeyError:
            self.log("⚠️ 未找到交易日期列，跳过排序")
        except Exception as e:
            self.log(f"⚠️ 排序异常：{str(e)}，跳过排序")

        # 保存文件
        final_df.to_excel(output_file, index=False, engine="openpyxl")

if __name__ == "__main__":
    root = tk.Tk()
    ExcelMergeGUI(root)
    root.mainloop()
