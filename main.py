import sys
import tkinter as tk
from tkinter import ttk
from tkinter import filedialog
from tkinter import messagebox
import pandas as pd

class MergeExcelApp:
    def __init__(self, root):
        self.root = root
        self.root.title('Excel文件合并工具')
        self.root.geometry('500x300')  # 减小窗口高度
        
        # 设置窗口最小尺寸
        self.root.minsize(400, 250)  # 减小最小高度
        
        # 创建主框架并添加内边距
        main_frame = ttk.Frame(root, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 状态标签
        self.status_var = tk.StringVar(value='未选择文件')
        status_label = ttk.Label(main_frame, textvariable=self.status_var)
        status_label.pack(fill=tk.X, pady=(0, 10))
        
        # 表头选项
        self.header_var = tk.BooleanVar(value=True)
        header_check = ttk.Checkbutton(
            main_frame, 
            text='包含表头行（勾选则按表头名合并，否则按列顺序合并）',
            variable=self.header_var
        )
        header_check.pack(fill=tk.X, pady=(0, 10))
        
        # 按钮样式
        style = ttk.Style()
        style.configure('Action.TButton', padding=10)
        
        # 选择文件按钮
        select_btn = ttk.Button(
            main_frame,
            text='选择Excel文件',
            command=self.select_files,
            style='Action.TButton'
        )
        select_btn.pack(fill=tk.X, pady=(0, 10))
        
        # 合并按钮
        merge_btn = ttk.Button(
            main_frame,
            text='合并文件',
            command=self.merge_files,
            style='Action.TButton'
        )
        merge_btn.pack(fill=tk.X)
        
        # 存储选中的文件
        self.selected_files = []

    def select_files(self):
        files = filedialog.askopenfilenames(
            title='选择Excel文件',
            filetypes=[('Excel文件', '*.xlsx *.xls')]
        )
        if files:
            self.selected_files = files
            self.status_var.set(f'已选择 {len(files)} 个文件')

    def merge_files(self):
        if not self.selected_files:
            messagebox.showwarning('警告', '请先选择要合并的Excel文件！')
            return

        try:
            has_header = self.header_var.get()
            # 读取第一个文件以获取列信息
            first_df = pd.read_excel(
                self.selected_files[0],
                header=0 if has_header else None,
                names=None if has_header else ['']*pd.read_excel(self.selected_files[0], nrows=1).shape[1]  # 如果没有表头，使用空字符串作为列名
            )
            
            # 合并所有文件
            all_dfs = []
            for file in self.selected_files:
                df = pd.read_excel(
                    file,
                    header=0 if has_header else None,
                    names=None if has_header else ['']*pd.read_excel(file, nrows=1).shape[1]  # 如果没有表头，使用空字符串作为列名
                )
                
                if has_header:
                    # 按表头名合并时检查列名
                    missing_cols = set(first_df.columns) - set(df.columns)
                    extra_cols = set(df.columns) - set(first_df.columns)
                    if missing_cols or extra_cols:
                        error_msg = []
                        if missing_cols:
                            error_msg.append(f'缺少列：{", ".join(missing_cols)}')
                        if extra_cols:
                            error_msg.append(f'多余列：{", ".join(extra_cols)}')
                        messagebox.showerror('错误', f'文件 {file} 的列名不匹配！\n' + '\n'.join(error_msg))
                        return
                    
                    # 按表头名对齐列
                    df = df[first_df.columns]
                else:
                    # 不包含表头时检查列数
                    if len(df.columns) != len(first_df.columns):
                        messagebox.showerror('错误', f'文件 {file} 的列数与其他文件不匹配！')
                        return
                
                all_dfs.append(df)

            # 合并数据框
            merged_df = pd.concat(all_dfs, ignore_index=True)

            # 保存合并后的文件
            save_path = filedialog.asksaveasfilename(
                title='保存合并后的文件',
                defaultextension='.xlsx',
                filetypes=[('Excel文件', '*.xlsx')]
            )

            if save_path:
                # 显示进度窗口
                progress_window = tk.Toplevel(self.root)
                progress_window.title('保存中')
                progress_window.geometry('300x100')
                progress_window.transient(self.root)
                progress_window.grab_set()
                
                progress_label = ttk.Label(
                    progress_window,
                    text='正在保存文件，请稍候...'
                )
                progress_label.pack(pady=20)
                
                # 更新UI
                self.root.update()
                
                # 保存文件，如果没有表头则不写入列名
                merged_df.to_excel(save_path, index=False, header=has_header)
                
                # 关闭进度窗口
                progress_window.destroy()
                
                messagebox.showinfo('成功', '文件合并完成！')

        except Exception as e:
            messagebox.showerror('错误', f'处理文件时出错：{str(e)}')

def main():
    root = tk.Tk()
    app = MergeExcelApp(root)
    # 设置窗口图标
    try:
        root.iconbitmap('icon.ico')
    except:
        pass  # 如果没有图标文件就忽略
    root.mainloop()

if __name__ == "__main__":
    main()
