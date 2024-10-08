import win32com.client
import os
from datetime import datetime
import tkinter as tk
import pytz
from tkinter import ttk, filedialog, messagebox
from tkcalendar import DateEntry
from PIL import Image, ImageTk

class OutlookApp:
    def __init__(self, master):
        self.master = master
        self.master.title("EmailMagic")
        self.master.geometry("800x700")  # Reduced size but still accommodates all elements
        self.outlook = None
        self.setup_ui()

    def setup_ui(self):
        self.master.configure(bg='#f0f0f0')
        style = ttk.Style()
        style.theme_use('clam')

        # 主框架设置
        main_frame = ttk.Frame(self.master, padding="5")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # 状态部分
        status_frame = ttk.LabelFrame(main_frame, text="状态", padding="5")
        status_frame.pack(fill=tk.X, pady=(0, 5))

        self.status_var = tk.StringVar(value="未连接")
        ttk.Label(status_frame, textvariable=self.status_var).pack(side=tk.LEFT)

        self.login_button = ttk.Button(status_frame, text="登录", command=self.login, style='Accent.TButton')
        self.login_button.pack(side=tk.LEFT, padx=(5, 0))

        self.account_label = ttk.Label(status_frame, text="")
        self.account_label.pack(side=tk.LEFT, padx=(5, 0))

        # 搜索部分
        search_frame = ttk.LabelFrame(main_frame, text="搜索选项", padding="5")
        search_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 5))

        # 文件夹选择
        folder_frame = ttk.Frame(search_frame)
        folder_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 5))
        
        tree_frame = ttk.Frame(folder_frame)
        tree_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        self.folder_tree = ttk.Treeview(tree_frame, height=4)
        self.folder_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.folder_tree.heading("#0", text="文件夹", anchor="w")

        folder_scrollbar = ttk.Scrollbar(tree_frame, orient="vertical", command=self.folder_tree.yview)
        folder_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.folder_tree.configure(yscrollcommand=folder_scrollbar.set)

        # 关键词和日期选择
        options_frame = ttk.Frame(search_frame)
        options_frame.pack(fill=tk.X, pady=(0, 5))

        ttk.Label(options_frame, text="关键词:").grid(row=0, column=0, sticky="w", pady=2)
        self.keyword_entry = ttk.Entry(options_frame, width=40)
        self.keyword_entry.grid(row=0, column=1, columnspan=3, sticky="ew", pady=2)

        ttk.Label(options_frame, text="开始日期:").grid(row=1, column=0, sticky="w", pady=2)
        self.start_date_entry = DateEntry(options_frame, width=15, background='darkblue', foreground='white', borderwidth=2)
        self.start_date_entry.grid(row=1, column=1, sticky="w", pady=2)

        ttk.Label(options_frame, text="结束日期:").grid(row=1, column=2, sticky="w", pady=2)
        self.end_date_entry = DateEntry(options_frame, width=15, background='darkblue', foreground='white', borderwidth=2)
        self.end_date_entry.grid(row=1, column=3, sticky="w", pady=2)

        options_frame.grid_columnconfigure(1, weight=1)
        options_frame.grid_columnconfigure(3, weight=1)

        # 下载选项
        download_frame = ttk.Frame(search_frame)
        download_frame.pack(fill=tk.X, pady=(0, 5))

        self.download_var = tk.BooleanVar()
        self.download_checkbox = ttk.Checkbutton(download_frame, text="下载附件", variable=self.download_var, command=self.toggle_download_path)
        self.download_checkbox.pack(side=tk.LEFT)

        self.download_path_label = ttk.Label(download_frame, text="下载路径:")
        self.download_path_entry = ttk.Entry(download_frame, width=30)
        self.download_path_button = ttk.Button(download_frame, text="浏览", command=self.browse_download_path)

        # 搜索按钮
        self.search_button = ttk.Button(search_frame, text="搜索", command=self.search, style='Accent.TButton')
        self.search_button.pack(pady=(0, 5))

        # 结果部分
        result_frame = ttk.LabelFrame(main_frame, text="搜索结果", padding="5")
        result_frame.pack(fill=tk.BOTH, expand=True)

        self.result_text = tk.Text(result_frame, height=15, width=80)
        self.result_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.result_text.config(state=tk.DISABLED)

        result_scrollbar = ttk.Scrollbar(result_frame, orient="vertical", command=self.result_text.yview)
        result_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.result_text.configure(yscrollcommand=result_scrollbar.set)

        # 工具提示和信息
        self.add_tooltips_and_info()

    def toggle_download_path(self):
        if self.download_var.get():
            self.download_path_label.pack(side=tk.LEFT, padx=(5, 0))
            self.download_path_entry.pack(side=tk.LEFT, padx=(2, 0), expand=True, fill=tk.X)
            self.download_path_button.pack(side=tk.LEFT, padx=(2, 0))
        else:
            self.download_path_label.pack_forget()
            self.download_path_entry.pack_forget()
            self.download_path_button.pack_forget()

    def browse_download_path(self):
        download_path = filedialog.askdirectory()
        if download_path:
            self.download_path_entry.delete(0, tk.END)
            self.download_path_entry.insert(0, download_path)

    def login(self):
        try:
            self.outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
            self.status_var.set("已连接")
            self.show_folders()
            
            account = self.outlook.CurrentUser.Address
            account_name = account.split('-')[-1]
            self.account_label.config(text=f"当前账户: {account_name}")
        except Exception as e:
            messagebox.showerror("错误", f"连接Outlook失败: {str(e)}")
            self.manual_login()

    def show_folders(self):
        self.folder_tree.delete(*self.folder_tree.get_children())
        root_folder = self.outlook.GetDefaultFolder(6)  # 6是收件箱的索引
        self.add_folder_to_tree("", root_folder)

    def add_folder_to_tree(self, parent, folder):
        folder_id = self.folder_tree.insert(parent, "end", text=folder.Name)
        for subfolder in folder.Folders:
            self.add_folder_to_tree(folder_id, subfolder)

    def search(self):
        keywords = [keyword.strip() for keyword in self.keyword_entry.get().split(',')]
        start_date = self.start_date_entry.get_date()
        end_date = self.end_date_entry.get_date()

        from datetime import datetime, time

        start_datetime = datetime.combine(start_date, time.min)
        end_datetime = datetime.combine(end_date, time.max)

        start_date_str = start_datetime.strftime('%m/%d/%Y %I:%M %p')
        end_date_str = end_datetime.strftime('%m/%d/%Y %I:%M %p')

        selected_folder = self.folder_tree.focus()
        if not selected_folder:
            messagebox.showerror("错误", "请选择要搜索的文件夹。")
            return

        folder_path = self.get_folder_path(selected_folder)
        folder = self.get_folder_by_path(folder_path)

        messages = folder.Items
        messages.Sort("[ReceivedTime]", True)

        date_filter = f"[ReceivedTime] >= '{start_date_str}' AND [ReceivedTime] <= '{end_date_str}'"
        messages = messages.Restrict(date_filter)

        if keywords:
            keyword_filter = " OR ".join([f"@SQL=\"urn:schemas:httpmail:subject\" LIKE '%{keyword}%'" for keyword in keywords])
            messages = messages.Restrict(keyword_filter)

        results = []
        for msg in messages:
            results.append(msg)

        self.display_results(results)

        if results and self.download_var.get():
            self.download_attachments(results)

    def display_results(self, results):
        self.result_text.config(state=tk.NORMAL)
        self.result_text.delete(1.0, tk.END)
        
        if results:
            self.result_text.insert(tk.END, f"找到 {len(results)} 封匹配的邮件:\n\n")
            for i, msg in enumerate(results, 1):
                self.result_text.insert(tk.END, f"--------------- 第{i}封 -------------\n")
                self.result_text.insert(tk.END, f"主题: {msg.Subject}\n")
                self.result_text.insert(tk.END, f"发件人: {msg.SenderName}\n")
                self.result_text.insert(tk.END, f"日期: {msg.ReceivedTime.strftime('%Y-%m-%d %H:%M:%S')}\n")
                attachments = [att.FileName for att in msg.Attachments if self.is_valid_attachment(att.FileName)]
                self.result_text.insert(tk.END, f"附件: {', '.join(attachments)}\n\n")
        else:
            self.result_text.insert(tk.END, "未找到标题中包含给定关键词的邮件。")
        
        self.result_text.config(state=tk.DISABLED)

    def get_folder_path(self, item):
        path = []
        while item:
            path.append(self.folder_tree.item(item)["text"])
            item = self.folder_tree.parent(item)
        return "/".join(reversed(path))

    def get_folder_by_path(self, path):
        folders = path.split("/")
        current_folder = self.outlook.GetDefaultFolder(6)  # 从收件箱开始
        for folder_name in folders[1:]:  # 跳过第一个元素（收件箱）
            current_folder = current_folder.Folders[folder_name]
        return current_folder

    def is_valid_attachment(self, filename):
        valid_extensions = ('.xlsx', '.xls', '.docx', '.doc', '.pdf', '.txt', '.csv')
        return filename.lower().endswith(valid_extensions)

    def download_attachments(self, messages):
        output_path = self.download_path_entry.get()
        if not output_path:
            messagebox.showerror("错误", "请选择下载路径。")
            return

        downloaded_files = []
        for message in messages:
            received_time = message.ReceivedTime.strftime('%Y%m%d%H%M%S')
            for attachment in message.Attachments:
                if self.is_valid_attachment(attachment.FileName):
                    file_name, file_extension = os.path.splitext(attachment.FileName)
                    new_file_name = f"{file_name}_{received_time}{file_extension}"
                    file_path = os.path.join(output_path, new_file_name)
                    attachment.SaveAsFile(file_path)
                    downloaded_files.append(file_path)

        if downloaded_files:
            messagebox.showinfo("下载完成", f"附件已保存到 {output_path}")
            self.result_text.config(state=tk.NORMAL)
            self.result_text.insert(tk.END, f"\n已下载的附件:\n")
            for file in downloaded_files:
                self.result_text.insert(tk.END, f"{file}\n")
            self.result_text.config(state=tk.DISABLED)

    def add_tooltips_and_info(self):
        # 使用说明
        instruction_frame = ttk.Frame(self.master)
        instruction_frame.pack(fill=tk.X, padx=5, pady=(0, 2))
        
        # 问号按钮
        ttk.Button(instruction_frame, text="使用指南", command=self.show_info).pack(side='right')

    def show_info(self):
        info = ("""
                1. 程序将自动连接Outlook，如果连接失败，请手动登录
                2. 选择要搜索的文件夹
                3. 输入搜索关键词（多个关键词用逗号分隔）
                4. 选择日期范围
                5. 如需下载附件，勾选"下载附件"并选择保存路径
                6. 点击"搜索"按钮开始搜索并下载

                        Version: 1.24.929
                如有任何问题，请联系: Eason  (yan.1024@icloud.com)""")
        messagebox.showinfo("使用说明", info)


def main():
    root = tk.Tk()
    app = OutlookApp(root)
    root.mainloop()
if __name__ == "__main__":
    main()


# pyinstaller --onefile --windowed --name "EmailMagic" --icon="搜查_find-one.ico" --add-data "hook-babel.py:." --hidden-import babel.numbers --hidden-import win32timezone 邮箱附件下载.py