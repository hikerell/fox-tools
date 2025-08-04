import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox, filedialog
import smtplib
from email.header import Header
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email import utils
import threading
import json
import os
import pandas as pd


class EmailSenderGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("令狐会展科技批量邮件发送工具")
        self.root.geometry("900x750")
        self.root.resizable(True, True)

        # 配置文件路径
        self.config_file = "email_config.json"

        # 固定的SMTP配置
        self.smtp_server_value = "smtp.exmail.qq.com"
        self.smtp_port_value = "587"

        self.setup_ui()
        self.load_config()

    def setup_ui(self):
        # 创建主框架
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        # 配置根窗口的网格权重
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)

        # 发件人设置区域
        sender_frame = ttk.LabelFrame(main_frame, text="发件人设置", padding="5")
        sender_frame.grid(
            row=0, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10)
        )
        sender_frame.columnconfigure(1, weight=1)

        # SMTP信息显示
        ttk.Label(sender_frame, text="SMTP服务器:").grid(
            row=0, column=0, sticky=tk.W, padx=(0, 5)
        )
        ttk.Label(
            sender_frame,
            text=f"{self.smtp_server_value}:{self.smtp_port_value}",
            foreground="white",
        ).grid(row=0, column=1, sticky=tk.W)

        ttk.Label(sender_frame, text="发件人邮箱:").grid(
            row=1, column=0, sticky=tk.W, padx=(0, 5), pady=(5, 0)
        )
        self.sender_email = ttk.Entry(sender_frame)
        self.sender_email.grid(
            row=1, column=1, sticky=(tk.W, tk.E), padx=(0, 10), pady=(5, 0)
        )

        ttk.Label(sender_frame, text="邮箱密码:").grid(
            row=1, column=2, sticky=tk.W, padx=(0, 5), pady=(5, 0)
        )
        self.sender_password = ttk.Entry(sender_frame, show="*", width=20)
        self.sender_password.grid(row=1, column=3, sticky=tk.W, pady=(5, 0))

        # 员工名称设置
        name_frame = ttk.Frame(sender_frame)
        name_frame.grid(
            row=2, column=0, columnspan=4, sticky=(tk.W, tk.E), pady=(10, 0)
        )
        name_frame.columnconfigure(1, weight=1)
        name_frame.columnconfigure(3, weight=1)

        ttk.Label(name_frame, text="员工中文名:").grid(
            row=0, column=0, sticky=tk.W, padx=(0, 5)
        )
        self.chinese_name = ttk.Entry(name_frame)
        self.chinese_name.grid(row=0, column=1, sticky=(tk.W, tk.E), padx=(0, 20))

        ttk.Label(name_frame, text="员工英文名:").grid(
            row=0, column=2, sticky=tk.W, padx=(0, 5)
        )
        self.english_name = ttk.Entry(name_frame)
        self.english_name.grid(row=0, column=3, sticky=(tk.W, tk.E))

        # 收件人设置区域
        receivers_frame = ttk.LabelFrame(main_frame, text="收件人设置", padding="5")
        receivers_frame.grid(
            row=1, column=0, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10)
        )
        receivers_frame.columnconfigure(0, weight=1)
        receivers_frame.rowconfigure(2, weight=1)

        # 收件人输入方式选择
        input_method_frame = ttk.Frame(receivers_frame)
        input_method_frame.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=(0, 5))

        self.input_method = tk.StringVar(value="manual")
        ttk.Radiobutton(
            input_method_frame,
            text="手动输入",
            variable=self.input_method,
            value="manual",
            command=self.toggle_input_method,
        ).pack(side=tk.LEFT)
        ttk.Radiobutton(
            input_method_frame,
            text="TXT文件导入",
            variable=self.input_method,
            value="txt",
            command=self.toggle_input_method,
        ).pack(side=tk.LEFT, padx=(10, 0))
        ttk.Radiobutton(
            input_method_frame,
            text="Excel文件导入",
            variable=self.input_method,
            value="excel",
            command=self.toggle_input_method,
        ).pack(side=tk.LEFT, padx=(10, 0))

        # 文件导入按钮和说明
        file_info_frame = ttk.Frame(receivers_frame)
        file_info_frame.grid(row=1, column=0, sticky=(tk.W, tk.E), pady=(5, 5))

        self.file_button = ttk.Button(
            file_info_frame, text="选择文件", command=self.load_receivers_from_file
        )
        self.file_button.pack(side=tk.LEFT)

        self.file_info_label = ttk.Label(file_info_frame, text="", foreground="gray")
        self.file_info_label.pack(side=tk.LEFT, padx=(10, 0))

        # 收件人文本框
        receivers_text_frame = ttk.Frame(receivers_frame)
        receivers_text_frame.grid(row=2, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        receivers_text_frame.columnconfigure(0, weight=1)
        receivers_text_frame.rowconfigure(1, weight=1)

        self.receivers_label = ttk.Label(
            receivers_text_frame, text="收件人邮箱 (每行一个):"
        )
        self.receivers_label.grid(row=0, column=0, sticky=tk.W, pady=(0, 5))

        self.receivers_text = scrolledtext.ScrolledText(receivers_text_frame, height=8)
        self.receivers_text.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        # 邮件内容区域
        content_frame = ttk.LabelFrame(main_frame, text="邮件内容", padding="5")
        content_frame.grid(
            row=2, column=0, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10)
        )
        content_frame.columnconfigure(0, weight=1)
        content_frame.rowconfigure(2, weight=1)

        # 邮件语言选择
        language_frame = ttk.Frame(content_frame)
        language_frame.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=(0, 10))

        ttk.Label(language_frame, text="邮件语言:").pack(side=tk.LEFT)
        self.email_language = tk.StringVar(value="chinese")
        ttk.Radiobutton(
            language_frame,
            text="中文",
            variable=self.email_language,
            value="chinese",
            command=self.update_signature_preview,
        ).pack(side=tk.LEFT, padx=(10, 0))
        ttk.Radiobutton(
            language_frame,
            text="英文",
            variable=self.email_language,
            value="english",
            command=self.update_signature_preview,
        ).pack(side=tk.LEFT, padx=(10, 0))

        # 签名预览
        self.signature_label = ttk.Label(language_frame, text="", foreground="white")
        self.signature_label.pack(side=tk.RIGHT)

        ttk.Label(content_frame, text="主题:").grid(row=1, column=0, sticky=tk.W)
        self.subject_entry = ttk.Entry(content_frame)
        self.subject_entry.grid(row=1, column=0, sticky=(tk.W, tk.E), pady=(5, 10))

        ttk.Label(content_frame, text="正文:").grid(row=2, column=0, sticky=tk.W)
        self.content_text = scrolledtext.ScrolledText(content_frame, height=10)
        self.content_text.grid(
            row=2, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(5, 0)
        )

        # 控制按钮区域
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=3, column=0, columnspan=2, pady=(10, 0))

        ttk.Button(button_frame, text="保存配置", command=self.save_config).pack(
            side=tk.LEFT, padx=(0, 10)
        )
        ttk.Button(button_frame, text="测试连接", command=self.test_connection).pack(
            side=tk.LEFT, padx=(0, 10)
        )
        self.send_button = ttk.Button(
            button_frame, text="发送邮件", command=self.send_emails_thread
        )
        self.send_button.pack(side=tk.LEFT, padx=(0, 10))

        # 进度条
        self.progress = ttk.Progressbar(main_frame, mode="determinate")
        self.progress.grid(
            row=4, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(10, 5)
        )

        # 状态标签
        self.status_label = ttk.Label(main_frame, text="就绪")
        self.status_label.grid(row=5, column=0, columnspan=2, pady=(5, 0))

        # 配置网格权重
        main_frame.rowconfigure(1, weight=1)
        main_frame.rowconfigure(2, weight=2)

        # 初始化界面状态
        self.toggle_input_method()
        self.update_signature_preview()

    def toggle_input_method(self):
        """切换收件人输入方式"""
        method = self.input_method.get()

        if method == "manual":
            self.file_button.config(state="disabled")
            self.receivers_text.config(state="normal")
            self.file_info_label.config(text="")
            self.receivers_label.config(text="收件人邮箱 (每行一个):")
        elif method == "txt":
            self.file_button.config(state="normal")
            self.receivers_text.config(state="normal")
            self.file_info_label.config(text="TXT格式：每行一个邮箱地址")
            self.receivers_label.config(text="从TXT文件导入的收件人:")
        elif method == "excel":
            self.file_button.config(state="normal")
            self.receivers_text.config(state="normal")
            self.file_info_label.config(text="Excel格式：需要包含'email'列")
            self.receivers_label.config(text="从Excel文件导入的收件人:")

    def update_signature_preview(self):
        """更新签名预览"""
        chinese_name = self.chinese_name.get().strip()
        english_name = self.english_name.get().strip()

        if self.email_language.get() == "chinese":
            if chinese_name:
                self.signature_label.config(text=f"署名将使用: {chinese_name}")
            else:
                self.signature_label.config(text="请输入中文名称、邮件主题和邮件正文")
        else:
            if english_name:
                self.signature_label.config(text=f"Signature will use: {english_name}")
            else:
                self.signature_label.config(
                    text="Please enter English name、subject and content"
                )

    def load_receivers_from_file(self):
        """从文件加载收件人"""
        method = self.input_method.get()

        if method == "txt":
            file_path = filedialog.askopenfilename(
                title="选择TXT文件",
                filetypes=[("文本文件", "*.txt"), ("所有文件", "*.*")],
            )
            if file_path:
                self.load_txt_file(file_path)
        elif method == "excel":
            file_path = filedialog.askopenfilename(
                title="选择Excel文件",
                filetypes=[("Excel文件", "*.xlsx"), ("所有文件", "*.*")],
            )
            if file_path:
                self.load_excel_file(file_path)

    def load_txt_file(self, file_path):
        """加载TXT文件"""
        try:
            with open(file_path, "r", encoding="utf-8") as f:
                content = f.read()

            # 处理内容，过滤空行
            emails = [line.strip() for line in content.split("\n") if line.strip()]

            self.receivers_text.config(state="normal")
            self.receivers_text.delete("1.0", tk.END)
            self.receivers_text.insert("1.0", "\n".join(emails))
            self.receivers_text.config(state="normal")
            # self.receivers_text.config(state="disabled")

            self.status_label.config(text=f"已从TXT文件加载 {len(emails)} 个邮箱地址")

        except Exception as e:
            messagebox.showerror("错误", f"读取TXT文件失败: {str(e)}")

    def load_excel_file(self, file_path):
        """加载Excel文件"""
        try:
            # 读取Excel文件
            df = pd.read_excel(file_path)

            # 检查是否有email列
            email_columns = [col for col in df.columns if "email" in col.lower()]
            if not email_columns:
                messagebox.showerror(
                    "错误", "Excel文件中未找到'email'列\n请确保有一列名为'email'的数据"
                )
                return

            # 使用第一个找到的email列
            email_column = email_columns[0]
            emails = df[email_column].dropna().astype(str).tolist()

            # 过滤无效邮箱
            valid_emails = [
                email.strip() for email in emails if email.strip() and "@" in email
            ]

            self.receivers_text.config(state="normal")
            self.receivers_text.delete("1.0", tk.END)
            self.receivers_text.insert("1.0", "\n".join(valid_emails))
            self.receivers_text.config(state="normal")
            # self.receivers_text.config(state="disabled")

            self.status_label.config(
                text=f"已从Excel文件加载 {len(valid_emails)} 个邮箱地址 (使用列: {email_column})"
            )

        except ImportError:
            messagebox.showerror(
                "错误",
                "需要安装pandas库来读取Excel文件\n请运行: pip install pandas openpyxl",
            )
        except Exception as e:
            messagebox.showerror("错误", f"读取Excel文件失败: {str(e)}")

    def save_config(self):
        """保存配置到文件"""
        config = {
            "sender_email": self.sender_email.get(),
            "sender_password": self.sender_password.get(),  # 注意：实际应用中不建议保存密码
            "chinese_name": self.chinese_name.get(),
            "english_name": self.english_name.get(),
        }
        try:
            with open(self.config_file, "w", encoding="utf-8") as f:
                json.dump(config, f, ensure_ascii=False, indent=2)
            messagebox.showinfo("成功", "配置已保存")
        except Exception as e:
            messagebox.showerror("错误", f"保存配置失败: {str(e)}")

    def load_config(self):
        """从文件加载配置"""
        if os.path.exists(self.config_file):
            try:
                with open(self.config_file, "r", encoding="utf-8") as f:
                    config = json.load(f)
                self.sender_email.insert(0, config.get("sender_email", ""))
                self.sender_password.insert(0, config.get("sender_password", ""))
                self.chinese_name.insert(0, config.get("chinese_name", ""))
                self.english_name.insert(0, config.get("english_name", ""))

                # 绑定事件更新签名预览
                self.chinese_name.bind(
                    "<KeyRelease>", lambda e: self.update_signature_preview()
                )
                self.english_name.bind(
                    "<KeyRelease>", lambda e: self.update_signature_preview()
                )

            except Exception as e:
                print(f"加载配置失败: {str(e)}")
        else:
            # 绑定事件更新签名预览
            self.chinese_name.bind(
                "<KeyRelease>", lambda e: self.update_signature_preview()
            )
            self.english_name.bind(
                "<KeyRelease>", lambda e: self.update_signature_preview()
            )

    def test_connection(self):
        """测试SMTP连接"""
        try:
            server = smtplib.SMTP(self.smtp_server_value, int(self.smtp_port_value))
            server.starttls()
            server.login(self.sender_email.get(), self.sender_password.get())
            server.quit()
            messagebox.showinfo("成功", "SMTP连接测试成功！")
        except Exception as e:
            messagebox.showerror("错误", f"SMTP连接失败: {str(e)}")

    def get_receivers_list(self):
        """获取收件人列表"""
        receivers_content = self.receivers_text.get("1.0", tk.END).strip()
        if not receivers_content:
            return []

        # 按行分割并过滤空行
        receivers = [
            email.strip() for email in receivers_content.split("\n") if email.strip()
        ]
        return receivers

    def get_sender_name(self):
        """根据邮件语言获取发件人姓名"""
        if self.email_language.get() == "chinese":
            return self.chinese_name.get().strip()
        else:
            return self.english_name.get().strip()

    def send_emails(self, receivers: list[str], subject: str, content: str):
        """发送邮件的核心函数"""
        try:
            # 连接SMTP服务器
            server = smtplib.SMTP(self.smtp_server_value, int(self.smtp_port_value))
            server.starttls()
            server.login(self.sender_email.get(), self.sender_password.get())

            success_count = 0
            total_count = len(receivers)
            sender_name = self.get_sender_name()

            for i, receiver in enumerate(receivers):
                try:
                    # 创建邮件
                    msg = MIMEMultipart()

                    # 设置发件人信息（包含姓名，处理中文编码）
                    if sender_name:
                        # 使用Header来正确编码中文姓名
                        encoded_name = Header(sender_name, "utf-8").encode()
                        msg["From"] = f"{encoded_name} <{self.sender_email.get()}>"
                    else:
                        msg["From"] = self.sender_email.get()

                    msg["To"] = receiver

                    # 正确编码中文主题
                    msg["Subject"] = Header(subject, "utf-8")

                    # 设置邮件头部信息
                    msg["Date"] = utils.formatdate(localtime=True)
                    msg["Message-ID"] = utils.make_msgid()

                    # 添加邮件正文
                    msg.attach(MIMEText(content, "plain", "utf-8"))

                    # 发送邮件
                    text = msg.as_string()
                    server.sendmail(self.sender_email.get(), receiver, text)
                    success_count += 1

                    # 更新进度
                    progress_value = (i + 1) / total_count * 100
                    self.progress["value"] = progress_value
                    self.status_label.config(
                        text=f"正在发送... ({i + 1}/{total_count})"
                    )
                    self.root.update()

                except Exception as e:
                    print(f"发送到 {receiver} 失败: {str(e)}")
                    continue

            server.quit()

            # 显示结果
            self.status_label.config(
                text=f"发送完成！成功: {success_count}, 失败: {total_count - success_count}"
            )
            messagebox.showinfo(
                "完成",
                f"邮件发送完成！\n成功发送: {success_count} 封\n发送失败: {total_count - success_count} 封",
            )

        except Exception as e:
            self.status_label.config(text="发送失败")
            messagebox.showerror("错误", f"发送邮件时出错: {str(e)}")
        finally:
            self.send_button.config(state="normal")
            self.progress["value"] = 0

    def send_emails_thread(self):
        """在新线程中发送邮件"""
        # 验证输入
        if not self.sender_email.get():
            messagebox.showerror("错误", "请输入发件人邮箱")
            return

        if not self.sender_password.get():
            messagebox.showerror("错误", "请输入邮箱密码")
            return

        # 检查员工姓名
        sender_name = self.get_sender_name()
        if not sender_name:
            language = "中文" if self.email_language.get() == "chinese" else "英文"
            if not messagebox.askyesno(
                "警告",
                f"未设置员工{language}姓名，邮件将不显示发件人姓名。\n是否继续发送？",
            ):
                return

        receivers = self.get_receivers_list()
        if not receivers:
            messagebox.showerror("错误", "请输入收件人邮箱")
            return

        subject = self.subject_entry.get()
        if not subject:
            messagebox.showerror("错误", "请输入邮件主题")
            return

        content = self.content_text.get("1.0", tk.END).strip()
        if not content:
            messagebox.showerror("错误", "请输入邮件内容")
            return

        # 确认发送信息
        language = "中文" if self.email_language.get() == "chinese" else "英文"
        sender_info = (
            f"发件人: {sender_name} <{self.sender_email.get()}>"
            if sender_name
            else f"发件人: {self.sender_email.get()}"
        )
        confirm_msg = (
            f"确定要发送{language}邮件给 {len(receivers)} 个收件人吗？\n\n{sender_info}"
        )

        if not messagebox.askyesno("确认", confirm_msg):
            return

        # 禁用发送按钮
        self.send_button.config(state="disabled")
        self.status_label.config(text="准备发送...")

        # 在新线程中发送邮件
        thread = threading.Thread(
            target=self.send_emails, args=(receivers, subject, content)
        )
        thread.daemon = True
        thread.start()


def main():
    root = tk.Tk()
    app = EmailSenderGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()
