import yagmail
import datetime
import tkinter as tk
import os
import configparser
import sys
import winshell
from tkinter import messagebox
from win32com.client import Dispatch


class AutoMail:
    """快速发送每日邮件工具"""

    def __init__(self):
        """初始化配置文件"""
        # 检查D盘更目录是否有auto_mail.ini文件，如果没有则创建
        self.config_file = "D:\\auto_mail.ini"
        self.config = configparser.ConfigParser()  # 创建配置文件解析器

        if not os.path.exists(self.config_file):  # 如果配置文件不存在
            print("配置文件不存在，开始创建")
            self.create_config()  # 创建配置文件
            print("配置文件创建成功，位于D:\\auto_mail.ini")

        # 读取配置文件，获取邮箱账号,授权码，收件人邮箱
        self.config.read(self.config_file)  # 读取配置文件
        self.my_email = self.config["email"].get("my_email", "")
        self.my_token = self.config["email"].get("my_token", "")
        self.my_receiver = self.config["email"].get("my_receiver", "")
        self.my_cc = self.config["email"].get("my_cc", "")

    def is_in_startup(self):
        startup_folder = winshell.startup()
        shortcut_path = os.path.join(startup_folder, "AutoMail.lnk")
        return os.path.exists(shortcut_path)

    def add_to_startup(self):
        script_path = os.path.abspath(sys.argv[0])
        startup_folder = winshell.startup()
        shortcut_path = os.path.join(startup_folder, "AutoMail.lnk")

        shell = Dispatch('WScript.Shell')
        shortcut = shell.CreateShortCut(shortcut_path)
        shortcut.Targetpath = sys.executable
        shortcut.Arguments = script_path
        shortcut.WorkingDirectory = os.path.dirname(script_path)
        shortcut.save()

    def remove_from_startup(self):
        startup_folder = winshell.startup()
        shortcut_path = os.path.join(startup_folder, "AutoMail.lnk")
        if os.path.exists(shortcut_path):
            os.remove(shortcut_path)

    def toggle_startup(self):
        if self.is_in_startup():
            if messagebox.askyesno("开机启动", "程序已在开机启动项中。是否要移除？"):
                try:
                    self.remove_from_startup()
                    messagebox.showinfo("开机启动", "程序已从开机启动项中移除。")
                    self.startup_button.config(text="添加到开机启动")
                except Exception as e:
                    messagebox.showerror("错误", f"从开机启动项移除时出错: {e}")
        else:
            if messagebox.askyesno("开机启动", "程序未在开机启动项中。是否要添加？"):
                try:
                    self.add_to_startup()
                    messagebox.showinfo("开机启动", "程序已成功添加到开机启动项。")
                    self.startup_button.config(text="从开机启动中移除")
                except Exception as e:
                    messagebox.showerror("错误", f"添加到开机启动项时出错: {e}")

    def create_config(self):
        """创建配置文件"""
        self.config["email"] = {
            "my_email": "",
            "my_token": "",
            "my_receiver": "",
            "my_cc": ""
        }
        try:
            with open(self.config_file, "w") as configfile:  # 写入配置文件
                self.config.write(configfile)  # 写入配置文件
            messagebox.showinfo("提示", "配置文件创建成功，位于D盘根目录")
        except Exception as e:
            messagebox.showerror("错误", f"配置文件创建失败: {e}")

    def ui(self):
        """创建UI界面"""
        self.root = tk.Tk()  # 创建窗口
        self.root.title("快速发送每日邮件")  # 设置窗口标题
        self.root.geometry("600x630")  # 设置窗口大小

        """subtitle"""
        tk.Label(self.root, text="一次设置后续无需再次填写！！！", bg="yellow").pack(anchor='w', padx=10,
                                                                                  pady=5)  # 创建label

        """邮箱账户"""
        tk.Label(self.root, text="你的公司邮箱账户").pack(anchor='w', padx=10, pady=5)  # 创建label
        self.email = tk.Entry(self.root, width=40)  # 创建输入框
        self.email.pack(anchor='w', padx=10)  # 放置输入框
        self.email.insert(0, self.my_email)  # 默认填入账号

        """邮箱授权码"""
        tk.Label(self.root, text="你的公司邮箱授权码").pack(anchor='w', padx=10, pady=5)  # 创建label
        self.token = tk.Entry(self.root, width=40)  # 创建输入框
        self.token.pack(anchor='w', padx=10)  # 放置输入框
        self.token.insert(0, self.my_token)  # 默认填入授权码

        """收件人邮箱"""
        tk.Label(self.root, text="发送给（多账户请用英文逗号隔开）：").pack(anchor='w', padx=10, pady=5)  # 创建label
        self.receiver = tk.Entry(self.root, width=80)  # 创建输入框
        self.receiver.pack(anchor='w', padx=10)  # 放置输入框
        self.receiver.insert(0, self.my_receiver)  # 默认填入收件人邮箱

        """抄送人邮箱"""
        tk.Label(self.root, text="抄送给（多账户请用英文逗号隔开）：").pack(anchor='w', padx=10, pady=5)  # 创建label
        self.cc = tk.Entry(self.root, width=80)  # 创建输入框
        self.cc.pack(anchor='w', padx=10)  # 放置输入框
        self.cc.insert(0, self.my_cc)  # 默认填入抄送人邮箱

        """邮件主题"""
        tk.Label(self.root, text="主题：").pack(anchor='w', padx=10, pady=5)  # 创建label
        self.title = tk.Entry(self.root, width=70)  # 创建输入框
        self.title.pack(anchor='w', padx=10)  # 放置输入框
        today = datetime.date.today()  # 获取今天的日期
        self.title.insert(0, f"{today} Work Plan")  # 默认填入邮件主题

        """邮件内容"""
        tk.Label(self.root, text="内容：").pack(anchor='w', padx=10, pady=5)  # 创建label
        self.content = tk.Text(self.root, width=70, height=15)  # 创建输入框
        self.content.pack(anchor='w', padx=10)  # 放置输入框

        send_button = tk.Button(self.root, text="发送邮件", command=self.send_mail)
        send_button.pack(pady=10)

        # 更新开机启动按钮
        button_text = "从开机启动中移除" if self.is_in_startup() else "添加到开机启动"
        self.startup_button = tk.Button(self.root, text=button_text, command=self.toggle_startup)
        self.startup_button.pack(pady=10)

        send_button = tk.Button(self.root, text="发送邮件", command=self.send_mail)
        send_button.pack(pady=10)

        self.root.mainloop()

    def send_mail(self):
        """发送邮件"""
        print("准备开始发送邮件")

        # 写入配置文件
        self.config["email"] = {
            "my_email": self.email.get(),
            "my_token": self.token.get(),
            "my_receiver": self.receiver.get(),
            "my_cc": self.cc.get()
        }
        try:
            with open(self.config_file, "w") as configfile:  # 写入配置文件
                self.config.write(configfile)  # 写入配置文件
        except Exception as e:
            messagebox.showerror("错误", f"配置文件更新失败: {e}")

        try:
            yag = yagmail.SMTP(user=self.email.get(),
                               password=self.token.get(),
                               host="smtp.qiye.aliyun.com",
                               port=465,
                               smtp_ssl=True)  # 登录邮箱

            # 准备邮件参数
            email_params = {
                "to": self.receiver.get().split(","),
                "subject": self.title.get(),
                "contents": self.content.get("1.0", tk.END)
            }

            # 只有在抄送列表不为空时才添加抄送
            cc_list = self.cc.get().strip()
            if cc_list:
                email_params["cc"] = cc_list.split(",")

            # 发送邮件
            yag.send(**email_params)

            messagebox.showinfo("提示", "邮件发送完成")
        except Exception as e:
            error_message = f"邮件发送失败: {str(e)}\n\n"
            error_message += f"邮箱账户: {self.email.get()}\n"
            error_message += f"收件人: {self.receiver.get()}\n"
            error_message += f"抄送: {self.cc.get()}\n"
            error_message += "请检查以上信息是否正确，以及网络连接是否正常。"
            messagebox.showerror("错误", error_message)


if __name__ == '__main__':
    a = AutoMail()
    a.ui()
