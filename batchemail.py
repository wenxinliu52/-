import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.image import MIMEImage
from email.mime.base import MIMEBase
from email.header import Header
from email.utils import formatdate
from email import encoders
import pandas as pd
import time
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import threading
import json
import os
import random
import base64
from PIL import Image, ImageTk
import io

class EmailSenderGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("吉吉国邮件批发王_bylwx v3.2")
        
        # 获取屏幕尺寸，设置合适的窗口大小
        screen_w = root.winfo_screenwidth()
        screen_h = root.winfo_screenheight()
        win_w = min(1000, screen_w - 100)
        win_h = min(900, screen_h - 100)
        x = (screen_w - win_w) // 2
        y = (screen_h - win_h) // 2
        self.root.geometry(f"{win_w}x{win_h}+{x}+{y}")
        self.root.minsize(600, 500)  # 设置最小窗口大小
        self.root.resizable(True, True)
        
        self.config_file = "email_config.json"
        self.images = []
        self.attachments = []
        
        self.create_scrollable_frame()
        self.create_widgets()
        self.load_config()
        
    def create_scrollable_frame(self):
        """创建可滚动的主框架"""
        # 外层容器
        self.outer_frame = ttk.Frame(self.root)
        self.outer_frame.pack(fill=tk.BOTH, expand=True)
        
        # 创建Canvas和滚动条
        self.canvas = tk.Canvas(self.outer_frame, highlightthickness=0)
        self.v_scrollbar = ttk.Scrollbar(self.outer_frame, orient=tk.VERTICAL, command=self.canvas.yview)
        self.h_scrollbar = ttk.Scrollbar(self.outer_frame, orient=tk.HORIZONTAL, command=self.canvas.xview)
        
        # 可滚动框架
        self.scrollable_frame = ttk.Frame(self.canvas)
        
        self.scrollable_frame.bind("<Configure>", lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))
        
        self.canvas_window = self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        
        self.canvas.configure(yscrollcommand=self.v_scrollbar.set, xscrollcommand=self.h_scrollbar.set)
        
        # 布局
        self.v_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.h_scrollbar.pack(side=tk.BOTTOM, fill=tk.X)
        self.canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # 绑定鼠标滚轮
        self.canvas.bind_all("<MouseWheel>", self._on_mousewheel)
        self.canvas.bind_all("<Button-4>", self._on_mousewheel)
        self.canvas.bind_all("<Button-5>", self._on_mousewheel)
        
        # 让canvas内容宽度跟随窗口
        self.canvas.bind("<Configure>", self._on_canvas_configure)
    
    def _on_canvas_configure(self, event):
        # 让scrollable_frame的宽度至少等于canvas宽度
        min_width = event.width
        self.canvas.itemconfig(self.canvas_window, width=min_width)
    
    def _on_mousewheel(self, event):
        # Windows/Mac
        if event.num == 4:
            self.canvas.yview_scroll(-1, "units")
        elif event.num == 5:
            self.canvas.yview_scroll(1, "units")
        else:
            self.canvas.yview_scroll(int(-1*(event.delta/120)), "units")
        
    def create_widgets(self):
        # 主框架 - 使用scrollable_frame
        main_frame = ttk.Frame(self.scrollable_frame, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # ========== 邮箱配置区域 ==========
        config_frame = ttk.LabelFrame(main_frame, text="邮箱配置", padding="5")
        config_frame.pack(fill=tk.X, pady=(0, 5))
        
        # 使用grid布局，但让第1列（输入框）可扩展
        config_frame.columnconfigure(1, weight=1)
        
        ttk.Label(config_frame, text="发件人邮箱:").grid(row=0, column=0, sticky=tk.W, pady=2)
        self.email_entry = ttk.Entry(config_frame)
        self.email_entry.grid(row=0, column=1, sticky=(tk.W, tk.E), padx=5, pady=2)
        
        ttk.Label(config_frame, text="发件人姓名:").grid(row=1, column=0, sticky=tk.W, pady=2)
        name_frame = ttk.Frame(config_frame)
        name_frame.grid(row=1, column=1, sticky=(tk.W, tk.E), padx=5, pady=2)
        name_frame.columnconfigure(0, weight=1)
        
        self.sender_name_entry = ttk.Entry(name_frame)
        self.sender_name_entry.grid(row=0, column=0, sticky=(tk.W, tk.E), padx=(0, 5))
        self.sender_name_entry.insert(0, "Vincent")
        
        ttk.Label(name_frame, text="(显示为：姓名 <邮箱>)", foreground="gray", font=('', 8)).grid(row=0, column=1)
        
        ttk.Label(config_frame, text="邮箱密码:").grid(row=2, column=0, sticky=tk.W, pady=2)
        self.password_entry = ttk.Entry(config_frame, show="*")
        self.password_entry.grid(row=2, column=1, sticky=(tk.W, tk.E), padx=5, pady=2)
        
        ttk.Label(config_frame, text="抄送邮箱(CC):").grid(row=3, column=0, sticky=tk.W, pady=2)
        self.cc_entry = ttk.Entry(config_frame)
        self.cc_entry.grid(row=3, column=1, sticky=(tk.W, tk.E), padx=5, pady=2)
        
        ttk.Label(config_frame, text="邮箱类型:").grid(row=4, column=0, sticky=tk.W, pady=2)
        self.email_type = ttk.Combobox(config_frame, state="readonly")
        self.email_type.grid(row=4, column=1, sticky=(tk.W, tk.E), padx=5, pady=2)
        self.email_type['values'] = ('QQ邮箱', '阿里企业邮箱', '163邮箱', '126邮箱', 'Gmail', '自定义')
        self.email_type.current(0)
        self.email_type.bind('<<ComboboxSelected>>', self.on_email_type_change)
        
        ttk.Label(config_frame, text="SMTP服务器:").grid(row=5, column=0, sticky=tk.W, pady=2)
        self.smtp_entry = ttk.Entry(config_frame)
        self.smtp_entry.grid(row=5, column=1, sticky=(tk.W, tk.E), padx=5, pady=2)
        self.smtp_entry.insert(0, "smtp.qq.com")
        
        ttk.Label(config_frame, text="SMTP端口:").grid(row=6, column=0, sticky=tk.W, pady=2)
        port_frame = ttk.Frame(config_frame)
        port_frame.grid(row=6, column=1, sticky=(tk.W, tk.E), padx=5, pady=2)
        
        self.port_entry = ttk.Entry(port_frame, width=10)
        self.port_entry.pack(side=tk.LEFT)
        self.port_entry.insert(0, "587")
        
        self.use_ssl = tk.BooleanVar(value=False)
        ttk.Checkbutton(port_frame, text="使用SSL (465端口选中)", variable=self.use_ssl).pack(side=tk.LEFT, padx=10)
        
        button_frame = ttk.Frame(config_frame)
        button_frame.grid(row=7, column=1, sticky=tk.E, pady=2)
        ttk.Button(button_frame, text="测试连接", command=self.test_connection).pack(side=tk.RIGHT, padx=2)
        ttk.Button(button_frame, text="保存配置", command=self.save_config).pack(side=tk.RIGHT, padx=2)
        
        # ========== 文件选择区域 ==========
        file_frame = ttk.LabelFrame(main_frame, text="Excel文件与发送设置", padding="5")
        file_frame.pack(fill=tk.X, pady=(0, 5))
        file_frame.columnconfigure(1, weight=1)
        
        ttk.Label(file_frame, text="Excel文件:").grid(row=0, column=0, sticky=tk.W, pady=2)
        file_inner = ttk.Frame(file_frame)
        file_inner.grid(row=0, column=1, sticky=(tk.W, tk.E), padx=5, pady=2)
        file_inner.columnconfigure(0, weight=1)
        
        self.file_path = tk.StringVar()
        ttk.Entry(file_inner, textvariable=self.file_path, state="readonly").grid(row=0, column=0, sticky=(tk.W, tk.E))
        ttk.Button(file_inner, text="选择", command=self.select_file).grid(row=0, column=1, padx=(5, 0))
        
        ttk.Label(file_frame, text="发送间隔(秒):").grid(row=1, column=0, sticky=tk.W, pady=2)
        delay_frame = ttk.Frame(file_frame)
        delay_frame.grid(row=1, column=1, sticky=tk.W, padx=5, pady=2)
        
        self.delay_min_entry = ttk.Entry(delay_frame, width=8)
        self.delay_min_entry.pack(side=tk.LEFT)
        self.delay_min_entry.insert(0, "15")
        ttk.Label(delay_frame, text=" 至 ").pack(side=tk.LEFT)
        self.delay_max_entry = ttk.Entry(delay_frame, width=8)
        self.delay_max_entry.pack(side=tk.LEFT)
        self.delay_max_entry.insert(0, "30")
        
        ttk.Label(file_frame, text="自动分批:").grid(row=2, column=0, sticky=tk.W, pady=2)
        batch_frame = ttk.Frame(file_frame)
        batch_frame.grid(row=2, column=1, sticky=tk.W, padx=5, pady=2)
        
        self.enable_batch = tk.BooleanVar(value=True)
        ttk.Checkbutton(batch_frame, text="启用", variable=self.enable_batch, command=self.toggle_batch_settings).pack(side=tk.LEFT)
        
        self.batch_size_label = ttk.Label(batch_frame, text=" 每批:")
        self.batch_size_label.pack(side=tk.LEFT)
        self.batch_size_entry = ttk.Entry(batch_frame, width=6)
        self.batch_size_entry.pack(side=tk.LEFT)
        self.batch_size_entry.insert(0, "15")
        ttk.Label(batch_frame, text="封").pack(side=tk.LEFT)
        
        self.batch_delay_label = ttk.Label(batch_frame, text=" 间隔:")
        self.batch_delay_label.pack(side=tk.LEFT)
        self.batch_delay_entry = ttk.Entry(batch_frame, width=6)
        self.batch_delay_entry.pack(side=tk.LEFT)
        self.batch_delay_entry.insert(0, "10")
        ttk.Label(batch_frame, text="分钟").pack(side=tk.LEFT)
        
        # ========== 邮件模板区域 ==========
        template_frame = ttk.LabelFrame(main_frame, text="邮件模板（{name}替换姓名，{img1},{img2}等替换图片位置）", padding="5")
        template_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 5))
        
        # 主题行
        subject_frame = ttk.Frame(template_frame)
        subject_frame.pack(fill=tk.X, pady=(0, 5))
        ttk.Label(subject_frame, text="邮件主题:").pack(side=tk.LEFT)
        self.subject_entry = ttk.Entry(subject_frame)
        self.subject_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(5, 0))
        self.subject_entry.insert(0, "Inquiry Follow-up from KAT VR")
        
        # 图片和附件管理
        media_frame = ttk.Frame(template_frame)
        media_frame.pack(fill=tk.X, pady=5)
        
        ttk.Button(media_frame, text="添加图片", command=self.add_image).pack(side=tk.LEFT, padx=2)
        ttk.Button(media_frame, text="添加附件", command=self.add_attachment).pack(side=tk.LEFT, padx=2)
        ttk.Button(media_frame, text="清空图片/附件", command=self.clear_media).pack(side=tk.LEFT, padx=2)
        
        self.media_label = ttk.Label(media_frame, text="已添加: 0张图片, 0个附件", foreground="blue")
        self.media_label.pack(side=tk.LEFT, padx=10)
        
        ttk.Label(template_frame, text="邮件正文:").pack(anchor=tk.W)
        
        self.body_text = scrolledtext.ScrolledText(template_frame, height=10, wrap=tk.WORD)
        self.body_text.pack(fill=tk.BOTH, expand=True, pady=2)
        
        default_body = """Dear {name},

This is Vincent, the sales manager from KAT VR.

{img1}

Our company received your inquiries about our treadmills before, may I know if we still have the chance to cooperate?

{img2}

Would you like to accept our latest prices and catalog?

Thank you. Waiting for your reply. :)

Best regards,
Vincent
Sales Manager
KAT VR"""
        self.body_text.insert("1.0", default_body)
        
        # 图片预览区
        preview_frame = ttk.LabelFrame(template_frame, text="图片预览（点击可删除）", padding="3")
        preview_frame.pack(fill=tk.X, pady=5)
        
        self.preview_canvas = tk.Canvas(preview_frame, height=70, bg='#f0f0f0')
        self.preview_canvas.pack(fill=tk.X, expand=True)
        
        # ========== 操作按钮区域 ==========
        action_frame = ttk.Frame(main_frame)
        action_frame.pack(fill=tk.X, pady=5)
        
        self.send_button = ttk.Button(action_frame, text="开始发送", command=self.start_sending)
        self.send_button.pack(side=tk.LEFT, padx=3)
        
        self.stop_button = ttk.Button(action_frame, text="停止发送", command=self.stop_sending, state=tk.DISABLED)
        self.stop_button.pack(side=tk.LEFT, padx=3)
        
        ttk.Button(action_frame, text="测试邮件", command=self.send_test_email).pack(side=tk.LEFT, padx=3)
        ttk.Button(action_frame, text="预览邮件", command=self.preview_email).pack(side=tk.LEFT, padx=3)
        
        # ========== 日志区域 ==========
        log_frame = ttk.LabelFrame(main_frame, text="发送日志", padding="5")
        log_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 5))
        
        self.log_text = scrolledtext.ScrolledText(log_frame, height=5, state=tk.DISABLED)
        self.log_text.pack(fill=tk.BOTH, expand=True)
        
        # 进度条和状态
        self.progress = ttk.Progressbar(main_frame, mode='determinate')
        self.progress.pack(fill=tk.X, pady=(0, 5))
        
        self.status_label = ttk.Label(main_frame, text="就绪", relief=tk.SUNKEN)
        self.status_label.pack(fill=tk.X)
        
        self.stop_flag = False
        self.image_refs = []
    
    def add_image(self):
        filetypes = [("图片文件", "*.png *.jpg *.jpeg *.gif *.bmp"), ("所有文件", "*.*")]
        filepath = filedialog.askopenfilename(title="选择图片", filetypes=filetypes)
        
        if filepath:
            img_id = f"img{len(self.images) + 1}"
            cid = f"image{len(self.images) + 1}_{int(time.time())}"
            self.images.append({
                'path': filepath, 'cid': cid, 'id': img_id,
                'filename': os.path.basename(filepath)
            })
            self.update_media_label()
            self.update_preview()
            self.log_message(f"已添加图片: {os.path.basename(filepath)} (使用 {{{img_id}}} 在正文中插入)")
            messagebox.showinfo("提示", f"图片已添加！\n\n在邮件正文中使用 {{{img_id}}} 来插入这张图片")
    
    def add_attachment(self):
        filepath = filedialog.askopenfilename(title="选择附件")
        if filepath:
            self.attachments.append({'path': filepath, 'filename': os.path.basename(filepath)})
            self.update_media_label()
            self.log_message(f"已添加附件: {os.path.basename(filepath)}")
    
    def clear_media(self):
        if self.images or self.attachments:
            if messagebox.askyesno("确认", "确定要清空所有图片和附件吗？"):
                self.images = []
                self.attachments = []
                self.update_media_label()
                self.update_preview()
                self.log_message("已清空所有图片和附件")
    
    def update_media_label(self):
        self.media_label.config(text=f"已添加: {len(self.images)}张图片, {len(self.attachments)}个附件")
    
    def update_preview(self):
        self.preview_canvas.delete("all")
        self.image_refs = []
        x = 10
        for i, img_data in enumerate(self.images):
            try:
                img = Image.open(img_data['path'])
                img.thumbnail((60, 60))
                photo = ImageTk.PhotoImage(img)
                self.image_refs.append(photo)
                self.preview_canvas.create_image(x, 35, image=photo, anchor=tk.W)
                self.preview_canvas.create_text(x + 30, 65, text=f"{{{img_data['id']}}}", font=('', 7))
                x += 80
            except Exception as e:
                self.log_message(f"预览图片失败: {str(e)}")
    
    def preview_email(self):
        preview_window = tk.Toplevel(self.root)
        preview_window.title("邮件预览")
        preview_window.geometry("600x500")
        
        html_content = self.build_html_body(self.body_text.get("1.0", tk.END).strip().replace("{name}", "【收件人姓名】"))
        
        text_widget = scrolledtext.ScrolledText(preview_window, wrap=tk.WORD)
        text_widget.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        preview_text = f"""━━━━ 邮件预览 ━━━━

【主题】{self.subject_entry.get()}
【发件人】{self.sender_name_entry.get()} <{self.email_entry.get()}>
【图片】{len(self.images)} 张
"""
        for img in self.images:
            preview_text += f"  - {img['filename']} ({{{img['id']}}})\n"
        
        preview_text += f"【附件】{len(self.attachments)} 个\n"
        for att in self.attachments:
            preview_text += f"  - {att['filename']}\n"
        
        preview_text += f"\n━━━━ 正文内容 ━━━━\n\n{self.body_text.get('1.0', tk.END).strip().replace('{name}', '【收件人姓名】')}"
        
        text_widget.insert("1.0", preview_text)
        text_widget.config(state=tk.DISABLED)
        ttk.Button(preview_window, text="关闭", command=preview_window.destroy).pack(pady=10)
    
    def build_html_body(self, body_text):
        html_body = body_text.replace('\n', '<br>\n')
        for img_data in self.images:
            placeholder = f"{{{img_data['id']}}}"
            img_html = f'<br><img src="cid:{img_data["cid"]}" style="max-width:100%; height:auto; margin:10px 0;"><br>'
            html_body = html_body.replace(placeholder, img_html)
        
        fonts = ['Arial, sans-serif', 'Helvetica, Arial, sans-serif']
        colors = ['#333333', '#2c2c2c']
        
        return f"""<html>
<head><meta http-equiv="Content-Type" content="text/html; charset=utf-8"></head>
<body style="font-family: {random.choice(fonts)}; font-size: 14px; color: {random.choice(colors)}; line-height: 1.6;">
<div style="max-width: 600px; margin: 0 auto;">{html_body}</div>
</body></html>"""
    
    def toggle_batch_settings(self):
        state = tk.NORMAL if self.enable_batch.get() else tk.DISABLED
        self.batch_size_entry.config(state=state)
        self.batch_delay_entry.config(state=state)
    
    def on_email_type_change(self, event):
        smtp_configs = {
            'QQ邮箱': ('smtp.qq.com', '587', False),
            '阿里企业邮箱': ('smtp.mxhichina.com', '465', True),
            '163邮箱': ('smtp.163.com', '465', True),
            '126邮箱': ('smtp.126.com', '465', True),
            'Gmail': ('smtp.gmail.com', '587', False),
            '自定义': ('', '587', False)
        }
        email_type = self.email_type.get()
        if email_type in smtp_configs:
            smtp, port, ssl = smtp_configs[email_type]
            self.smtp_entry.delete(0, tk.END)
            self.smtp_entry.insert(0, smtp)
            self.port_entry.delete(0, tk.END)
            self.port_entry.insert(0, port)
            self.use_ssl.set(ssl)
    
    def select_file(self):
        filename = filedialog.askopenfilename(title="选择Excel文件", filetypes=[("Excel files", "*.xlsx *.xls")])
        if filename:
            self.file_path.set(filename)
            self.log_message(f"已选择文件: {filename}")
    
    def test_connection(self):
        if not self.email_entry.get() or not self.password_entry.get():
            messagebox.showerror("错误", "请填写邮箱和密码！")
            return
        
        self.log_message("\n=== 测试SMTP连接 ===")
        try:
            if self.use_ssl.get():
                with smtplib.SMTP_SSL(self.smtp_entry.get(), int(self.port_entry.get()), timeout=15) as srv:
                    srv.login(self.email_entry.get(), self.password_entry.get())
            else:
                with smtplib.SMTP(self.smtp_entry.get(), int(self.port_entry.get()), timeout=15) as srv:
                    srv.ehlo()
                    srv.starttls()
                    srv.ehlo()
                    srv.login(self.email_entry.get(), self.password_entry.get())
            self.log_message("✓ 连接成功！")
            messagebox.showinfo("成功", "SMTP连接测试通过！")
        except Exception as e:
            self.log_message(f"✗ 连接失败: {str(e)}")
            messagebox.showerror("失败", f"连接失败: {str(e)}")
    
    def log_message(self, message):
        self.log_text.config(state=tk.NORMAL)
        self.log_text.insert(tk.END, f"{message}\n")
        self.log_text.see(tk.END)
        self.log_text.config(state=tk.DISABLED)
    
    def save_config(self):
        config = {
            "email": self.email_entry.get(),
            "sender_name": self.sender_name_entry.get(),
            "cc_email": self.cc_entry.get(),
            "email_type": self.email_type.get(),
            "smtp": self.smtp_entry.get(),
            "port": self.port_entry.get(),
            "use_ssl": self.use_ssl.get()
        }
        with open(self.config_file, 'w', encoding='utf-8') as f:
            json.dump(config, f, ensure_ascii=False, indent=2)
        messagebox.showinfo("成功", "配置已保存！")
    
    def load_config(self):
        if os.path.exists(self.config_file):
            try:
                with open(self.config_file, 'r', encoding='utf-8') as f:
                    config = json.load(f)
                self.email_entry.insert(0, config.get("email", ""))
                self.sender_name_entry.delete(0, tk.END)
                self.sender_name_entry.insert(0, config.get("sender_name", "Vincent"))
                self.cc_entry.insert(0, config.get("cc_email", ""))
                email_type = config.get("email_type", "QQ邮箱")
                if email_type in self.email_type['values']:
                    self.email_type.set(email_type)
                self.smtp_entry.delete(0, tk.END)
                self.smtp_entry.insert(0, config.get("smtp", "smtp.qq.com"))
                self.port_entry.delete(0, tk.END)
                self.port_entry.insert(0, config.get("port", "587"))
                self.use_ssl.set(config.get("use_ssl", False))
            except:
                pass
    
    def send_test_email(self):
        if not self.email_entry.get() or not self.password_entry.get():
            messagebox.showerror("错误", "请填写邮箱和密码！")
            return
        
        test_dialog = tk.Toplevel(self.root)
        test_dialog.title("发送测试邮件")
        test_dialog.geometry("350x120")
        
        ttk.Label(test_dialog, text="发送测试邮件到：").pack(pady=10)
        test_email_var = tk.StringVar()
        ttk.Entry(test_dialog, textvariable=test_email_var, width=35).pack(pady=5)
        
        def do_send():
            test_email = test_email_var.get().strip()
            if not test_email or '@' not in test_email:
                messagebox.showerror("错误", "请输入有效邮箱！")
                return
            test_dialog.destroy()
            thread = threading.Thread(target=self._send_test_email_thread, args=(test_email,))
            thread.daemon = True
            thread.start()
        
        ttk.Button(test_dialog, text="发送", command=do_send).pack(pady=10)
    
    def _send_test_email_thread(self, test_email):
        try:
            self.log_message(f"\n发送测试邮件到: {test_email}")
            
            if self.use_ssl.get():
                server = smtplib.SMTP_SSL(self.smtp_entry.get(), int(self.port_entry.get()), timeout=30)
            else:
                server = smtplib.SMTP(self.smtp_entry.get(), int(self.port_entry.get()), timeout=30)
                server.ehlo()
                server.starttls()
                server.ehlo()
            
            server.login(self.email_entry.get(), self.password_entry.get())
            
            subject = self.subject_entry.get() + " [测试]"
            body = self.body_text.get("1.0", tk.END).strip().replace("{name}", "测试用户")
            
            if self.send_single_email_with_server(server, self.email_entry.get(), self.sender_name_entry.get(), 
                                                   "", "测试用户", test_email, subject, body):
                self.log_message("✓ 测试邮件发送成功！")
                messagebox.showinfo("成功", f"测试邮件已发送到：{test_email}")
            else:
                self.log_message("✗ 测试邮件发送失败")
            
            server.quit()
        except Exception as e:
            self.log_message(f"✗ 错误: {str(e)}")
            messagebox.showerror("错误", str(e))
    
    def start_sending(self):
        if not self.email_entry.get() or not self.password_entry.get():
            messagebox.showerror("错误", "请填写邮箱和密码！")
            return
        if not self.file_path.get():
            messagebox.showerror("错误", "请选择Excel文件！")
            return
        
        self.send_button.config(state=tk.DISABLED)
        self.stop_button.config(state=tk.NORMAL)
        self.stop_flag = False
        
        self.log_text.config(state=tk.NORMAL)
        self.log_text.delete(1.0, tk.END)
        self.log_text.config(state=tk.DISABLED)
        
        thread = threading.Thread(target=self.send_emails_thread)
        thread.daemon = True
        thread.start()
    
    def stop_sending(self):
        self.stop_flag = True
        self.log_message("\n用户请求停止...")
    
    def send_emails_thread(self):
        try:
            df = pd.read_excel(self.file_path.get())
            df.columns = df.columns.str.strip()
            
            name_col = email_col = None
            for col in df.columns:
                cl = col.lower()
                if cl in ['name', '姓名', '名字', 'names']:
                    name_col = col
                if cl in ['email', 'e-mail', 'mail', '邮箱', 'emails']:
                    email_col = col
            
            if not name_col or not email_col:
                self.log_message(f"错误：找不到姓名或邮箱列！列名: {list(df.columns)}")
                messagebox.showerror("错误", "找不到姓名或邮箱列！")
                self.reset_buttons()
                return
            
            total = len(df)
            success_count = 0
            self.progress['maximum'] = total
            self.progress['value'] = 0
            
            sender_email = self.email_entry.get()
            sender_name = self.sender_name_entry.get()
            sender_password = self.password_entry.get()
            smtp_server = self.smtp_entry.get()
            smtp_port = int(self.port_entry.get())
            use_ssl = self.use_ssl.get()
            delay_min = float(self.delay_min_entry.get())
            delay_max = float(self.delay_max_entry.get())
            batch_size = int(self.batch_size_entry.get()) if self.enable_batch.get() else 999999
            batch_delay = float(self.batch_delay_entry.get()) * 60 if self.enable_batch.get() else 0
            
            subject = self.subject_entry.get()
            body_template = self.body_text.get("1.0", tk.END).strip()
            cc_email = self.cc_entry.get().strip()
            
            self.log_message(f"开始发送，共 {total} 人，{len(self.images)} 张图片")
            
            if use_ssl:
                server = smtplib.SMTP_SSL(smtp_server, smtp_port, timeout=30)
            else:
                server = smtplib.SMTP(smtp_server, smtp_port, timeout=30)
                server.ehlo()
                server.starttls()
                server.ehlo()
            
            server.login(sender_email, sender_password)
            self.log_message("✓ SMTP连接成功\n")
            
            batch_count = 0
            current_batch = 1
            
            for index, row in df.iterrows():
                if self.stop_flag:
                    self.log_message("\n已停止！")
                    break
                
                name = row[name_col]
                email = row[email_col]
                
                if pd.isna(name) or pd.isna(email):
                    continue
                
                if self.enable_batch.get() and batch_count > 0 and batch_count % batch_size == 0:
                    self.log_message(f"\n第{current_batch}批完成，等待{batch_delay/60:.0f}分钟...")
                    
                    for remaining in range(int(batch_delay), 0, -30):
                        if self.stop_flag:
                            break
                        self.status_label.config(text=f"批次间隔: {remaining//60}分{remaining%60}秒...")
                        time.sleep(min(30, remaining))
                    
                    if self.stop_flag:
                        break
                    
                    current_batch += 1
                    try:
                        server.quit()
                    except:
                        pass
                    
                    if use_ssl:
                        server = smtplib.SMTP_SSL(smtp_server, smtp_port, timeout=30)
                    else:
                        server = smtplib.SMTP(smtp_server, smtp_port, timeout=30)
                        server.ehlo()
                        server.starttls()
                        server.ehlo()
                    server.login(sender_email, sender_password)
                    self.log_message("✓ 重新连接成功\n")
                
                self.log_message(f"[{index+1}/{total}] {name} ({email})...")
                
                body = body_template.replace("{name}", str(name))
                
                if self.send_single_email_with_server(server, sender_email, sender_name, cc_email, str(name), str(email), subject, body):
                    success_count += 1
                    batch_count += 1
                    self.log_message("  ✓ 成功")
                else:
                    self.log_message("  ✗ 失败")
                
                self.progress['value'] = index + 1
                self.status_label.config(text=f"进度: {index+1}/{total}")
                
                if index < total - 1 and not self.stop_flag:
                    delay = random.uniform(delay_min, delay_max)
                    time.sleep(delay)
            
            try:
                server.quit()
            except:
                pass
            
            self.log_message(f"\n完成！成功: {success_count}/{total}")
            self.status_label.config(text=f"完成: {success_count}/{total}")
            messagebox.showinfo("完成", f"发送完成！\n成功: {success_count}\n总数: {total}")
            
        except Exception as e:
            self.log_message(f"\n错误: {str(e)}")
            messagebox.showerror("错误", str(e))
        finally:
            self.reset_buttons()
    
    def send_single_email_with_server(self, server, sender_email, sender_name, cc_email, recipient_name, recipient_email, subject, body):
        try:
            msg = MIMEMultipart('related')
            
            # 使用 email.utils.formataddr 正确格式化地址
            from email.utils import formataddr
            from email.header import Header
            
            # 发件人地址
            if sender_name:
                msg['From'] = formataddr((sender_name, sender_email))
            else:
                msg['From'] = sender_email
            
            # 收件人地址 - 正确处理中文名称
            if recipient_name:
                msg['To'] = formataddr((str(recipient_name), recipient_email))
            else:
                msg['To'] = recipient_email
            if cc_email:
                msg['Cc'] = cc_email
            
            msg['Subject'] = Header(subject, 'utf-8')
            msg['Reply-To'] = sender_email
            msg['Date'] = formatdate(localtime=True)
            msg['MIME-Version'] = '1.0'
            
            timestamp = int(time.time() * 1000)
            random_part = ''.join(random.choices('0123456789abcdef', k=12))
            domain = sender_email.split("@")[1]
            msg['Message-ID'] = f'<{random_part}.{timestamp}@{domain}>'
            
            mailers = ['Mozilla Thunderbird 115.0', 'Microsoft Outlook 16.0', 'Apple Mail (16.0)']
            msg['X-Mailer'] = random.choice(mailers)
            msg['X-Priority'] = '3'
            
            msg_alternative = MIMEMultipart('alternative')
            msg.attach(msg_alternative)
            
            invisible = ''.join(random.choices(['\u200B', '\u200C', '\u200D'], k=random.randint(1, 2)))
            body_unique = body + invisible
            
            text_body = body_unique
            for img_data in self.images:
                text_body = text_body.replace(f"{{{img_data['id']}}}", f"[图片: {img_data['filename']}]")
            
            text_part = MIMEText(text_body, 'plain', 'utf-8')
            msg_alternative.attach(text_part)
            
            html_content = self.build_html_body(body_unique)
            html_part = MIMEText(html_content, 'html', 'utf-8')
            msg_alternative.attach(html_part)
            
            for img_data in self.images:
                try:
                    with open(img_data['path'], 'rb') as f:
                        img_content = f.read()
                    
                    ext = os.path.splitext(img_data['path'])[1].lower()
                    subtype_map = {'.jpg': 'jpeg', '.jpeg': 'jpeg', '.png': 'png', '.gif': 'gif', '.bmp': 'bmp'}
                    subtype = subtype_map.get(ext, 'jpeg')
                    
                    img_part = MIMEImage(img_content, _subtype=subtype)
                    img_part.add_header('Content-ID', f'<{img_data["cid"]}>')
                    img_part.add_header('Content-Disposition', 'inline', filename=img_data['filename'])
                    msg.attach(img_part)
                except Exception as e:
                    self.log_message(f"  图片加载失败 {img_data['filename']}: {str(e)}")
            
            for att_data in self.attachments:
                try:
                    with open(att_data['path'], 'rb') as f:
                        att_content = f.read()
                    
                    att_part = MIMEBase('application', 'octet-stream')
                    att_part.set_payload(att_content)
                    encoders.encode_base64(att_part)
                    att_part.add_header('Content-Disposition', 'attachment', filename=att_data['filename'])
                    msg.attach(att_part)
                except Exception as e:
                    self.log_message(f"  附件加载失败 {att_data['filename']}: {str(e)}")
            
            recipients = [recipient_email]
            if cc_email:
                recipients.append(cc_email)
            
            server.sendmail(sender_email, recipients, msg.as_string())
            return True
            
        except Exception as e:
            self.log_message(f"  发送错误: {str(e)}")
            return False
    
    def reset_buttons(self):
        self.send_button.config(state=tk.NORMAL)
        self.stop_button.config(state=tk.DISABLED)


if __name__ == "__main__":
    root = tk.Tk()
    app = EmailSenderGUI(root)
    root.mainloop()