import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from ctypes import windll
import threading

# 确保使用逻辑像素而不是物理像素
windll.shcore.SetProcessDpiAwareness(1)

class Application(tk.Tk):
    def __init__(self, process_function):
        super().__init__()
        
        self.process_function = process_function
        self.progress_value = 0
        
        # 设置窗口标题和尺寸
        self.title("ChatExcel")
        self.geometry("1000x500")  # 宽度大于高度
        
        # 文件上传按钮
        self.upload_button = ttk.Button(self, text="上传文件", command=self.open_file)
        self.upload_button.place(relx=0.05, rely=0.05, anchor='w')
        
        # 显示文件路径的文本框
        self.file_path_label = ttk.Label(self, text="文件路径：")
        self.file_path_label.place(relx=0.05, rely=0.12, anchor='w')
        self.file_path_display = tk.Text(self, wrap=tk.WORD, height=1, width=60)
        self.file_path_display.place(relx=0.2, rely=0.12, anchor='w')
        self.file_path_display.config(state='disabled')  # 禁用编辑
        
        # 用户输入文本框
        self.user_input_label = ttk.Label(self, text="请输入内容：")
        self.user_input_label.place(relx=0.05, rely=0.20, anchor='w')
        self.user_input = ttk.Entry(self, width=50)
        self.user_input.place(relx=0.2, rely=0.20, anchor='w')
        
        # 发送按钮
        self.send_button = ttk.Button(self, text="发送", command=self.start_processing)
        self.send_button.place(relx=0.05, rely=0.30, anchor='w')
        
        # 进度条和当前步骤信息
        self.progress_and_step_frame = ttk.Frame(self)
        self.progress_and_step_frame.place(relx=0.2, rely=0.30, anchor='w')
        
        self.progress = ttk.Progressbar(self.progress_and_step_frame, orient="horizontal", length=200, mode="determinate")
        self.progress.pack(side=tk.LEFT, padx=5)
        
        self.current_step = ttk.Label(self.progress_and_step_frame, text="")
        self.current_step.pack(side=tk.RIGHT, padx=5)
        
        # AI输出文本框
        self.ai_output_label = ttk.Label(self, text="AI 输出：")
        self.ai_output_label.place(relx=0.05, rely=0.40, anchor='w')
        self.ai_output = tk.Text(self, wrap=tk.WORD, height=5, width=80)
        self.ai_output.place(relx=0.05, rely=0.47, anchor='nw')
        
        # 初始化进度条值
        self.progress["value"] = 0
        self.progress["maximum"] = 100
        
    def open_file(self):
        filename = filedialog.askopenfilename()
        if filename:
            # 更新文件路径文本框
            self.file_path_display.config(state='normal')
            self.file_path_display.delete('1.0', tk.END)
            self.file_path_display.insert(tk.END, filename)
            self.file_path_display.config(state='disabled')
            
    def update_progress(self, value):
        self.progress["value"] = value
        if value <= 10:
            self.current_step.config(text="请求体已构建好")
        if 10 < value and value <= 45:
            self.current_step.config(text="得到大模型回应")
        if 45 < value and value <= 65:
            self.current_step.config(text="程序正在调用函数处理")
        if 65 < value and value <= 100:
            self.current_step.config(text="处理完毕")
        self.update_idletasks()  # 更新GUI
    
    def start_processing(self):
        user_input_text = self.user_input.get().strip()
        file_path = self.file_path_display.get('1.0', tk.END).strip()
        
        if not user_input_text:
            messagebox.showwarning("警告", "请输入内容后再发送。")
            return
        
        if not file_path:
            messagebox.showwarning("警告", "请先上传文件。")
            return
        
        # 清空之前的输出
        self.ai_output.delete('1.0', tk.END)
        
        # 启动新线程来运行处理函数
        self.thread = threading.Thread(target=self.run_process, args=(user_input_text, file_path))
        self.thread.start()
        
    def run_process(self, input_text, file_path):
        result = self.process_function(input_text, file_path, self.update_progress)
        
        # 处理完成后更新UI
        self.after(0, lambda: self.on_process_complete(result))
    
    def on_process_complete(self, result):
        self.ai_output.insert(tk.END, f"{result}\n")
        messagebox.showinfo("通知", "已完成操作")

        # 清空用户输入
        self.user_input.delete(0, tk.END)