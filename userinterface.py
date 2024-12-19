import sys
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QComboBox,
    QPushButton, QProgressBar, QLabel, QFileDialog, QLineEdit, QScrollArea
)
from PyQt5.QtCore import Qt, QTimer, QObject, QThread, pyqtSignal
import time
from utils import OUTPUT_STATE

class Worker(QObject):
    progress = pyqtSignal(int, str)
    finished = pyqtSignal(str, object)

    def __init__(self, user_input: str, file_path: str, function_called: str, func):
        super().__init__()
        self.user_input = user_input
        self.file_path = file_path
        self.function_called = function_called
        self.func = func

    def run(self):
        try:
            result, state = self.func(
                self.user_input,
                self.file_path,
                self.function_called,
                self.progress_callback
            )
            self.finished.emit(result, state)
        except Exception as e:
            self.finished.emit(str(e), OUTPUT_STATE.MESSAGE_SUCCESS)

    def progress_callback(self, percentage: int, step: str):
        self.progress.emit(percentage, step)


class Application(QMainWindow):
    def __init__(self ,func):
        super().__init__()

        self.func = func

        self.setWindowTitle("AI交互界面")
        self.setGeometry(100, 100, 800, 600)
        
        # 主部件
        self.central_widget = QWidget()
        self.setCentralWidget(self.central_widget)
        
        # 主布局
        self.main_layout = QVBoxLayout()
        self.main_layout.setSpacing(15)
        self.central_widget.setLayout(self.main_layout)

        # 警告标签
        self.warning_label = QLabel("")
        self.warning_label.setStyleSheet("background-color:rgb(255, 0, 0); color:rgb(0, 0, 0); padding: 10px;")
        self.warning_label.setVisible(False)
        self.main_layout.addWidget(self.warning_label)

        # 提醒标签
        self.information_label =QLabel("")
        self.information_label.setStyleSheet("background-color:rgb(4, 255, 0); color:rgb(0, 0, 0); padding: 10px;")
        self.information_label.setVisible(False)
        self.main_layout.addWidget(self.information_label)
        
        # 消息显示区域
        self.scroll_area = QScrollArea()
        self.scroll_area.setWidgetResizable(True)
        self.messages_widget = QWidget()
        self.messages_layout = QVBoxLayout()
        self.messages_layout.setAlignment(Qt.AlignTop)
        self.messages_widget.setLayout(self.messages_layout)
        self.scroll_area.setWidget(self.messages_widget)
        self.main_layout.addWidget(self.scroll_area)
        
        # 文件上传区域布局
        self.file_layout = QHBoxLayout()

        # 文件上传按钮
        self.upload_button = QPushButton("上传文件")
        self.upload_button.clicked.connect(self.upload_file)
        self.file_layout.addWidget(self.upload_button)
        
        # 显示文件路径的文本框
        self.file_path_textfield = QLineEdit()
        self.file_path_textfield.setReadOnly(True)
        self.file_layout.addWidget(self.file_path_textfield)

        self.main_layout.addLayout(self.file_layout)

         # 用户输入区域布局
        self.input_layout = QHBoxLayout()
        
        # 用户输入文本框
        self.user_input_textfield = QLineEdit()
        self.user_input_textfield.setPlaceholderText("请输入您的消息...")
        self.input_layout.addWidget(self.user_input_textfield)
        
        # 发送按钮
        self.send_button = QPushButton("发送")
        self.send_button.clicked.connect(self.start_progress)
        self.input_layout.addWidget(self.send_button)

        # 函数选择下拉列表
        self.func_select = QComboBox()
        func_opera = ["read", "write", "auto", "none"]
        self.func_select.addItems(func_opera)
        self.input_layout.addWidget(self.func_select)
        
        self.main_layout.addLayout(self.input_layout)
        
        # 状态布局
        self.status_layout = QHBoxLayout()
        
        # 进度条
        self.progress_bar = QProgressBar()
        self.progress_bar.setMinimum(0)
        self.progress_bar.setMaximum(100)
        self.progress_bar.setValue(0)
        self.status_layout.addWidget(self.progress_bar)
        
        # 状态文本
        self.status_label = QLabel("状态：空闲")
        self.status_layout.addWidget(self.status_label)
        
        self.main_layout.addLayout(self.status_layout)
    
    def hide_warning(self):
        self.warning_label.setVisible(False)
    
    def show_warning(self, message: str):
        self.warning_label.setText(message)
        self.warning_label.setVisible(True)
        QTimer.singleShot(1000, self.hide_warning)  
    
    def hide_information(self):
        self.information_label.setVisible(False)
    
    def show_information(self, message: str):
        self.information_label.setText(message)
        self.information_label.setVisible(True)
        QTimer.singleShot(1000, self.hide_information)  

    def upload_file(self):
        options = QFileDialog.Options()
        options |= QFileDialog.ReadOnly
        file_name, _ = QFileDialog.getOpenFileName(self, "选择文件", "", "Excel文件 (*.xlsx)", options=options)
        if file_name:
            self.file_path_textfield.setText(file_name)
    
    def update_progress(self, percentage: int, step: str):
        self.progress_bar.setValue(percentage)
        self.status_label.setText(f"{step}")
        # self.addmessage(f"进度更新: {percentage}% - 步骤: {step}")

    def on_finished(self, result: str, state: OUTPUT_STATE):
        self.send_button.setEnabled(True)
        if state == OUTPUT_STATE.FUNCTION_CALLED_SUCCESS:
            self.progress_bar.setValue(100)
            self.status_label.setText("状态：完成")
            self.add_message(f"结果: {result}", "ai")
            self.show_information("调用函数成功")
        elif state == OUTPUT_STATE.FUNCTION_CALLED_FAIL:
            self.progress_bar.setValue(100)
            self.status_label.setText("状态：完成")
            self.add_message(f"结果: {result} (失败)", "ai")
            
        elif state == OUTPUT_STATE.INCORRECT_PARAMETER:
            self.progress_bar.setValue(0)
            self.status_label.setText("状态：空闲")
            self.add_message(f"结果: {result} (参数错误)", "ai")
            
        elif state == OUTPUT_STATE.FILE_PATH_ERROR:
            self.progress_bar.setValue(0)
            self.status_label.setText("状态：空闲")
            self.add_message(f"结果: {result} (文件路径错误)", "ai")
            
        elif state == OUTPUT_STATE.MESSAGE_SUCCESS:
            self.progress_bar.setValue(0)
            self.status_label.setText("状态：空闲")
            self.add_message(f"结果: {result} (消息)", "ai")
            
        else:
            self.progress_bar.setValue(0)
            self.status_label.setText("状态：空闲")
            self.add_message(f"结果: {result} (未知状态)", "ai")
            

    def start_progress(self):
        user_text = self.user_input_textfield.text().strip()
        if not user_text:
            self.show_warning("用户输入不能为空")
            return
        
        file_path = self.file_path_textfield.text().strip()
        if not file_path:
            self.show_warning("文件不能为空")
            return
        
        self.send_button.setEnabled(False)
        
        # 显示用户消息
        self.add_message(user_text, sender="user")
        self.user_input_textfield.clear()

        self.thread = QThread()
        self.worker = Worker(
            user_input = user_text,
            file_path = file_path,
            function_called = self.func_select.currentText(),
            func = self.func
        )
        self.worker.moveToThread(self.thread)
        # 连接信号和槽
        self.thread.started.connect(self.worker.run)
        self.worker.progress.connect(self.update_progress)
        self.worker.finished.connect(self.on_finished)
        self.worker.finished.connect(self.thread.quit)
        self.worker.finished.connect(self.worker.deleteLater)
        self.thread.finished.connect(self.thread.deleteLater)
        
        # 启动线程
        self.thread.start()

        # # 模拟AI回复
        # self.add_message("AI正在处理您的请求...", sender="ai")

        # self.show_information("处理成功")
        
        # 这里可以添加实际的AI处理逻辑
    
    def add_message(self, text, sender="user"):
        message_layout = QHBoxLayout()
        if sender == "user":
            # 用户消息左对齐
            message_label = QLabel(text)
            message_label.setStyleSheet("background-color: #DCF8C6; padding: 5px; border-radius: 5px;")
            message_layout.addWidget(message_label, alignment=Qt.AlignLeft)
        elif sender == "ai":
            # AI消息右对齐
            message_label = QLabel(text)
            message_label.setStyleSheet("background-color:rgb(145, 127, 127); padding: 5px; border-radius: 5px;")
            message_layout.addWidget(message_label, alignment=Qt.AlignRight)
        
        self.messages_layout.addLayout(message_layout)
        
        # 自动滚动到底部
        self.scroll_area.verticalScrollBar().setValue(self.scroll_area.verticalScrollBar().maximum())