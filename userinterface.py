import sys
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QProgressBar, QLabel, QFileDialog, QLineEdit, QScrollArea
)
from PyQt5.QtCore import Qt, QTimer

class ChatWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("AI交互界面")
        self.setGeometry(100, 100, 800, 600)
        
        # 主部件
        self.central_widget = QWidget()
        self.setCentralWidget(self.central_widget)
        
        # 主布局
        self.main_layout = QVBoxLayout()
        self.main_layout.setSpacing(15)
        self.central_widget.setLayout(self.main_layout)

        # 警告标签（初始隐藏）
        self.warning_label = QLabel("")
        self.warning_label.setStyleSheet("background-color: #FFCCCC; color: #FF0000; padding: 10px;")
        self.warning_label.setVisible(False)
        self.main_layout.addWidget(self.warning_label)
        
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
        self.send_button.clicked.connect(self.send_message)
        self.input_layout.addWidget(self.send_button)
        
        self.main_layout.addLayout(self.input_layout)
        
        # 状态布局
        self.status_layout = QHBoxLayout()
        
        # 进度条
        self.progress_bar = QProgressBar()
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
        QTimer.singleShot(5000, self.hide_warning)  

    def upload_file(self):
        options = QFileDialog.Options()
        options |= QFileDialog.ReadOnly
        file_name, _ = QFileDialog.getOpenFileName(self, "选择文件", "", "Excel文件 (*.xlsx)", options=options)
        if file_name:
            self.file_path_textfield.setText(file_name)
    
    def send_message(self):
        user_text = self.user_input_textfield.text().strip()
        if user_text:
            # 显示用户消息
            self.add_message(user_text, sender="user")
            self.user_input_textfield.clear()
            
            # 模拟AI回复
            self.add_message("AI正在处理您的请求...", sender="ai")
            
            # 这里可以添加实际的AI处理逻辑
        else:
            self.show_warning("用户输入不能为空")
    
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
            message_label.setStyleSheet("background-color: #FFFFFF; padding: 5px; border-radius: 5px;")
            message_layout.addWidget(message_label, alignment=Qt.AlignRight)
        
        self.messages_layout.addLayout(message_layout)
        
        # 自动滚动到底部
        self.scroll_area.verticalScrollBar().setValue(self.scroll_area.verticalScrollBar().maximum())

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = ChatWindow()
    window.show()
    sys.exit(app.exec_())