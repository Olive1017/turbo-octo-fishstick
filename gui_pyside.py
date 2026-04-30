import os
import time
import threading
import sys
from PySide6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout,
                                 QHBoxLayout, QLabel, QPushButton, QLineEdit,
                                 QCheckBox, QFrame, QMessageBox, QFileDialog, QTextEdit)
from PySide6.QtCore import Qt, Signal, QObject
from PySide6.QtGui import QTextCursor

# 延迟导入，避免启动时加载可能影响UI的库
PrintingManager = None


class LogRedirect:
    """将标准输出重定向到GUI日志窗口"""
    def __init__(self, signal_emitter):
        self.signal_emitter = signal_emitter

    def write(self, text):
        if text.strip():  # 只写入非空内容
            # 通过信号发射日志，而不是直接操作GUI
            self.signal_emitter.log_signal.emit(text.strip())

    def flush(self):
        pass


class SignalEmitter(QObject):
    """信号发射器，用于跨线程通信"""
    show_message = Signal(str, str, str)  # 标题、内容、类型（info/error）
    update_status_signal = Signal(str)  # 更新状态
    log_signal = Signal(str)  # 日志输出


# 更新状态


class PrintingApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("壳牌自动打印工具")
        self.resize(700, 520)

        # 默认账号密码
        self.default_username = "LUOJIAN"
        self.default_password = "Aa123456@@@"

        # 状态变量
        self.save_path = None
        self.use_default_auth = True
        self.is_running = False
        self.should_continue = False
        self.incremental_mode = False  # 增量下载模式

        # 打印管理器
        self.print_manager = None

        # 创建信号发射器
        self.signal_emitter = SignalEmitter()
        self.signal_emitter.show_message.connect(self.handle_show_message)
        self.signal_emitter.update_status_signal.connect(self.update_status)

        # 创建界面
        self.create_widgets()

        # 连接日志信号
        self.signal_emitter.log_signal.connect(self.handle_log)

        # 重定向标准输出到日志窗口
        self.log_redirect = LogRedirect(self.signal_emitter)
        sys.stdout = self.log_redirect

    def create_widgets(self):
        # 主窗口部件
        central_widget = QWidget()
        self.setCentralWidget(central_widget)

        # 主布局
        main_layout = QVBoxLayout()
        main_layout.setSpacing(15)
        main_layout.setContentsMargins(20, 20, 20, 20)
        central_widget.setLayout(main_layout)

        # 标题
        title_label = QLabel("壳牌自动打印工具")
        title_label.setStyleSheet("font-size: 24px; font-weight: bold; color: #2196F3;")
        title_label.setAlignment(Qt.AlignCenter)
        main_layout.addWidget(title_label)

        # 注意事项框架
        notice_frame = QFrame()
        notice_frame.setFrameShape(QFrame.StyledPanel)
        notice_layout = QVBoxLayout()
        notice_frame.setLayout(notice_layout)

        notice_title = QLabel("⚠️ 注意事项")
        notice_title.setStyleSheet("font-size: 16px; font-weight: bold;")
        notice_layout.addWidget(notice_title)

        notice_text = """
• 请确保已安装Edge浏览器或谷歌浏览器。
• 程序自动打印时，请勿人工操作电脑。
• 使用某一账号登录系统自动打印时，不可以在其他电脑同时登录。
"""
        notice_label = QLabel(notice_text)
        notice_label.setStyleSheet("font-size: 12px; color: #555;")
        notice_layout.addWidget(notice_label)

        main_layout.addWidget(notice_frame)

        # 保存路径和账号密码框架（横向布局）
        settings_layout = QHBoxLayout()
        settings_layout.setSpacing(15)

        # 账号密码框架（左边）
        auth_frame = QFrame()
        auth_frame.setFrameShape(QFrame.StyledPanel)
        auth_layout = QVBoxLayout()
        auth_frame.setLayout(auth_layout)

        auth_label = QLabel("账号密码:")
        auth_label.setStyleSheet("font-size: 14px; font-weight: bold;")
        auth_layout.addWidget(auth_label)

        # 使用默认账号密码复选框
        self.use_default_check = QCheckBox("使用默认账号密码")
        self.use_default_check.setChecked(True)
        self.use_default_check.stateChanged.connect(self.toggle_auth_inputs)
        self.use_default_check.setStyleSheet("font-size: 12px;")
        auth_layout.addWidget(self.use_default_check)

        # 账号输入
        username_layout = QHBoxLayout()
        username_label = QLabel("账号:")
        username_label.setFixedWidth(50)
        self.username_entry = QLineEdit(self.default_username)
        self.username_entry.setEnabled(False)
        self.username_entry.setStyleSheet("padding: 8px; border: 1px solid #ccc; border-radius: 4px;")
        username_layout.addWidget(username_label)
        username_layout.addWidget(self.username_entry)
        auth_layout.addLayout(username_layout)

        # 密码输入
        password_layout = QHBoxLayout()
        password_label = QLabel("密码:")
        password_label.setFixedWidth(50)
        self.password_entry = QLineEdit(self.default_password)
        self.password_entry.setEchoMode(QLineEdit.Password)
        self.password_entry.setEnabled(False)
        self.password_entry.setStyleSheet("padding: 8px; border: 1px solid #ccc; border-radius: 4px;")
        password_layout.addWidget(password_label)
        password_layout.addWidget(self.password_entry)
        auth_layout.addLayout(password_layout)

        settings_layout.addWidget(auth_frame, 1)

        # 保存路径框架（右边）
        path_frame = QFrame()
        path_frame.setFrameShape(QFrame.StyledPanel)
        path_layout = QVBoxLayout()
        path_frame.setLayout(path_layout)

        path_label = QLabel("保存路径:")
        path_label.setStyleSheet("font-size: 14px; font-weight: bold;")
        path_layout.addWidget(path_label)

        path_input_layout = QHBoxLayout()

        self.path_entry = QLineEdit()
        self.path_entry.setStyleSheet("padding: 10px; border: 1px solid #ccc; border-radius: 4px; font-size: 13px;")
        self.path_entry.setPlaceholderText("请点击浏览按钮选择保存路径")
        self.path_entry.setReadOnly(True)
        path_input_layout.addWidget(self.path_entry)

        browse_button = QPushButton("浏览")
        browse_button.setStyleSheet("""
            QPushButton {
                background-color: #2196F3;
                color: white;
                border: none;
                padding: 10px 16px;
                border-radius: 4px;
                font-weight: bold;
                font-size: 13px;
            }
            QPushButton:hover {
                background-color: #1976D2;
            }
        """)
        browse_button.clicked.connect(self.browse_folder)
        browse_button.setFixedWidth(80)
        path_input_layout.addWidget(browse_button)

        path_layout.addLayout(path_input_layout)

        # 增量下载模式复选框
        self.incremental_check = QCheckBox("增量下载模式（仅下载新订单）")
        self.incremental_check.setChecked(False)
        self.incremental_check.setStyleSheet("font-size: 12px;")
        self.incremental_check.stateChanged.connect(lambda: setattr(self, 'incremental_mode', self.incremental_check.isChecked()))
        path_layout.addWidget(self.incremental_check)

        settings_layout.addWidget(path_frame, 1)
        main_layout.addLayout(settings_layout)

        # 按钮框架
        button_layout = QHBoxLayout()

        self.start_button = QPushButton("开始打印")
        self.start_button.setStyleSheet("""
            QPushButton {
                background-color: #2196F3;
                color: white;
                border: none;
                padding: 12px;
                border-radius: 4px;
                font-size: 14px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #1976D2;
            }
            QPushButton:disabled {
                background-color: #BDBDBD;
            }
        """)
        self.start_button.clicked.connect(self.start_printing)
        button_layout.addWidget(self.start_button)

        self.continue_button = QPushButton("继续")
        self.continue_button.setStyleSheet("""
            QPushButton {
                background-color: #4CAF50;
                color: white;
                border: none;
                padding: 12px;
                border-radius: 4px;
                font-size: 14px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #388E3C;
            }
            QPushButton:disabled {
                background-color: #BDBDBD;
            }
        """)
        self.continue_button.clicked.connect(self.continue_printing)
        self.continue_button.setEnabled(False)
        button_layout.addWidget(self.continue_button)

        exit_button = QPushButton("退出")
        exit_button.setStyleSheet("""
            QPushButton {
                background-color: #F44336;
                color: white;
                border: none;
                padding: 12px;
                border-radius: 4px;
                font-size: 14px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #D32F2F;
            }
        """)
        exit_button.clicked.connect(self.close)
        button_layout.addWidget(exit_button)

        main_layout.addLayout(button_layout)

        # 日志输出窗口
        log_frame = QFrame()
        log_frame.setFrameShape(QFrame.StyledPanel)
        log_layout = QVBoxLayout()
        log_frame.setLayout(log_layout)

        log_label = QLabel("运行日志:")
        log_label.setStyleSheet("font-size: 14px; font-weight: bold;")
        log_layout.addWidget(log_label)

        self.log_output = QTextEdit()
        self.log_output.setReadOnly(True)
        self.log_output.setStyleSheet("""
            QTextEdit {
                background-color: #f5f5f5;
                border: 1px solid #ccc;
                border-radius: 4px;
                font-family: Consolas, 'Courier New', monospace;
                font-size: 11px;
                padding: 8px;
            }
        """)
        self.log_output.setMinimumHeight(120)
        self.log_output.setMaximumHeight(150)
        log_layout.addWidget(self.log_output)

        main_layout.addWidget(log_frame)

        # 状态标签
        self.status_label = QLabel("就绪")
        self.status_label.setStyleSheet("font-size: 12px; color: #666;")
        self.status_label.setAlignment(Qt.AlignCenter)
        main_layout.addWidget(self.status_label)

    def browse_folder(self):
        """选择保存路径"""
        print("开始浏览文件夹...")
        try:
            # 获取初始路径
            initial_path = self.path_entry.text()
            if os.path.exists(initial_path):
                start_dir = initial_path
            else:
                start_dir = ""

            print(f"打开文件夹对话框，起始路径: {start_dir}")

            # 打开文件夹选择对话框
            folder_path = QFileDialog.getExistingDirectory(
                self,
                "选择保存路径",
                start_dir,
                QFileDialog.ShowDirsOnly | QFileDialog.DontResolveSymlinks
            )

            print(f"选择的路径: {folder_path}")

            if folder_path:
                self.path_entry.setText(folder_path)
                print("路径已更新")

        except Exception as e:
            print(f"浏览文件夹出错: {e}")
            import traceback
            traceback.print_exc()
            QMessageBox.warning(self, "警告", f"选择文件夹失败：{str(e)}")

    def toggle_auth_inputs(self):
        """切换账号密码输入框的启用/禁用状态"""
        use_default = self.use_default_check.isChecked()
        self.username_entry.setEnabled(not use_default)
        self.password_entry.setEnabled(not use_default)

    def handle_show_message(self, title, content, msg_type):
        """在主线程中显示消息框"""
        if msg_type == "info":
            QMessageBox.information(self, title, content)
        elif msg_type == "error":
            QMessageBox.critical(self, title, content)

    def update_status(self, message):
        """更新状态标签"""
        self.status_label.setText(message)
        QApplication.processEvents()

    def handle_log(self, text):
        """在主线程中添加日志到日志窗口"""
        self.log_output.append(text)

    def start_printing(self):
        """开始打印流程"""
        # 延迟导入，只在需要时才加载
        global PrintingManager
        if PrintingManager is None:
            from printer import PrintingManager as PM
            PrintingManager = PM

        # 获取用户选择的保存路径
        self.save_path = self.path_entry.text().strip()
        self.is_running = True
        self.should_continue = False

        # 验证保存路径
        if not self.save_path:
            QMessageBox.critical(self, "错误", "请选择保存路径！")
            return

        if not os.path.exists(self.save_path):
            QMessageBox.critical(self, "错误", f"保存路径不存在：\n{self.save_path}")
            return

        # 获取账号密码
        if self.use_default_check.isChecked():
            username = self.default_username
            password = self.default_password
        else:
            username = self.username_entry.text().strip()
            password = self.password_entry.text().strip()
            if not username or not password:
                QMessageBox.critical(self, "错误", "请输入账号和密码！")
                return

        # 禁用开始按钮，启用继续按钮
        self.start_button.setEnabled(False)
        self.path_entry.setEnabled(False)
        self.continue_button.setEnabled(True)

        # 在单个线程中运行所有Playwright操作
        self.printing_thread = threading.Thread(
            target=self.run_all_operations,
            args=(username, password),
            daemon=True
        )
        self.printing_thread.start()

    def run_all_operations(self, username, password):
        """在同一个线程中运行所有Playwright操作，避免跨线程问题"""
        try:
            # 第一步：启动浏览器并登录
            self.signal_emitter.update_status_signal.emit("正在启动浏览器...")

            self.print_manager = PrintingManager()

            if not self.print_manager.start_browser_and_login(username, password, browser_type="chromium"):
                self.on_exit()
                return

            self.signal_emitter.update_status_signal.emit("请在浏览器中查询所需订单，然后点击'继续'按钮")

            # 等待用户点击继续按钮
            while not self.should_continue:
                time.sleep(0.5)

            # 第二步：根据模式选择打印流程
            if self.incremental_mode:
                # 增量模式：先获取新订单列表
                self.signal_emitter.update_status_signal.emit("正在扫描需要下载的新订单...")

                from utils import get_new_orders
                new_orders = get_new_orders(self.save_path, self.print_manager.page)

                if not new_orders:
                    self.signal_emitter.update_status_signal.emit("没有需要下载的新订单")
                    self.signal_emitter.show_message.emit("提示", "没有需要下载的新订单！\n所有订单都已存在。", "info")
                    self.on_exit()
                    return

                self.signal_emitter.update_status_signal.emit(f"开始增量下载，共 {len(new_orders)} 个新订单...")

                # 重新导航到装运单页面（因为扫描过程可能改变了页面状态）
                self.print_manager.page.locator("a").filter(has_text="装运单").click()
                time.sleep(2)

                # 执行增量下载
                result = self.print_manager.start_printing_incremental(self.save_path, new_orders)
            else:
                # 全量模式：下载所有订单
                self.signal_emitter.update_status_signal.emit("开始自动下载流程...")

                result = self.print_manager.start_printing(self.save_path)

            # 完成提示
            result_message = (
                f"全部处理完成！\n\n"
                f"总页数: {result['total_pages']}\n"
                f"总订单数(TO): {result['total_orders']}\n"
                f"成功处理: {result['success_orders']}\n"
                f"失败订单: {result['failed_orders']}"
            )

            self.signal_emitter.update_status_signal.emit("处理完成！")
            self.signal_emitter.show_message.emit("完成", result_message, "info")

        except Exception as e:
            self.signal_emitter.update_status_signal.emit(f"处理失败: {str(e)}")
            self.signal_emitter.show_message.emit("错误", f"处理失败:\n{str(e)}", "error")

        finally:
            self.on_exit()

    def continue_printing(self):
        """继续打印流程"""
        self.continue_button.setEnabled(False)
        self.should_continue = True

        # 最小化窗口，让浏览器窗口成为活动窗口
        self.showMinimized()

    def on_exit(self):
        """退出程序"""
        try:
            if self.print_manager:
                self.print_manager.close()
        except:
            pass

        # 恢复窗口显示
        self.showNormal()
        self.raise_()
        self.activateWindow()

        self.start_button.setEnabled(True)
        self.path_entry.setEnabled(True)
        self.continue_button.setEnabled(False)
        self.signal_emitter.update_status_signal.emit("就绪")

        # 清空日志（可选）
        # self.log_output.clear()

    def closeEvent(self, event):
        """窗口关闭事件"""
        # 恢复标准输出
        sys.stdout = sys.__stdout__
        try:
            if self.print_manager:
                self.print_manager.close()
        except:
            pass
        event.accept()


def main():
    app = QApplication()

    # 设置应用样式
    app.setStyle('Fusion')

    window = PrintingApp()
    window.show()

    app.exec()


if __name__ == "__main__":
    main()
