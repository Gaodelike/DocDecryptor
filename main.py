import sys
import logging
import os
from pathlib import Path
from typing import Optional

from PyQt5.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QHBoxLayout, QLabel, QLineEdit,
    QPushButton, QTextEdit, QFileDialog, QMessageBox, QFrame, QStyleFactory,
    QListWidget, QProgressBar
)
from PyQt5.QtCore import Qt, pyqtSignal, QRunnable, QThreadPool, QObject
from PyQt5.QtGui import QTextCursor, QTextCharFormat, QColor

# 从我们定义的核心逻辑文件导入函数
from OpenSesame import process_file_unified, dummy_decrypt_function, ESAFENET_MARKER
# 导入信号代理类
from worker_signals import WorkerSignals


# --- 辅助类：日志重定向 ---
class QTextEditLogger(QObject, logging.Handler):  # <<< 修改：继承 QObject 和 logging.Handler
    # 定义不同日志级别的颜色
    LEVEL_COLORS = {
        logging.INFO: QColor("black"),
        logging.WARNING: QColor("orange"),
        logging.ERROR: QColor("red"),
        logging.CRITICAL: QColor("darkred"),
        logging.DEBUG: QColor("gray"),
    }

    log_record_signal = pyqtSignal(str, int)  # (message, levelno)

    def __init__(self, parent_text_edit, parent=None):  # <<< 修改：增加 parent 参数用于QObject
        super().__init__(parent)  # <<< 修改：初始化 QObject
        logging.Handler.__init__(self)  # <<< 修改：手动初始化 logging.Handler

        self.widget = parent_text_edit
        self.widget.setReadOnly(True)

        self.log_record_signal.connect(self._append_log_message)

    def emit(self, record):
        # 在任何线程中，当日志产生时，发出信号
        msg = self.format(record)
        self.log_record_signal.emit(msg, record.levelno)

    def _append_log_message(self, msg: str, levelno: int):
        # 这个方法只在主线程中执行，安全地更新 GUI
        format = QTextCharFormat()
        format.setForeground(self.LEVEL_COLORS.get(levelno, QColor("black")))

        self.widget.moveCursor(QTextCursor.End)
        self.widget.setCurrentCharFormat(format)
        self.widget.insertPlainText(msg + '\n')
        self.widget.verticalScrollBar().setValue(self.widget.verticalScrollBar().maximum())  # 自动滚动到底部


# --- 辅助类：文件处理线程（QRunnable） ---
class FileProcessorRunnable(QRunnable):
    def __init__(self, original_file_path_str: str, output_dir_str: str):
        super().__init__()
        self.original_file_path_str = original_file_path_str
        self.output_dir_str = output_dir_str
        self.setAutoDelete(True)
        self.signals = WorkerSignals()

    def run(self):
        logging.info(f"开始处理文件: {self.original_file_path_str}")
        try:
            self.signals.status_update.emit(f"正在处理: {os.path.basename(self.original_file_path_str)}")

            result_path = process_file_unified(self.original_file_path_str, self.output_dir_str)

            self.signals.finished.emit(self.original_file_path_str, result_path)
            logging.info(f"文件处理完成: {self.original_file_path_str}")
        except ImportError as ie:
            logging.error(f"处理文件 {self.original_file_path_str} 失败：缺少依赖库。", exc_info=True)
            self.signals.error.emit(self.original_file_path_str, f"处理失败：缺少依赖库。\n请安装相应的库：{ie}")
        except Exception as e:
            logging.exception(f"处理文件 {self.original_file_path_str} 时发生未预期的错误: {e}")
            self.signals.error.emit(self.original_file_path_str, f"处理失败：发生未知错误。\n详细信息请查看日志。")


# --- 主应用程序类 ---
class FileProcessorApp(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("文件处理工具")
        self.setGeometry(100, 100, 800, 650)
        self.init_ui()

        self.thread_pool = QThreadPool.globalInstance()
        self.thread_pool.setMaxThreadCount(os.cpu_count() or 4)

        self.processing_files_count = 0
        self.processed_files_count = 0
        self.total_files_to_process = 0

    def init_ui(self):
        QApplication.setStyle(QStyleFactory.create('Fusion'))

        main_layout = QVBoxLayout()

        # --- 输入文件选择区域 ---
        input_file_frame = QFrame(self)
        input_file_frame.setFrameShape(QFrame.StyledPanel)
        input_file_layout = QHBoxLayout(input_file_frame)

        input_file_layout.addWidget(QLabel("选择文件:"))
        self.browse_input_button = QPushButton("浏览...")
        self.browse_input_button.clicked.connect(self.browse_input_files)
        input_file_layout.addWidget(self.browse_input_button)
        main_layout.addWidget(input_file_frame)

        # --- 输出目录选择区域 ---
        output_dir_frame = QFrame(self)
        output_dir_frame.setFrameShape(QFrame.StyledPanel)
        output_dir_layout = QHBoxLayout(output_dir_frame)

        output_dir_layout.addWidget(QLabel("输出目录:"))
        default_output_path = Path.home() / "Processed_Files"
        default_output_path.mkdir(parents=True, exist_ok=True)
        self.output_dir_input = QLineEdit(str(default_output_path))
        output_dir_layout.addWidget(self.output_dir_input)
        self.browse_output_button = QPushButton("选择目录...")
        self.browse_output_button.clicked.connect(self.browse_output_directory)
        output_dir_layout.addWidget(self.browse_output_button)
        main_layout.addWidget(output_dir_frame)

        # --- 文件列表区域 ---
        file_list_frame = QFrame(self)
        file_list_frame.setFrameShape(QFrame.StyledPanel)
        file_list_layout = QVBoxLayout(file_list_frame)
        file_list_layout.addWidget(QLabel("待处理文件列表:"))
        self.file_list_widget = QListWidget()
        self.file_list_widget.setMinimumHeight(120)
        file_list_layout.addWidget(self.file_list_widget)

        file_list_buttons_layout = QHBoxLayout()
        self.clear_list_button = QPushButton("清空列表")
        self.clear_list_button.clicked.connect(self.file_list_widget.clear)
        file_list_buttons_layout.addWidget(self.clear_list_button)
        file_list_buttons_layout.addStretch(1)  # 撑开，让按钮靠左
        file_list_layout.addLayout(file_list_buttons_layout)
        main_layout.addWidget(file_list_frame)

        # --- 按钮区域 ---
        button_layout = QHBoxLayout()
        self.process_button = QPushButton("开始处理所有文件")
        self.process_button.clicked.connect(self.start_multiple_file_processing)
        self.process_button.setEnabled(False)
        button_layout.addWidget(self.process_button)

        # “模拟在线解密”按钮，只针对非标准文件（.info）
        self.online_decrypt_button = QPushButton("模拟在线解密 (针对 .info 文件)")
        self.online_decrypt_button.clicked.connect(self.simulate_online_decryption)
        self.online_decrypt_button.setEnabled(True)  # 用户可随时选择 .info 文件进行解密
        button_layout.addWidget(self.online_decrypt_button)

        # >>> 移除“恢复原始后缀”按钮
        # self.restore_suffix_button = QPushButton("恢复原始后缀 (.info -> 原后缀)")
        # self.restore_suffix_button.clicked.connect(self.restore_original_suffix)
        # self.restore_suffix_button.setEnabled(True)
        # button_layout.addWidget(self.restore_suffix_button)
        # <<< 移除结束

        main_layout.addLayout(button_layout)

        # --- 进度条和状态区域 ---
        progress_status_frame = QFrame(self)
        progress_status_frame.setFrameShape(QFrame.StyledPanel)
        progress_status_layout = QVBoxLayout(progress_status_frame)

        self.status_label = QLabel("状态: 等待选择文件...")
        progress_status_layout.addWidget(self.status_label)

        self.overall_progress_bar = QProgressBar(self)
        self.overall_progress_bar.setAlignment(Qt.AlignCenter)
        self.overall_progress_bar.setTextVisible(True)
        self.overall_progress_bar.setValue(0)
        progress_status_layout.addWidget(self.overall_progress_bar)

        main_layout.addWidget(progress_status_frame)

        # --- 日志区域 ---
        log_frame = QFrame(self)
        log_frame.setFrameShape(QFrame.StyledPanel)
        log_layout = QVBoxLayout(log_frame)

        log_layout.addWidget(QLabel("日志:"))
        self.log_text_edit = QTextEdit()
        self.log_text_edit.setMinimumHeight(150)
        log_layout.addWidget(self.log_text_edit)

        main_layout.addWidget(log_frame)

        self.setLayout(main_layout)

        self.log_handler = QTextEditLogger(self.log_text_edit)
        logging.getLogger().addHandler(self.log_handler)
        logging.getLogger().setLevel(logging.INFO)

    def browse_input_files(self):
        file_paths, _ = QFileDialog.getOpenFileNames(
            self, "选择要处理的文件", "",
            "所有文件 (*.*);;"
            "Office 文档 (*.docx *.pdf *.xlsx *.pptx);;"
            "图像文件 (*.jpg *.jpeg *.png *.bmp *.gif);;"
            "CSV 文件 (*.csv);;"
            "二进制文件 (*.bin *.dat)"
        )
        if file_paths:
            self.file_list_widget.clear()
            self.processed_files_count = 0
            self.overall_progress_bar.setValue(0)

            for path in file_paths:
                self.file_list_widget.addItem(path)
            self.update_status(f"已选择 {len(file_paths)} 个文件。", "black")
            self.process_button.setEnabled(True)
            self.log_text_edit.clear()

    def browse_output_directory(self):
        initial_dir = self.output_dir_input.text()
        if not initial_dir or not Path(initial_dir).is_dir():
            initial_dir = str(Path.home())

        dir_path = QFileDialog.getExistingDirectory(self, "选择输出目录", initial_dir)
        if dir_path:
            self.output_dir_input.setText(dir_path)
            Path(dir_path).mkdir(parents=True, exist_ok=True)

    def start_multiple_file_processing(self):
        if self.file_list_widget.count() == 0:
            QMessageBox.warning(self, "警告", "请先选择要处理的文件！")
            return

        output_dir_str = self.output_dir_input.text()
        if not output_dir_str:
            QMessageBox.warning(self, "警告", "请指定一个输出目录！")
            return

        try:
            Path(output_dir_str).mkdir(parents=True, exist_ok=True)
        except Exception as e:
            QMessageBox.critical(self, "目录错误", f"无法创建或访问输出目录：{output_dir_str}\n错误：{e}")
            return

        self.process_button.setEnabled(False)
        self.browse_input_button.setEnabled(False)
        self.browse_output_button.setEnabled(False)
        self.clear_list_button.setEnabled(False)
        self.online_decrypt_button.setEnabled(False)  # 处理过程中禁用
        # self.restore_suffix_button.setEnabled(False) # 移除

        self.processing_files_count = 0
        self.processed_files_count = 0
        self.total_files_to_process = self.file_list_widget.count()
        self.overall_progress_bar.setValue(0)
        self.update_status("开始处理所有文件...", "black")

        for i in range(self.total_files_to_process):
            file_path_str = self.file_list_widget.item(i).text()
            runnable = FileProcessorRunnable(file_path_str, output_dir_str)

            runnable.signals.finished.connect(self.on_single_file_processing_finished)
            runnable.signals.error.connect(self.on_single_file_processing_error)
            runnable.signals.status_update.connect(self.update_status_for_file)

            self.thread_pool.start(runnable)
            self.processing_files_count += 1

    def update_status_for_file(self, message: str):
        self.status_label.setText(
            f"状态: {message} (已处理 {self.processed_files_count}/{self.total_files_to_process})")
        QApplication.processEvents()

    def on_single_file_processing_finished(self, original_file_path_str: str, result_path: Optional[Path]):
        self.processed_files_count += 1
        self.update_overall_progress_bar()

        if result_path:
            logging.info(f"文件处理成功: {os.path.basename(original_file_path_str)} -> {result_path}")
            # 如果是 .info 后缀的非标准文件，才启用“模拟在线解密”按钮
            if result_path.suffix.lower() == ".info":
                self.online_decrypt_button.setEnabled(True)
        else:
            logging.warning(f"文件 {original_file_path_str} 处理未返回有效路径。")

        if self.processed_files_count == self.total_files_to_process:
            self.on_all_files_processed()

    def on_single_file_processing_error(self, original_file_path: str, message: str):
        self.processed_files_count += 1
        self.update_overall_progress_bar()

        logging.error(f"文件 {original_file_path} 处理失败: {message}")
        if self.processed_files_count == self.total_files_to_process:
            self.on_all_files_processed()

    def update_overall_progress_bar(self):
        if self.total_files_to_process > 0:
            progress = int((self.processed_files_count / self.total_files_to_process) * 100)
            self.overall_progress_bar.setValue(progress)
            self.status_label.setText(
                f"状态: 整体进度 {progress}% (已处理 {self.processed_files_count}/{self.total_files_to_process})")
        QApplication.processEvents()

    def on_all_files_processed(self):
        self.process_button.setEnabled(True)
        self.browse_input_button.setEnabled(True)
        self.browse_output_button.setEnabled(True)
        self.clear_list_button.setEnabled(True)
        # 仅当有 .info 文件时才启用模拟解密按钮
        # 更好的做法是遍历已处理文件列表或检查输出目录
        # 这里简单地在每次处理结束时启用，如果用户没有 .info 文件，点击也会有警告
        self.online_decrypt_button.setEnabled(True)
        # self.restore_suffix_button.setEnabled(True) # 移除
        self.update_status("所有文件处理完成！", "green")
        QMessageBox.information(self, "处理完成", f"所有 {self.total_files_to_process} 个文件已处理完成。")
        self.processing_files_count = 0
        self.processed_files_count = 0
        self.total_files_to_process = 0

    def update_status(self, message, color="black"):
        self.status_label.setText(f"状态: {message}")
        self.status_label.setStyleSheet(f"color: {color};")
        QApplication.processEvents()

    def simulate_online_decryption(self):
        initial_dir = self.output_dir_input.text()
        if not initial_dir or not Path(initial_dir).is_dir():
            initial_dir = str(Path.home())

        file_path_to_decrypt, _ = QFileDialog.getOpenFileName(
            self, "选择要模拟解密的文件", initial_dir, "信息文件 (*.info);;所有文件 (*.*)"
        )
        if not file_path_to_decrypt:
            return

        input_path = Path(file_path_to_decrypt)
        # 模拟在线解密只针对 .info 文件
        if input_path.suffix.lower() != ".info":
            QMessageBox.warning(self, "警告", f"请选择一个 .info 文件进行模拟解密。\n当前文件: {input_path.name}")
            return

        logging.info(f"模拟在线解密请求针对文件: {file_path_to_decrypt}")
        QMessageBox.information(self, "模拟在线解密", "模拟在线解密请求已发送！\n（此处应集成实际的在线解密逻辑）")

        try:
            # 这里的模拟解密逻辑，假定将输入文件读入并应用解密函数，然后移除标志
            with open(input_path, "rb") as src_file:
                content = src_file.read()

            content = dummy_decrypt_function(content)

            # 移除原始加密标志和模拟解密标志，模拟“完全解密”
            content = content.replace(ESAFENET_MARKER, b"")
            content = content.replace(b"DecryptedData_Marker", b"")

            # 模拟解密后，仍然保持 .info 后缀
            output_file_path = input_path.parent / f"{input_path.stem}_final_decrypted.info"

            output_file_path.parent.mkdir(parents=True, exist_ok=True)

            with open(output_file_path, "wb") as dest_file:
                dest_file.write(content)

            QMessageBox.information(self, "模拟在线解密结果",
                                    f"模拟在线解密成功！\n解密后文件保存到: {output_file_path}")
            logging.info(f"模拟在线解密文件保存到: {output_file_path}")
        except Exception as e:
            QMessageBox.critical(self, "模拟在线解密错误", f"模拟在线解密过程中发生错误: {e}")
            logging.error(f"模拟在线解密 {file_path_to_decrypt} 失败: {e}")

    # >>> 移除 restore_original_suffix 方法
    # def restore_original_suffix(self):
    #     ...
    # <<< 移除结束

    def closeEvent(self, event):
        if self.thread_pool.activeThreadCount() > 0:
            reply = QMessageBox.question(self, '退出程序',
                                         "有文件正在处理，您确定要退出吗？未完成的任务将被中断。",
                                         QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
            if reply == QMessageBox.Yes:
                self.thread_pool.clear()
                self.thread_pool.waitForDone()
                self.clean_up_and_exit()
                event.accept()
            else:
                event.ignore()
        else:
            self.clean_up_and_exit()
            event.accept()

    def clean_up_and_exit(self):
        logger = logging.getLogger()
        if self.log_handler in logger.handlers:
            logger.removeHandler(self.log_handler)
        QApplication.quit()


if __name__ == "__main__":
    app = QApplication(sys.argv)
    main_window = FileProcessorApp()
    main_window.show()
    sys.exit(app.exec_())
