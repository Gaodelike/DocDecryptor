import sys
import logging
from PyQt5.QtWidgets import QApplication, QWidget, QVBoxLayout, QPushButton, QFileDialog, QLabel, QMessageBox, \
    QListWidget, QTextEdit
from file_processor import process_file  # 导入文件处理函数

logging.basicConfig(filename='app.log', level=logging.ERROR,
                    format='%(asctime)s - %(levelname)s - %(message)s', encoding='utf-8')


class FileProcessorApp(QWidget):
    def __init__(self):
        super().__init__()
        self.init_ui()

    def init_ui(self):
        self.layout = QVBoxLayout()

        self.label = QLabel("Select files to process:")
        self.layout.addWidget(self.label)

        self.file_button = QPushButton("Choose Files")
        self.file_button.clicked.connect(self.choose_files)
        self.layout.addWidget(self.file_button)

        self.file_list_widget = QListWidget()
        self.layout.addWidget(self.file_list_widget)

        self.process_button = QPushButton("Process Files")
        self.process_button.clicked.connect(self.process_files)
        self.layout.addWidget(self.process_button)

        self.error_display = QTextEdit()
        self.error_display.setReadOnly(True)
        self.layout.addWidget(self.error_display)

        self.setLayout(self.layout)
        self.setWindowTitle("DocDecryptor")

        self.selected_files = []

    def choose_files(self):
        options = QFileDialog.Options()
        files, _ = QFileDialog.getOpenFileNames(
            self, "Choose Files", "",
            "All Files (*);;Docx Files (*.docx);;PDF Files (*.pdf);;Image Files (*.jpg *.jpeg *.png *.bmp *.gif);;Excel Files (*.xlsx);;PowerPoint Files (*.pptx)",
            options=options
        )
        if files:
            self.selected_files.extend(files)
            self.update_file_list()

    def update_file_list(self):
        self.file_list_widget.clear()
        for file in self.selected_files:
            self.file_list_widget.addItem(file)

    def process_files(self):
        if self.selected_files:
            errors = []
            for file_path in self.selected_files:
                try:
                    process_file(file_path)
                except Exception as e:
                    error_message = f"Error processing file {file_path}: {e}"
                    errors.append(error_message)
                    self.log_error(error_message)
            if errors:
                self.show_errors(errors)
            else:
                self.label.setText(f"Processed {len(self.selected_files)} files successfully.")
            self.selected_files.clear()
            self.update_file_list()

    def show_errors(self, errors):
        self.error_display.clear()
        self.error_display.append("Errors occurred while processing files:\n")
        for error in errors:
            self.error_display.append(error)

    def log_error(self, message):
        logging.error(message)

    def show_error_message(self, message):
        error_dialog = QMessageBox()
        error_dialog.setIcon(QMessageBox.Critical)
        error_dialog.setText("An error occurred")
        error_dialog.setInformativeText(message)
        error_dialog.setWindowTitle("Error")
        error_dialog.exec_()


if __name__ == "__main__":
    app = QApplication(sys.argv)
    ex = FileProcessorApp()
    ex.show()
    sys.exit(app.exec_())
