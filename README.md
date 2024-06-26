# DocDecryptor

DocDecryptor is a Python application that processes various document files and saves them in a specified format. This application supports `.docx`, `.pdf`, `.jpg`, `.jpeg`, `.png`, `.bmp`, `.gif`, `.xlsx`, and `.pptx` file formats.

## Features

- Read and save .docx files
- Read and save .pdf files
- Read and save image files (.jpg, .jpeg, .png, .bmp, .gif)
- Read and save .xlsx files
- Read and save .pptx files

## Installation

To run this application, you need to have Python installed on your system. Additionally, the following Python packages are required:

- docx
- PyPDF2
- PIL (Pillow)
- openpyxl
- python-pptx

You can install these packages using `pip`:

```bash
pip install python-docx PyPDF2 Pillow openpyxl python-pptx

Usage
Running the Application
To run the application, you can execute the following command:

python main.py
Packaging the Application
To package the application into a standalone executable, you can use PyInstaller:

pyinstaller --onefile --noconsole --name winword main.py
This will generate an executable file named winword.exe in the dist directory.

Contributing
Contributions are welcome! Please open an issue or submit a pull request for any improvements or bug fixes.

License
This project is licensed under the MIT License - see the LICENSE file for details.
