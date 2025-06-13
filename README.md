# DocDecryptor
在线解密亿赛通加密的word、excel等常见文档和图片
核心思想是伪装亿赛通的白名单进程，通过亿赛通自己来解密

`dist/winword.exe`可直接运行解密工具，前提是必须在可打开加密文档的电脑上运行
执行完毕后，会在指定目录生成.info后缀的解密文件，记得修改后缀后再使用对应的应用程序打开。
如果不能运行，请按下面的环境配置步骤进行配置，然后使用pyinstaller打包成可执行文件。

## Features
This application supports `.docx`, `.pdf`, `.jpg`, `.jpeg`, `.png`, `.bmp`, `.gif`, `.xlsx`, and `.pptx` file formats.
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

## Usage
Running the Application
To run the application, you can execute the following command:

`python main.py`

Packaging the Application
To package the application into a standalone executable, you can use PyInstaller:

`pyinstaller --onefile --noconsole --name winword main.py`

This will generate an executable file named winword.exe in the dist directory.

You can install these packages using `pip`:

```bash
pip install python-docx PyPDF2 Pillow openpyxl python-pptx
```

## Contributing
Contributions are welcome! Please open an issue or submit a pull request for any improvements or bug fixes.

## License
This project is licensed under the MIT License - see the LICENSE file for details.




