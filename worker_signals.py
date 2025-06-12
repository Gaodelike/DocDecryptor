from PyQt5.QtCore import QObject, pyqtSignal
from pathlib import Path
from typing import Optional


class WorkerSignals(QObject):
    '''
    定义工作线程发出的信号
    '''
    # 发送原始文件路径和处理结果 (Path 或 None)
    finished = pyqtSignal(str, object)
    # 发送原始文件路径和错误消息
    error = pyqtSignal(str, str)
    # 发送状态更新
    status_update = pyqtSignal(str)
    # 可以添加更细粒度的文件处理进度信号 (0-100)
    # file_progress_update = pyqtSignal(str, int) # 文件名，进度
