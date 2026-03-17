import sys
import time
import datetime
import threading
import keyboard

from PyQt5.QtWidgets import *
from PyQt5.QtCore import *
from PyQt5.QtGui import *

# 定时发送子线程类
class ClockThread(QThread):
    def __init__(self):
        super().__init__()
        self.time_counting = False
        self.send_func = None
        self.clocks = None
        self.prevent_offline = False
        self.prevent_func = None
        self.prevent_count = 60
        self.executed_tasks = set()
        self._prevent_timer = 0

    def __del__(self):
        self.wait()

    def run(self):
        import uiautomation as auto
        with auto.UIAutomationInitializerInThread():
            self._prevent_timer = self.prevent_count * 60

            while self.time_counting:
                now = datetime.datetime.now()
                next_event_time = None

                # 1. 遍历列表，查找最近的下一个闹钟时间
                try:
                    for i in range(self.clocks.count()):
                        task_id = self.clocks.item(i).text()
                        if task_id in self.executed_tasks:
                            continue

                        parts = task_id.split(" ")
                        clock_str = " ".join(parts[:5])
                        dt_obj = datetime.datetime.strptime(clock_str, "%Y %m %d %H %M")

                        if dt_obj > now:
                            if next_event_time is None or dt_obj < next_event_time:
                                next_event_time = dt_obj
                except Exception as e:
                    print(f"读取闹钟列表时出错: {e}")
                    time.sleep(1)
                    continue

                # 2. 计算休眠时间
                sleep_seconds = 0
                if next_event_time:
                    delta = (next_event_time - now).total_seconds()
                    sleep_seconds = max(0, delta)

                # 3. 整合“防止掉线”的逻辑
                if self.prevent_offline:
                    sleep_seconds = min(sleep_seconds, self._prevent_timer)

                # 4. 执行休眠
                time.sleep(sleep_seconds)

                self._prevent_timer -= sleep_seconds
                if self._prevent_timer <= 0:
                    self._prevent_timer = 0

                # 5. 休眠结束，检查并执行到期的任务
                now = datetime.datetime.now()

                try:
                    for i in range(self.clocks.count()):
                        task_id = self.clocks.item(i).text()
                        if task_id in self.executed_tasks:
                            continue

                        parts = task_id.split(" ")
                        st_ed = parts[5]
                        st, ed = st_ed.split('-')
                        clock_str = " ".join(parts[:5])
                        dt_obj = datetime.datetime.strptime(clock_str, "%Y %m %d %H %M")

                        time_diff = (now - dt_obj).total_seconds()
                        if 0 <= time_diff <= 60:
                            if self.send_func:
                                self.send_func(st=int(st), ed=int(ed))
                            self.executed_tasks.add(task_id)
                        elif time_diff > 60:
                            self.executed_tasks.add(task_id)

                except Exception as e:
                    print(f"执行任务时读取闹钟列表出错: {e}")

                # 检查并执行防止掉线
                if self.prevent_offline and self._prevent_timer <= 0:
                    if self.prevent_func:
                        self.prevent_func()
                    self._prevent_timer = self.prevent_count * 60

class MyListWidget(QListWidget):
    """支持双击可编辑的QListWidget"""
    def __init__(self, parent=None) -> None:
        super().__init__(parent)
        self.setSelectionMode(QAbstractItemView.ExtendedSelection)

        self.edited_item = self.currentItem()
        self.close_flag = True
        self.doubleClicked.connect(self.item_double_clicked)
        self.currentItemChanged.connect(self.close_edit)

    def keyPressEvent(self, e: QKeyEvent) -> None:
        super().keyPressEvent(e)
        if e.key() == Qt.Key_Return:
            if self.close_flag:
                self.close_edit()
            self.close_flag = True

    def edit_new_item(self) -> None:
        self.close_flag = False
        self.close_edit()
        count = self.count()
        self.addItem('')
        item = self.item(count)
        self.edited_item = item
        self.openPersistentEditor(item)
        self.editItem(item)

    def item_double_clicked(self, modelindex: QModelIndex) -> None:
        self.close_edit()
        item = self.item(modelindex.row())
        self.edited_item = item
        self.openPersistentEditor(item)
        self.editItem(item)

    def close_edit(self, *_) -> None:
        if self.edited_item and self.isPersistentEditorOpen(self.edited_item):
            self.closePersistentEditor(self.edited_item)

class MultiInputDialog(QDialog):
    """用于用户输入的输入框，可以根据传入的参数自动创建输入框"""
    def __init__(self, inputs: list, default_values: list = None, parent=None) -> None:
        super().__init__(parent)
        
        layout = QVBoxLayout(self)
        self.inputs = []
        for n, i in enumerate(inputs):
            layout.addWidget(QLabel(i))
            input = QLineEdit(self)

            # 【修复】防止 default_values 长度不足导致数组越界崩溃
            if default_values is not None and n < len(default_values):
                input.setText(str(default_values[n]))

            layout.addWidget(input)
            self.inputs.append(input)
            
        ok_button = QPushButton("确认")
        ok_button.clicked.connect(self.accept)
        
        cancel_button = QPushButton("取消")
        cancel_button.clicked.connect(self.reject)
        
        button_layout = QHBoxLayout()
        button_layout.addWidget(ok_button)
        button_layout.addWidget(cancel_button)
        layout.addLayout(button_layout)
    
    def get_input(self):
        return [i.text() for i in self.inputs]

class FileDialog(QDialog):
    """文件选择框"""
    def __init__(self, parent=None) -> None:
        super().__init__(parent)
        self.inputs = []
        layout = QVBoxLayout(self)
        
        layout.addWidget(QLabel("请指定发送给哪些用户(1,2,3代表发送给前三位用户)，如需全部发送请忽略此项"))
        input = QLineEdit(self)
        layout.addWidget(input)
        self.inputs.append(input)
        
        choose_layout = QHBoxLayout()

        path = QLineEdit(self)
        choose_layout.addWidget(path)
        self.inputs.append(path)

        file_button = QPushButton("选择文件")
        file_button.clicked.connect(self.select)
        choose_layout.addWidget(file_button)

        layout.addLayout(choose_layout)
        
        ok_button = QPushButton("确认")
        ok_button.clicked.connect(self.accept)

        cancel_button = QPushButton("取消")
        cancel_button.clicked.connect(self.reject)

        button_layout = QHBoxLayout()
        button_layout.addWidget(ok_button)
        button_layout.addWidget(cancel_button)
        layout.addLayout(button_layout)
    
    def select(self):
        path_input = self.inputs[1]
        paths = QFileDialog.getOpenFileNames(self, '打开文件', '/home')[0]
        if paths:
            path_input.setText(";".join(paths))
    
    def get_input(self):
        return [i.text() for i in self.inputs]

class MySpinBox(QWidget):
    def __init__(self, desc: str, **kwargs):
        super().__init__(**kwargs)
        layout = QHBoxLayout()

        self.desc = desc
        self.label = QLabel(desc)
        self.spin_box = QSpinBox()

        layout.addWidget(self.label)
        layout.addWidget(self.spin_box)
        self.setLayout(layout)