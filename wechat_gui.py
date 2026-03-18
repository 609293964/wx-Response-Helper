import sys
import time
import os
import itertools
import json
import datetime

from PyQt5.QtWidgets import *
from PyQt5.QtCore import *
from PyQt5.QtGui import *
from ui_auto_wechat import WeChat
from module import *
from wechat_locale import WeChatLocale


class WechatGUI(QWidget):
    # 定义信号，用于将后台线程的UI更新操作切回主线程
    new_message_signal = pyqtSignal(str)
    alert_message_signal = pyqtSignal(str)

    def __init__(self):
        super().__init__()

        # 读取之前保存的配置文件，如果没有则新建一个
        self.config_path = "wechat_config.json"
        if os.path.exists(self.config_path):
            with open(self.config_path, "r", encoding="utf-8") as r:
                try:
                    self.config = json.load(r)
                except json.JSONDecodeError:
                    self.config = {}
                
                # 优化：确保基础配置结构存在，防止旧版本升级或配置文件损坏导致缺少键值而闪退
                if "settings" not in self.config: self.config["settings"] = {}
                if "contacts" not in self.config: self.config["contacts"] = []
                if "messages" not in self.config: self.config["messages"] = []
                if "schedules" not in self.config: self.config["schedules"] = []
                
                # 确保关键设置有默认值
                settings = self.config["settings"]
                settings.setdefault("wechat_path", "")
                settings.setdefault("send_interval", 0)
                settings.setdefault("system_version", "new")
                settings.setdefault("language", "zh-CN")

        else:
            # 默认配置
            self.config = {
                "settings": {
                    "wechat_path": "",
                    "send_interval": 0,
                    "system_version": "new",
                    "language": "zh-CN",
                },
                "contacts": [],
                "messages": [],
                "schedules": [],
            }
            self.save_config()

        self.wechat = WeChat(
            path=self.config["settings"].get("wechat_path", ""),
            locale=self.config["settings"].get("language", "zh-CN"),
        )
        self.clock = ClockThread()

        # 发消息的用户列表
        self.contacts = []

        # 连接跨线程UI更新信号
        self.new_message_signal.connect(self._append_monitor_message)
        self.alert_message_signal.connect(self._append_monitor_message)

        # 初始化图形界面
        self.initUI()

        # 判断全局热键是否被按下
        self.hotkey_pressed = False
        keyboard.add_hotkey('ctrl+alt+q', self.hotkey_press)
        
        # 自动打开提示
        self.show_wechat_open_notice()

    # 重写关闭事件，确保程序完全退出
    def closeEvent(self, event):
        # 停止所有定时器
        if hasattr(self, 'clock') and self.clock.time_counting:
            self.clock.time_counting = False
            self.clock.quit()
            self.clock.wait()
        
        if hasattr(self, 'sync_folder_timer'):
            self.sync_folder_timer.stop()
        
        # 移除热键
        keyboard.unhook_all()
        
        # 退出应用程序
        QApplication.quit()
        event.accept()

    # 显示微信打开方式变更提示
    def show_wechat_open_notice(self):
        msg_box = QMessageBox(self)
        msg_box.setIcon(QMessageBox.Information)
        msg_box.setWindowTitle("重要提示")
        msg_box.setText("微信打开方式已变更")
        msg_box.setInformativeText(
            "由于微信版本更新，我们现在使用微信内置的快捷键来打开/隐藏微信窗口，请确保你的微信打开快捷键为Ctrl+Alt+w。具体查看方式为“设置”->“快捷键”->“显示/隐藏窗口”\n\n"
            "⚠️ 注意事项：\n"
            "• 如果微信已经打开且在前台，再次按快捷键会导致微信窗口被隐藏\n"
            "• 为避免此问题，建议在使用定时发送功能前，先手动关闭或最小化微信窗口\n"
            "• 这样可以确保程序能够正常打开微信并发送消息\n\n"
        )
        msg_box.setStandardButtons(QMessageBox.Ok)
        msg_box.exec_()

    # 保存当前的配置
    def save_config(self):
        with open(self.config_path, "w", encoding="utf8") as w:
            json.dump(self.config, w, indent=4, ensure_ascii=False)

    def hotkey_press(self):
        print("hotkey pressed")
        self.hotkey_pressed = True

    # 选择用户界面的初始化
    def init_choose_contacts(self):
        # 在联系人有变化后更新配置文件
        def update_contacts():
            contacts = []
            for i in range(self.contacts_view.count()):
                contacts.append(self.contacts_view.item(i).text())

            self.config["contacts"] = contacts
            self.save_config()

        # 读取联系人列表并保存
        def save_contacts():
            # 先弹出一个提示词告诉用户这个提取并不保证可靠，因为微信组织信息的方式本身就有歧义
            QMessageBox.information(self, "注意", "提取联系人列表功能并不保证完全可靠，因为微信组织信息的方式本身就有歧义。"
                                                  "如果想要提取更可靠，请不需要在给用户的备注和设置的分组标签里面加空格。")

            path = QFileDialog.getSaveFileName(self, "保存联系人列表", "contacts.csv", "表格文件(*.csv)")[0]
            if not path == "":
                contacts = self.wechat.find_all_contacts()
                contacts.to_csv(path, index=False, encoding='utf_8_sig')
                QMessageBox.information(self, "保存成功", "联系人列表保存成功！")

        # 保存群聊列表
        def save_groups():
            path = QFileDialog.getSaveFileName(self, "保存群聊列表", "groups.txt", "文本文件(*.txt)")[0]
            if not path == "":
                contacts = self.wechat.find_all_groups()
                with open(path, 'w', encoding='utf-8') as f:
                    for contact in contacts:
                        f.write(contact + '\n')

                QMessageBox.information(self, "保存成功", "群聊列表保存成功！")

        # 读取联系人列表并加载
        def load_contacts():
            path = QFileDialog.getOpenFileName(self, "加载联系人列表", "", "文本文件(*.txt)")[0]
            if not path == "":
                with open(path, 'r', encoding='utf-8') as f:
                    for line in f.readlines():
                        self.contacts_view.addItem(f"{self.contacts_view.count()+1}:{line.strip()}")

                update_contacts()
                QMessageBox.information(self, "加载成功", "联系人列表加载成功！")

        # 增加用户列表信息
        def add_contact():
            name_list, ok = QInputDialog.getText(self, '添加用户', '输入添加的用户名(可添加多个人名，用英文逗号,分隔):')
            if ok:
                if name_list != "":
                    names = name_list.split(',')
                    for name in names:
                        id = f"{self.contacts_view.count() + 1}"
                        self.contacts_view.addItem(f"{id}:{str(name).strip()}")
                    update_contacts()

        # 删除用户信息
        def del_contact():
            # 删除选中的用户
            for i in range(self.contacts_view.count()-1, -1, -1):
                if self.contacts_view.item(i).isSelected():
                    self.contacts_view.takeItem(i)

            # 为所有剩余的用户重新编号
            for i in range(self.contacts_view.count()):
                self.contacts_view.item(i).setText(f"{i+1}:{self.contacts_view.item(i).text().split(':')[1]}")

            update_contacts()

        hbox = QHBoxLayout()

        # 左边的用户列表
        self.contacts_view = MyListWidget()

        # 加载配置文件里保存的用户
        for contact in self.config["contacts"]:
            self.contacts_view.addItem(contact)

        self.clock.contacts = self.contacts_view
        for name in self.contacts:
            self.contacts_view.addItem(name)

        hbox.addWidget(self.contacts_view)

        # 右边的按钮界面
        vbox = QVBoxLayout()
        vbox.stretch(1)

        # 用户界面的按钮
        info = QLabel("待发送用户列表")

        save_btn = QPushButton("保存微信好友列表")
        save_btn.clicked.connect(save_contacts)

        save_group_btn = QPushButton("保存微信群聊列表")
        save_group_btn.clicked.connect(save_groups)

        load_btn = QPushButton("加载用户txt文件")
        load_btn.clicked.connect(load_contacts)

        add_btn = QPushButton("添加用户")
        add_btn.clicked.connect(add_contact)

        del_btn = QPushButton("删除用户")
        del_btn.clicked.connect(del_contact)

        vbox.addWidget(info)
        vbox.addWidget(save_btn)
        vbox.addWidget(save_group_btn)
        vbox.addWidget(load_btn)
        vbox.addWidget(add_btn)
        vbox.addWidget(del_btn)
        hbox.addLayout(vbox)

        return hbox

    # 定时功能界面的初始化
    def init_clock(self):
        # 在定时列表有变化后更新配置文件
        def update_schedules():
            schedules = []
            for i in range(self.time_view.count()):
                schedules.append(self.time_view.item(i).text())

            self.config["schedules"] = schedules
            self.save_config()
            
        # 按钮响应：增加时间
        def add_contact():
            inputs = [
                "注：在每一个时间输入框内都可以使用英文逗号“,“来一次性区分多个数值进行多次定时。\n(例：分钟框输入 10,20,30,40)",
                "年 (例如: 2026)",
                "月 (1~12)",
                "日 (1~31)",
                "小时（0~23）",
                "分钟 (0~59)",
            ]

            # 设置默认值为当前时间
            local_time = time.localtime(time.time())
            default_values = [
                str(local_time.tm_year),
                str(local_time.tm_mon),
                str(local_time.tm_mday),
                str(local_time.tm_hour),
                str(local_time.tm_min),
            ]

            dialog = MultiInputDialog(inputs, default_values)
            if dialog.exec_() == QDialog.Accepted:
                input_values = dialog.get_input()
                if len(input_values) != 6:
                    QMessageBox.warning(self, "输入错误", "输入项不匹配！")
                    return
                    
                note, year, month, day, hour, min = input_values
                if year == "" or month == "" or day == "" or hour == "" or min == "":
                    QMessageBox.warning(self, "输入错误", "输入不能为空！")
                    return

                else:
                    year_list = [y.strip() for y in year.split(',')]
                    month_list = [m.strip() for m in month.split(',')]
                    day_list = [d.strip() for d in day.split(',')]
                    hour_list = [h.strip() for h in hour.split(',')]
                    min_list = [m.strip() for m in min.split(',')]
                    
                    # 默认发送全部内容
                    send_range = f"1-{max(1, self.msg.count())}"

                    for year, month, day, hour, min in itertools.product(year_list, month_list, day_list, hour_list, min_list):
                        input = f"{year} {month} {day} {hour} {min} {send_range}"
                        self.time_view.addItem(input)
                    
                    update_schedules()

        # 按钮响应：删除时间
        def del_contact():
            for i in range(self.time_view.count() - 1, -1, -1):
                if self.time_view.item(i).isSelected():
                    self.time_view.takeItem(i)
            
            update_schedules()

        # 按钮响应：开始定时
        def start_counting():
            if self.clock.time_counting is True:
                return
            else:
                self.clock.time_counting = True

            info.setStyleSheet("color:red")
            info.setText("定时发送（目前已开始）")
            self.clock.start()

        # 按钮响应：结束定时
        def end_counting():
            self.clock.time_counting = False
            info.setStyleSheet("color:black")
            info.setText("定时发送（目前未开始）")

        # 按钮相应：开启防止自动下线
        def prevent_offline():
            if self.clock.prevent_offline is True:
                self.clock.prevent_offline = False
                prevent_btn.setStyleSheet("color:black")
                prevent_btn.setText("防止自动下线：（目前关闭）")

            else:
                QMessageBox.information(self, "防止自动下线", "防止自动下线已开启！每隔一小时自动点击微信窗口，防"
                                                              "止自动下线。请不要在正常使用电脑时使用该策略。")

                self.clock.prevent_offline = True
                prevent_btn.setStyleSheet("color:red")
                prevent_btn.setText("防止自动下线：（目前开启）")

        hbox = QHBoxLayout()

        # 左边的时间列表
        self.time_view = MyListWidget()
        # 加载配置文件里保存的时间
        for schedule in self.config["schedules"]:
            self.time_view.addItem(schedule)
            
        self.clock.clocks = self.time_view
        hbox.addWidget(self.time_view)

        # 右边的按钮界面
        vbox = QVBoxLayout()
        vbox.stretch(1)

        info = QLabel("定时发送（目前未开始）")
        add_btn = QPushButton("添加时间")
        add_btn.clicked.connect(add_contact)
        del_btn = QPushButton("删除时间")
        del_btn.clicked.connect(del_contact)
        start_btn = QPushButton("开始定时")
        start_btn.clicked.connect(start_counting)
        end_btn = QPushButton("结束定时")
        end_btn.clicked.connect(end_counting)
        prevent_btn = QPushButton("防止自动下线：（目前关闭）")
        prevent_btn.clicked.connect(prevent_offline)

        # 新增：智能随机批量生成定时时间的函数
        def add_batch_time():
            import random
            msg_count = max(1, self.msg.count())
            inputs = [
                f"当前有 {msg_count} 条待发送内容，将生成对应数量的定时任务",
                "开始时间段（例如: 19:30 表示从晚上7点30分开始）",
                "结束时间段（例如: 23:00 表示到晚上11点结束）",
                "随机波动范围（分钟，例如: 15 表示每个时间点在范围内随机）",
                "要生成的天数（例如: 7 表示生成未来7天的定时任务）",
            ]
            default_values = ["", "19:30", "23:00", "15", "1"]
            dialog = MultiInputDialog(inputs, default_values)
            if dialog.exec_() == QDialog.Accepted:
                input_values = dialog.get_input()
                if len(input_values) == 5:
                    note, start_time, end_time, random_range, days = input_values
                else:
                    # 兼容旧版本输入
                    start_time, end_time, _, random_range, days = input_values
                
                try:
                    # 解析输入
                    random_range = int(random_range)
                    days = int(days)
                    msg_count = max(1, self.msg.count())
                    
                    # 解析开始和结束时间
                    start_hour, start_minute = map(int, start_time.split(':'))
                    end_hour, end_minute = map(int, end_time.split(':'))
                    
                    # 转换为总分钟数便于计算
                    start_total = start_hour * 60 + start_minute
                    end_total = end_hour * 60 + end_minute
                    total_duration = end_total - start_total
                    
                    if end_total <= start_total:
                        QMessageBox.warning(self, "输入错误", "结束时间必须晚于开始时间！")
                        return
                    
                    # 获取当前日期
                    today = datetime.datetime.now().date()
                    
                    # 生成所有时间点
                    added_count = 0
                    for day_offset in range(days):
                        current_date = today + datetime.timedelta(days=day_offset)
                        
                        # 计算每条内容的平均间隔
                        if msg_count == 1:
                            # 只有1条内容时，在时间段内随机一个时间
                            base_time = start_total + random.randint(0, total_duration)
                            offset = random.randint(-random_range, random_range)
                            actual_time = max(start_total, min(base_time + offset, end_total))
                            hour = actual_time // 60
                            minute = actual_time % 60
                            # 每条内容对应一个时间点，发送对应序号的内容
                            time_str = f"{current_date.year} {current_date.month} {current_date.day} {hour} {minute} 1-1"
                            self.time_view.addItem(time_str)
                            added_count += 1
                        else:
                            # 多条内容时，平均分配时间
                            interval = total_duration // (msg_count - 1)
                            for msg_index in range(msg_count):
                                # 计算基础时间点（平均分布）
                                base_time = start_total + msg_index * interval
                                # 生成随机偏移量
                                offset = random.randint(-random_range, random_range)
                                actual_time = max(start_total, min(base_time + offset, end_total))
                                
                                # 转换为小时和分钟
                                hour = actual_time // 60
                                minute = actual_time % 60
                                
                                # 每条内容对应一个时间点，发送对应序号的内容
                                msg_num = msg_index + 1
                                time_str = f"{current_date.year} {current_date.month} {current_date.day} {hour} {minute} {msg_num}-{msg_num}"
                                self.time_view.addItem(time_str)
                                added_count += 1
                    
                    # 更新配置
                    update_schedules()
                    QMessageBox.information(self, "生成成功", f"已成功生成 {added_count} 个随机定时任务！\n\n提示：\n• 有 {msg_count} 条待发送内容，生成了对应数量的时间点\n• 每个时间点在平均时间前后{random_range}分钟内随机\n• 所有时间都在你设定的时间段范围内\n• 每个时间点发送对应序号的1条内容，发完即止")
                    
                except Exception as e:
                    QMessageBox.warning(self, "输入错误", f"输入格式不正确，请检查！\n错误信息：{e}")

        vbox.addWidget(info)
        vbox.addWidget(add_btn)
        vbox.addWidget(del_btn)
        # 一键删除所有时间按钮
        clear_all_btn = QPushButton("清空所有定时")
        def clear_all_schedules():
            reply = QMessageBox.question(self, "确认清空", "确定要删除所有定时任务吗？", 
                                       QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
            if reply == QMessageBox.Yes:
                self.time_view.clear()
                update_schedules()
                QMessageBox.information(self, "清空成功", "所有定时任务已删除！")
        clear_all_btn.clicked.connect(clear_all_schedules)
        vbox.addWidget(clear_all_btn)
        batch_btn = QPushButton("智能生成随机定时")
        batch_btn.clicked.connect(add_batch_time)
        vbox.addWidget(batch_btn)
        vbox.addWidget(start_btn)
        vbox.addWidget(end_btn)
        vbox.addWidget(prevent_btn)
        hbox.addLayout(vbox)

        return hbox

    # 发送消息内容界面的初始化
    def init_send_msg(self):
        # 在发送消息有变化后更新配置文件
        def update_messages():
            messages = []
            for i in range(self.msg.count()):
                messages.append(self.msg.item(i).text())

            self.config["messages"] = messages
            self.save_config()

        # 从txt中加载消息内容
        def load_text():
            path = QFileDialog.getOpenFileName(self, "加载内容文本", "", "文本文件(*.txt)")[0]
            if not path == "":
                with open(path, 'r', encoding='utf-8') as f:
                    for line in f.readlines():
                        self.msg.addItem(f"{self.msg.count()+1}:text:{line.strip()}")

                QMessageBox.information(self, "加载成功", "内容文本加载成功！")

        # 增加一条文本信息
        def add_text():
            inputs = [
                "是否需要at他人(无则不填，有则填写所有你要at的人名，用英文逗号分隔。要at所有人就填写'所有人')",
                "请输入发送的文本内容(如果需要换行则输入\\n，例如你好\\n吃饭了吗？)",
                "请指定发送给哪些用户(1,2,3代表发送给前三位用户)，如需全部发送请忽略此项",
            ]
            dialog = MultiInputDialog(inputs)
            if dialog.exec_() == QDialog.Accepted:
                at, text, to = dialog.get_input()
                to = "all" if to == "" else to
                if text != "":
                    # 消息的序号
                    rank = self.msg.count() + 1

                    self.msg.addItem(f"{rank}:text:{to}:{at}:{str(text)}")
                    update_messages()

        # 增加一个文件
        def add_file():
            dialog = FileDialog()
            if dialog.exec_() == QDialog.Accepted:
                to, paths = dialog.get_input()
                to = "all" if to == "" else to
                if paths != "":
                    # 将多个文件路径按分号分隔
                    path_list = paths.split(";")
                    # 循环添加每个文件
                    for path in path_list:
                        path = path.strip()
                        if path != "":
                            self.msg.addItem(f"{self.msg.count()+1}:file:{to}:{str(path)}")
                    update_messages()

        # 删除一条发送的信息
        def del_content():
            # 删除选中的信息
            for i in range(self.msg.count() - 1, -1, -1):
                if self.msg.item(i).isSelected():
                    self.msg.takeItem(i)

            # 为所有剩余的信息重新设置编号
            for i in range(self.msg.count()):
                self.msg.item(i).setText(f"{i+1}:"+self.msg.item(i).text().split(':', 1)[1])

            update_messages()

        # 发送按钮响应事件
        def send_msg(gap=None, st=None, ed=None):
            # 在每次发送时进行初始化
            self.hotkey_pressed = False

            # 获取发送间隔
            interval = send_interval.spin_box.value()

            try:
                # 如果未定义范围的开头和结尾，则默认发送全部信息
                if st is None:
                    st = 1
                    ed = self.msg.count()

                # 获得用户编号列表
                for user_i in range(self.contacts_view.count()):
                    # 等待间隔时间
                    time.sleep(int(interval))

                    rank, name = self.contacts_view.item(user_i).text().split(':', 1)
                    # For the first message, we need to search user
                    search_user = True

                    for msg_i in range(st - 1, ed):
                        # 如果全局热键被按下，则停止发送
                        if self.hotkey_pressed is True:
                            QMessageBox.warning(self, "发送失败", f"热键已按下，已停止发送！")
                            return

                        msg = self.msg.item(msg_i).text().replace("\\n", "\n")

                        _, type, to, content = msg.split(':', 3)
                        # 判断是否需要发送给该用户
                        if to == "all" or str(rank) in to.split(','):
                            # 判断为文本内容
                            if type == "text":
                                # 分出at的人和发送的文本内容
                                at_names, text = content.split(":", 1)
                                at_names = at_names.split(",")
                                self.wechat.send_msg(name, at_names, text, search_user)

                            # 判断为文件内容
                            elif type == "file":
                                self.wechat.send_file(name, content, search_user)
                                # 如果开启了自动删除文件功能，则删除源文件
                                if self.auto_delete_files.isChecked():
                                    try:
                                        import os
                                        if os.path.exists(content):
                                            os.remove(content)
                                            print(f"已自动删除文件: {content}")
                                    except Exception as e:
                                        print(f"删除文件失败: {content}, 错误: {e}")

                            # 搜索用户只在第一次发送时进行
                            search_user = False

            except Exception as e:
                QMessageBox.warning(self, "发送失败", f"发送失败！请检查内容格式或是否有遗漏步骤！\n错误信息：{e}")
                return

        # 左边的布局
        vbox_left = QVBoxLayout()

        # 提示信息
        info = QLabel("添加要发送的内容（程序将按顺序发送）")

        # 输入内容框
        self.msg = MyListWidget()
        # 加载配置文件里保存的内容
        for message in self.config["messages"]:
            self.msg.addItem(message)

        self.clock.send_func = send_msg
        self.clock.prevent_func = self.wechat.prevent_offline

        # 发送按钮
        send_btn = QPushButton("发送")
        send_btn.clicked.connect(send_msg)

        # 发送后自动删除文件的复选框
        self.auto_delete_files = QCheckBox("发送文件后自动删除源文件")
        # 优化：安全读取配置
        self.auto_delete_files.setChecked(self.config["settings"].get("auto_delete_files", False))
        
        # 保存设置到配置
        def toggle_auto_delete(state):
            self.config["settings"]["auto_delete_files"] = (state == Qt.Checked)
            self.save_config()
        
        self.auto_delete_files.stateChanged.connect(toggle_auto_delete)

        # 同步文件夹功能
        self.sync_folder_enabled = QCheckBox("启用文件夹同步到待发送列表")
        self.sync_folder_path = QLineEdit()
        self.sync_folder_path.setPlaceholderText("选择要同步的文件夹路径")
        self.sync_folder_btn = QPushButton("选择文件夹")
        self.sync_folder_manual_btn = QPushButton("立即同步")
        self.sync_folder_timer = QTimer(self)
        self.sync_folder_timer.setInterval(10000)  # 每10秒检查一次文件夹变化，资源占用极低
        self.sync_folder_last_mtime = 0  # 记录文件夹上次修改时间
        
        # 优化：安全读取配置
        self.sync_folder_enabled.setChecked(self.config["settings"].get("sync_folder_enabled", False))
        self.sync_folder_path.setText(self.config["settings"].get("sync_folder_path", ""))
        
        # 选择文件夹按钮响应
        def choose_sync_folder():
            folder_path = QFileDialog.getExistingDirectory(self, "选择要同步的文件夹")
            if folder_path:
                self.sync_folder_path.setText(folder_path)
                self.config["settings"]["sync_folder_path"] = folder_path
                self.save_config()
                # 如果同步功能已开启，立即同步一次
                if self.sync_folder_enabled.isChecked():
                    sync_folder_files()
        
        # 同步文件夹文件到待发送列表
        def sync_folder_files():
            if not self.sync_folder_enabled.isChecked():
                return
                
            folder_path = self.sync_folder_path.text().strip()
            if not folder_path or not os.path.exists(folder_path):
                return
                
            try:
                # 先检查文件夹是否有变化，没有变化则不同步
                current_mtime = os.path.getmtime(folder_path)
                if current_mtime == self.sync_folder_last_mtime:
                    return
                self.sync_folder_last_mtime = current_mtime
                
                # 获取文件夹下的所有文件（不包含子文件夹）
                file_list = []
                for filename in os.listdir(folder_path):
                    file_path = os.path.join(folder_path, filename)
                    if os.path.isfile(file_path):
                        file_list.append(file_path)
                
                # 获取当前选中的项和滚动位置，同步后恢复
                selected_items = [item.text() for item in self.msg.selectedItems()]
                scroll_pos = self.msg.verticalScrollBar().value()
                
                # 先清除当前消息列表中不是来自同步文件夹的文件项保留，同步文件夹的文件项更新
                current_items = []
                for i in range(self.msg.count()):
                    item_text = self.msg.item(i).text()
                    # 检查是否是文件类型
                    if item_text.split(':')[1] == 'file':
                        # 提取文件路径
                        _, _, to, content = item_text.split(':', 3)
                        # 如果不是同步文件夹中的文件，保留；否则添加到保留列表
                        if os.path.dirname(content) == folder_path.rstrip(os.sep):
                            continue
                    current_items.append(item_text)
                
                # 添加同步文件夹中的所有文件
                for file_path in file_list:
                    # 默认发送给所有人
                    current_items.append(f"{len(current_items)+1}:file:all:{file_path}")
                
                # 重新加载列表
                self.msg.clear()
                for item in current_items:
                    list_item = QListWidgetItem(item)
                    self.msg.addItem(list_item)
                    # 恢复选中状态
                    if item in selected_items:
                        list_item.setSelected(True)
                
                # 恢复滚动位置
                self.msg.verticalScrollBar().setValue(scroll_pos)
                
                # 更新配置
                update_messages()
                
            except Exception as e:
                print(f"同步文件夹失败: {e}")
        
        # 切换同步功能开关
        def toggle_sync_folder(state):
            enabled = (state == Qt.Checked)
            self.config["settings"]["sync_folder_enabled"] = enabled
            self.save_config()
            
            if enabled:
                # 立即同步一次
                sync_folder_files()
                # 启动定时器
                self.sync_folder_timer.start()
                QMessageBox.information(self, "同步提示", "文件夹同步已开启！\n\n提示：\n1. 每10秒自动检查文件夹变化，资源占用极低\n2. 也可以点击「立即同步」按钮手动同步\n3. 同步不会删除你手动添加的消息和文件\n4. 选择的项和滚动位置会自动保留，不会影响操作")
            else:
                # 停止定时器
                self.sync_folder_timer.stop()
                QMessageBox.information(self, "同步提示", "文件夹同步已关闭！")
        
        # 绑定事件
        self.sync_folder_btn.clicked.connect(choose_sync_folder)
        self.sync_folder_manual_btn.clicked.connect(sync_folder_files)
        self.sync_folder_enabled.stateChanged.connect(toggle_sync_folder)
        self.sync_folder_timer.timeout.connect(sync_folder_files)
        
        # 如果启动时已开启同步，启动定时器
        if self.sync_folder_enabled.isChecked():
            self.sync_folder_timer.start()

        # 发送不同用户时的间隔
        send_interval = MySpinBox("发送不同用户时的间隔（秒）")
        # 优化：安全读取配置
        send_interval.spin_box.setValue(self.config["settings"].get("send_interval", 0))

        # 添加修改间隔的响应
        def change_spin_box():
            interval = send_interval.spin_box.value()
            self.config["settings"]["send_interval"] = interval
            self.save_config()

        send_interval.spin_box.valueChanged.connect(change_spin_box)

        vbox_left.addWidget(info)
        vbox_left.addWidget(self.msg)
        vbox_left.addWidget(send_interval)
        vbox_left.addWidget(self.auto_delete_files)
        
        # 同步文件夹布局
        sync_folder_layout = QHBoxLayout()
        sync_folder_layout.addWidget(self.sync_folder_enabled)
        sync_folder_layout.addWidget(self.sync_folder_path)
        sync_folder_layout.addWidget(self.sync_folder_btn)
        sync_folder_layout.addWidget(self.sync_folder_manual_btn)
        vbox_left.addLayout(sync_folder_layout)
        
        vbox_left.addWidget(send_btn)

        # 右边的选择内容界面
        vbox_right = QVBoxLayout()
        vbox_right.stretch(1)


        load_btn = QPushButton("加载内容txt文件")
        load_btn.clicked.connect(load_text)

        text_btn = QPushButton("添加文本内容")
        text_btn.clicked.connect(add_text)

        file_btn = QPushButton("添加文件")
        file_btn.clicked.connect(add_file)

        del_btn = QPushButton("删除内容")
        del_btn.clicked.connect(del_content)

        vbox_right.addWidget(text_btn)
        vbox_right.addWidget(file_btn)
        vbox_right.addWidget(del_btn)
        vbox_right.addWidget(load_btn)

        # 整体布局
        hbox = QHBoxLayout()
        hbox.addLayout(vbox_left)
        hbox.addLayout(vbox_right)

        return hbox

    # 提供选择微信语言版本的按钮
    def init_language_choose(self):
        def switch_language():
            if lang_zh_CN_btn.isChecked():
                self.wechat.lc = WeChatLocale("zh-CN")
                self.config["settings"]["language"] = "zh-CN"

            elif lang_zh_TW_btn.isChecked():
                self.wechat.lc = WeChatLocale("zh-TW")
                self.config["settings"]["language"] = "zh-TW"

            elif lang_en_btn.isChecked():
                self.wechat.lc = WeChatLocale("en-US")
                self.config["settings"]["language"] = "en-US"

            self.save_config()

        # 提示信息
        info = QLabel("请选择你的微信系统语言")

        # 选择按钮
        lang_zh_CN_btn = QRadioButton("简体中文")
        lang_zh_TW_btn = QRadioButton("繁体中文")
        lang_en_btn = QRadioButton("English")

        # 优化：安全读取语言配置
        lang = self.config["settings"].get("language", "zh-CN")
        if lang == "zh-CN":
            lang_zh_CN_btn.setChecked(True)
        elif lang == "zh-TW":
            lang_zh_TW_btn.setChecked(True)
        elif lang == "en-US":
            lang_en_btn.setChecked(True)

        # 选择按钮的响应事件
        lang_zh_CN_btn.clicked.connect(switch_language)
        lang_zh_TW_btn.clicked.connect(switch_language)
        lang_en_btn.clicked.connect(switch_language)

        # 整体布局
        hbox = QHBoxLayout()
        hbox.addWidget(lang_zh_CN_btn)
        hbox.addWidget(lang_zh_TW_btn)
        hbox.addWidget(lang_en_btn)

        vbox = QVBoxLayout()
        vbox.addWidget(info)
        vbox.addLayout(hbox)

        return vbox

    def initUI(self):
        # 垂直布局
        vbox = QVBoxLayout()

        # 关于自动打开微信界面的按钮
        self.wechat_notice_btn = QPushButton("关于自动打开微信界面", self)
        self.wechat_notice_btn.resize(self.wechat_notice_btn.sizeHint())
        self.wechat_notice_btn.clicked.connect(self.show_wechat_open_notice)

        # 选择微信语言界面
        lang = self.init_language_choose()

        # 用户选择界面
        contacts = self.init_choose_contacts()

        # 发送内容界面
        msg_widget = self.init_send_msg()

        # 定时界面
        clock = self.init_clock()
        
        # 消息监控界面
        def init_monitor():
            # 创建滚动区域
            scroll_area = QScrollArea()
            scroll_area.setWidgetResizable(True)
            scroll_content = QWidget()
            
            hbox = QHBoxLayout(scroll_content)
            
            # 左边的消息显示区域（占70%宽度）
            vbox_left = QVBoxLayout()
            info = QLabel("消息监控（实时显示收到的新消息）")
            self.monitor_view = QListWidget()
            self.monitor_view.setSelectionMode(QAbstractItemView.ExtendedSelection)
            vbox_left.addWidget(info)
            vbox_left.addWidget(self.monitor_view)
            hbox.addLayout(vbox_left, stretch=7)
            
            # 右边的按钮和关键词警报区域（占30%宽度，支持滚动）
            scroll_right = QScrollArea()
            scroll_right.setWidgetResizable(True)
            content_right = QWidget()
            vbox_right = QVBoxLayout(content_right)
            vbox_right.addStretch(1)
            
            # ========== 普通消息监控 ==========
            # 监控开关
            self.monitor_btn = QPushButton("开始监控消息")
            self.monitor_running = False
            
            # 槽函数：在主线程中添加监控消息
            def _append_monitor_message(msg):
                self.monitor_view.addItem(msg)
                # 自动滚动到底部
                self.monitor_view.scrollToBottom()
                # 最多保存100条消息
                if self.monitor_view.count() > 100:
                    self.monitor_view.takeItem(0)

            # 消息回调函数（在后台线程调用，通过信号切到主线程更新UI）
            def on_new_message(sender, content, time, msg_type):
                # 在监控列表中显示新消息
                msg = f"[{time}] [{msg_type}] {sender}: {content}"
                self.new_message_signal.emit(msg)
            
            def toggle_monitor():
                if not self.monitor_running:
                    # 开始监控
                    self.wechat.set_message_callback(on_new_message)
                    self.wechat.start_monitor(check_interval=2)
                    self.monitor_running = True
                    self.monitor_btn.setText("停止监控消息")
                    self.monitor_btn.setStyleSheet("color:red")
                    QMessageBox.information(self, "监控已启动", "微信消息监控已启动！\n\n提示：\n• 基于UIAutomation技术\n• 支持文本、图片、文件、语音识别\n• 微信窗口可以在后台运行")
                else:
                    # 停止监控
                    self.wechat.stop_monitor()
                    self.monitor_running = False
                    self.monitor_btn.setText("开始监控消息")
                    self.monitor_btn.setStyleSheet("color:black")
                    QMessageBox.information(self, "监控已停止", "微信消息监控已停止！")
            
            self.monitor_btn.clicked.connect(toggle_monitor)
            
            # 清空消息按钮
            clear_monitor_btn = QPushButton("清空消息列表")
            def clear_monitor():
                self.monitor_view.clear()
            clear_monitor_btn.clicked.connect(clear_monitor)
            
            # 获取聊天列表按钮
            get_chat_list_btn = QPushButton("获取聊天列表")
            def show_chat_list():
                chat_list = self.wechat.get_chat_list()
                if chat_list:
                    QMessageBox.information(self, "聊天列表", "\n".join(chat_list[:20]) + ("\n..." if len(chat_list) > 20 else ""))
                else:
                    QMessageBox.warning(self, "获取失败", "未获取到聊天列表，请确保微信窗口已打开！")
            get_chat_list_btn.clicked.connect(show_chat_list)
            
            vbox_right.addWidget(self.monitor_btn)
            vbox_right.addWidget(clear_monitor_btn)
            vbox_right.addWidget(get_chat_list_btn)
            
            # ========== 精准最后一条消息监控（关键词警报） ==========
            # 添加分隔线
            vbox_right.addWidget(QLabel("------------------------"))
            vbox_right.addWidget(QLabel("🔔 单窗口关键词警报"))
            
            # 关键词输入框
            keyword_label = QLabel("触发关键词（多个用|分隔，例如 !|！|加急）")
            self.keyword_input = QLineEdit()
            # 默认值：匹配感叹号
            self.keyword_input.setText("!|！")
            keyword_hint = QLabel("提示：保持聊天窗口打开即可持续监控")
            keyword_hint.setStyleSheet("color:gray; font-size: 10px")
            
            # 警报开关
            self.alert_btn = QPushButton("启动单窗口警报")
            self.alert_running = False
            self.last_alert_state = False
            
            def on_last_message_change(last_text, current_time):
                # 获取当前设置的关键词
                keywords = self.keyword_input.text().strip()
                if not keywords:
                    return
                    
                # 检查是否包含任何关键词
                has_keyword = False
                for kw in keywords.split('|'):
                    if kw.strip() and kw in last_text:
                        has_keyword = True
                        break
                
                # 每一条新消息只要包含关键词就触发警报
                if has_keyword:
                    # 添加到消息列表（通过信号切到主线程）
                    msg = f"[{current_time}] 🚨 警报: 检测到关键词 '{last_text}'"
                    self.alert_message_signal.emit(msg)
                    # 弹出提示框 - 使用 QTimer.singleShot 保证在主线程安全执行
                    def show_alert():
                        QMessageBox.warning(self, "关键词警报", 
                            f"检测到包含关键词的新消息！\n\n时间: {current_time}\n内容: {last_text}\n\n"
                            "请及时查看处理。")
                    QTimer.singleShot(0, show_alert)
                
                # 状态记录（只用于记录解除信息，不阻塞重复警报）
                if has_keyword != self.last_alert_state:
                    if not has_keyword and self.last_alert_state:
                        # 警报解除
                        msg = f"[{current_time}] ✅ 警报解除: 当前内容 '{last_text}'"
                        self.alert_message_signal.emit(msg)
                    self.last_alert_state = has_keyword
            
            def toggle_alert():
                if not self.alert_running:
                    # 检查关键词是否为空
                    keywords = self.keyword_input.text().strip()
                    if not keywords:
                        QMessageBox.warning(self, "输入错误", "请先输入触发关键词！")
                        return
                    
                    # 开始监控
                    self.wechat.start_last_message_monitor(callback=on_last_message_change, check_interval=1)
                    self.alert_running = True
                    self.last_alert_state = False
                    self.alert_btn.setText("停止单窗口警报")
                    self.alert_btn.setStyleSheet("color:red")
                    QMessageBox.information(self, "警报已启动", 
                        "单窗口精准关键词警报已启动！\n\n"
                        "说明：\n"
                        "• 监控当前打开聊天窗口的最后一条消息\n"
                        "• 当消息包含你设置的关键词时会弹窗提醒\n"
                        "• 关键词用 | 分隔，支持多关键词匹配\n"
                        "• 每条新消息只要含关键词都会触发警报\n"
                        "• 保持聊天窗口打开即可持续监控")
                else:
                    # 停止监控
                    self.wechat.stop_last_message_monitor()
                    self.alert_running = False
                    self.alert_btn.setText("启动单窗口警报")
                    self.alert_btn.setStyleSheet("color:black")
                    QMessageBox.information(self, "警报已停止", "单窗口关键词警报已停止！")
            
            self.alert_btn.clicked.connect(toggle_alert)
            
            vbox_right.addWidget(keyword_label)
            vbox_right.addWidget(self.keyword_input)
            vbox_right.addWidget(keyword_hint)
            vbox_right.addWidget(self.alert_btn)
            vbox_right.addStretch()
            
            # 完成右侧滚动区域
            content_right.setLayout(vbox_right)
            scroll_right.setWidget(content_right)
            hbox.addWidget(scroll_right, stretch=3)
            
            # 设置滚动区域内容
            scroll_content.setLayout(hbox)
            scroll_area.setWidget(scroll_content)
            
            # 返回滚动区域布局（占满全屏）
            vbox_wrapper = QVBoxLayout()
            vbox_wrapper.addWidget(scroll_area)
            return vbox_wrapper
        
        monitor = init_monitor()

        vbox.addWidget(self.wechat_notice_btn)
        vbox.addLayout(lang)
        vbox.addLayout(contacts)
        vbox.addStretch(5)
        vbox.addLayout(msg_widget)
        vbox.addStretch(5)
        vbox.addLayout(clock)
        vbox.addStretch(5)
        vbox.addLayout(monitor)
        vbox.addStretch(1)

        #获取显示器分辨率
        desktop = QApplication.desktop()
        screenRect = desktop.screenGeometry()
        height = screenRect.height()
        width = screenRect.width()

        self.setLayout(vbox)
        # self.setFixedSize(width*0.2, height*0.6)
        self.setWindowTitle('EasyChat微信助手(作者：LTEnjoy)')
        self.show()


if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = WechatGUI()
    sys.exit(app.exec_())