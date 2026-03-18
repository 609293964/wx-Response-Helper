import sys
import time
import os
import random
import json
import datetime
import threading

from PyQt5.QtWidgets import *
from PyQt5.QtCore import *
from PyQt5.QtGui import *
from ui_auto_wechat import WeChat
from wechat_locale import WeChatLocale


class MomoReplyGUI(QWidget):
    add_log_signal = pyqtSignal(str)

    def __init__(self):
        super().__init__()

        self.config_path = "wechat_config_momo.json"
        default_material_folder = os.path.join(os.path.expanduser("~"), "Desktop", "素材")
        
        if os.path.exists(self.config_path):
            with open(self.config_path, "r", encoding="utf-8") as r:
                self.config = json.load(r)
                if "settings" not in self.config:
                    self.config["settings"] = {}
        else:
            self.config = {
                "settings": {
                    "wechat_path": "",
                    "language": "zh-CN",
                    "material_folder": default_material_folder,
                    "trigger_sender": "momo",
                    "trigger_keywords": "!,！",
                }
            }
            self.save_config()

        self.wechat = WeChat(
            path=self.config.get("settings", {}).get("wechat_path", ""),
            locale=self.config.get("settings", {}).get("language", "zh-CN"),
        )
        
        self.monitoring = False
        self.last_triggered = False
        self.auto_timer = None

        self.add_log_signal.connect(self._do_add_log)
        self.initUI()
        
        if self.config.get("settings", {}).get("enable_auto_timer", False):
            self.enable_auto_timer.setChecked(True)
            self.start_auto_timer_check()
        
        self.show_wechat_open_notice()

    def save_config(self):
        with open(self.config_path, "w", encoding="utf8") as w:
            json.dump(self.config, w, indent=4, ensure_ascii=False)

    def get_valid_images(self, folder):
        if not os.path.exists(folder):
            return []
        image_extensions = {'.jpg', '.jpeg', '.png', '.gif', '.bmp', '.webp'}
        images = []
        for file in os.listdir(folder):
            ext = os.path.splitext(file)[1].lower()
            if ext in image_extensions:
                images.append(os.path.join(folder, file))
        return images

    def closeEvent(self, event):
        self.monitoring = False
        self.last_triggered = False
        
        if hasattr(self, 'wechat') and hasattr(self.wechat, 'stop_last_message_monitor'):
            self.wechat.stop_last_message_monitor()
            
        if self.auto_timer is not None:
            self.stop_auto_timer_check()
            
        event.accept()

    def show_wechat_open_notice(self):
        msg_box = QMessageBox(self)
        msg_box.setIcon(QMessageBox.Information)
        msg_box.setWindowTitle("重要提示")
        msg_box.setText("微信打开方式提示")
        msg_box.setInformativeText(
            "由于微信版本更新，我们现在使用微信内置的快捷键来打开/隐藏微信窗口，请确保你的微信打开快捷键为Ctrl+Alt+w。具体查看方式为“设置”->“快捷键”->“显示/隐藏窗口”\n\n"
            "⚠️ 使用说明：\n"
            "• 请打开并保持与指定联系人的聊天窗口在前台\n"
            "• 当对方发送满足【触发关键词】条件的消息时，自动随机回复素材文件夹中的一张图片\n"
            "• 图片发送后会自动从素材文件夹中删除，避免重复发送\n\n"
        )
        msg_box.setStandardButtons(QMessageBox.Ok)
        msg_box.exec_()

    def init_language_choose(self):
        def switch_language():
            lang = "zh-CN" if lang_zh_CN_btn.isChecked() else "zh-TW" if lang_zh_TW_btn.isChecked() else "en-US"
            self.wechat.lc = WeChatLocale(lang)
            self.config["settings"]["language"] = lang
            self.save_config()

        info = QLabel("请选择你的微信系统语言")
        lang_zh_CN_btn = QRadioButton("简体中文")
        lang_zh_TW_btn = QRadioButton("繁体中文")
        lang_en_btn = QRadioButton("English")

        current_lang = self.config.get("settings", {}).get("language", "zh-CN")
        if current_lang == "zh-CN": lang_zh_CN_btn.setChecked(True)
        elif current_lang == "zh-TW": lang_zh_TW_btn.setChecked(True)
        elif current_lang == "en-US": lang_en_btn.setChecked(True)

        lang_zh_CN_btn.clicked.connect(switch_language)
        lang_zh_TW_btn.clicked.connect(switch_language)
        lang_en_btn.clicked.connect(switch_language)

        hbox = QHBoxLayout()
        hbox.addWidget(lang_zh_CN_btn)
        hbox.addWidget(lang_zh_TW_btn)
        hbox.addWidget(lang_en_btn)

        vbox = QVBoxLayout()
        vbox.addWidget(info)
        vbox.addLayout(hbox)
        return vbox

    def init_settings(self):
        settings_config = self.config.get("settings", {})
        
        def choose_wechat_path():
            path, _ = QFileDialog.getOpenFileName(self, "选择微信.exe", "", "可执行文件(*.exe)")
            if path:
                wechat_path_input.setText(path)
                self.config["settings"]["wechat_path"] = path
                self.save_config()

        def choose_material_folder():
            folder_path = QFileDialog.getExistingDirectory(self, "选择素材文件夹")
            if folder_path:
                material_folder_input.setText(folder_path)
                self.config["settings"]["material_folder"] = folder_path
                self.save_config()
                update_image_count()

        def update_image_count():
            folder = material_folder_input.text().strip()
            if os.path.exists(folder):
                images = self.get_valid_images(folder)
                image_count_label.setText(f"当前素材文件夹中有 {len(images)} 张图片")
            else:
                image_count_label.setText("素材文件夹不存在")

        form_layout = QFormLayout()

        wechat_path_input = QLineEdit()
        wechat_path_input.setText(settings_config.get("wechat_path", ""))
        wechat_path_btn = QPushButton("浏览...")
        wechat_path_btn.clicked.connect(choose_wechat_path)
        hbox_wechat = QHBoxLayout()
        hbox_wechat.addWidget(wechat_path_input)
        hbox_wechat.addWidget(wechat_path_btn)
        form_layout.addRow("微信exe路径:", hbox_wechat)

        trigger_sender_input = QLineEdit()
        trigger_sender_input.setText(settings_config.get("trigger_sender", "momo"))
        trigger_sender_input.editingFinished.connect(
            lambda: self.config["settings"].update({"trigger_sender": trigger_sender_input.text()}) or self.save_config()
        )
        form_layout.addRow("触发者昵称:", trigger_sender_input)

        trigger_keywords_input = QLineEdit()
        trigger_keywords_input.setText(settings_config.get("trigger_keywords", "!,！"))
        trigger_keywords_input.editingFinished.connect(
            lambda: self.config["settings"].update({"trigger_keywords": trigger_keywords_input.text()}) or self.save_config()
        )
        form_layout.addRow("触发关键词(用逗号分隔):", trigger_keywords_input)

        default_folder = os.path.join(os.path.expanduser("~"), "Desktop", "素材")
        material_folder_input = QLineEdit()
        material_folder_input.setText(settings_config.get("material_folder", default_folder))
        
        def on_material_folder_edited():
            self.config["settings"]["material_folder"] = material_folder_input.text()
            self.save_config()
            update_image_count()
            
        material_folder_input.editingFinished.connect(on_material_folder_edited)
        
        material_folder_btn = QPushButton("浏览...")
        material_folder_btn.clicked.connect(choose_material_folder)
        hbox_folder = QHBoxLayout()
        hbox_folder.addWidget(material_folder_input)
        hbox_folder.addWidget(material_folder_btn)
        form_layout.addRow("素材文件夹:", hbox_folder)

        image_count_label = QLabel()
        update_image_count()
        form_layout.addRow("", image_count_label)
        
        delay_label = QLabel("检测到触发后基础延迟（分钟）:")
        self.delay_spin = QDoubleSpinBox()
        self.delay_spin.setRange(0, 60)
        self.delay_spin.setValue(settings_config.get("send_delay", 0))
        self.delay_spin.setSingleStep(0.5)
        self.delay_spin.setDecimals(1)
        self.delay_spin.valueChanged.connect(lambda v: self.config["settings"].update({"send_delay": v}) or self.save_config())
        form_layout.addRow(delay_label, self.delay_spin)
        
        random_label = QLabel("随机浮动范围（分钟）:")
        self.random_delay_spin = QDoubleSpinBox()
        self.random_delay_spin.setRange(0, 30)
        self.random_delay_spin.setValue(settings_config.get("random_delay", 0))
        self.random_delay_spin.setSingleStep(0.5)
        self.random_delay_spin.setDecimals(1)
        self.random_delay_spin.valueChanged.connect(lambda v: self.config["settings"].update({"random_delay": v}) or self.save_config())
        form_layout.addRow(random_label, self.random_delay_spin)
        random_hint = QLabel("提示：最终延迟 = 基础延迟 ± (浮动范围/2)，总宽度为你输入的浮动值，随机取值")
        random_hint.setStyleSheet("color:gray; font-size: 10px")
        form_layout.addRow("", random_hint)
        
        form_layout.addRow(QLabel("------------------------"))
        
        # 【修改点1】：将界面文案修改为泛用型的“触发关键词”
        self.trigger_mode_exact = QRadioButton("完全匹配单独的触发关键词")
        self.trigger_mode_contains = QRadioButton("只要包含触发关键词就触发")
        
        if settings_config.get("trigger_mode", "exact") == "exact":
            self.trigger_mode_exact.setChecked(True)
        else:
            self.trigger_mode_contains.setChecked(True)
            
        def update_trigger_mode():
            self.config["settings"]["trigger_mode"] = "exact" if self.trigger_mode_exact.isChecked() else "contains"
            self.save_config()
            
        self.trigger_mode_exact.clicked.connect(update_trigger_mode)
        self.trigger_mode_contains.clicked.connect(update_trigger_mode)
        form_layout.addRow("", self.trigger_mode_exact)
        form_layout.addRow("", self.trigger_mode_contains)
        
        form_layout.addRow(QLabel("------------------------"))
        form_layout.addRow(QLabel("定时自动启停（可选）:"))
        
        start_hbox = QHBoxLayout()
        self.start_hour = QSpinBox()
        self.start_hour.setRange(0, 23)
        self.start_hour.setValue(settings_config.get("auto_start_hour", 10))
        self.start_minute = QSpinBox()
        self.start_minute.setRange(0, 59)
        self.start_minute.setValue(settings_config.get("auto_start_minute", 0))
        
        def update_start_time():
            self.config["settings"].update({"auto_start_hour": self.start_hour.value(), "auto_start_minute": self.start_minute.value()})
            self.save_config()
            
        self.start_hour.valueChanged.connect(update_start_time)
        self.start_minute.valueChanged.connect(update_start_time)
        start_hbox.addWidget(QLabel("每日开始:"))
        start_hbox.addWidget(self.start_hour)
        start_hbox.addWidget(QLabel("时"))
        start_hbox.addWidget(self.start_minute)
        start_hbox.addWidget(QLabel("分"))
        form_layout.addRow(start_hbox)
        
        end_hbox = QHBoxLayout()
        self.end_hour = QSpinBox()
        self.end_hour.setRange(0, 23)
        self.end_hour.setValue(settings_config.get("auto_end_hour", 12))
        self.end_minute = QSpinBox()
        self.end_minute.setRange(0, 59)
        self.end_minute.setValue(settings_config.get("auto_end_minute", 0))
        
        def update_end_time():
            self.config["settings"].update({"auto_end_hour": self.end_hour.value(), "auto_end_minute": self.end_minute.value()})
            self.save_config()
            
        self.end_hour.valueChanged.connect(update_end_time)
        self.end_minute.valueChanged.connect(update_end_time)
        end_hbox.addWidget(QLabel("每日结束:"))
        end_hbox.addWidget(self.end_hour)
        end_hbox.addWidget(QLabel("时"))
        end_hbox.addWidget(self.end_minute)
        end_hbox.addWidget(QLabel("分"))
        form_layout.addRow(end_hbox)
        
        self.enable_auto_timer = QCheckBox("启用每日定时自动启停")
        self.enable_auto_timer.setChecked(settings_config.get("enable_auto_timer", False))
        
        def toggle_auto_timer(state):
            self.config["settings"]["enable_auto_timer"] = (state == Qt.Checked)
            self.save_config()
            if state == Qt.Checked:
                self.start_auto_timer_check()
            else:
                self.stop_auto_timer_check()
                
        self.enable_auto_timer.stateChanged.connect(toggle_auto_timer)
        form_layout.addRow("", self.enable_auto_timer)

        return form_layout

    def init_monitor_log(self):
        vbox = QVBoxLayout()
        info = QLabel("监控日志")
        self.log_view = QListWidget()
        self.log_view.setSelectionMode(QAbstractItemView.ExtendedSelection)
        vbox.addWidget(info)
        vbox.addWidget(self.log_view)
        return vbox

    def add_log(self, message):
        self.add_log_signal.emit(message)

    def _do_add_log(self, message):
        current_time = time.strftime("%H:%M:%S")
        self.log_view.addItem(f"[{current_time}] {message}")
        self.log_view.scrollToBottom()
        if self.log_view.count() > 200:
            self.log_view.takeItem(0)

    def on_last_message_change(self, last_text, current_time):
        settings_config = self.config.get("settings", {})
        trigger_sender = settings_config.get("trigger_sender", "momo")
        trigger_keywords = settings_config.get("trigger_keywords", "!,！")
        material_folder = settings_config.get("material_folder", "")
        trigger_mode = settings_config.get("trigger_mode", "exact")
        
        keywords = [k.strip() for k in trigger_keywords.split(',') if k.strip()]
        triggered = False
        clean_text = str(last_text).strip()
        
        for keyword in keywords:
            if trigger_mode == "exact":
                if clean_text == keyword:
                    triggered = True
                    break
            else:
                if keyword in clean_text:
                    triggered = True
                    break
        
        if triggered and not self.last_triggered:
            self.last_triggered = True
            # 【修改点2】：将日志警告中的感叹号替换为泛用名称
            self.add_log(f"🚨🚨🚨 【高危警报】检测到触发关键词！抓取内容: '{last_text}'")
            
            base_delay = settings_config.get("send_delay", 0)
            random_range = settings_config.get("random_delay", 0)
            
            if base_delay > 0 or random_range > 0:
                half_range = random_range / 2
                actual_delay = max(0, random.uniform(base_delay - half_range, base_delay + half_range) if random_range > 0 else base_delay)
                self.add_log(f"⏳ 将在 {actual_delay:.1f} 分钟后发送图片...")
                delay_seconds = int(actual_delay * 60)
                
                def delayed_send():
                    wait_time = delay_seconds
                    while wait_time > 0:
                        if not getattr(self, 'monitoring', False):
                            self.add_log("⏹️ 监控已停止，取消本次延迟发送计划")
                            self.last_triggered = False
                            return
                        time.sleep(1)
                        wait_time -= 1
                        
                    if getattr(self, 'monitoring', False):
                        self._do_send_image(trigger_sender, material_folder, current_time)
                        
                threading.Thread(target=delayed_send, daemon=True).start()
            else:
                self._do_send_image(trigger_sender, material_folder, current_time)
                
        elif not triggered and self.last_triggered:
            self.last_triggered = False
            self.add_log(f"✅ 警报解除：最后一条消息变成了: '{last_text}'")
    
    def _do_send_image(self, trigger_sender, material_folder, trigger_time):
        if not self.monitoring:
            return
            
        self.add_log(f"📡 开始执行发送...")
        images = self.get_valid_images(material_folder)
        
        if not images:
            self.add_log("❌ 素材文件夹中没有图片，无法发送")
            self.last_triggered = False
            return
        
        selected_image = random.choice(images)
        self.add_log(f"🎲 选择图片: {os.path.basename(selected_image)}")
        
        try:
            self.wechat.send_file(trigger_sender, selected_image, search_user=False)
            self.add_log(f"📤 图片操作发送完毕")
            
            time.sleep(1.0) 
            try:
                os.remove(selected_image)
                self.add_log(f"🗑️ 已删除图片: {os.path.basename(selected_image)}")
            except PermissionError:
                self.add_log(f"⚠️ 警告: 图片正被微信占用，未能删除，请稍后手动清理。")
                
        except Exception as e:
            self.add_log(f"❌ 发送失败: {str(e)}")
        finally:
            self.last_triggered = False

    def start_monitoring(self):
        if self.monitoring:
            return
            
        material_folder = self.config.get("settings", {}).get("material_folder", "")
        if not os.path.exists(material_folder) or not self.get_valid_images(material_folder):
            QMessageBox.warning(self, "错误", "素材文件夹不存在或为空，请准备好图片！")
            return
        
        self.monitoring = True
        self.last_triggered = False
        self.monitor_start_time = time.strftime("%Y-%m-%d %H:%M:%S")
        self.add_log(f"🚀 [{self.monitor_start_time}] 启动精准监控")
        
        self.wechat.start_last_message_monitor(callback=self.on_last_message_change, check_interval=1)
        
        self.start_btn.setEnabled(False)
        self.stop_btn.setEnabled(True)
        self.start_btn.setStyleSheet("color:gray; padding: 10px;")
        self.stop_btn.setStyleSheet("color:red; font-size: 14px; padding: 10px;")

    def stop_monitoring(self):
        if self.monitoring:
            self.monitoring = False
            self.wechat.stop_last_message_monitor()
            self.add_log(f"⏹️ [{time.strftime('%Y-%m-%d %H:%M:%S')}] 监控已停止")
            
            self.start_btn.setEnabled(True)
            self.stop_btn.setEnabled(False)
            self.start_btn.setStyleSheet("color:green; font-size: 14px; padding: 10px;")
            self.stop_btn.setStyleSheet("color:gray; padding: 10px;")
    
    def start_auto_timer_check(self):
        if self.auto_timer is None:
            self.auto_timer = QTimer(self)
            self.auto_timer.timeout.connect(self.auto_check_time)
            self.auto_timer.start(60000)
            self.add_log("⏰ 定时自动启停检查已开启")
    
    def stop_auto_timer_check(self):
        if self.auto_timer is not None:
            self.auto_timer.stop()
            self.auto_timer = None
            self.add_log("⏹️ 定时自动启停检查已关闭")
    
    def auto_check_time(self):
        now = datetime.datetime.now()
        settings = self.config.get("settings", {})
        
        current_total = now.hour * 60 + now.minute
        start_total = settings.get("auto_start_hour", 10) * 60 + settings.get("auto_start_minute", 0)
        end_total = settings.get("auto_end_hour", 12) * 60 + settings.get("auto_end_minute", 0)
        
        should_be_monitoring = start_total <= current_total < end_total
        
        if should_be_monitoring and not self.monitoring:
            self.add_log("🤖 到达设定时间，自动启动")
            self.start_monitoring()
        elif not should_be_monitoring and self.monitoring:
            self.add_log("🤖 到达设定时间，自动停止")
            self.stop_monitoring()

    def initUI(self):
        vbox = QVBoxLayout()

        self.wechat_notice_btn = QPushButton("使用说明", self)
        self.wechat_notice_btn.clicked.connect(self.show_wechat_open_notice)

        hbox_controls = QHBoxLayout()
        self.start_btn = QPushButton("开始监控")
        self.start_btn.setStyleSheet("color:green; font-size: 14px; padding: 10px;")
        self.start_btn.clicked.connect(self.start_monitoring)
        
        self.stop_btn = QPushButton("停止监控")
        self.stop_btn.setStyleSheet("color:gray; padding: 10px;")
        self.stop_btn.setEnabled(False)
        self.stop_btn.clicked.connect(self.stop_monitoring)
        
        hbox_controls.addWidget(self.start_btn)
        hbox_controls.addWidget(self.stop_btn)

        vbox.addWidget(self.wechat_notice_btn)
        vbox.addLayout(self.init_language_choose())
        vbox.addLayout(self.init_settings())
        vbox.addLayout(self.init_monitor_log())
        vbox.addLayout(hbox_controls)

        self.setLayout(vbox)
        
        screen_rect = QApplication.primaryScreen().geometry()
        self.setFixedSize(int(screen_rect.width() * 0.5), int(screen_rect.height() * 0.7))
        self.setWindowTitle('Momo自动回复 - 专属定制版')
        self.show()


if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = MomoReplyGUI()
    sys.exit(app.exec_())