import sys
import time
import os
import random
import json
import datetime
import threading
import winsound 

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
        self.log_file_path = "Momo运行日志.txt"
        
        if os.path.exists(self.config_path):
            with open(self.config_path, "r", encoding="utf-8") as r:
                self.config = json.load(r)
                if "settings" not in self.config:
                    self.config["settings"] = {}
        else:
            self.config = {"settings": {"language": "zh-CN", "trigger_sender": "momo", "rules": [], "rule_count": 1}}

        # 配置文件兼容与升级逻辑
        settings = self.config["settings"]
        
        if not settings.get("wechat_path"):
            settings["wechat_path"] = r"C:\Program Files\Tencent\Weixin.exe"
            
        old_kw = settings.pop("trigger_keywords", None)
        old_folder = settings.pop("material_folder", None)
        old_global_match = settings.pop("trigger_mode", "contains")
        
        if "rules" not in settings:
            settings["rules"] = []
            if old_kw or old_folder:
                settings["rules"].append({
                    "keywords": old_kw or "", 
                    "match": old_global_match, 
                    "type": "image", 
                    "content": old_folder or ""
                })
        
        while len(settings["rules"]) < 5:
            settings["rules"].append({"keywords": "", "match": "contains", "type": "text", "content": ""})
            
        for rule in settings["rules"]:
            if "match" not in rule:
                rule["match"] = "contains"
                
        if "rule_count" not in settings:
            settings["rule_count"] = 1
            
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
            
        self.add_log("🚀 欢迎使用多规则自动回复助手！")

    def save_config(self):
        with open(self.config_path, "w", encoding="utf8") as w:
            json.dump(self.config, w, indent=4, ensure_ascii=False)

    def get_valid_images(self, folder):
        if not os.path.exists(folder):
            return []
        image_extensions = {'.jpg', '.jpeg', '.png', '.gif', '.bmp', '.webp'}
        return [os.path.join(folder, file) for file in os.listdir(folder) if os.path.splitext(file)[1].lower() in image_extensions]

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
        msg_box.setText("新版多规则自动回复")
        msg_box.setInformativeText(
            "⚠️ 进阶功能说明：\n"
            "• 【独立匹配模式】：现在每条规则都可以单独设置是“精确匹配”还是“包含匹配”了！\n"
            "• 【自动记录日志】：所有的触发和发送记录都会永久保存在软件目录下的 Momo运行日志.txt 中。\n"
            "• 【一键重启唤醒】：如果发现抓不到消息，点击[一键重启并唤醒微信]按钮，软件会自动：关掉当前微信 -> 开讲述人 -> 重新开微信 -> 关讲述人。完美避开微信的检测盲区！\n"
        )
        msg_box.setStandardButtons(QMessageBox.Ok)
        msg_box.exec_()

    # --- 修正版：彻底关闭微信 -> 开讲述人 -> 开微信 -> 关讲述人 ---
    def auto_wake_wechat(self):
        # 防误触二次确认
        reply = QMessageBox.question(self, '确认重启微信',
                                     "此操作将强制关闭您当前正在运行的微信，\n并在后台开启无障碍环境后重新启动它。\n\n重启后您可能需要重新点击登录。\n是否继续执行？",
                                     QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
        if reply == QMessageBox.No:
            return

        self.add_log("🛠️ 开始执行一键重启并唤醒微信流程...")
        self.wake_btn.setEnabled(False)
        self.wake_btn.setText("唤醒中...")
        
        def task():
            try:
                # 1. 彻底杀掉微信进程
                self.add_log("🛑 正在强制关闭微信...")
                os.system("taskkill /F /IM WeChat.exe >nul 2>&1")
                time.sleep(2.0)
                
                # 2. 启动讲述人
                self.add_log("👁️ 正在开启 Windows 讲述人...")
                os.system("start narrator")
                time.sleep(3.0) # 等待讲述人完全就绪
                
                # 3. 启动微信
                wechat_path = self.config.get("settings", {}).get("wechat_path", "")
                if wechat_path and os.path.exists(wechat_path):
                    self.add_log("🚀 正在以无障碍环境重新启动微信...")
                    import subprocess
                    subprocess.Popen(wechat_path)
                    time.sleep(6.0) # 给微信充足的启动和加载时间
                else:
                    self.add_log("❌ 唤醒失败：微信exe路径配置不正确！")
                    
                # 4. 关闭讲述人
                self.add_log("🤫 微信已重新调起，正在关闭讲述人...")
                os.system("taskkill /F /IM narrator.exe >nul 2>&1")
                
                self.add_log("✅ 唤醒全流程结束！请您登录微信并打开目标聊天窗口后，再点击[开始监控]。")
                
            except Exception as e:
                self.add_log(f"❌ 唤醒过程出错: {e}")
            finally:
                QMetaObject.invokeMethod(self.wake_btn, "setEnabled", Qt.QueuedConnection, Q_ARG(bool, True))
                QMetaObject.invokeMethod(self.wake_btn, "setText", Qt.QueuedConnection, Q_ARG(str, "🛠️ 一键重启并唤醒微信"))
                
        threading.Thread(target=task, daemon=True).start()

    def init_language_choose(self):
        def switch_language():
            lang = "zh-CN" if lang_zh_CN_btn.isChecked() else "zh-TW" if lang_zh_TW_btn.isChecked() else "en-US"
            self.wechat.lc = WeChatLocale(lang)
            self.config["settings"]["language"] = lang
            self.save_config()

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
        hbox.addWidget(QLabel("微信语言:"))
        hbox.addWidget(lang_zh_CN_btn)
        hbox.addWidget(lang_zh_TW_btn)
        hbox.addWidget(lang_en_btn)
        hbox.addStretch(1)
        return hbox

    def init_settings(self):
        settings_config = self.config.get("settings", {})
        form_layout = QFormLayout()

        wechat_path_input = QLineEdit()
        wechat_path_input.setText(settings_config.get("wechat_path", r"C:\Program Files\Tencent\Weixin.exe"))
        wechat_path_input.editingFinished.connect(
            lambda: self.config["settings"].update({"wechat_path": wechat_path_input.text().strip()}) or self.save_config()
        )
        
        wechat_path_btn = QPushButton("浏览...")
        def choose_wechat_path():
            path, _ = QFileDialog.getOpenFileName(self, "选择微信.exe", "", "可执行文件(*.exe)")
            if path:
                wechat_path_input.setText(path)
                self.config["settings"]["wechat_path"] = path
                self.save_config()
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

        rules_group = QGroupBox("触发规则设置 (从上到下优先级递减)")
        rules_vbox = QVBoxLayout()
        
        count_hbox = QHBoxLayout()
        count_hbox.addWidget(QLabel("启用规则数量:"))
        self.rule_count_cb = QComboBox()
        self.rule_count_cb.addItems(["1 组", "2 组", "3 组", "4 组", "5 组"])
        current_count = settings_config.get("rule_count", 1)
        self.rule_count_cb.setCurrentIndex(current_count - 1)
        count_hbox.addWidget(self.rule_count_cb)
        count_hbox.addStretch(1)
        rules_vbox.addLayout(count_hbox)

        grid = QGridLayout()
        grid.addWidget(QLabel("触发关键词(逗号分隔)"), 0, 0)
        grid.addWidget(QLabel("匹配模式"), 0, 1)    
        grid.addWidget(QLabel("回复方式"), 0, 2)
        grid.addWidget(QLabel("回复内容 (文本/文件夹)"), 0, 3)
        
        grid.setColumnStretch(0, 3)
        grid.setColumnStretch(1, 2)
        grid.setColumnStretch(2, 2)
        grid.setColumnStretch(3, 4)
        
        self.rule_rows_ui = [] 
        
        for i in range(5):
            rule = settings_config["rules"][i]
            
            kw_inp = QLineEdit(rule.get("keywords", ""))
            kw_inp.setPlaceholderText(f"规则 {i+1} 关键词")
            
            match_cb = QComboBox()
            match_cb.addItems(["包含匹配", "精确匹配"])
            match_cb.setCurrentIndex(0 if rule.get("match", "contains") == "contains" else 1)
            
            type_cb = QComboBox()
            type_cb.addItems(["回复固定文本", "回复随机图片"])
            type_cb.setCurrentIndex(0 if rule.get("type", "text") == "text" else 1)
            
            content_inp = QLineEdit(rule.get("content", ""))
            browse_btn = QPushButton("📁")
            browse_btn.setFixedWidth(30)
            browse_btn.setVisible(type_cb.currentIndex() == 1)
            
            content_layout = QHBoxLayout()
            content_layout.setContentsMargins(0, 0, 0, 0)
            content_layout.addWidget(content_inp)
            content_layout.addWidget(browse_btn)
            content_widget = QWidget()
            content_widget.setLayout(content_layout)
            
            grid.addWidget(kw_inp, i+1, 0)
            grid.addWidget(match_cb, i+1, 1)
            grid.addWidget(type_cb, i+1, 2)
            grid.addWidget(content_widget, i+1, 3)
            
            self.rule_rows_ui.append((kw_inp, match_cb, type_cb, content_widget))
            
            def bind_events(idx, k_w, m_w, t_w, c_w, b_w):
                def update_cfg():
                    self.config["settings"]["rules"][idx]["keywords"] = k_w.text().strip()
                    self.config["settings"]["rules"][idx]["match"] = "contains" if m_w.currentIndex() == 0 else "exact"
                    self.config["settings"]["rules"][idx]["type"] = "text" if t_w.currentIndex() == 0 else "image"
                    self.config["settings"]["rules"][idx]["content"] = c_w.text().strip()
                    b_w.setVisible(t_w.currentIndex() == 1)
                    self.save_config()
                    
                k_w.editingFinished.connect(update_cfg)
                c_w.editingFinished.connect(update_cfg)
                m_w.currentIndexChanged.connect(update_cfg)
                t_w.currentIndexChanged.connect(update_cfg)
                
                def browse():
                    path = QFileDialog.getExistingDirectory(self, "选择素材文件夹")
                    if path:
                        c_w.setText(path)
                        update_cfg()
                b_w.clicked.connect(browse)
                
            bind_events(i, kw_inp, match_cb, type_cb, content_inp, browse_btn)
            
        rules_vbox.addLayout(grid)
        rules_group.setLayout(rules_vbox)
        form_layout.addRow(rules_group)
        
        def on_rule_count_changed(idx):
            count = idx + 1
            self.config["settings"]["rule_count"] = count
            self.save_config()
            for i, (w1, w2, w3, w4) in enumerate(self.rule_rows_ui):
                is_visible = i < count
                w1.setVisible(is_visible)
                w2.setVisible(is_visible)
                w3.setVisible(is_visible)
                w4.setVisible(is_visible)
                
        self.rule_count_cb.currentIndexChanged.connect(on_rule_count_changed)
        on_rule_count_changed(self.rule_count_cb.currentIndex()) 
        
        delay_hbox = QHBoxLayout()
        self.delay_spin = QDoubleSpinBox()
        self.delay_spin.setRange(0, 60)
        self.delay_spin.setValue(settings_config.get("send_delay", 0))
        self.delay_spin.setSingleStep(0.5)
        self.delay_spin.valueChanged.connect(lambda v: self.config["settings"].update({"send_delay": v}) or self.save_config())
        delay_hbox.addWidget(QLabel("基础延迟(分):"))
        delay_hbox.addWidget(self.delay_spin)
        
        self.random_delay_spin = QDoubleSpinBox()
        self.random_delay_spin.setRange(0, 30)
        self.random_delay_spin.setValue(settings_config.get("random_delay", 0))
        self.random_delay_spin.setSingleStep(0.5)
        self.random_delay_spin.valueChanged.connect(lambda v: self.config["settings"].update({"random_delay": v}) or self.save_config())
        delay_hbox.addWidget(QLabel("浮动范围(分):"))
        delay_hbox.addWidget(self.random_delay_spin)
        delay_hbox.addStretch(1)
        form_layout.addRow(delay_hbox)
        
        self.enable_sound_cb = QCheckBox("触发时播放系统提示音 (防漏接)")
        self.enable_sound_cb.setChecked(settings_config.get("enable_sound", True))
        self.enable_sound_cb.stateChanged.connect(
            lambda state: self.config["settings"].update({"enable_sound": state == Qt.Checked}) or self.save_config()
        )
        form_layout.addRow("", self.enable_sound_cb)
        
        form_layout.addRow(QLabel("------------------------"))
        
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
        start_hbox.addStretch(1)
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
        end_hbox.addStretch(1)
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
        self.log_view = QListWidget()
        self.log_view.setSelectionMode(QAbstractItemView.ExtendedSelection)
        vbox.addWidget(QLabel("监控日志 (已开启自动保存到文件)"))
        vbox.addWidget(self.log_view)
        return vbox

    def add_log(self, message):
        self.add_log_signal.emit(message)

    def _do_add_log(self, message):
        current_time = time.strftime("%H:%M:%S")
        full_message = f"[{current_time}] {message}"
        
        self.log_view.addItem(full_message)
        self.log_view.scrollToBottom()
        if self.log_view.count() > 200:
            self.log_view.takeItem(0)
            
        try:
            with open(self.log_file_path, "a", encoding="utf-8") as f:
                f.write(f"[{time.strftime('%Y-%m-%d')}] {full_message}\n")
        except Exception:
            pass

    def on_last_message_change(self, last_text, current_time):
        settings_config = self.config.get("settings", {})
        trigger_sender = settings_config.get("trigger_sender", "momo") 
        rules = settings_config.get("rules", [])
        rule_count = settings_config.get("rule_count", 1) 
        
        clean_text = str(last_text).strip()
        matched_rule = None
        matched_index = -1
        
        for idx in range(rule_count):
            rule = rules[idx]
            kw_str = rule.get("keywords", "").strip()
            content = rule.get("content", "").strip()
            rule_match_mode = rule.get("match", "contains") 
            
            if not kw_str or not content:
                continue
                
            keywords = [k.strip() for k in kw_str.split(',') if k.strip()]
            for keyword in keywords:
                if rule_match_mode == "exact":
                    if clean_text == keyword:
                        matched_rule = rule
                        matched_index = idx + 1
                        break
                else: 
                    if keyword in clean_text:
                        matched_rule = rule
                        matched_index = idx + 1
                        break
            if matched_rule:
                break
        
        if matched_rule and not self.last_triggered:
            self.last_triggered = True
            self.add_log(f"🚨 触发【规则{matched_index}】! 抓取内容: '{last_text}'")
            
            if settings_config.get("enable_sound", True):
                try:
                    winsound.PlaySound("SystemAsterisk", winsound.SND_ALIAS | winsound.SND_ASYNC)
                except:
                    pass
            
            base_delay = settings_config.get("send_delay", 0)
            random_range = settings_config.get("random_delay", 0)
            
            if base_delay > 0 or random_range > 0:
                half_range = random_range / 2
                actual_delay = max(0, random.uniform(base_delay - half_range, base_delay + half_range) if random_range > 0 else base_delay)
                self.add_log(f"⏳ 将在 {actual_delay:.1f} 分钟后执行回复...")
                delay_seconds = int(actual_delay * 60)
                
                def delayed_send():
                    wait_time = delay_seconds
                    while wait_time > 0:
                        if not getattr(self, 'monitoring', False):
                            self.add_log("⏹️ 监控已停止，取消本次延迟发送")
                            self.last_triggered = False
                            return
                        time.sleep(1)
                        wait_time -= 1
                        
                    if getattr(self, 'monitoring', False):
                        self._do_send_action(trigger_sender, matched_rule)
                        
                threading.Thread(target=delayed_send, daemon=True).start()
            else:
                self._do_send_action(trigger_sender, matched_rule)
                
        elif not matched_rule and self.last_triggered:
            self.last_triggered = False
            self.add_log(f"✅ 状态重置：对方最新消息变成了: '{last_text}'")
    
    def _do_send_action(self, trigger_sender, rule):
        if not self.monitoring:
            return
            
        r_type = rule.get("type", "text")
        content = rule.get("content", "")
        
        try:
            if r_type == "text":
                self.add_log(f"📡 正在发送文本回复...")
                self.wechat.send_text(trigger_sender, content, search_user=False)
                self.add_log(f"📤 文本发送完毕: {content}")
                
            elif r_type == "image":
                self.add_log(f"📡 正在获取随机图片...")
                images = self.get_valid_images(content)
                if not images:
                    self.add_log("❌ 指定的素材文件夹中没有图片，无法发送")
                    return
                
                selected_image = random.choice(images)
                self.add_log(f"🎲 抽中图片: {os.path.basename(selected_image)}")
                
                self.wechat.send_file(trigger_sender, selected_image, search_user=False)
                self.add_log(f"📤 图片发送完毕")
                
                time.sleep(1.0) 
                try:
                    os.remove(selected_image)
                    self.add_log(f"🗑️ 已安全删除: {os.path.basename(selected_image)}")
                except PermissionError:
                    self.add_log(f"⚠️ 图片正被微信占用未能删除，请稍后手动清理。")
                    
        except Exception as e:
            self.add_log(f"❌ 发送失败: {str(e)}")
        finally:
            self.last_triggered = False

    def start_monitoring(self):
        if self.monitoring: return
        
        settings = self.config.get("settings", {})
        rules = settings.get("rules", [])
        rule_count = settings.get("rule_count", 1)
        
        valid_count = sum(1 for i in range(rule_count) if rules[i].get("keywords") and rules[i].get("content"))
        if valid_count == 0:
            QMessageBox.warning(self, "错误", "在当前启用的规则中，没有检测到有效配置！请确保至少有一组填写了关键词和回复内容。")
            return
        
        self.monitoring = True
        self.last_triggered = False
        self.add_log(f"🚀 [{time.strftime('%Y-%m-%d %H:%M:%S')}] 启动精准监控 (共检测到 {valid_count} 组有效规则)")
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
            self.add_log(f"🤖 到达设定开始时间 {settings.get('auto_start_hour', 10):02d}:{settings.get('auto_start_minute', 0):02d}，自动启动监控")
            self.start_monitoring()
        elif not should_be_monitoring and self.monitoring:
            self.add_log(f"🤖 到达设定结束时间 {settings.get('auto_end_hour', 12):02d}:{settings.get('auto_end_minute', 0):02d}，自动停止监控")
            self.stop_monitoring()
        
    def initUI(self):
        vbox = QVBoxLayout()
        
        header_hbox = QHBoxLayout()
        header_hbox.addLayout(self.init_language_choose())
        
        # 修正的唤醒按钮
        self.wake_btn = QPushButton("🛠️ 一键重启并唤醒微信", self)
        self.wake_btn.setStyleSheet("background-color: #FFFACD; padding: 5px;")
        self.wake_btn.clicked.connect(self.auto_wake_wechat)
        header_hbox.addWidget(self.wake_btn)
        
        self.wechat_notice_btn = QPushButton("查看说明", self)
        self.wechat_notice_btn.clicked.connect(self.show_wechat_open_notice)
        header_hbox.addWidget(self.wechat_notice_btn)

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

        vbox.addLayout(header_hbox)
        vbox.addLayout(self.init_settings())
        vbox.addLayout(self.init_monitor_log())
        vbox.addLayout(hbox_controls)

        self.setLayout(vbox)
        screen_rect = QApplication.primaryScreen().geometry()
        self.setFixedSize(int(screen_rect.width() * 0.55), int(screen_rect.height() * 0.90))
        self.setWindowTitle('多规则自动回复助手')
        self.show()

if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = MomoReplyGUI()
    sys.exit(app.exec_())