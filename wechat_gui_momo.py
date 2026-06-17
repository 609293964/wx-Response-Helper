import datetime
import ctypes
import csv
import json
import os
import random
import secrets
import shutil
import socket as network_socket
import subprocess
import sys
import threading
import time
import traceback
from pathlib import Path

# Keep Qt controls sharp and correctly sized on 4K/high-DPI displays while
# still allowing users to override Qt scaling with their own environment.
os.environ.setdefault("QT_AUTO_SCREEN_SCALE_FACTOR", "1")
os.environ.setdefault("QT_ENABLE_HIGHDPI_SCALING", "1")
os.environ.setdefault("QT_SCALE_FACTOR_ROUNDING_POLICY", "PassThrough")

from windows_dpi import configure_windows_dpi_awareness


configure_windows_dpi_awareness()

APP_WINDOW_TITLE = "EasyChat Momo - 微信自动回复助手"
ONE_CLICK_LOGIN_CONFIRM_TIMEOUT_SECONDS = 180


def is_running_as_admin():
    if os.name != "nt":
        return True
    try:
        return bool(ctypes.windll.shell32.IsUserAnAdmin())
    except Exception:
        return False


def relaunch_as_admin():
    if os.name != "nt":
        return False

    if getattr(sys, "frozen", False):
        executable = sys.executable
        params = subprocess.list2cmdline(sys.argv[1:])
    else:
        executable = sys.executable
        script = str(Path(sys.argv[0]).resolve())
        params = subprocess.list2cmdline([script, *sys.argv[1:]])

    try:
        result = ctypes.windll.shell32.ShellExecuteW(
            None,
            "runas",
            executable,
            params,
            str(Path.cwd()),
            1,
        )
    except Exception:
        return False
    return result > 32


def show_admin_launch_error():
    message = "EasyChat Momo needs administrator permission to start."
    try:
        ctypes.windll.user32.MessageBoxW(None, message, "EasyChat Momo", 0x10)
    except Exception:
        print(message, file=sys.stderr)


def activate_existing_app_window():
    if os.name != "nt":
        return False
    try:
        hwnd = ctypes.windll.user32.FindWindowW(None, APP_WINDOW_TITLE)
        if not hwnd:
            return False
        if getattr(sys, "frozen", False):
            base_dir = Path(sys.executable).resolve().parent
        else:
            base_dir = Path(__file__).resolve().parent
        request_path = base_dir / "logs" / "activate.request"
        request_path.parent.mkdir(parents=True, exist_ok=True)
        request_path.write_text(str(time.time()), encoding="ascii")
        ctypes.windll.user32.ShowWindow(hwnd, 9)
        ctypes.windll.user32.SetForegroundWindow(hwnd)
        return True
    except Exception:
        return False


def report_unhandled_exception(exc_type, exc_value, exc_traceback):
    if issubclass(exc_type, KeyboardInterrupt):
        sys.__excepthook__(exc_type, exc_value, exc_traceback)
        return

    if getattr(sys, "frozen", False):
        base_dir = Path(sys.executable).resolve().parent
    else:
        base_dir = Path(__file__).resolve().parent
    error_path = base_dir / "logs" / "startup_error.log"
    try:
        error_path.parent.mkdir(parents=True, exist_ok=True)
        with error_path.open("a", encoding="utf-8") as handle:
            handle.write(f"\n[{datetime.datetime.now():%Y-%m-%d %H:%M:%S}]\n")
            traceback.print_exception(
                exc_type,
                exc_value,
                exc_traceback,
                file=handle,
            )
    except Exception:
        pass

    message = f"程序运行异常：{exc_value}\n\n详细信息已写入：\n{error_path}"
    try:
        ctypes.windll.user32.MessageBoxW(None, message, "EasyChat Momo", 0x10)
    except Exception:
        pass


def ensure_admin_or_exit():
    if os.name != "nt" or is_running_as_admin():
        return
    if relaunch_as_admin():
        sys.exit(0)
    show_admin_launch_error()
    sys.exit(1)


if __name__ == "__main__":
    ensure_admin_or_exit()
    if activate_existing_app_window():
        sys.exit(0)


import uiautomation as auto
from PyQt5.QtCore import Qt, QTimer, pyqtSignal
from PyQt5.QtWidgets import (
    QAbstractItemView,
    QApplication,
    QCheckBox,
    QComboBox,
    QDoubleSpinBox,
    QFileDialog,
    QFormLayout,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QListWidget,
    QMessageBox,
    QPushButton,
    QRadioButton,
    QScrollArea,
    QSizePolicy,
    QSpinBox,
    QTabWidget,
    QVBoxLayout,
    QWidget,
)

from ui_auto_wechat import WeChat
from remote_control_server import API_VERSION, SERVICE_NAME, RemoteControlServer
from wechat_locale import WeChatLocale


class MomoReplyGUI(QWidget):
    add_log_signal = pyqtSignal(str)
    update_img_count_signal = pyqtSignal(int)
    update_scheduled_img_count_signal = pyqtSignal()
    status_signal = pyqtSignal(str, str)
    pause_scheduled_image_sender_signal = pyqtSignal(str)
    update_scheduled_calculation_signal = pyqtSignal()
    remote_command_signal = pyqtSignal(str)
    target_window_action_finished_signal = pyqtSignal(bool, str)
    one_click_target_finished_signal = pyqtSignal(bool, str)

    def __init__(self):
        super().__init__()
        self.base_dir = self.get_app_base_dir()
        self.config_path = self.base_dir / "wechat_config_momo.json"
        self.logs_dir = self.base_dir / "logs"
        self.app_log_path = self.logs_dir / "app.log"
        self.activation_request_path = self.logs_dir / "activate.request"
        self.config_load_notes = []

        self.config = self.load_config()
        self.wechat = WeChat(locale=self.config.get("settings", {}).get("language", "zh-CN"))

        self.monitoring = False
        self.last_triggered = False
        self.trigger_state_lock = threading.Lock()
        self.trigger_token = 0
        self.auto_timer = None
        self.scheduled_image_timer = None
        self.scheduled_image_next_at = None
        self.scheduled_image_sending = False
        self.scheduled_image_schedule_mode = None
        self.scheduled_image_schedule_base_at = None
        self.scheduled_image_pause_reason = ""
        self.scheduled_image_waiting_for_reply = False
        self.scheduled_image_waiting_text = ""
        self.scheduled_image_waiting_since = None
        self.scheduled_image_resume_after = None
        self.scheduled_image_deferred_due_at = None
        self.expected_self_message_until = None
        self.conversation_guard_lock = threading.Lock()
        self.status_timer = None
        self.rule_img_count_labels = {}
        self.scheduled_img_count_label = None
        self.scheduled_folder_input = None
        self.scheduled_image_calculation_label = None
        self.scheduled_pause_on_contact_checkbox = None
        self.scheduled_resume_delay_spin = None
        self.scheduled_mark_replied_btn = None
        self.status_labels = {}
        self.status_values = {}
        self.last_remote_error = ""
        self.last_startup_check_summary = "尚未自检"
        self.remote_control_server = None
        self.remote_service_status_label = None
        self.remote_port_spin = None
        self.remote_enabled_checkbox = None
        self.wechat_launch_path_input = None
        self.one_click_run_btn = None
        self.one_click_run_active = False
        self.one_click_narrator_was_running = False
        self.one_click_stage_deadline = 0
        self.one_click_login_wait_deadline = 0
        self.one_click_tree_stable_count = 0
        self.one_click_login_button_clicked = False
        self.open_target_window_btn = None
        self.open_target_window_active = False

        self.add_log_signal.connect(self._do_add_log)
        self.update_img_count_signal.connect(self._do_update_img_count)
        self.update_scheduled_img_count_signal.connect(self._do_update_scheduled_img_count)
        self.status_signal.connect(self._do_update_status)
        self.pause_scheduled_image_sender_signal.connect(self._pause_scheduled_image_sender)
        self.update_scheduled_calculation_signal.connect(self._update_scheduled_image_calculation)
        self.remote_command_signal.connect(self._handle_remote_command)
        self.target_window_action_finished_signal.connect(self._finish_open_target_window)
        self.one_click_target_finished_signal.connect(self._finish_one_click_target_window)

        self.initUI()
        self.start_status_timer()
        self.start_remote_control_server()

        if self.config.get("settings", {}).get("enable_auto_timer", False):
            self.enable_auto_timer.setChecked(True)
            self.start_auto_timer_check()
        if self.config.get("settings", {}).get("enable_scheduled_image_sender", False):
            self.enable_scheduled_image_sender.setChecked(True)
            self.start_scheduled_image_sender()

        self.show_wechat_open_notice()
        QTimer.singleShot(100, self.run_startup_check)

    def get_app_base_dir(self):
        if getattr(sys, "frozen", False):
            return Path(sys.executable).resolve().parent
        return Path(__file__).resolve().parent

    def default_config(self):
        return {
            "settings": {
                "language": "zh-CN",
                "trigger_sender": "momo",
                "wechat_launch_path": "",
                "active_rules_count": 1,
                "send_delay": 2.0,
                "random_delay": 0,
                "enable_scheduled_image_sender": False,
                "scheduled_image_folder": "",
                "scheduled_image_start_hour": 9,
                "scheduled_image_start_minute": 0,
                "scheduled_image_first_delay_minutes": 30,
                "scheduled_image_end_hour": 23,
                "scheduled_image_end_minute": 0,
                "scheduled_image_interval_minutes": 60,
                "scheduled_image_random_window_minutes": 15,
                "scheduled_image_pause_on_contact_message": True,
                "scheduled_image_resume_delay_minutes": 5,
                "remote_control_enabled": True,
                "remote_control_bind": "0.0.0.0",
                "remote_control_port": 8765,
                "remote_control_token": secrets.token_urlsafe(24),
            },
            "rules": [
                {"keywords": "!,！", "reply_type": "image", "folder": "", "reply_text": "", "mode": "exact"},
                {"keywords": "", "reply_type": "image", "folder": "", "reply_text": "", "mode": "contains"},
                {"keywords": "", "reply_type": "image", "folder": "", "reply_text": "", "mode": "contains"},
                {"keywords": "", "reply_type": "image", "folder": "", "reply_text": "", "mode": "contains"},
                {"keywords": "", "reply_type": "image", "folder": "", "reply_text": "", "mode": "contains"},
            ],
        }

    def load_config(self):
        if not self.config_path.exists():
            config = self.default_config()
            self.config_load_notes.append("未找到配置文件，已创建默认配置")
            self.config = config
            self.save_config()
            return config

        try:
            with self.config_path.open("r", encoding="utf-8") as handle:
                config = json.load(handle)
        except Exception as exc:
            backup_path = self.config_path.with_suffix(f".broken_{time.strftime('%Y%m%d_%H%M%S')}.json")
            try:
                shutil.copy2(self.config_path, backup_path)
                self.config_load_notes.append(f"配置文件读取失败，已备份到 {backup_path.name}")
            except Exception:
                self.config_load_notes.append("配置文件读取失败，且备份失败")
            self.config_load_notes.append(f"已使用默认配置: {exc}")
            config = self.default_config()

        changed = self.normalize_config(config)
        if changed:
            self.config_load_notes.append("已自动修复配置字段")
            self.config = config
            self.save_config()
        return config

    def normalize_config(self, config):
        changed = False
        defaults = self.default_config()

        if not isinstance(config, dict):
            config.clear()
            config.update(defaults)
            return True

        settings = config.setdefault("settings", {})
        if not isinstance(settings, dict):
            config["settings"] = defaults["settings"].copy()
            settings = config["settings"]
            changed = True

        legacy_wechat_path = settings.get("wechat_path")
        if (
            "wechat_launch_path" not in settings
            and isinstance(legacy_wechat_path, str)
            and legacy_wechat_path.strip()
        ):
            settings["wechat_launch_path"] = legacy_wechat_path.strip()
            changed = True

        for stale_key in ("wechat_path", "material_folder", "trigger_keywords", "monitor_backend"):
            if stale_key in settings:
                settings.pop(stale_key, None)
                changed = True

        if "scheduled_image_first_delay_minutes" not in settings:
            legacy_first_delay = settings.get(
                "scheduled_image_start_random_window_minutes",
                settings.get("scheduled_image_min_lead_minutes", 30),
            )
            settings["scheduled_image_first_delay_minutes"] = legacy_first_delay
            changed = True

        for stale_key in (
            "scheduled_image_start_random_window_minutes",
            "scheduled_image_min_lead_minutes",
            "scheduled_image_min_lead_random_window_minutes",
        ):
            if stale_key in settings:
                settings.pop(stale_key, None)
                changed = True

        for key, value in defaults["settings"].items():
            if key not in settings:
                settings[key] = value
                changed = True

        try:
            active_count = int(settings.get("active_rules_count", 1))
        except (TypeError, ValueError):
            active_count = 1
        active_count = max(1, min(5, active_count))
        if settings.get("active_rules_count") != active_count:
            settings["active_rules_count"] = active_count
            changed = True

        numeric_ranges = {
            "scheduled_image_start_hour": (0, 23, 9),
            "scheduled_image_start_minute": (0, 59, 0),
            "scheduled_image_first_delay_minutes": (0, 1440, 30),
            "scheduled_image_end_hour": (0, 23, 23),
            "scheduled_image_end_minute": (0, 59, 0),
            "scheduled_image_interval_minutes": (1, 1440, 60),
            "scheduled_image_random_window_minutes": (0, 720, 15),
            "scheduled_image_resume_delay_minutes": (0, 120, 5),
            "remote_control_port": (1024, 65535, 8765),
        }
        for key, (minimum, maximum, default) in numeric_ranges.items():
            try:
                value = int(settings.get(key, default))
            except (TypeError, ValueError):
                value = default
            value = max(minimum, min(maximum, value))
            if settings.get(key) != value:
                settings[key] = value
                changed = True

        if not isinstance(settings.get("enable_scheduled_image_sender"), bool):
            settings["enable_scheduled_image_sender"] = bool(settings.get("enable_scheduled_image_sender"))
            changed = True
        if not isinstance(settings.get("scheduled_image_pause_on_contact_message"), bool):
            settings["scheduled_image_pause_on_contact_message"] = bool(
                settings.get("scheduled_image_pause_on_contact_message")
            )
            changed = True
        if not isinstance(settings.get("remote_control_enabled"), bool):
            settings["remote_control_enabled"] = bool(settings.get("remote_control_enabled"))
            changed = True
        if not isinstance(settings.get("remote_control_bind"), str):
            settings["remote_control_bind"] = "0.0.0.0"
            changed = True
        if not isinstance(settings.get("remote_control_token"), str) or not settings.get("remote_control_token"):
            settings["remote_control_token"] = secrets.token_urlsafe(24)
            changed = True
        if not isinstance(settings.get("scheduled_image_folder"), str):
            settings["scheduled_image_folder"] = ""
            changed = True
        if not isinstance(settings.get("wechat_launch_path"), str):
            settings["wechat_launch_path"] = ""
            changed = True

        rules = config.setdefault("rules", [])
        if not isinstance(rules, list):
            rules = []
            config["rules"] = rules
            changed = True

        while len(rules) < 5:
            rules.append(defaults["rules"][len(rules)].copy())
            changed = True
        if len(rules) > 5:
            del rules[5:]
            changed = True

        for index, rule in enumerate(rules):
            if not isinstance(rule, dict):
                rules[index] = defaults["rules"][index].copy()
                changed = True
                continue
            for key, value in defaults["rules"][index].items():
                if key not in rule:
                    rule[key] = value
                    changed = True
            if rule.get("reply_type") not in ("image", "text"):
                rule["reply_type"] = "image"
                changed = True
            if rule.get("mode") not in ("exact", "contains"):
                rule["mode"] = "exact"
                changed = True

        return changed

    def save_config(self):
        self.config_path.parent.mkdir(parents=True, exist_ok=True)
        with self.config_path.open("w", encoding="utf-8") as handle:
            json.dump(self.config, handle, indent=4, ensure_ascii=False)

    def find_default_wechat_launch_path(self):
        if os.name != "nt":
            return ""

        desktop_candidates = [
            Path.home() / "Desktop",
            Path(os.environ.get("PUBLIC", r"C:\Users\Public")) / "Desktop",
        ]
        one_drive = os.environ.get("OneDrive")
        if one_drive:
            desktop_candidates.extend([Path(one_drive) / "Desktop", Path(one_drive) / "桌面"])

        shortcut_names = ("微信.lnk", "WeChat.lnk", "Weixin.lnk")
        for desktop in desktop_candidates:
            for name in shortcut_names:
                candidate = desktop / name
                if candidate.is_file():
                    return str(candidate)

        program_files = [
            os.environ.get("ProgramFiles"),
            os.environ.get("ProgramFiles(x86)"),
            os.environ.get("LOCALAPPDATA"),
        ]
        relative_paths = (
            Path("Tencent") / "Weixin" / "Weixin.exe",
            Path("Tencent") / "WeChat" / "WeChat.exe",
        )
        for root in program_files:
            if not root:
                continue
            for relative_path in relative_paths:
                candidate = Path(root) / relative_path
                if candidate.is_file():
                    return str(candidate)
        return ""

    def get_wechat_launch_path(self):
        configured = self.config.get("settings", {}).get("wechat_launch_path", "").strip().strip('"')
        if configured and Path(configured).is_file():
            return configured

        detected = self.find_default_wechat_launch_path()
        if detected:
            self.config.setdefault("settings", {})["wechat_launch_path"] = detected
            if self.wechat_launch_path_input is not None:
                self.wechat_launch_path_input.setText(detected)
            self.save_config()
        return detected

    def get_lan_ipv4_addresses(self):
        addresses = set()
        try:
            hostname = network_socket.gethostname()
            for item in network_socket.getaddrinfo(hostname, None, network_socket.AF_INET):
                address = item[4][0]
                if address and not address.startswith("127."):
                    addresses.add(address)
        except OSError:
            pass
        return sorted(addresses)

    def get_remote_status(self):
        settings = self.config.get("settings", {})
        try:
            narrator_running = self.is_narrator_running()
            narrator_status = "已开启" if narrator_running else "未开启"
        except Exception as exc:
            narrator_running = False
            narrator_status = f"检查失败: {exc}"

        next_send_at = ""
        if self.scheduled_image_next_at is not None:
            next_send_at = self.scheduled_image_next_at.astimezone().isoformat(timespec="seconds")

        return {
            "service": SERVICE_NAME,
            "api_version": API_VERSION,
            "remote_enabled": bool(settings.get("remote_control_enabled", True)),
            "server_time": datetime.datetime.now().astimezone().isoformat(timespec="seconds"),
            "monitoring": bool(self.monitoring),
            "narrator_running": narrator_running,
            "narrator_status": narrator_status,
            "target": settings.get("trigger_sender", ""),
            "scheduled_enabled": bool(settings.get("enable_scheduled_image_sender", False)),
            "scheduled_sending": bool(self.scheduled_image_sending),
            "next_send_at": next_send_at,
            "timed_send_status": self.status_values.get("timed_send", "未启用"),
            "last_message": self.status_values.get("last_message", "-"),
            "last_trigger": self.status_values.get("last_trigger", "-"),
            "conversation_guard": self.status_values.get("conversation_guard", "未触发"),
            "send_result": self.status_values.get("send_result", "-"),
            "schedule_calculation": self._build_scheduled_image_calculation_text(),
            "last_error": self.last_remote_error,
            "startup_check_summary": self.last_startup_check_summary,
        }

    def start_remote_control_server(self):
        self.stop_remote_control_server()
        settings = self.config.get("settings", {})
        if not settings.get("remote_control_enabled", True):
            self._set_remote_service_status("未启用")
            self.last_remote_error = ""
            return

        host = settings.get("remote_control_bind", "0.0.0.0").strip() or "0.0.0.0"
        port = int(settings.get("remote_control_port", 8765))
        token = settings.get("remote_control_token", "")
        if not token:
            self.last_remote_error = "访问令牌为空，远程控制服务未启动"
            self._set_remote_service_status("启动失败: 访问令牌为空")
            self.add_log("手机远程控制服务启动失败: 访问令牌为空")
            return
        try:
            self.remote_control_server = RemoteControlServer(
                host=host,
                port=port,
                token=token,
                status_provider=self.get_remote_status,
                command_handler=self.remote_command_signal.emit,
                logger=self.add_log,
            )
            self.remote_control_server.start()
            addresses = self.get_lan_ipv4_addresses()
            endpoint = ", ".join(f"http://{address}:{port}" for address in addresses)
            if not endpoint:
                endpoint = f"http://本机IP:{port}"
            self._set_remote_service_status("运行中: " + endpoint)
            self.add_log("手机远程控制服务已启动: " + endpoint)
            self.last_remote_error = ""
        except Exception as exc:
            self.remote_control_server = None
            detail = self.format_remote_start_error(exc, host, port)
            self.last_remote_error = detail
            self._set_remote_service_status("启动失败: " + detail)
            self.add_log("手机远程控制服务启动失败: " + detail)

    def format_remote_start_error(self, exc, host, port):
        winerror = getattr(exc, "winerror", None)
        errno = getattr(exc, "errno", None)
        raw_message = str(exc).strip() or exc.__class__.__name__
        if winerror == 10048 or errno == 98:
            return f"端口 {port} 已被占用，请换一个监听端口后重试"
        if winerror == 10013 or errno == 13:
            return "绑定被系统拒绝，请检查 Windows 防火墙、管理员权限或专用网络授权"
        if winerror == 10049 or errno == 99:
            return f"绑定地址 {host} 不可用，请改回 0.0.0.0 或使用本机局域网地址"
        return raw_message

    def get_mobile_connection_package(self):
        settings = self.config.get("settings", {})
        address_list = self.get_lan_ipv4_addresses()
        host = address_list[0] if address_list else "本机IP"
        port = self.remote_port_spin.value() if self.remote_port_spin is not None else int(settings.get("remote_control_port", 8765))
        return {
            "service": SERVICE_NAME,
            "host": host,
            "port": port,
            "token": settings.get("remote_control_token", ""),
            "api_version": API_VERSION,
        }

    def copy_mobile_connection_package(self):
        package = self.get_mobile_connection_package()
        QApplication.clipboard().setText(json.dumps(package, ensure_ascii=False))
        self.add_log("已复制手机连接信息")
        QMessageBox.information(self, "已复制", "手机连接信息已复制，可直接粘贴到 HarmonyOS 应用设置页。")

    def stop_remote_control_server(self):
        server = self.remote_control_server
        self.remote_control_server = None
        if server is not None:
            server.stop()

    def restart_remote_control_server(self):
        settings = self.config.setdefault("settings", {})
        if self.remote_enabled_checkbox is not None:
            settings["remote_control_enabled"] = self.remote_enabled_checkbox.isChecked()
        if self.remote_port_spin is not None:
            settings["remote_control_port"] = self.remote_port_spin.value()
        self.save_config()
        self.start_remote_control_server()

    def _set_remote_service_status(self, message):
        if self.remote_service_status_label is not None:
            self.remote_service_status_label.setText(message)

    def _handle_remote_command(self, command):
        self.add_log(f"收到手机远程命令: {command}")
        if command == "narrator_start":
            self.start_narrator()
        elif command == "narrator_stop":
            self.stop_narrator()
        elif command == "monitoring_start":
            if not self.monitoring:
                self.start_monitoring()
        elif command == "monitoring_stop":
            if self.monitoring:
                self.stop_monitoring()
        elif command == "scheduled_start":
            self.config.setdefault("settings", {})["enable_scheduled_image_sender"] = True
            self.save_config()
            if hasattr(self, "enable_scheduled_image_sender"):
                if not self.enable_scheduled_image_sender.isChecked():
                    self.enable_scheduled_image_sender.setChecked(True)
                else:
                    self.start_scheduled_image_sender()
            else:
                self.start_scheduled_image_sender()
        elif command == "scheduled_stop":
            self.config.setdefault("settings", {})["enable_scheduled_image_sender"] = False
            self.save_config()
            if hasattr(self, "enable_scheduled_image_sender"):
                if self.enable_scheduled_image_sender.isChecked():
                    self.enable_scheduled_image_sender.setChecked(False)
                else:
                    self.stop_scheduled_image_sender()
            else:
                self.stop_scheduled_image_sender()
        elif command == "run_check":
            self.run_startup_check()

    def init_remote_control_settings(self):
        settings = self.config.get("settings", {})
        group = QGroupBox("手机远程控制")
        layout = QFormLayout()

        self.remote_enabled_checkbox = QCheckBox("允许局域网内的 HarmonyOS 应用控制")
        self.remote_enabled_checkbox.setChecked(settings.get("remote_control_enabled", True))

        self.remote_port_spin = QSpinBox()
        self.remote_port_spin.setRange(1024, 65535)
        self.remote_port_spin.setValue(settings.get("remote_control_port", 8765))

        token_input = QLineEdit(settings.get("remote_control_token", ""))
        token_input.setReadOnly(True)
        token_input.setEchoMode(QLineEdit.Password)

        show_token_checkbox = QCheckBox("显示令牌")
        show_token_checkbox.stateChanged.connect(
            lambda state: token_input.setEchoMode(QLineEdit.Normal if state == Qt.Checked else QLineEdit.Password)
        )

        copy_token_btn = QPushButton("复制令牌")
        copy_token_btn.clicked.connect(lambda: QApplication.clipboard().setText(token_input.text()))

        copy_connection_btn = QPushButton("复制手机连接信息")
        copy_connection_btn.clicked.connect(self.copy_mobile_connection_package)

        token_row = QHBoxLayout()
        token_row.addWidget(token_input, 1)
        token_row.addWidget(copy_token_btn)

        token_options = QHBoxLayout()
        token_options.addWidget(show_token_checkbox)
        token_options.addStretch()

        address_list = self.get_lan_ipv4_addresses()
        address_text = "、".join(address_list) if address_list else "请查看 Windows 当前局域网 IP"
        address_label = QLabel(address_text)
        address_label.setWordWrap(True)

        self.remote_service_status_label = QLabel("等待启动")
        self.remote_service_status_label.setWordWrap(True)

        apply_btn = QPushButton("应用并重启远程服务")
        apply_btn.clicked.connect(self.restart_remote_control_server)
        self.remote_enabled_checkbox.stateChanged.connect(lambda _state: self.restart_remote_control_server())

        layout.addRow(self.remote_enabled_checkbox)
        layout.addRow("监听端口:", self.remote_port_spin)
        layout.addRow("访问令牌:", token_row)
        layout.addRow("", token_options)
        layout.addRow("", copy_connection_btn)
        layout.addRow("电脑局域网 IP:", address_label)
        layout.addRow("服务状态:", self.remote_service_status_label)
        layout.addRow(apply_btn)
        group.setLayout(layout)
        return group

    def sync_scheduled_image_settings_from_ui(self, save=True):
        settings = self.config.setdefault("settings", {})
        if self.scheduled_folder_input is not None:
            settings["scheduled_image_folder"] = self.scheduled_folder_input.text().strip()
        if hasattr(self, "scheduled_start_hour"):
            settings.update(
                {
                    "scheduled_image_start_hour": self.scheduled_start_hour.value(),
                    "scheduled_image_start_minute": self.scheduled_start_minute.value(),
                    "scheduled_image_first_delay_minutes": self.scheduled_first_delay_spin.value(),
                    "scheduled_image_end_hour": self.scheduled_end_hour.value(),
                    "scheduled_image_end_minute": self.scheduled_end_minute.value(),
                }
            )
        if hasattr(self, "scheduled_interval_spin"):
            settings.update(
                {
                    "scheduled_image_interval_minutes": self.scheduled_interval_spin.value(),
                    "scheduled_image_random_window_minutes": self.scheduled_random_spin.value(),
                }
            )
        if self.scheduled_pause_on_contact_checkbox is not None:
            settings["scheduled_image_pause_on_contact_message"] = (
                self.scheduled_pause_on_contact_checkbox.isChecked()
            )
        if self.scheduled_resume_delay_spin is not None:
            settings["scheduled_image_resume_delay_minutes"] = self.scheduled_resume_delay_spin.value()
        if hasattr(self, "enable_scheduled_image_sender"):
            settings["enable_scheduled_image_sender"] = self.enable_scheduled_image_sender.isChecked()
        if save:
            self.save_config()

    def normalize_folder_path(self, folder):
        if not folder:
            return ""
        return os.path.expandvars(os.path.expanduser(str(folder).strip().strip('"')))

    def _is_windows_drive_path(self, folder):
        return os.name == "nt" and len(folder) >= 2 and folder[1] == ":"

    def _mapped_drive_to_unc(self, folder):
        if not self._is_windows_drive_path(folder):
            return ""
        try:
            import winreg

            drive_letter = folder[0].upper()
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER, fr"Network\{drive_letter}") as key:
                remote_path, _ = winreg.QueryValueEx(key, "RemotePath")
        except Exception:
            return ""

        relative = folder[2:].lstrip("\\/")
        return os.path.join(remote_path, relative) if relative else remote_path

    def resolve_accessible_folder_path(self, folder):
        folder = self.normalize_folder_path(folder)
        if not folder:
            return ""
        if os.path.isdir(folder):
            return folder

        unc_path = self._mapped_drive_to_unc(folder)
        if unc_path and os.path.isdir(unc_path):
            return unc_path
        return folder

    def get_folder_access_message(self, folder):
        folder = self.normalize_folder_path(folder)
        if not folder:
            return ""
        if os.path.isdir(folder):
            return ""

        unc_path = self._mapped_drive_to_unc(folder)
        if unc_path and os.path.isdir(unc_path):
            return ""
        if unc_path:
            return f"目录不可访问: {folder}；已尝试映射到共享路径 {unc_path}，但仍无法读取。"
        if self._is_windows_drive_path(folder):
            return (
                f"目录不可访问: {folder}。如果这是共享盘映射盘符，管理员模式可能看不到它；"
                r"请改用 \\服务器\共享\目录 这样的 UNC 路径，或在管理员会话中重新映射该盘符。"
            )
        if folder.startswith("\\\\"):
            return f"共享目录不可访问: {folder}。请确认网络连接、共享权限和账号访问权限。"
        return f"目录不存在或不可访问: {folder}"

    def image_count_text(self, folder):
        problem = self.get_folder_access_message(folder)
        if problem:
            return "图片数量: 目录不可访问"
        return f"图片数量: {len(self.get_valid_images(folder))}"

    def get_valid_images(self, folder):
        folder = self.resolve_accessible_folder_path(folder)
        if not folder or not os.path.isdir(folder):
            return []
        image_extensions = {".jpg", ".jpeg", ".png", ".gif", ".bmp", ".webp"}
        images = []
        try:
            for entry in os.scandir(folder):
                if not entry.is_file():
                    continue
                ext = os.path.splitext(entry.name)[1].lower()
                if ext in image_extensions:
                    images.append(entry.path)
        except OSError:
            return []
        return images

    def closeEvent(self, event):
        self.stop_remote_control_server()
        if self.monitoring:
            self.stop_monitoring()
        if self.auto_timer is not None:
            self.stop_auto_timer_check()
        if self.scheduled_image_timer is not None:
            self.stop_scheduled_image_sender()
        if self.status_timer is not None:
            self.status_timer.stop()
        QApplication.quit()
        event.accept()

    def show_wechat_open_notice(self):
        self.add_log(
            "操作说明: 可使用“一键启动监控”自动准备讲述人、微信、"
            "目标独立窗口和关键词监控；图片发送成功后会移动到素材目录下的 sent 文件夹。"
        )

    def init_language_choose(self):
        def switch_language():
            if lang_zh_cn_btn.isChecked():
                self.wechat.lc = WeChatLocale("zh-CN")
                self.config["settings"]["language"] = "zh-CN"
            elif lang_zh_tw_btn.isChecked():
                self.wechat.lc = WeChatLocale("zh-TW")
                self.config["settings"]["language"] = "zh-TW"
            elif lang_en_btn.isChecked():
                self.wechat.lc = WeChatLocale("en-US")
                self.config["settings"]["language"] = "en-US"
            self.save_config()

        lang_group = QGroupBox("系统语言设置")
        lang_layout = QHBoxLayout()
        lang_zh_cn_btn = QRadioButton("简体中文")
        lang_zh_tw_btn = QRadioButton("繁体中文")
        lang_en_btn = QRadioButton("English")

        current_lang = self.config.get("settings", {}).get("language", "zh-CN")
        if current_lang == "zh-CN":
            lang_zh_cn_btn.setChecked(True)
        elif current_lang == "zh-TW":
            lang_zh_tw_btn.setChecked(True)
        elif current_lang == "en-US":
            lang_en_btn.setChecked(True)

        lang_zh_cn_btn.clicked.connect(switch_language)
        lang_zh_tw_btn.clicked.connect(switch_language)
        lang_en_btn.clicked.connect(switch_language)

        lang_layout.addWidget(lang_zh_cn_btn)
        lang_layout.addWidget(lang_zh_tw_btn)
        lang_layout.addWidget(lang_en_btn)
        lang_group.setLayout(lang_layout)
        return lang_group

    def init_settings(self):
        settings_config = self.config.get("settings", {})
        main_layout = QVBoxLayout()

        base_group = QGroupBox("基础设置")
        base_layout = QFormLayout()

        trigger_sender_input = QLineEdit(settings_config.get("trigger_sender", "momo"))
        trigger_sender_input.setPlaceholderText("填入聊天窗口的名字(如: momo)")
        trigger_sender_input.editingFinished.connect(
            lambda: self.config["settings"].update({"trigger_sender": trigger_sender_input.text().strip()})
            or self.save_config()
            or self.refresh_runtime_status()
        )
        base_layout.addRow("目标对话/触发者昵称:", trigger_sender_input)

        initial_wechat_path = settings_config.get("wechat_launch_path", "").strip()
        if not initial_wechat_path or not Path(initial_wechat_path).is_file():
            initial_wechat_path = self.get_wechat_launch_path()

        wechat_path_widget = QWidget()
        wechat_path_layout = QHBoxLayout(wechat_path_widget)
        wechat_path_layout.setContentsMargins(0, 0, 0, 0)
        self.wechat_launch_path_input = QLineEdit(initial_wechat_path)
        self.wechat_launch_path_input.setPlaceholderText("选择微信程序或桌面快捷方式")
        wechat_path_btn = QPushButton("选择文件")
        wechat_path_btn.setFixedWidth(92)

        def save_wechat_launch_path():
            path = self.wechat_launch_path_input.text().strip().strip('"')
            self.wechat_launch_path_input.setText(path)
            self.config.setdefault("settings", {})["wechat_launch_path"] = path
            self.save_config()

        def browse_wechat_launch_path():
            current_path = self.wechat_launch_path_input.text().strip()
            start_dir = str(Path(current_path).parent) if Path(current_path).is_file() else str(Path.home() / "Desktop")
            selected, _selected_filter = QFileDialog.getOpenFileName(
                self,
                "选择微信程序或快捷方式",
                start_dir,
                "微信程序或快捷方式 (*.exe *.lnk);;所有文件 (*)",
            )
            if not selected:
                return
            self.wechat_launch_path_input.setText(selected)
            save_wechat_launch_path()
            self.add_log(f"微信启动路径已保存: {selected}")

        self.wechat_launch_path_input.editingFinished.connect(save_wechat_launch_path)
        wechat_path_btn.clicked.connect(browse_wechat_launch_path)
        wechat_path_layout.addWidget(self.wechat_launch_path_input, 1)
        wechat_path_layout.addWidget(wechat_path_btn)
        base_layout.addRow("微信启动路径:", wechat_path_widget)

        base_group.setLayout(base_layout)
        main_layout.addWidget(base_group)

        rule_group = QGroupBox("触发与回复规则设置")
        rule_layout = QVBoxLayout()

        combo_layout = QHBoxLayout()
        combo_layout.addWidget(QLabel("启用的规则数量 (1~5):"))
        self.rule_count_combo = QComboBox()
        self.rule_count_combo.addItems(["1", "2", "3", "4", "5"])
        self.rule_count_combo.setCurrentIndex(settings_config.get("active_rules_count", 1) - 1)
        combo_layout.addWidget(self.rule_count_combo)
        combo_layout.addStretch()
        rule_layout.addLayout(combo_layout)

        self.rules_tabs = QTabWidget()
        self.rule_img_count_labels = {}

        for i in range(5):
            tab = QWidget()
            tab_layout = QFormLayout()
            rule_data = self.config["rules"][i]

            kw_input = QLineEdit(rule_data.get("keywords", ""))
            kw_input.setPlaceholderText("用英文逗号分隔，例如: !,！")
            kw_input.editingFinished.connect(
                lambda idx=i, edit=kw_input: self.config["rules"][idx].update({"keywords": edit.text()})
                or self.save_config()
            )
            tab_layout.addRow("触发关键词:", kw_input)

            type_combo = QComboBox()
            type_combo.addItems(["发送素材图片", "发送指定文本"])

            folder_w = QWidget()
            folder_layout = QHBoxLayout(folder_w)
            folder_layout.setContentsMargins(0, 0, 0, 0)
            folder_input = QLineEdit(rule_data.get("folder", ""))
            folder_input.setPlaceholderText("选择或粘贴图片素材文件夹路径")
            folder_btn = QPushButton("选择目录")
            folder_btn.setFixedWidth(92)
            folder_layout.addWidget(folder_input)
            folder_layout.addWidget(folder_btn)
            folder_info = QLabel("图片数量: 0")
            self.rule_img_count_labels[i] = folder_info

            text_w = QWidget()
            text_layout = QHBoxLayout(text_w)
            text_layout.setContentsMargins(0, 0, 0, 0)
            text_input = QLineEdit(rule_data.get("reply_text", ""))
            text_input.setPlaceholderText("填入触发后需要回复的文本内容")
            text_layout.addWidget(text_input)

            tab_layout.addRow("回复类型:", type_combo)
            tab_layout.addRow("图片目录:", folder_w)
            tab_layout.addRow("", folder_info)
            tab_layout.addRow("回复文本:", text_w)

            def update_visibility(idx, cb, fw, cl, tw):
                is_img = cb.currentIndex() == 0
                fw.setVisible(is_img)
                cl.setVisible(is_img)
                tw.setVisible(not is_img)
                self.config["rules"][idx]["reply_type"] = "image" if is_img else "text"
                self.save_config()

            def make_folder_browser(idx, f_input, f_info):
                def browse():
                    current_folder = self.resolve_accessible_folder_path(f_input.text().strip())
                    if not os.path.isdir(current_folder):
                        desktop_folder = Path.home() / "Desktop"
                        current_folder = str(desktop_folder if desktop_folder.exists() else Path.home())

                    dialog = QFileDialog(self, f"选择规则 {idx + 1} 的图片素材目录")
                    dialog.setFileMode(QFileDialog.Directory)
                    dialog.setOption(QFileDialog.ShowDirsOnly, True)
                    dialog.setDirectory(current_folder)

                    if dialog.exec_() != QFileDialog.Accepted:
                        return

                    selected = dialog.selectedFiles()
                    if not selected:
                        return

                    folder = selected[0]
                    f_input.setText(folder)
                    self.config["rules"][idx]["folder"] = folder
                    self.save_config()
                    f_info.setText(self.image_count_text(folder))
                    self.add_log(f"规则 {idx + 1} 图片目录已选择: {folder}")
                return browse

            def update_count(idx, f_input, f_info):
                self.config["rules"][idx]["folder"] = f_input.text().strip()
                self.save_config()
                f_info.setText(self.image_count_text(f_input.text().strip()))

            def update_text(idx, edit):
                self.config["rules"][idx]["reply_text"] = edit.text()
                self.save_config()

            type_combo.currentIndexChanged.connect(
                lambda _, idx=i, cb=type_combo, fw=folder_w, cl=folder_info, tw=text_w: update_visibility(
                    idx, cb, fw, cl, tw
                )
            )
            folder_btn.clicked.connect(make_folder_browser(i, folder_input, folder_info))
            folder_input.editingFinished.connect(
                lambda idx=i, f_in=folder_input, f_inf=folder_info: update_count(idx, f_in, f_inf)
            )
            text_input.editingFinished.connect(lambda idx=i, edit=text_input: update_text(idx, edit))

            curr_type = rule_data.get("reply_type", "image")
            type_combo.setCurrentIndex(0 if curr_type == "image" else 1)
            update_visibility(i, type_combo, folder_w, folder_info, text_w)
            update_count(i, folder_input, folder_info)

            mode_layout = QHBoxLayout()
            rb_exact = QRadioButton("只匹配单独关键词(精确)")
            rb_contains = QRadioButton("包含关键词即可触发(模糊)")
            if rule_data.get("mode", "exact") == "exact":
                rb_exact.setChecked(True)
            else:
                rb_contains.setChecked(True)

            def update_mode(idx, exact_btn):
                self.config["rules"][idx]["mode"] = "exact" if exact_btn.isChecked() else "contains"
                self.save_config()

            rb_exact.clicked.connect(lambda checked, idx=i, btn=rb_exact: update_mode(idx, btn))
            rb_contains.clicked.connect(lambda checked, idx=i, btn=rb_exact: update_mode(idx, btn))
            mode_layout.addWidget(rb_exact)
            mode_layout.addWidget(rb_contains)
            tab_layout.addRow("匹配模式:", mode_layout)

            tab.setLayout(tab_layout)
            self.rules_tabs.addTab(tab, f"规则 {i + 1}")

        def on_rule_count_changed():
            count = int(self.rule_count_combo.currentText())
            self.config["settings"]["active_rules_count"] = count
            self.save_config()
            for index in range(5):
                self.rules_tabs.setTabEnabled(index, index < count)

        self.rule_count_combo.currentIndexChanged.connect(on_rule_count_changed)
        on_rule_count_changed()

        rule_layout.addWidget(self.rules_tabs)
        rule_group.setLayout(rule_layout)
        main_layout.addWidget(rule_group)

        return main_layout

    def init_keyword_timing_settings(self):
        settings_config = self.config.get("settings", {})
        main_layout = QVBoxLayout()

        time_group = QGroupBox("回复延迟与监控时段")
        time_layout = QFormLayout()

        self.delay_spin = QDoubleSpinBox()
        self.delay_spin.setRange(0, 60)
        self.delay_spin.setValue(settings_config.get("send_delay", 0))
        self.delay_spin.setSingleStep(0.5)
        self.delay_spin.valueChanged.connect(
            lambda: self.config["settings"].update({"send_delay": self.delay_spin.value()}) or self.save_config()
        )

        self.random_delay_spin = QDoubleSpinBox()
        self.random_delay_spin.setRange(0, 30)
        self.random_delay_spin.setValue(settings_config.get("random_delay", 0))
        self.random_delay_spin.setSingleStep(0.5)
        self.random_delay_spin.valueChanged.connect(
            lambda: self.config["settings"].update({"random_delay": self.random_delay_spin.value()})
            or self.save_config()
        )

        delay_layout = QHBoxLayout()
        delay_layout.addWidget(QLabel("基础延迟(分):"))
        delay_layout.addWidget(self.delay_spin)
        delay_layout.addWidget(QLabel("随机浮动(分):"))
        delay_layout.addWidget(self.random_delay_spin)
        time_layout.addRow(delay_layout)

        start_hbox = QHBoxLayout()
        self.start_hour = QSpinBox()
        self.start_hour.setRange(0, 23)
        self.start_hour.setValue(settings_config.get("auto_start_hour", 10))
        self.start_minute = QSpinBox()
        self.start_minute.setRange(0, 59)
        self.start_minute.setValue(settings_config.get("auto_start_minute", 0))
        self.end_hour = QSpinBox()
        self.end_hour.setRange(0, 23)
        self.end_hour.setValue(settings_config.get("auto_end_hour", 12))
        self.end_minute = QSpinBox()
        self.end_minute.setRange(0, 59)
        self.end_minute.setValue(settings_config.get("auto_end_minute", 0))

        def update_time():
            self.config["settings"].update(
                {
                    "auto_start_hour": self.start_hour.value(),
                    "auto_start_minute": self.start_minute.value(),
                    "auto_end_hour": self.end_hour.value(),
                    "auto_end_minute": self.end_minute.value(),
                }
            )
            self.save_config()

        for widget in [self.start_hour, self.start_minute, self.end_hour, self.end_minute]:
            widget.valueChanged.connect(update_time)

        start_hbox.addWidget(QLabel("每日开始:"))
        start_hbox.addWidget(self.start_hour)
        start_hbox.addWidget(QLabel("时"))
        start_hbox.addWidget(self.start_minute)
        start_hbox.addWidget(QLabel("分"))
        start_hbox.addStretch()
        time_layout.addRow(start_hbox)

        end_hbox = QHBoxLayout()
        end_hbox.addWidget(QLabel("每日结束:"))
        end_hbox.addWidget(self.end_hour)
        end_hbox.addWidget(QLabel("时"))
        end_hbox.addWidget(self.end_minute)
        end_hbox.addWidget(QLabel("分"))
        end_hbox.addStretch()
        time_layout.addRow(end_hbox)

        self.enable_auto_timer = QCheckBox("启用每日定时自动启停监控")
        self.enable_auto_timer.setChecked(settings_config.get("enable_auto_timer", False))

        def toggle_auto_timer(state):
            self.config["settings"]["enable_auto_timer"] = state == Qt.Checked
            self.save_config()
            if state == Qt.Checked:
                self.start_auto_timer_check()
            else:
                self.stop_auto_timer_check()

        self.enable_auto_timer.stateChanged.connect(toggle_auto_timer)
        time_layout.addRow(self.enable_auto_timer)

        time_group.setLayout(time_layout)
        main_layout.addWidget(time_group)
        main_layout.addStretch()

        return main_layout

    def init_scheduled_image_settings(self):
        settings_config = self.config.get("settings", {})
        main_layout = QVBoxLayout()

        scheduled_group = QGroupBox("发送素材")
        scheduled_layout = QFormLayout()

        scheduled_folder_w = QWidget()
        scheduled_folder_layout = QHBoxLayout(scheduled_folder_w)
        scheduled_folder_layout.setContentsMargins(0, 0, 0, 0)
        scheduled_folder_input = QLineEdit(settings_config.get("scheduled_image_folder", ""))
        self.scheduled_folder_input = scheduled_folder_input
        scheduled_folder_input.setPlaceholderText("选择或粘贴定时发送的图片素材文件夹路径")
        scheduled_folder_btn = QPushButton("选择目录")
        scheduled_folder_btn.setFixedWidth(92)
        scheduled_folder_layout.addWidget(scheduled_folder_input)
        scheduled_folder_layout.addWidget(scheduled_folder_btn)
        self.scheduled_img_count_label = QLabel("图片数量: 0")

        def refresh_scheduled_count():
            self.sync_scheduled_image_settings_from_ui()
            folder = self.config.get("settings", {}).get("scheduled_image_folder", "")
            self.scheduled_img_count_label.setText(self.image_count_text(folder))

        def browse_scheduled_folder():
            current_folder = self.resolve_accessible_folder_path(scheduled_folder_input.text().strip())
            if not os.path.isdir(current_folder):
                desktop_folder = Path.home() / "Desktop"
                current_folder = str(desktop_folder if desktop_folder.exists() else Path.home())

            dialog = QFileDialog(self, "选择定时发送图片素材目录")
            dialog.setFileMode(QFileDialog.Directory)
            dialog.setOption(QFileDialog.ShowDirsOnly, True)
            dialog.setDirectory(current_folder)

            if dialog.exec_() != QFileDialog.Accepted:
                return

            selected = dialog.selectedFiles()
            if not selected:
                return

            folder = selected[0]
            scheduled_folder_input.setText(folder)
            refresh_scheduled_count()
            self.add_log(f"定时发图素材目录已选择: {folder}")

        scheduled_folder_btn.clicked.connect(browse_scheduled_folder)
        scheduled_folder_input.editingFinished.connect(refresh_scheduled_count)
        scheduled_layout.addRow("图片目录:", scheduled_folder_w)
        scheduled_layout.addRow("", self.scheduled_img_count_label)

        scheduled_group.setLayout(scheduled_layout)
        main_layout.addWidget(scheduled_group)

        scheduled_window_group = QGroupBox("发送窗口")
        scheduled_window_layout = QFormLayout()

        scheduled_window_hbox = QHBoxLayout()
        self.scheduled_start_hour = QSpinBox()
        self.scheduled_start_hour.setRange(0, 23)
        self.scheduled_start_hour.setValue(settings_config.get("scheduled_image_start_hour", 9))
        self.scheduled_start_minute = QSpinBox()
        self.scheduled_start_minute.setRange(0, 59)
        self.scheduled_start_minute.setValue(settings_config.get("scheduled_image_start_minute", 0))
        self.scheduled_first_delay_spin = QSpinBox()
        self.scheduled_first_delay_spin.setRange(0, 1440)
        self.scheduled_first_delay_spin.setValue(
            settings_config.get("scheduled_image_first_delay_minutes", 30)
        )
        self.scheduled_end_hour = QSpinBox()
        self.scheduled_end_hour.setRange(0, 23)
        self.scheduled_end_hour.setValue(settings_config.get("scheduled_image_end_hour", 23))
        self.scheduled_end_minute = QSpinBox()
        self.scheduled_end_minute.setRange(0, 59)
        self.scheduled_end_minute.setValue(settings_config.get("scheduled_image_end_minute", 0))

        def update_scheduled_time():
            self.sync_scheduled_image_settings_from_ui()
            self.reset_scheduled_image_next_send()

        for widget in [
            self.scheduled_start_hour,
            self.scheduled_start_minute,
            self.scheduled_first_delay_spin,
            self.scheduled_end_hour,
            self.scheduled_end_minute,
        ]:
            widget.valueChanged.connect(update_scheduled_time)

        scheduled_window_hbox.addWidget(QLabel("从"))
        scheduled_window_hbox.addWidget(self.scheduled_start_hour)
        scheduled_window_hbox.addWidget(QLabel("时"))
        scheduled_window_hbox.addWidget(self.scheduled_start_minute)
        scheduled_window_hbox.addWidget(QLabel("分 到"))
        scheduled_window_hbox.addWidget(self.scheduled_end_hour)
        scheduled_window_hbox.addWidget(QLabel("时"))
        scheduled_window_hbox.addWidget(self.scheduled_end_minute)
        scheduled_window_hbox.addWidget(QLabel("分"))
        scheduled_window_hbox.addStretch()
        scheduled_window_layout.addRow("发送窗口:", scheduled_window_hbox)

        scheduled_first_delay_hbox = QHBoxLayout()
        scheduled_first_delay_hbox.addWidget(QLabel("窗口开始后"))
        scheduled_first_delay_hbox.addWidget(self.scheduled_first_delay_spin)
        scheduled_first_delay_hbox.addWidget(QLabel("分钟发送第一次"))
        scheduled_first_delay_hbox.addStretch()
        scheduled_window_layout.addRow("首次延迟:", scheduled_first_delay_hbox)

        scheduled_window_group.setLayout(scheduled_window_layout)
        main_layout.addWidget(scheduled_window_group)

        scheduled_interval_group = QGroupBox("发送节奏")
        scheduled_interval_layout = QFormLayout()

        scheduled_interval_hbox = QHBoxLayout()
        self.scheduled_interval_spin = QSpinBox()
        self.scheduled_interval_spin.setRange(1, 1440)
        self.scheduled_interval_spin.setValue(settings_config.get("scheduled_image_interval_minutes", 60))
        self.scheduled_random_spin = QSpinBox()
        self.scheduled_random_spin.setRange(0, 720)
        self.scheduled_random_spin.setValue(settings_config.get("scheduled_image_random_window_minutes", 15))

        def update_scheduled_interval():
            self.sync_scheduled_image_settings_from_ui()
            self.reset_scheduled_image_next_send()

        self.scheduled_interval_spin.valueChanged.connect(update_scheduled_interval)
        self.scheduled_random_spin.valueChanged.connect(update_scheduled_interval)
        scheduled_interval_hbox.addWidget(QLabel("后续间隔(分):"))
        scheduled_interval_hbox.addWidget(self.scheduled_interval_spin)
        scheduled_interval_hbox.addStretch()
        scheduled_interval_layout.addRow(scheduled_interval_hbox)

        scheduled_random_hbox = QHBoxLayout()
        scheduled_random_hbox.addWidget(QLabel("间隔随机浮动(分):"))
        scheduled_random_hbox.addWidget(self.scheduled_random_spin)
        scheduled_random_hbox.addStretch()
        scheduled_interval_layout.addRow(scheduled_random_hbox)

        self.enable_scheduled_image_sender = QCheckBox("启用定时主动发送图片")
        self.enable_scheduled_image_sender.setChecked(settings_config.get("enable_scheduled_image_sender", False))

        def toggle_scheduled_image_sender(state):
            self.sync_scheduled_image_settings_from_ui()
            if state == Qt.Checked:
                self.start_scheduled_image_sender()
            else:
                self.stop_scheduled_image_sender()

        self.enable_scheduled_image_sender.stateChanged.connect(toggle_scheduled_image_sender)
        scheduled_interval_layout.addRow(self.enable_scheduled_image_sender)

        scheduled_interval_group.setLayout(scheduled_interval_layout)
        main_layout.addWidget(scheduled_interval_group)

        scheduled_guard_group = QGroupBox("会话保护")
        scheduled_guard_layout = QFormLayout()
        self.scheduled_pause_on_contact_checkbox = QCheckBox("对方新消息未回复时暂停定时发图")
        self.scheduled_pause_on_contact_checkbox.setChecked(
            settings_config.get("scheduled_image_pause_on_contact_message", True)
        )
        self.scheduled_resume_delay_spin = QSpinBox()
        self.scheduled_resume_delay_spin.setRange(0, 120)
        self.scheduled_resume_delay_spin.setValue(settings_config.get("scheduled_image_resume_delay_minutes", 5))
        self.scheduled_mark_replied_btn = QPushButton("我已回复，恢复排程")

        def update_scheduled_guard_options():
            self.sync_scheduled_image_settings_from_ui()
            if not self.scheduled_pause_on_contact_checkbox.isChecked():
                self.clear_scheduled_conversation_guard("会话保护已关闭", reset_schedule=False)
            self._update_scheduled_image_status()

        self.scheduled_pause_on_contact_checkbox.stateChanged.connect(lambda _state: update_scheduled_guard_options())
        self.scheduled_resume_delay_spin.valueChanged.connect(lambda _value: update_scheduled_guard_options())
        self.scheduled_mark_replied_btn.clicked.connect(
            lambda: self.mark_scheduled_conversation_replied(manual=True)
        )

        scheduled_guard_delay_hbox = QHBoxLayout()
        scheduled_guard_delay_hbox.addWidget(QLabel("回复后等待(分):"))
        scheduled_guard_delay_hbox.addWidget(self.scheduled_resume_delay_spin)
        scheduled_guard_delay_hbox.addStretch()

        scheduled_guard_layout.addRow(self.scheduled_pause_on_contact_checkbox)
        scheduled_guard_layout.addRow(scheduled_guard_delay_hbox)
        scheduled_guard_layout.addRow(self.scheduled_mark_replied_btn)
        scheduled_guard_group.setLayout(scheduled_guard_layout)
        main_layout.addWidget(scheduled_guard_group)

        calculation_group = QGroupBox("排程计算")
        calculation_layout = QVBoxLayout()
        self.scheduled_image_calculation_label = QLabel()
        self.scheduled_image_calculation_label.setObjectName("scheduleCalculation")
        self.scheduled_image_calculation_label.setWordWrap(True)
        self.scheduled_image_calculation_label.setTextInteractionFlags(Qt.TextSelectableByMouse)
        calculation_layout.addWidget(self.scheduled_image_calculation_label)
        calculation_group.setLayout(calculation_layout)
        main_layout.addWidget(calculation_group)

        refresh_scheduled_count()
        self._update_scheduled_image_calculation()
        main_layout.addStretch()

        return main_layout

    def init_status_panel(self):
        status_group = QWidget()
        status_group.setObjectName("sidePanel")
        layout = QVBoxLayout(status_group)
        layout.setContentsMargins(12, 12, 12, 12)
        layout.setSpacing(8)
        rows = [
            ("one_click", "一键监控", "待启动"),
            ("monitor", "监控状态", "未启动"),
            ("narrator", "讲述人", "未检查"),
            ("target", "目标窗口", "未检查"),
            ("last_message", "最后消息", "-"),
            ("last_trigger", "最近触发", "-"),
            ("delay", "延时发送", "空闲"),
            ("conversation_guard", "会话保护", "未触发"),
            ("timed_send", "定时发图", "未启用"),
            ("send_result", "发送结果", "-"),
        ]
        for key, label, value in rows:
            card = QWidget()
            card.setObjectName("statusCard")
            card_layout = QVBoxLayout(card)
            card_layout.setContentsMargins(12, 9, 12, 9)
            card_layout.setSpacing(4)
            title_label = QLabel(label)
            title_label.setObjectName("statusTitle")
            value_label = QLabel(value)
            value_label.setWordWrap(True)
            value_label.setObjectName("statusValue")
            self.status_labels[key] = value_label
            card_layout.addWidget(title_label)
            card_layout.addWidget(value_label)
            layout.addWidget(card)

        layout.addStretch()
        return status_group

    def init_monitor_log(self):
        vbox = QVBoxLayout()
        info_layout = QHBoxLayout()
        info_layout.addWidget(QLabel("监控运行日志"))
        export_btn = QPushButton("保存日志并清空面板")
        export_btn.setStyleSheet("padding: 2px 10px; font-size: 11px;")
        export_btn.clicked.connect(lambda: self.export_logs(manual=True))
        info_layout.addStretch()
        info_layout.addWidget(export_btn)
        self.log_view = QListWidget()
        self.log_view.setSelectionMode(QAbstractItemView.ExtendedSelection)
        vbox.addLayout(info_layout)
        vbox.addWidget(self.log_view)
        return vbox

    def export_logs(self, manual=False):
        if self.log_view.count() == 0:
            if manual:
                QMessageBox.information(self, "提示", "当前没有日志可以导出。")
            return
        self.logs_dir.mkdir(parents=True, exist_ok=True)
        timestamp = time.strftime("%Y%m%d_%H%M%S")
        filename = self.logs_dir / f"momo_log_{timestamp}.txt"
        try:
            with filename.open("w", encoding="utf-8") as handle:
                for i in range(self.log_view.count()):
                    handle.write(self.log_view.item(i).text() + "\n")
            self.log_view.clear()
            msg = f"已将面板日志保存至 {filename}"
            self.add_log(msg)
            if manual:
                QMessageBox.information(self, "清理成功", msg)
        except Exception as exc:
            if manual:
                QMessageBox.warning(self, "导出失败", str(exc))

    def add_log(self, message):
        self.add_log_signal.emit(str(message))

    def _do_add_log(self, message):
        current_time = time.strftime("%H:%M:%S")
        line = f"[{current_time}] {message}"
        self.log_view.addItem(line)
        self.log_view.scrollToBottom()
        self.write_persistent_log(line)
        if self.log_view.count() > 300:
            self.export_logs(manual=False)

    def write_persistent_log(self, line):
        try:
            self.logs_dir.mkdir(parents=True, exist_ok=True)
            self.rotate_app_log_if_needed()
            with self.app_log_path.open("a", encoding="utf-8") as handle:
                handle.write(f"{datetime.datetime.now():%Y-%m-%d} {line}\n")
        except Exception:
            pass

    def rotate_app_log_if_needed(self):
        if self.app_log_path.exists() and self.app_log_path.stat().st_size > 2 * 1024 * 1024:
            archive = self.logs_dir / f"app_{time.strftime('%Y%m%d_%H%M%S')}.log"
            self.app_log_path.replace(archive)

        cutoff = time.time() - 7 * 24 * 60 * 60
        for path in self.logs_dir.glob("app_*.log"):
            try:
                if path.stat().st_mtime < cutoff:
                    path.unlink()
            except Exception:
                pass

    def set_status(self, key, value):
        self.status_values[key] = str(value)
        self.status_signal.emit(key, str(value))

    def _do_update_status(self, key, value):
        label = self.status_labels.get(key)
        if label is not None:
            label.setText(value)

    def start_status_timer(self):
        self.status_timer = QTimer(self)
        self.status_timer.timeout.connect(self.refresh_runtime_status)
        self.status_timer.start(3000)
        self.refresh_runtime_status()

    def refresh_runtime_status(self):
        self.handle_activation_request()
        self.set_status("monitor", "运行中" if self.monitoring else "未启动")
        self.set_status("narrator", self.get_narrator_status_text())
        target_name = self.config.get("settings", {}).get("trigger_sender", "").strip()
        result = self.wechat.check_target_window(target_name)
        self.set_status("target", result.message)
        self._update_conversation_guard_status()
        self.update_scheduled_calculation_signal.emit()

    def handle_activation_request(self):
        if not self.activation_request_path.exists():
            return
        try:
            self.activation_request_path.unlink()
        except OSError:
            return
        self.bring_main_window_to_front()

    def run_startup_check(self):
        self.add_log("开始启动自检")
        check_issues = []
        for note in self.config_load_notes:
            self.add_log(note)
            check_issues.append(note)

        dependency_checks = [
            ("PyQt5", "PyQt5"),
            ("uiautomation", "uiautomation"),
            ("pyperclip", "pyperclip"),
            ("pywin32/win32clipboard", "win32clipboard"),
        ]
        for name, module_name in dependency_checks:
            try:
                __import__(module_name)
                self.add_log(f"依赖正常: {name}")
            except Exception as exc:
                message = f"依赖异常: {name} - {exc}"
                self.add_log(message)
                check_issues.append(message)

        self.add_log(f"Windows 讲述人状态: {self.get_narrator_status_text()}")
        self.add_log("请确认 Windows 讲述人模式已开启，否则微信控件可能无法识别")
        target_name = self.config.get("settings", {}).get("trigger_sender", "").strip()
        target_result = self.wechat.check_target_window(target_name)
        self.set_status("target", target_result.message)
        self.add_log(target_result.message)
        if getattr(target_result, "success", True) is False:
            check_issues.append(target_result.message)

        active_count = self.config.get("settings", {}).get("active_rules_count", 1)
        for index in range(active_count):
            rule = self.config["rules"][index]
            keywords = [item.strip() for item in rule.get("keywords", "").split(",") if item.strip()]
            if not keywords:
                message = f"规则 {index + 1} 未配置关键词"
                self.add_log(message)
                check_issues.append(message)
            if rule.get("reply_type") == "image":
                folder = rule.get("folder", "").strip()
                if not folder:
                    message = f"规则 {index + 1} 未配置图片目录"
                    self.add_log(message)
                    check_issues.append(message)
                elif self.get_folder_access_message(folder):
                    message = f"规则 {index + 1} {self.get_folder_access_message(folder)}"
                    self.add_log(message)
                    check_issues.append(message)
                else:
                    self.add_log(f"规则 {index + 1} 图片数量: {len(self.get_valid_images(folder))}")
            else:
                if not rule.get("reply_text", "").strip():
                    message = f"规则 {index + 1} 未配置回复文本"
                    self.add_log(message)
                    check_issues.append(message)

        self.last_startup_check_summary = "正常" if not check_issues else "；".join(check_issues[:3])
        self.add_log("启动自检完成")

    def _do_update_img_count(self, rule_idx):
        if rule_idx in self.rule_img_count_labels:
            folder = self.config["rules"][rule_idx].get("folder", "")
            self.rule_img_count_labels[rule_idx].setText(self.image_count_text(folder))

    def _do_update_scheduled_img_count(self):
        if self.scheduled_img_count_label is not None:
            folder = self.config.get("settings", {}).get("scheduled_image_folder", "")
            self.scheduled_img_count_label.setText(self.image_count_text(folder))

    def _try_activate_trigger(self):
        with self.trigger_state_lock:
            if self.last_triggered:
                return None
            self.trigger_token += 1
            self.last_triggered = True
            return self.trigger_token

    def _clear_trigger_if_active(self):
        with self.trigger_state_lock:
            if not self.last_triggered:
                return False
            self.trigger_token += 1
            self.last_triggered = False
            return True

    def _invalidate_trigger(self):
        with self.trigger_state_lock:
            self.trigger_token += 1
            self.last_triggered = False

    def _is_trigger_active(self, trigger_token):
        with self.trigger_state_lock:
            return self.last_triggered and self.trigger_token == trigger_token

    def _finish_trigger(self, trigger_token):
        with self.trigger_state_lock:
            if self.trigger_token == trigger_token:
                self.last_triggered = False

    def _find_matching_rule_index(self, clean_text, active_count):
        for i in range(active_count):
            rule = self.config["rules"][i]
            keywords = [k.strip() for k in rule.get("keywords", "").split(",") if k.strip()]
            mode = rule.get("mode", "exact")

            if mode == "exact":
                if clean_text in keywords:
                    return i
            elif any(keyword in clean_text for keyword in keywords):
                return i

        return -1

    def _get_delay_seconds(self, settings_config):
        base_delay = settings_config.get("send_delay", 0)
        random_range = settings_config.get("random_delay", 0)

        if base_delay <= 0 and random_range <= 0:
            return 0, 0

        if random_range > 0:
            half_range = random_range / 2
            actual_delay = max(0, random.uniform(base_delay - half_range, base_delay + half_range))
        else:
            actual_delay = max(0, base_delay)

        return int(actual_delay * 60), actual_delay

    def _message_direction_label(self, direction):
        return {
            "contact": "对方",
            "assumed_contact": "未知/按对方",
            "self": "我",
            "unknown": "未知",
        }.get(direction or "unknown", "未知")

    def _mark_expected_self_message(self, seconds=12):
        with self.conversation_guard_lock:
            self.expected_self_message_until = datetime.datetime.now() + datetime.timedelta(seconds=seconds)

    def _consume_expected_self_message(self):
        now = datetime.datetime.now()
        with self.conversation_guard_lock:
            expected_until = self.expected_self_message_until
            if expected_until is None:
                return False
            if now <= expected_until:
                self.expected_self_message_until = None
                return True
            self.expected_self_message_until = None
        return False

    def _clear_expected_self_message(self):
        with self.conversation_guard_lock:
            self.expected_self_message_until = None

    def _resolve_message_direction(self, clean_text, direction):
        if direction == "self":
            self._clear_expected_self_message()
            return direction
        if direction == "contact":
            return direction
        if not clean_text:
            return "unknown"
        if self._consume_expected_self_message():
            return "self"
        if self._scheduled_conversation_guard_enabled():
            return "assumed_contact"
        return "unknown"

    def _scheduled_conversation_guard_enabled(self):
        return bool(
            self.config.get("settings", {}).get("scheduled_image_pause_on_contact_message", True)
        )

    def _get_scheduled_resume_delay_minutes(self):
        try:
            value = int(self.config.get("settings", {}).get("scheduled_image_resume_delay_minutes", 5))
        except (TypeError, ValueError):
            value = 5
        return max(0, min(120, value))

    def _schedule_after_conversation_reply(self, now, resume_after):
        if not self.config.get("settings", {}).get("enable_scheduled_image_sender", False):
            return

        schedule_at = resume_after or now
        if self.scheduled_image_next_at is not None and self.scheduled_image_next_at > schedule_at:
            self._update_scheduled_image_status()
            return

        start, end = self._get_scheduled_image_window(now)
        active_start = self._get_scheduled_image_active_start(start)
        if schedule_at < active_start:
            schedule_at = active_start
        if schedule_at >= end:
            self.pause_scheduled_image_sender_signal.emit(
                "回复后已超过今天可安排时间，定时主动发图已暂停；明天需要手动启用"
            )
            return

        self.scheduled_image_schedule_mode = "conversation_resume"
        self.scheduled_image_schedule_base_at = now
        self.scheduled_image_next_at = schedule_at
        self._update_scheduled_image_status()
        self.add_log(f"会话已回复，定时发图恢复到: {schedule_at:%Y-%m-%d %H:%M}")

    def _get_scheduled_conversation_block_reason(self, now=None):
        if not self._scheduled_conversation_guard_enabled():
            return ""

        now = now or datetime.datetime.now()
        with self.conversation_guard_lock:
            if self.scheduled_image_waiting_for_reply:
                return "等待人工回复"
            if self.scheduled_image_resume_after is not None:
                if now < self.scheduled_image_resume_after:
                    return f"回复后等待到 {self.scheduled_image_resume_after:%H:%M}"
                self.scheduled_image_resume_after = None
        return ""

    def _update_conversation_guard_status(self):
        if not self._scheduled_conversation_guard_enabled():
            self.set_status("conversation_guard", "未启用")
            return

        reason = self._get_scheduled_conversation_block_reason()
        if reason:
            self.set_status("conversation_guard", reason)
        else:
            self.set_status("conversation_guard", "未触发")

    def _mark_scheduled_contact_message_pending(self, clean_text):
        if not self._scheduled_conversation_guard_enabled():
            return

        now = datetime.datetime.now()
        with self.conversation_guard_lock:
            was_waiting = self.scheduled_image_waiting_for_reply
            previous_text = self.scheduled_image_waiting_text
            self.scheduled_image_waiting_for_reply = True
            self.scheduled_image_waiting_text = clean_text
            self.scheduled_image_waiting_since = now
            self.scheduled_image_resume_after = None

        if not was_waiting or previous_text != clean_text:
            self.add_log(f"检测到对方新消息，定时发图进入等待人工回复: {clean_text}")
        self._update_scheduled_image_status()

    def mark_scheduled_conversation_replied(self, manual=False):
        if not self._scheduled_conversation_guard_enabled():
            if manual:
                self.add_log("会话保护未启用，无需恢复")
            self._update_conversation_guard_status()
            return

        now = datetime.datetime.now()
        delay_minutes = self._get_scheduled_resume_delay_minutes()
        resume_after = now + datetime.timedelta(minutes=delay_minutes) if delay_minutes > 0 else None
        with self.conversation_guard_lock:
            had_state = (
                self.scheduled_image_waiting_for_reply
                or self.scheduled_image_resume_after is not None
                or self.scheduled_image_deferred_due_at is not None
            )
            self.scheduled_image_waiting_for_reply = False
            self.scheduled_image_waiting_text = ""
            self.scheduled_image_waiting_since = None
            self.scheduled_image_resume_after = resume_after if had_state else None
            self.scheduled_image_deferred_due_at = None

        if had_state:
            if delay_minutes > 0:
                self.add_log(f"检测到已回复，定时发图将在 {delay_minutes} 分钟后恢复")
            else:
                self.add_log("检测到已回复，定时发图立即恢复")
            self._schedule_after_conversation_reply(now, resume_after)
        elif manual:
            self.add_log("当前没有待回复消息，已刷新会话保护状态")

        self._update_conversation_guard_status()

    def clear_scheduled_conversation_guard(self, reason="", reset_schedule=True):
        with self.conversation_guard_lock:
            had_state = (
                self.scheduled_image_waiting_for_reply
                or self.scheduled_image_resume_after is not None
                or self.scheduled_image_deferred_due_at is not None
            )
            self.scheduled_image_waiting_for_reply = False
            self.scheduled_image_waiting_text = ""
            self.scheduled_image_waiting_since = None
            self.scheduled_image_resume_after = None
            self.scheduled_image_deferred_due_at = None
            self.expected_self_message_until = None

        if had_state and reason:
            self.add_log(reason)
        if reset_schedule and self.config.get("settings", {}).get("enable_scheduled_image_sender", False):
            self.reset_scheduled_image_next_send()
        else:
            self._update_scheduled_image_status()
        self._update_conversation_guard_status()

    def _handle_scheduled_conversation_message(self, clean_text, direction):
        if direction in ("contact", "assumed_contact"):
            self._mark_scheduled_contact_message_pending(clean_text)
        elif direction == "self":
            self.mark_scheduled_conversation_replied(manual=False)

    def on_last_message_change(self, last_text, _current_time, message_direction="unknown"):
        settings_config = self.config.get("settings", {})
        trigger_sender = settings_config.get("trigger_sender", "momo")
        active_count = settings_config.get("active_rules_count", 1)

        clean_text = str(last_text).strip()
        direction = self._resolve_message_direction(clean_text, message_direction or "unknown")
        direction_label = self._message_direction_label(direction)
        self.set_status("last_message", f"{direction_label}: {clean_text}" if clean_text else "-")
        self._handle_scheduled_conversation_message(clean_text, direction)
        matched_rule_idx = -1 if direction == "self" else self._find_matching_rule_index(clean_text, active_count)

        if matched_rule_idx != -1:
            trigger_token = self._try_activate_trigger()
            if trigger_token is None:
                return

            trigger_text = f"规则 {matched_rule_idx + 1}: {last_text}"
            self.set_status("last_trigger", trigger_text)
            self.add_log(f"警报触发: 命中规则 {matched_rule_idx + 1}, 内容: '{last_text}'")

            delay_seconds, actual_delay = self._get_delay_seconds(settings_config)

            if delay_seconds > 0:
                self.add_log(f"延迟 {actual_delay:.1f} 分钟后开始发送")
                threading.Thread(
                    target=self._delayed_send_action,
                    args=(delay_seconds, trigger_sender, matched_rule_idx, trigger_token),
                    daemon=True,
                ).start()
            else:
                self._do_send_action(trigger_sender, matched_rule_idx, trigger_token)

        elif self._clear_trigger_if_active():
            self.set_status("last_trigger", "已解除")
            self.set_status("delay", "空闲")
            self.add_log(f"警报解除: 最后一条消息变成了: '{last_text}'")

    def _format_seconds(self, seconds):
        minutes, sec = divmod(max(0, int(seconds)), 60)
        return f"{minutes:02d}:{sec:02d}"

    def _delayed_send_action(self, delay_seconds, trigger_sender, rule_idx, trigger_token):
        _uia_init = auto.UIAutomationInitializerInThread()
        try:
            wait_time = delay_seconds
            while wait_time > 0:
                if not self.monitoring:
                    self.set_status("delay", "已取消")
                    self.add_log("监控已停止，取消延时发送")
                    self._finish_trigger(trigger_token)
                    return
                if not self._is_trigger_active(trigger_token):
                    self.set_status("delay", "已取消")
                    self.add_log("警报已解除，取消本次延时发送")
                    return
                self.set_status("delay", f"倒计时 {self._format_seconds(wait_time)}")
                time.sleep(1)
                wait_time -= 1

            if self.monitoring and self._is_trigger_active(trigger_token):
                self.set_status("delay", "发送中")
                self._do_send_action(trigger_sender, rule_idx, trigger_token)
            elif self.monitoring:
                self.set_status("delay", "已取消")
                self.add_log("警报已解除，取消本次延时发送")
        finally:
            _uia_init = None

    def _move_to_sent_folder(self, image_path):
        source = Path(image_path)
        sent_dir = source.parent / "sent"
        sent_dir.mkdir(parents=True, exist_ok=True)
        destination = sent_dir / source.name
        if destination.exists():
            timestamp = time.strftime("%Y%m%d_%H%M%S")
            destination = sent_dir / f"{source.stem}_{timestamp}{source.suffix}"
            counter = 1
            while destination.exists():
                destination = sent_dir / f"{source.stem}_{timestamp}_{counter}{source.suffix}"
                counter += 1
        shutil.move(str(source), str(destination))
        return destination

    def _do_send_action(self, trigger_sender, rule_idx, trigger_token):
        if not self._is_trigger_active(trigger_token):
            self.add_log("当前触发已失效，跳过发送")
            return

        rule = self.config["rules"][rule_idx]
        reply_type = rule.get("reply_type", "image")
        self.add_log("开始执行发送动作")

        try:
            if reply_type == "image":
                material_folder = rule.get("folder", "")
                folder_problem = self.get_folder_access_message(material_folder)
                if folder_problem:
                    self.set_status("send_result", "失败: 图片目录不可访问")
                    self.add_log(folder_problem)
                    return

                images = self.get_valid_images(material_folder)

                if len(images) == 0:
                    self.set_status("send_result", "失败: 没有可发送图片")
                    self.add_log("指定素材文件夹中没有图片可发")
                    return

                selected_image = random.choice(images)
                self.add_log(f"已抽取图片: {os.path.basename(selected_image)}")

                self._mark_expected_self_message()
                result = self.wechat.send_file(trigger_sender, selected_image)
                if not result:
                    self._clear_expected_self_message()
                    self.set_status("send_result", f"失败: {result.message}")
                    self.add_log(f"图片发送失败: {result.message}")
                    return

                moved_path = self._move_to_sent_folder(selected_image)
                self.set_status("send_result", result.message)
                self.add_log(f"图片发送成功: {result.message}")
                self.add_log(f"已移动到已发送目录: {moved_path}")
                self.update_img_count_signal.emit(rule_idx)

            elif reply_type == "text":
                reply_text = rule.get("reply_text", "")
                if not reply_text:
                    self.set_status("send_result", "失败: 未配置回复文本")
                    self.add_log("规则未配置回复文本，无法发送")
                    return

                self.add_log("准备发送文本")
                self._mark_expected_self_message()
                result = self.wechat.send_msg(trigger_sender, text=reply_text)
                if not result:
                    self._clear_expected_self_message()
                    self.set_status("send_result", f"失败: {result.message}")
                    self.add_log(f"文本发送失败: {result.message}")
                    return

                self.set_status("send_result", result.message)
                self.add_log(f"文本发送成功: {result.message}")

        except Exception as exc:
            self.set_status("send_result", f"异常: {exc}")
            self.add_log(f"发送异常: {exc}")
        finally:
            self.set_status("delay", "空闲")
            self._finish_trigger(trigger_token)

    def open_target_chat_window(self):
        if self.open_target_window_active:
            return

        target_name = self.config.get("settings", {}).get("trigger_sender", "").strip()
        if not target_name:
            QMessageBox.warning(self, "目标联系人为空", "请先填写“目标对话/触发者昵称”。")
            return

        self.open_target_window_active = True
        if self.open_target_window_btn is not None:
            self.open_target_window_btn.setEnabled(False)
            self.open_target_window_btn.setText("正在查找联系人...")
        self.set_status("target", f"正在查找联系人: {target_name}")
        self.add_log(f"开始查找联系人并打开独立窗口: {target_name}")
        threading.Thread(
            target=self._open_target_chat_window_worker,
            args=(target_name,),
            daemon=True,
        ).start()

    def _open_target_chat_window_worker(self, target_name):
        _uia_init = auto.UIAutomationInitializerInThread()
        try:
            result = self.wechat.open_independent_chat_window(target_name)
            self.target_window_action_finished_signal.emit(bool(result), result.message)
        except Exception as exc:
            self.target_window_action_finished_signal.emit(False, f"打开目标聊天窗口异常: {exc}")
        finally:
            _uia_init = None

    def _finish_open_target_window(self, success, message):
        self.open_target_window_active = False
        if self.open_target_window_btn is not None:
            self.open_target_window_btn.setEnabled(True)
            self.open_target_window_btn.setText("打开并置顶目标窗口")
        self.set_status("target", message)
        self.add_log(message)
        if not success:
            QMessageBox.warning(self, "打开目标窗口失败", message)

    def start_monitoring(self, silent=False):
        if self.monitoring:
            if not silent:
                QMessageBox.information(self, "提示", "监控已经在运行中")
            return True

        start_time = time.strftime("%Y-%m-%d %H:%M:%S")
        self.add_log(f"[{start_time}] 启动精准多重规则监控")

        self._invalidate_trigger()
        trigger_sender = self.config.get("settings", {}).get("trigger_sender", "momo")
        result = self.wechat.start_last_message_monitor(
            target_name=trigger_sender,
            callback=self.on_last_message_change,
            check_interval=1,
            status_callback=lambda message: self.status_signal.emit("target", message),
        )

        if not result:
            self.add_log(result.message)
            return False

        self.monitoring = True
        self.set_status("monitor", "运行中")
        self.start_btn.setEnabled(False)
        self.stop_btn.setEnabled(True)
        return True

    def stop_monitoring(self):
        if self.monitoring:
            result = self.wechat.stop_last_message_monitor()
            self._invalidate_trigger()
            self.add_log(result.message)
            self.add_log("消息监控已手动停止")
            self.monitoring = False
            self.set_status("monitor", "未启动")
            self.set_status("delay", "空闲")
            self.start_btn.setEnabled(True)
            self.stop_btn.setEnabled(False)

    def start_auto_timer_check(self):
        if self.auto_timer is None:
            self.auto_timer = QTimer(self)
            self.auto_timer.timeout.connect(self.auto_check_time)
            self.auto_timer.start(60000)
            self.add_log("自动定时启停检查已就绪")

    def stop_auto_timer_check(self):
        if self.auto_timer is not None:
            self.auto_timer.stop()
            self.auto_timer = None
            self.add_log("自动定时启停已被关闭")

    def auto_check_time(self):
        now = datetime.datetime.now()
        settings_config = self.config.get("settings", {})
        start_h = settings_config.get("auto_start_hour", 10)
        start_m = settings_config.get("auto_start_minute", 0)
        end_h = settings_config.get("auto_end_hour", 12)
        end_m = settings_config.get("auto_end_minute", 0)

        current_total = now.hour * 60 + now.minute
        start_total = start_h * 60 + start_m
        end_total = end_h * 60 + end_m
        should_be_monitoring = start_total <= current_total < end_total

        if should_be_monitoring and not self.monitoring:
            self.add_log("到达设定区间，自动启动监控")
            self.start_monitoring()
        elif not should_be_monitoring and self.monitoring:
            self.add_log("离开设定区间，自动停止监控")
            self.stop_monitoring()

    def start_scheduled_image_sender(self):
        self.scheduled_image_pause_reason = ""
        if self.scheduled_image_timer is None:
            self.scheduled_image_timer = QTimer(self)
            self.scheduled_image_timer.timeout.connect(self.check_scheduled_image_send)
            self.scheduled_image_timer.start(30000)
            self.add_log("定时主动发图检查已就绪")
        self.reset_scheduled_image_next_send()
        QTimer.singleShot(1000, self.check_scheduled_image_send)

    def stop_scheduled_image_sender(self):
        if self.scheduled_image_timer is not None:
            self.scheduled_image_timer.stop()
            self.scheduled_image_timer = None
        self.scheduled_image_next_at = None
        self.scheduled_image_schedule_mode = None
        self.scheduled_image_schedule_base_at = None
        self.scheduled_image_pause_reason = ""
        with self.conversation_guard_lock:
            self.scheduled_image_waiting_for_reply = False
            self.scheduled_image_waiting_text = ""
            self.scheduled_image_waiting_since = None
            self.scheduled_image_resume_after = None
            self.scheduled_image_deferred_due_at = None
            self.expected_self_message_until = None
        self.set_status("timed_send", "未启用")
        self._update_conversation_guard_status()
        self.update_scheduled_calculation_signal.emit()
        self.add_log("定时主动发图已关闭")

    def _pause_scheduled_image_sender(self, reason):
        was_enabled = self.config.get("settings", {}).get("enable_scheduled_image_sender", False)
        self.config.setdefault("settings", {})["enable_scheduled_image_sender"] = False
        self.save_config()
        if self.scheduled_image_timer is not None:
            self.scheduled_image_timer.stop()
            self.scheduled_image_timer = None
        self.scheduled_image_next_at = None
        self.scheduled_image_schedule_mode = None
        self.scheduled_image_schedule_base_at = None
        self.scheduled_image_pause_reason = reason
        with self.conversation_guard_lock:
            self.scheduled_image_waiting_for_reply = False
            self.scheduled_image_waiting_text = ""
            self.scheduled_image_waiting_since = None
            self.scheduled_image_resume_after = None
            self.scheduled_image_deferred_due_at = None
            self.expected_self_message_until = None
        if hasattr(self, "enable_scheduled_image_sender"):
            was_blocked = self.enable_scheduled_image_sender.blockSignals(True)
            self.enable_scheduled_image_sender.setChecked(False)
            self.enable_scheduled_image_sender.blockSignals(was_blocked)
        self.set_status("timed_send", "已暂停")
        self._update_conversation_guard_status()
        self.update_scheduled_calculation_signal.emit()
        if was_enabled:
            self.add_log(reason)

    def reset_scheduled_image_next_send(self):
        if not self.config.get("settings", {}).get("enable_scheduled_image_sender", False):
            return
        now = datetime.datetime.now()
        self.scheduled_image_schedule_mode = "initial"
        self.scheduled_image_schedule_base_at = now
        self.scheduled_image_next_at = self._get_initial_scheduled_image_due_time(now)
        if self.scheduled_image_next_at is None:
            self._pause_scheduled_image_sender("今天已超过定时发图可安排时间，已暂停；明天需要手动启用")
            return
        self._update_scheduled_image_status()

    def _get_scheduled_image_window(self, now):
        settings_config = self.config.get("settings", {})
        start = now.replace(
            hour=settings_config.get("scheduled_image_start_hour", 9),
            minute=settings_config.get("scheduled_image_start_minute", 0),
            second=0,
            microsecond=0,
        )
        end = now.replace(
            hour=settings_config.get("scheduled_image_end_hour", 23),
            minute=settings_config.get("scheduled_image_end_minute", 0),
            second=0,
            microsecond=0,
        )
        if end <= start:
            end += datetime.timedelta(days=1)

        previous_start = start - datetime.timedelta(days=1)
        previous_end = end - datetime.timedelta(days=1)
        if previous_start <= now < previous_end:
            return previous_start, previous_end
        return start, end

    def _get_next_scheduled_image_window_start(self, now):
        start, _end = self._get_scheduled_image_window(now)
        if now < start:
            return start
        return start + datetime.timedelta(days=1)

    def _get_scheduled_image_active_start(self, start):
        return start

    def _get_scheduled_first_delay_minutes(self):
        try:
            value = int(self.config.get("settings", {}).get("scheduled_image_first_delay_minutes", 30))
        except (TypeError, ValueError):
            value = 30
        return max(0, min(1440, value))

    def _get_initial_scheduled_image_due_time(self, now):
        start, end = self._get_scheduled_image_window(now)
        first_due = start + datetime.timedelta(minutes=self._get_scheduled_first_delay_minutes())
        latest_due = end - datetime.timedelta(seconds=5)

        if first_due > latest_due:
            return None
        if now <= first_due:
            return first_due
        if now < end:
            return now
        return None

    def _get_next_scheduled_image_due_time(self, now):
        settings_config = self.config.get("settings", {})
        interval_minutes = max(1, int(settings_config.get("scheduled_image_interval_minutes", 60)))
        random_window = max(0, int(settings_config.get("scheduled_image_random_window_minutes", 0)))
        offset_minutes = interval_minutes
        if random_window > 0:
            offset_minutes += random.randint(-random_window, random_window)
        offset_minutes = max(1, offset_minutes)

        due_time = now + datetime.timedelta(minutes=offset_minutes)
        _start, end = self._get_scheduled_image_window(now)
        if due_time >= end:
            return None
        return due_time

    def _build_scheduled_image_calculation_text(self, now=None):
        now = now or datetime.datetime.now()
        settings = self.config.get("settings", {})
        start, end = self._get_scheduled_image_window(now)
        first_delay = self._get_scheduled_first_delay_minutes()
        first_due = start + datetime.timedelta(minutes=first_delay)
        interval = max(1, int(settings.get("scheduled_image_interval_minutes", 60)))
        interval_random = max(0, int(settings.get("scheduled_image_random_window_minutes", 0)))
        minimum_interval = max(1, interval - interval_random)
        maximum_interval = max(1, interval + interval_random)

        lines = [
            f"当前时间：{now:%m-%d %H:%M}",
            f"发送窗口：{start:%m-%d %H:%M} 至 {end:%m-%d %H:%M}",
            f"首次规则：窗口开始后 {first_delay} 分钟，固定为 {first_due:%m-%d %H:%M}",
            f"后续规则：每次成功发送后 {interval} ± {interval_random} 分钟，"
            f"即 {minimum_interval} 至 {maximum_interval} 分钟",
        ]

        pause_reason = getattr(self, "scheduled_image_pause_reason", "")
        enabled = settings.get("enable_scheduled_image_sender", False)
        next_at = getattr(self, "scheduled_image_next_at", None)
        mode = getattr(self, "scheduled_image_schedule_mode", None)
        base_at = getattr(self, "scheduled_image_schedule_base_at", None)
        guard_reason = self._get_scheduled_conversation_block_reason(now)
        with self.conversation_guard_lock:
            deferred_due_at = self.scheduled_image_deferred_due_at

        if pause_reason:
            lines.append(f"计算结果：无法在停止时间前完成等待，已暂停")
            lines.append(f"原因：{pause_reason}")
        elif not enabled:
            lines.append("当前状态：未启用，勾选后按上方规则计算")
        elif guard_reason:
            lines.append(f"会话保护：{guard_reason}")
            if deferred_due_at is not None:
                lines.append(f"原计划发送：{deferred_due_at:%m-%d %H:%M}，已延后到回复后恢复")
        elif next_at is None:
            lines.append("当前状态：正在发送或等待重新计算")
        elif mode == "interval" and base_at is not None:
            actual_minutes = max(1, round((next_at - base_at).total_seconds() / 60))
            lines.extend(
                [
                    "当前模式：后续循环",
                    f"间隔范围：{interval} ± {interval_random} 分钟 = "
                    f"{minimum_interval} 至 {maximum_interval} 分钟",
                    f"本次随机结果：{actual_minutes} 分钟",
                    f"计算过程：{base_at:%m-%d %H:%M} + {actual_minutes} 分钟 "
                    f"= {next_at:%m-%d %H:%M}",
                    f"最终结果：下一次发送 {next_at:%m-%d %H:%M}",
                ]
            )
        elif mode == "conversation_resume" and base_at is not None:
            wait_minutes = max(0, round((next_at - base_at).total_seconds() / 60))
            lines.extend(
                [
                    f"会话恢复：{base_at:%m-%d %H:%M} 已检测到回复",
                    f"回复后等待：{wait_minutes} 分钟",
                    f"最终结果：下一次发送 {next_at:%m-%d %H:%M}",
                ]
            )
        else:
            base_at = base_at or now
            next_min_at = next_at + datetime.timedelta(minutes=minimum_interval)
            next_max_at = next_at + datetime.timedelta(minutes=maximum_interval)
            lines.extend(
                [
                    "当前模式：首次发送",
                    f"首次计算：{start:%m-%d %H:%M} + {first_delay} 分钟 "
                    f"= {first_due:%m-%d %H:%M}",
                ]
            )
            if first_due < base_at:
                lines.append("首次时间已过：本次启用后立即安排一次")
            lines.append(
                f"下轮预览：首次成功后约 {next_min_at:%m-%d %H:%M} 至 "
                f"{min(next_max_at, end):%m-%d %H:%M}"
            )
            lines.append(f"最终结果：下一次发送 {next_at:%m-%d %H:%M}")

        lines.append(f"停止规则：到 {end:%m-%d %H:%M} 自动暂停，次日需手动启用")
        return "\n".join(lines)

    def _update_scheduled_image_calculation(self):
        label = getattr(self, "scheduled_image_calculation_label", None)
        if label is not None:
            label.setText(self._build_scheduled_image_calculation_text())

    def _update_scheduled_image_status(self):
        if not self.config.get("settings", {}).get("enable_scheduled_image_sender", False):
            self.set_status("timed_send", "未启用")
            self._update_conversation_guard_status()
            self.update_scheduled_calculation_signal.emit()
            return
        guard_reason = self._get_scheduled_conversation_block_reason()
        if guard_reason:
            self.set_status("timed_send", guard_reason)
            self._update_conversation_guard_status()
            self.update_scheduled_calculation_signal.emit()
            return
        if self.scheduled_image_sending:
            self.set_status("timed_send", "发送中")
            self._update_conversation_guard_status()
            self.update_scheduled_calculation_signal.emit()
            return
        if self.scheduled_image_next_at:
            self.set_status("timed_send", f"下次 {self.scheduled_image_next_at:%m-%d %H:%M}")
        else:
            self.set_status("timed_send", "待安排")
        self._update_conversation_guard_status()
        self.update_scheduled_calculation_signal.emit()

    def check_scheduled_image_send(self):
        settings_config = self.config.get("settings", {})
        if not settings_config.get("enable_scheduled_image_sender", False):
            return
        if self.scheduled_image_sending:
            self._update_scheduled_image_status()
            return

        now = datetime.datetime.now()
        start, end = self._get_scheduled_image_window(now)
        active_start = self._get_scheduled_image_active_start(start)
        if now >= end:
            self._pause_scheduled_image_sender("已超过停止发送时间，定时主动发图已暂停；明天需要手动启用")
            return
        if not (active_start <= now < end):
            if self.scheduled_image_next_at is None or self.scheduled_image_next_at <= now:
                self.scheduled_image_schedule_mode = "initial"
                self.scheduled_image_schedule_base_at = now
                self.scheduled_image_next_at = self._get_initial_scheduled_image_due_time(now)
                if self.scheduled_image_next_at is None:
                    self._pause_scheduled_image_sender("今天已超过定时发图可安排时间，已暂停；明天需要手动启用")
                    return
            self._update_scheduled_image_status()
            return

        guard_reason = self._get_scheduled_conversation_block_reason(now)
        if guard_reason:
            if self.scheduled_image_next_at is not None and now >= self.scheduled_image_next_at:
                due_time = self.scheduled_image_next_at
                self.scheduled_image_next_at = None
                with self.conversation_guard_lock:
                    self.scheduled_image_deferred_due_at = due_time
                self.add_log(f"定时发图到点，但{guard_reason}，已等待你回复后恢复")
            self._update_scheduled_image_status()
            return

        if self.scheduled_image_next_at is None:
            self.scheduled_image_schedule_mode = "initial"
            self.scheduled_image_schedule_base_at = now
            self.scheduled_image_next_at = self._get_initial_scheduled_image_due_time(now)
            if self.scheduled_image_next_at is None:
                self._pause_scheduled_image_sender("今天已超过定时发图可安排时间，已暂停；明天需要手动启用")
                return
            self._update_scheduled_image_status()
            return

        guard_reason = self._get_scheduled_conversation_block_reason(now)
        if guard_reason:
            if now >= self.scheduled_image_next_at:
                due_time = self.scheduled_image_next_at
                self.scheduled_image_next_at = None
                with self.conversation_guard_lock:
                    self.scheduled_image_deferred_due_at = due_time
                self.add_log(f"定时发图到点，但{guard_reason}，已等待你回复后恢复")
            self._update_scheduled_image_status()
            return

        if now < self.scheduled_image_next_at:
            self._update_scheduled_image_status()
            return

        due_time = self.scheduled_image_next_at
        self.scheduled_image_next_at = None
        self.scheduled_image_sending = True
        self.set_status("timed_send", "发送中")
        self.add_log(f"到达定时发图时间: {due_time:%Y-%m-%d %H:%M}")
        threading.Thread(target=self._scheduled_image_send_worker, daemon=True).start()

    def _scheduled_image_send_worker(self):
        _uia_init = auto.UIAutomationInitializerInThread()
        try:
            settings_config = self.config.get("settings", {})
            trigger_sender = settings_config.get("trigger_sender", "momo").strip()
            material_folder = settings_config.get("scheduled_image_folder", "").strip()
            if not trigger_sender:
                self.set_status("send_result", "失败: 未配置目标窗口")
                self.add_log("定时发图失败: 未配置目标对话/触发者昵称")
                return

            folder_problem = self.get_folder_access_message(material_folder)
            if folder_problem:
                self.set_status("send_result", "失败: 定时图片目录不可访问")
                self.add_log(f"定时发图失败: {folder_problem}")
                return

            images = self.get_valid_images(material_folder)
            if len(images) == 0:
                self.set_status("send_result", "失败: 没有可发送图片")
                self.add_log("定时发图失败: 指定素材文件夹中没有图片可发")
                return

            selected_image = random.choice(images)
            self.add_log(f"定时发图已抽取图片: {os.path.basename(selected_image)}")
            self._mark_expected_self_message()
            result = self.wechat.send_file(trigger_sender, selected_image)
            if not result:
                self._clear_expected_self_message()
                self.set_status("send_result", f"失败: {result.message}")
                self.add_log(f"定时发图失败: {result.message}")
                return

            moved_path = self._move_to_sent_folder(selected_image)
            self.set_status("send_result", result.message)
            self.add_log(f"定时发图成功: {result.message}")
            self.add_log(f"已移动到已发送目录: {moved_path}")
            self.update_scheduled_img_count_signal.emit()
        except Exception as exc:
            self.set_status("send_result", f"异常: {exc}")
            self.add_log(f"定时发图异常: {exc}")
        finally:
            self.scheduled_image_sending = False
            if self.config.get("settings", {}).get("enable_scheduled_image_sender", False):
                schedule_base = datetime.datetime.now()
                self.scheduled_image_schedule_mode = "interval"
                self.scheduled_image_schedule_base_at = schedule_base
                self.scheduled_image_next_at = self._get_next_scheduled_image_due_time(schedule_base)
                if self.scheduled_image_next_at is None:
                    self.pause_scheduled_image_sender_signal.emit(
                        "本次发送完成后已超过今天可安排时间，定时主动发图已暂停；明天需要手动启用"
                    )
                else:
                    self._update_scheduled_image_status()
                    self.add_log(f"下一次定时发图已安排: {self.scheduled_image_next_at:%Y-%m-%d %H:%M}")
            _uia_init = None

    def _get_windows_system_tool(self, exe_name):
        system_root = os.environ.get("SystemRoot") or os.environ.get("windir") or r"C:\Windows"
        candidates = [
            Path(system_root) / "Sysnative" / exe_name,
            Path(system_root) / "System32" / exe_name,
        ]
        for candidate in candidates:
            if candidate.exists():
                return str(candidate)
        return exe_name

    def _run_hidden_command(self, args, timeout=10):
        kwargs = {
            "capture_output": True,
            "text": True,
            "timeout": timeout,
        }
        if os.name == "nt":
            kwargs.update(
                {
                    "encoding": "mbcs",
                    "errors": "replace",
                    "creationflags": getattr(subprocess, "CREATE_NO_WINDOW", 0),
                }
            )
        return subprocess.run(args, **kwargs)

    def is_narrator_running(self):
        if os.name != "nt":
            return False

        tasklist_exe = self._get_windows_system_tool("tasklist.exe")
        result = self._run_hidden_command(
            [tasklist_exe, "/FI", "IMAGENAME eq Narrator.exe", "/NH"],
            timeout=5,
        )
        if result.returncode != 0:
            raise RuntimeError((result.stderr or result.stdout or "tasklist 执行失败").strip())
        return "narrator.exe" in result.stdout.lower()

    def get_narrator_status_text(self):
        if os.name != "nt":
            return "仅 Windows 可用"

        try:
            return "已开启" if self.is_narrator_running() else "未开启"
        except Exception as exc:
            return f"检查失败: {exc}"

    def start_narrator(self):
        if os.name != "nt":
            QMessageBox.warning(self, "不可用", "讲述人功能仅支持 Windows。")
            return

        try:
            if self.is_narrator_running():
                self.add_log("Windows 讲述人已开启")
                self.set_status("narrator", "已开启")
                return

            narrator_exe = self._get_windows_system_tool("Narrator.exe")
            try:
                creationflags = getattr(subprocess, "CREATE_NO_WINDOW", 0)
                subprocess.Popen([narrator_exe], close_fds=True, creationflags=creationflags)
                self.add_log("已请求开启 Windows 讲述人")
            except OSError as exc:
                if not self._is_permission_error(str(exc)):
                    raise
                self.add_log("开启 Windows 讲述人需要管理员权限，正在请求授权")
                elevated, elevated_message = self._run_executable_as_admin(narrator_exe)
                if not elevated:
                    raise RuntimeError(elevated_message)
                self.add_log("已发起管理员授权开启 Windows 讲述人")
            QTimer.singleShot(1500, lambda: self.set_status("narrator", self.get_narrator_status_text()))
        except Exception as exc:
            self.add_log(f"开启 Windows 讲述人失败: {exc}")
            QMessageBox.warning(self, "开启失败", f"开启 Windows 讲述人失败:\n{exc}")

    def stop_narrator(self):
        if os.name != "nt":
            QMessageBox.warning(self, "不可用", "讲述人功能仅支持 Windows。")
            return

        try:
            if not self.is_narrator_running():
                self.add_log("Windows 讲述人未开启")
                self.set_status("narrator", "未开启")
                return

            taskkill_exe = self._get_windows_system_tool("taskkill.exe")
            result = self._run_hidden_command([taskkill_exe, "/IM", "Narrator.exe", "/T"], timeout=8)
            if result.returncode != 0:
                result = self._run_hidden_command([taskkill_exe, "/IM", "Narrator.exe", "/T", "/F"], timeout=8)

            if result.returncode != 0:
                message = (result.stderr or result.stdout or "taskkill 执行失败").strip()
                raise RuntimeError(message)

            self.add_log("已请求关闭 Windows 讲述人")
            QTimer.singleShot(1200, lambda: self.set_status("narrator", self.get_narrator_status_text()))
        except Exception as exc:
            self.add_log(f"关闭 Windows 讲述人失败: {exc}")
            QMessageBox.warning(self, "关闭失败", f"关闭 Windows 讲述人失败:\n{exc}")

    def _get_wechat_process_ids(self):
        if os.name != "nt":
            return set()

        tasklist_exe = self._get_windows_system_tool("tasklist.exe")
        process_ids = set()
        for image_name in ("Weixin.exe", "WeChat.exe"):
            result = self._run_hidden_command(
                [tasklist_exe, "/FI", f"IMAGENAME eq {image_name}", "/FO", "CSV", "/NH"],
                timeout=5,
            )
            if result.returncode != 0:
                continue
            for row in csv.reader(result.stdout.splitlines()):
                if len(row) < 2 or row[0].strip().lower() != image_name.lower():
                    continue
                try:
                    process_ids.add(int(row[1].replace(",", "").strip()))
                except ValueError:
                    continue
        return process_ids

    def _is_wechat_control_tree_ready(self):
        process_ids = self._get_wechat_process_ids()
        if not process_ids:
            return False, "等待微信进程出现"

        try:
            root = auto.GetRootControl()
            top_level_controls = root.GetChildren()
        except Exception as exc:
            return False, f"读取控件树失败: {exc}"

        for control in top_level_controls:
            try:
                process_id = int(getattr(control, "ProcessId", 0) or 0)
                if process_id not in process_ids:
                    continue
                children = control.GetChildren()
                if children:
                    name = str(getattr(control, "Name", "") or "微信")
                    return True, f"已识别微信控件树: {name}"
            except Exception:
                continue
        return False, "微信进程已启动，等待 UIAutomation 控件树"

    def start_one_click_wechat(self):
        if self.one_click_run_active:
            return
        if self.monitoring:
            QMessageBox.information(self, "监控已启动", "关键词监控已经在运行中。")
            return
        if os.name != "nt":
            QMessageBox.warning(self, "不可用", "一键启动监控仅支持 Windows。")
            return

        target_name = self.config.get("settings", {}).get("trigger_sender", "").strip()
        if not target_name:
            QMessageBox.warning(self, "目标联系人为空", "请先填写“目标对话/触发者昵称”。")
            return

        if self.wechat_launch_path_input is not None:
            configured_path = self.wechat_launch_path_input.text().strip().strip('"')
            self.config.setdefault("settings", {})["wechat_launch_path"] = configured_path
            self.save_config()

        launch_path = self.get_wechat_launch_path()
        if not launch_path or not Path(launch_path).is_file():
            QMessageBox.warning(self, "微信路径无效", "请先在“微信启动路径”中选择微信程序或桌面快捷方式。")
            return

        try:
            self.one_click_narrator_was_running = self.is_narrator_running()
        except Exception as exc:
            QMessageBox.warning(self, "讲述人状态异常", f"无法检查讲述人状态:\n{exc}")
            return

        self.one_click_run_active = True
        self.one_click_tree_stable_count = 0
        self.one_click_login_button_clicked = False
        self.one_click_login_wait_deadline = 0
        if self.one_click_run_btn is not None:
            self.one_click_run_btn.setEnabled(False)
            self.one_click_run_btn.setText("正在准备微信...")
        if self.open_target_window_btn is not None:
            self.open_target_window_btn.setEnabled(False)
        if self.start_btn is not None:
            self.start_btn.setEnabled(False)
        self.set_status("one_click", "正在启动讲述人")
        self.add_log(f"一键启动监控开始，目标联系人: {target_name}")
        self.add_log("步骤 1/5: 启动并确认 Windows 讲述人")

        self.start_narrator()
        self.one_click_stage_deadline = time.monotonic() + 15
        QTimer.singleShot(500, self._wait_one_click_narrator_started)

    def _wait_one_click_narrator_started(self):
        if not self.one_click_run_active:
            return
        try:
            narrator_ready = self.is_narrator_running()
        except Exception as exc:
            self._finish_one_click_wechat(False, f"检查讲述人状态失败: {exc}")
            return

        if narrator_ready:
            self.set_status("narrator", "已开启")
            self.set_status("one_click", "讲述人已确认，准备启动微信")
            self.add_log("已确认讲述人进程运行，等待 2 秒后启动微信")
            QTimer.singleShot(2000, self._launch_one_click_wechat)
            return

        if time.monotonic() >= self.one_click_stage_deadline:
            self._finish_one_click_wechat(False, "讲述人启动超时，已停止一键启动监控")
            return
        QTimer.singleShot(500, self._wait_one_click_narrator_started)

    def _launch_one_click_wechat(self):
        if not self.one_click_run_active:
            return
        launch_path = self.get_wechat_launch_path()
        try:
            os.startfile(launch_path)
        except Exception as exc:
            if not self.one_click_narrator_was_running:
                self.stop_narrator()
            self._finish_one_click_wechat(False, f"启动微信失败: {exc}")
            return

        self.set_status("one_click", "微信已启动，等待控件树")
        self.add_log(f"已启动微信: {launch_path}")
        self.add_log("正在确认微信 UIAutomation 控件树，讲述人将在确认完成前保持开启")
        self.one_click_stage_deadline = time.monotonic() + 45
        self.one_click_tree_stable_count = 0
        QTimer.singleShot(1000, self._wait_one_click_wechat_control_tree)

    def _wait_one_click_wechat_control_tree(self):
        if not self.one_click_run_active:
            return

        ready, detail = self._is_wechat_control_tree_ready()
        if ready:
            self.one_click_tree_stable_count += 1
            self.set_status("one_click", f"控件树确认 {self.one_click_tree_stable_count}/3")
            if self.one_click_tree_stable_count >= 3:
                self.add_log(f"{detail}，已连续确认 3 次")
                self.set_status("one_click", "控件树已就绪，准备检查登录")
                self.add_log("步骤 2/5: 微信控件树已稳定")
                self.one_click_stage_deadline = time.monotonic() + 15
                QTimer.singleShot(1000, self._prepare_login_before_closing_narrator)
                return
        else:
            self.one_click_tree_stable_count = 0
            self.set_status("one_click", detail)

        if time.monotonic() >= self.one_click_stage_deadline:
            self._finish_one_click_wechat(
                False,
                "45 秒内未能确认微信控件树；为避免过早关闭，讲述人将保持开启",
            )
            return
        QTimer.singleShot(1000, self._wait_one_click_wechat_control_tree)

    def _prepare_login_before_closing_narrator(self):
        if not self.one_click_run_active:
            return

        process_ids = self._get_wechat_process_ids()
        state, detail = self.wechat.get_login_state(process_ids)
        self.set_status("one_click", detail)

        if state == "login_ready":
            self.add_log("已在讲述人开启期间找到微信登录按钮，正在点击并验证")
            click_result = self.wechat.click_login_button(process_ids)
            if not click_result:
                summary = self.wechat.get_login_ui_summary(process_ids)
                self.add_log(f"登录按钮点击失败控件摘要: {summary}")
                self._finish_one_click_wechat(False, click_result.message)
                return
            self.one_click_login_button_clicked = True
            self._begin_one_click_login_wait()
            self.add_log(f"步骤 3/5: {click_result.message}")
            self.set_status("one_click", "登录已触发，正在关闭讲述人")
            QTimer.singleShot(300, self._close_one_click_narrator)
            return

        if state == "waiting_mobile":
            self.one_click_login_button_clicked = True
            self._begin_one_click_login_wait()
            self.add_log("步骤 3/5: 微信已经在等待手机确认登录")
            QTimer.singleShot(300, self._close_one_click_narrator)
            return

        if state == "logged_in":
            self.add_log(f"步骤 3/5: {detail}")
            QTimer.singleShot(300, self._close_one_click_narrator)
            return

        if state == "qr_required":
            self.add_log("步骤 3/5: 已检测到微信扫码登录界面")
            QTimer.singleShot(300, self._close_one_click_narrator)
            return

        if time.monotonic() >= self.one_click_stage_deadline:
            summary = self.wechat.get_login_ui_summary(process_ids)
            self.add_log(f"登录界面识别超时控件摘要: {summary}")
            self._finish_one_click_wechat(
                False,
                f"15 秒内未能识别微信登录状态。{detail}",
            )
            return

        if self.one_click_run_btn is not None:
            self.one_click_run_btn.setText("正在识别登录按钮...")
        QTimer.singleShot(500, self._prepare_login_before_closing_narrator)

    def _close_one_click_narrator(self):
        if not self.one_click_run_active:
            return
        self.set_status("one_click", "正在关闭讲述人")
        self.add_log("微信控件树已稳定，开始关闭 Windows 讲述人")
        self.stop_narrator()
        self.one_click_stage_deadline = time.monotonic() + 10
        QTimer.singleShot(500, self._wait_one_click_narrator_stopped)

    def _begin_one_click_login_wait(self):
        self.one_click_login_wait_deadline = (
            time.monotonic() + ONE_CLICK_LOGIN_CONFIRM_TIMEOUT_SECONDS
        )

    def _wait_one_click_narrator_stopped(self):
        if not self.one_click_run_active:
            return
        try:
            narrator_running = self.is_narrator_running()
        except Exception as exc:
            self._finish_one_click_wechat(False, f"检查讲述人关闭状态失败: {exc}")
            return

        if not narrator_running:
            self.set_status("narrator", "未开启")
            self.add_log("Windows 讲述人已关闭")
            if self.one_click_login_button_clicked:
                self._begin_one_click_login_wait()
            else:
                self.one_click_stage_deadline = time.monotonic() + 15
            QTimer.singleShot(300, self._continue_one_click_login_flow)
            return
        if time.monotonic() >= self.one_click_stage_deadline:
            self.set_status("narrator", "关闭超时，流程继续")
            self.add_log(
                "Windows 讲述人关闭确认超时；不再中断一键启动，"
                "继续检测微信登录并打开目标联系人"
            )
            if self.one_click_login_button_clicked:
                self._begin_one_click_login_wait()
            else:
                self.one_click_stage_deadline = time.monotonic() + 15
            QTimer.singleShot(300, self._continue_one_click_login_flow)
            return
        QTimer.singleShot(500, self._wait_one_click_narrator_stopped)

    def _continue_one_click_login_flow(self):
        if not self.one_click_run_active:
            return

        process_ids = self._get_wechat_process_ids()
        state, detail = self.wechat.get_login_state(process_ids)
        if state == "logged_in":
            self.set_status("one_click", "微信已登录，正在打开目标联系人")
            self.add_log(f"步骤 4/5: {detail}")
            QTimer.singleShot(300, self._start_one_click_target_window)
            return

        if state == "login_ready":
            click_result = self.wechat.click_login_button(process_ids)
            if not click_result:
                self._finish_one_click_wechat(False, click_result.message)
                return
            self.one_click_login_button_clicked = True
            self._begin_one_click_login_wait()
            self.add_log(f"步骤 4/5: {click_result.message}")
            self.add_log("登录请求已发出，持续等待手机微信确认")
            self._wait_one_click_login()
            return

        if state == "waiting_mobile":
            if not self.one_click_login_button_clicked:
                self.add_log("检测到微信正在等待手机确认登录")
                self.one_click_login_button_clicked = True
            self._begin_one_click_login_wait()
            self._wait_one_click_login()
            return

        if state == "qr_required":
            self._finish_one_click_wechat(
                False,
                "检测到微信当前需要扫码登录。请先使用手机微信扫描二维码，登录完成后再次点击“一键启动监控”。",
                attention=True,
            )
            return

        if self.one_click_login_button_clicked:
            self._wait_one_click_login()
            return

        if time.monotonic() >= self.one_click_stage_deadline:
            summary = self.wechat.get_login_ui_summary(process_ids)
            self.add_log(f"未识别登录界面控件摘要: {summary}")
            self._finish_one_click_wechat(
                False,
                f"15 秒内未能识别微信登录状态。{detail}\n请查看日志中的控件摘要。",
            )
            return

        self.set_status("one_click", detail)
        if self.one_click_run_btn is not None:
            self.one_click_run_btn.setText("正在识别登录状态...")
        QTimer.singleShot(1000, self._continue_one_click_login_flow)

    def _wait_one_click_login(self):
        if not self.one_click_run_active:
            return

        if not self.one_click_login_wait_deadline:
            self._begin_one_click_login_wait()

        process_ids = self._get_wechat_process_ids()
        state, detail = self.wechat.get_login_state(process_ids)
        if state == "logged_in":
            self.set_status("one_click", "登录成功，正在打开目标联系人")
            self.add_log(f"已检测到登录完成: {detail}")
            QTimer.singleShot(500, self._start_one_click_target_window)
            return
        if state == "qr_required":
            self._finish_one_click_wechat(
                False,
                "微信已切换到扫码登录。请使用手机微信扫描二维码，登录完成后再次点击“一键启动监控”。",
                attention=True,
            )
            return

        if state == "login_ready":
            click_result = self.wechat.click_login_button(process_ids)
            if not click_result:
                self._finish_one_click_wechat(False, click_result.message)
                return
            self.one_click_login_button_clicked = True
            self._begin_one_click_login_wait()
            self.add_log(f"微信重新显示登录按钮，已再次触发: {click_result.message}")

        if time.monotonic() >= self.one_click_login_wait_deadline:
            summary = self.wechat.get_login_ui_summary(process_ids)
            self.add_log(f"等待手机确认登录超时控件摘要: {summary}")
            self._finish_one_click_wechat(
                False,
                (
                    f"{ONE_CLICK_LOGIN_CONFIRM_TIMEOUT_SECONDS} 秒内未检测到微信登录完成。"
                    f"{detail}\n如果手机已确认，请确认微信主界面已经显示后再次点击“一键启动监控”。"
                ),
            )
            return

        status_text = "等待手机确认登录"
        if state not in {"waiting_mobile", "login_ready", "starting"}:
            status_text = f"等待登录完成: {detail}"
        self.set_status("one_click", status_text)
        if self.one_click_run_btn is not None:
            self.one_click_run_btn.setText("等待手机确认")
        QTimer.singleShot(1000, self._wait_one_click_login)

    def _start_one_click_target_window(self):
        if not self.one_click_run_active:
            return

        target_name = self.config.get("settings", {}).get("trigger_sender", "").strip()
        if not target_name:
            self._finish_one_click_wechat(False, "目标联系人名称为空")
            return

        self.set_status("one_click", f"正在打开并置顶: {target_name}")
        if self.one_click_run_btn is not None:
            self.one_click_run_btn.setText("正在打开目标窗口...")
        self.add_log(f"正在搜索联系人并打开独立窗口: {target_name}")
        threading.Thread(
            target=self._one_click_target_window_worker,
            args=(target_name,),
            daemon=True,
        ).start()

    def _one_click_target_window_worker(self, target_name):
        _uia_init = auto.UIAutomationInitializerInThread()
        try:
            result = self.wechat.open_independent_chat_window(target_name)
            self.one_click_target_finished_signal.emit(bool(result), result.message)
        except Exception as exc:
            self.one_click_target_finished_signal.emit(
                False,
                f"打开目标聊天窗口异常: {exc}",
            )
        finally:
            _uia_init = None

    def _finish_one_click_target_window(self, success, message):
        if not self.one_click_run_active:
            return
        self.set_status("target", message)
        self.add_log(message)
        if not success:
            self._finish_one_click_wechat(False, message)
            return

        self.add_log("步骤 5/5: 目标独立窗口已打开并置顶，正在启动监控")
        if not self.start_monitoring(silent=True):
            self._finish_one_click_wechat(
                False,
                "目标窗口已打开，但启动消息监控失败，请查看日志。",
            )
            return
        self._finish_one_click_wechat(
            True,
            "微信已登录，目标独立窗口已打开并置顶，关键词监控已启动。",
        )

    def _finish_one_click_wechat(self, success, message, attention=False):
        self.one_click_run_active = False
        if self.one_click_run_btn is not None:
            self.one_click_run_btn.setEnabled(True)
            self.one_click_run_btn.setText("一键启动监控")
        if self.open_target_window_btn is not None:
            self.open_target_window_btn.setEnabled(True)
        if self.start_btn is not None:
            self.start_btn.setEnabled(not self.monitoring)
        self.set_status(
            "one_click",
            "监控已启动" if success else ("需要扫码" if attention else "未完成"),
        )
        self.add_log(f"一键启动监控{'完成' if success else '未完成'}: {message}")
        if not success:
            if attention:
                QMessageBox.information(self, "需要扫码登录", message)
            else:
                QMessageBox.warning(self, "一键启动监控未完成", message)

    def _get_current_session_identifier(self):
        query_exe = self._get_windows_system_tool("query.exe")
        try:
            result = self._run_hidden_command([query_exe, "user"], timeout=5)
        except Exception as exc:
            result = None
            parse_note = f"读取远程会话失败: {exc}"
        else:
            parse_note = ""
            session_id = self._parse_query_user_session_id(result.stdout)
            if session_id:
                if result.returncode != 0:
                    self.add_log("query user 返回码异常，但已从输出中识别到当前会话 ID")
                return session_id, ""

            raw_output = (result.stderr or result.stdout or "query user 执行失败").strip()
            if result.returncode != 0:
                parse_note = raw_output
            else:
                parse_note = raw_output or "未能从 query user 解析当前会话 ID"

        session_name = os.environ.get("SESSIONNAME", "").strip()
        if session_name and session_name.lower() != "console":
            return session_name, f"{parse_note}；已改用 SESSIONNAME={session_name}"

        return None, parse_note or "未检测到远程桌面会话"

    def _parse_query_user_session_id(self, output):
        lines = str(output or "").splitlines()
        current_lines = [line for line in lines if line.lstrip().startswith(">")]

        username = os.environ.get("USERNAME", "").lower()
        if username:
            current_lines.extend(
                line
                for line in lines[1:]
                if line.lstrip().lstrip(">").strip().lower().startswith(username)
            )

        seen = set()
        for line in current_lines:
            normalized = line.lstrip().lstrip(">").strip()
            if not normalized or normalized in seen:
                continue
            seen.add(normalized)
            tokens = normalized.split()
            for token in tokens[1:5]:
                if token.isdigit():
                    return token
        return ""

    def _is_permission_error(self, message):
        lower_message = message.lower()
        permission_keywords = (
            "access is denied",
            "requires elevation",
            "requested operation requires elevation",
            "winerror 740",
            "740",
            "拒绝访问",
            "权限",
            "需要提升",
            "请求的操作需要提升",
        )
        return any(keyword in lower_message or keyword in message for keyword in permission_keywords)

    def _format_tscon_error(self, message):
        message = message.strip() or "tscon 执行失败"
        if self._is_permission_error(message):
            return f"{message}\n\n请以管理员身份运行本软件后重试。"
        return message

    def _run_executable_as_admin(self, executable, params="", show=1):
        try:
            result = ctypes.windll.shell32.ShellExecuteW(
                None,
                "runas",
                str(executable),
                params,
                None,
                show,
            )
        except Exception as exc:
            return False, str(exc)

        if result <= 32:
            return False, f"管理员授权启动失败，错误码: {result}"
        return True, ""

    def _run_tscon_as_admin(self, tscon_exe, session_identifier):
        return self._run_executable_as_admin(tscon_exe, f"{session_identifier} /dest:console", show=0)

    def disconnect_remote_session(self):
        if os.name != "nt":
            QMessageBox.warning(self, "不可用", "断开远程连接功能仅支持 Windows。")
            return

        session_name = os.environ.get("SESSIONNAME", "").strip()
        if session_name.lower() == "console":
            self.add_log("当前已经是控制台会话，无需断开远程连接")
            QMessageBox.information(self, "无需操作", "当前已经是控制台会话，无需断开远程连接。")
            return

        session_identifier, note = self._get_current_session_identifier()
        if note:
            self.add_log(note)
        if not session_identifier:
            QMessageBox.warning(self, "断开失败", "无法识别当前远程桌面会话，请查看日志。")
            return

        tscon_exe = self._get_windows_system_tool("tscon.exe")
        self.add_log(f"准备断开远程桌面会话: {session_identifier}")
        try:
            result = self._run_hidden_command([tscon_exe, str(session_identifier), "/dest:console"], timeout=10)
        except Exception as exc:
            self.add_log(f"断开远程连接失败: {exc}")
            QMessageBox.warning(self, "断开失败", f"执行 tscon 失败:\n{exc}")
            return

        if result.returncode != 0:
            raw_message = (result.stderr or result.stdout or "tscon 执行失败").strip()
            if self._is_permission_error(raw_message):
                self.add_log("断开远程连接需要管理员权限，正在请求管理员授权")
                elevated, elevated_message = self._run_tscon_as_admin(tscon_exe, session_identifier)
                if elevated:
                    self.add_log("已发起管理员授权执行断开远程连接命令")
                    return
                raw_message = f"{raw_message}\n{elevated_message}"

            message = self._format_tscon_error(raw_message)
            self.add_log(f"断开远程连接失败: {message}")
            QMessageBox.warning(self, "断开失败", f"执行 tscon 失败:\n{message}")
            return

        self.add_log("已执行断开远程连接命令，桌面会话将保持在控制台")

    def apply_app_style(self):
        self.setStyleSheet(
            """
            QWidget {
                font-size: 13px;
                color: #172033;
            }
            QWidget#root {
                background: #f3f6fa;
            }
            QWidget#sidePanel, QWidget#actionPanel {
                background: #ffffff;
                border: 1px solid #dce3ed;
                border-radius: 8px;
            }
            QGroupBox {
                font-weight: 600;
                border: 1px solid #dce3ed;
                border-radius: 8px;
                margin-top: 12px;
                padding: 14px 12px 12px 12px;
                background: #ffffff;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 12px;
                padding: 0 5px;
                color: #344054;
            }
            QLineEdit, QComboBox, QSpinBox, QDoubleSpinBox {
                min-height: 32px;
                border: 1px solid #cad4e0;
                border-radius: 8px;
                padding: 3px 10px;
                background: #ffffff;
            }
            QLineEdit:focus, QComboBox:focus, QSpinBox:focus, QDoubleSpinBox:focus {
                border: 1px solid #16805a;
            }
            QPushButton {
                min-height: 32px;
                border: 1px solid #b8c2d0;
                border-radius: 8px;
                padding: 6px 12px;
                background: #f8fafc;
            }
            QPushButton:hover {
                background: #eef4f8;
            }
            QPushButton#primaryButton {
                color: #ffffff;
                background: #14724f;
                border-color: #14724f;
                font-weight: 700;
                min-height: 42px;
            }
            QPushButton#dangerButton {
                color: #ffffff;
                background: #a33a2f;
                border-color: #a33a2f;
                font-weight: 700;
                min-height: 42px;
            }
            QPushButton#secondaryButton {
                background: #ffffff;
                min-height: 36px;
            }
            QTabWidget::pane {
                border: 1px solid #dce3ed;
                border-radius: 8px;
                background: #ffffff;
            }
            QTabBar::tab {
                padding: 10px 22px;
                margin-right: 6px;
                border-radius: 9px;
                background: transparent;
                color: #586475;
            }
            QTabBar::tab:selected {
                background: #e8f2ee;
                color: #124d38;
                font-weight: 700;
            }
            QListWidget {
                border: 1px solid #dce3ed;
                border-radius: 8px;
                background: #fbfcfe;
                padding: 6px;
            }
            QWidget#statusCard {
                background: #f8fafc;
                border: 1px solid #e4eaf2;
                border-radius: 8px;
            }
            QLabel#statusTitle {
                color: #667085;
                font-size: 12px;
                font-weight: 600;
            }
            QLabel#statusValue {
                color: #172033;
                font-weight: 500;
            }
            QLabel#scheduleCalculation {
                color: #344054;
                background: #f8fafc;
                border: 1px solid #e4eaf2;
                border-radius: 8px;
                padding: 12px;
            }
            QScrollArea {
                border: none;
                background: transparent;
            }
            """
        )

    def available_screen_geometry(self):
        screen = self.screen() or QApplication.primaryScreen()
        if screen:
            return screen.availableGeometry()
        return QApplication.desktop().availableGeometry()

    def apply_initial_window_geometry(self):
        screen_rect = self.available_screen_geometry()
        margin = 40 if min(screen_rect.width(), screen_rect.height()) >= 900 else 16
        max_width = max(320, screen_rect.width() - margin)
        max_height = max(320, screen_rect.height() - margin)

        preferred_width = min(int(screen_rect.width() * 0.72), 1800)
        preferred_height = int(screen_rect.height() * 0.88)
        initial_width = min(max(1120, preferred_width), max_width)
        initial_height = min(max(720, preferred_height), max_height)
        self.setMinimumSize(min(760, max_width), min(520, max_height))
        self.resize(initial_width, initial_height)
        self.move(
            screen_rect.x() + max(0, (screen_rect.width() - initial_width) // 2),
            screen_rect.y() + max(0, (screen_rect.height() - initial_height) // 2),
        )

    def bring_main_window_to_front(self):
        self.showNormal()
        self.raise_()
        self.activateWindow()
        if os.name == "nt":
            try:
                hwnd = int(self.winId())
                ctypes.windll.user32.ShowWindow(hwnd, 9)
                ctypes.windll.user32.SetForegroundWindow(hwnd)
            except Exception:
                pass

    def initUI(self):
        self.setObjectName("root")
        self.apply_app_style()

        outer_layout = QHBoxLayout()
        outer_layout.setContentsMargins(14, 14, 14, 14)
        outer_layout.setSpacing(14)

        sidebar = QWidget()
        sidebar_layout = QVBoxLayout(sidebar)
        sidebar_layout.setContentsMargins(0, 0, 0, 0)
        sidebar_layout.setSpacing(12)
        sidebar.setMinimumWidth(260)
        sidebar.setMaximumWidth(340)
        sidebar.setSizePolicy(QSizePolicy.Preferred, QSizePolicy.Expanding)

        action_panel = QWidget()
        action_panel.setObjectName("actionPanel")
        action_layout = QVBoxLayout(action_panel)
        action_layout.setContentsMargins(12, 12, 12, 12)
        action_layout.setSpacing(10)

        check_btn = QPushButton("立即自检")
        check_btn.setObjectName("secondaryButton")
        check_btn.clicked.connect(self.run_startup_check)

        self.one_click_run_btn = QPushButton("一键启动监控")
        self.one_click_run_btn.setObjectName("primaryButton")
        self.one_click_run_btn.setToolTip(
            "自动准备微信控件树、完成登录检测、打开并置顶目标窗口，然后启动关键词监控"
        )
        self.one_click_run_btn.clicked.connect(self.start_one_click_wechat)

        self.open_target_window_btn = QPushButton("打开并置顶目标窗口")
        self.open_target_window_btn.setObjectName("secondaryButton")
        self.open_target_window_btn.setToolTip("搜索目标联系人，打开独立聊天窗口并设置置顶")
        self.open_target_window_btn.clicked.connect(self.open_target_chat_window)

        narrator_buttons = QWidget()
        narrator_layout = QHBoxLayout(narrator_buttons)
        narrator_layout.setContentsMargins(0, 0, 0, 0)
        narrator_layout.setSpacing(8)

        start_narrator_btn = QPushButton("开启讲述人")
        start_narrator_btn.setObjectName("secondaryButton")
        start_narrator_btn.setToolTip("启动 Windows 讲述人，帮助 uiautomation 识别微信控件")
        start_narrator_btn.clicked.connect(self.start_narrator)

        stop_narrator_btn = QPushButton("关闭讲述人")
        stop_narrator_btn.setObjectName("secondaryButton")
        stop_narrator_btn.setToolTip("关闭 Windows 讲述人")
        stop_narrator_btn.clicked.connect(self.stop_narrator)

        narrator_layout.addWidget(start_narrator_btn)
        narrator_layout.addWidget(stop_narrator_btn)

        self.start_btn = QPushButton("开始监控")
        self.start_btn.setObjectName("primaryButton")
        self.start_btn.clicked.connect(self.start_monitoring)

        self.stop_btn = QPushButton("停止监控")
        self.stop_btn.setObjectName("dangerButton")
        self.stop_btn.setEnabled(False)
        self.stop_btn.clicked.connect(self.stop_monitoring)

        disconnect_btn = QPushButton("断开远程连接")
        disconnect_btn.setObjectName("secondaryButton")
        disconnect_btn.setToolTip("使用 tscon 将当前 RDP 会话切回控制台，避免锁屏导致控件树失效")
        disconnect_btn.clicked.connect(self.disconnect_remote_session)

        action_layout.addWidget(check_btn)
        action_layout.addWidget(self.one_click_run_btn)
        action_layout.addWidget(self.open_target_window_btn)
        action_layout.addWidget(narrator_buttons)
        action_layout.addWidget(self.start_btn)
        action_layout.addWidget(self.stop_btn)
        action_layout.addWidget(disconnect_btn)

        status_panel = self.init_status_panel()
        sidebar_layout.addWidget(action_panel)
        sidebar_layout.addWidget(status_panel, 1)

        tabs = QTabWidget()
        tabs.setDocumentMode(True)

        config_scroll = QScrollArea()
        config_scroll.setWidgetResizable(True)
        config_scroll.setFrameShape(QScrollArea.NoFrame)
        config_inner = QWidget()
        config_layout = QVBoxLayout(config_inner)
        config_layout.setContentsMargins(12, 12, 12, 12)
        config_layout.setSpacing(10)
        config_layout.addWidget(self.init_language_choose())
        config_layout.addLayout(self.init_settings())
        config_layout.addLayout(self.init_keyword_timing_settings())
        config_layout.addWidget(self.init_remote_control_settings())
        config_layout.addStretch()
        config_scroll.setWidget(config_inner)
        tabs.addTab(config_scroll, "关键词触发回复")

        schedule_scroll = QScrollArea()
        schedule_scroll.setWidgetResizable(True)
        schedule_scroll.setFrameShape(QScrollArea.NoFrame)
        schedule_inner = QWidget()
        schedule_layout = QVBoxLayout(schedule_inner)
        schedule_layout.setContentsMargins(12, 12, 12, 12)
        schedule_layout.setSpacing(10)
        schedule_layout.addLayout(self.init_scheduled_image_settings())
        schedule_layout.addStretch()
        schedule_scroll.setWidget(schedule_inner)
        tabs.addTab(schedule_scroll, "主动定时发送")

        log_page = QWidget()
        log_layout = QVBoxLayout(log_page)
        log_layout.setContentsMargins(12, 12, 12, 12)
        log_layout.addLayout(self.init_monitor_log())
        tabs.addTab(log_page, "日志")

        outer_layout.addWidget(sidebar)
        outer_layout.addWidget(tabs, 1)

        self.setLayout(outer_layout)
        self.apply_initial_window_geometry()
        self.setWindowTitle(APP_WINDOW_TITLE)
        self.show()
        QTimer.singleShot(100, self.bring_main_window_to_front)


def configure_high_dpi():
    if hasattr(Qt, "AA_EnableHighDpiScaling"):
        QApplication.setAttribute(Qt.AA_EnableHighDpiScaling, True)
    if hasattr(Qt, "AA_UseHighDpiPixmaps"):
        QApplication.setAttribute(Qt.AA_UseHighDpiPixmaps, True)

    rounding_policy = getattr(Qt, "HighDpiScaleFactorRoundingPolicy", None)
    if rounding_policy and hasattr(QApplication, "setHighDpiScaleFactorRoundingPolicy"):
        QApplication.setHighDpiScaleFactorRoundingPolicy(rounding_policy.PassThrough)


if __name__ == "__main__":
    sys.excepthook = report_unhandled_exception
    configure_high_dpi()
    app = QApplication(sys.argv)
    ex = MomoReplyGUI()
    sys.exit(app.exec_())
