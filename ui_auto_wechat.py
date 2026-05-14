import os
import re
import threading
import time
from ctypes import windll
from dataclasses import dataclass
from typing import Callable, Optional, Tuple

import pyperclip
import uiautomation as auto

from clipboard import setClipboardFiles
from wechat_locale import WeChatLocale


@dataclass
class AutomationResult:
    ok: bool
    message: str

    def __bool__(self):
        return self.ok


def ok(message: str = "操作成功") -> AutomationResult:
    return AutomationResult(True, message)


def fail(message: str) -> AutomationResult:
    return AutomationResult(False, message)


def move(element):
    x, y = element.GetPosition()
    auto.SetCursorPos(x, y)


def click(element):
    x, y = element.GetPosition()
    auto.Click(x, y)


class WeChat:
    def __init__(self, locale="zh-CN"):
        assert locale in WeChatLocale.getSupportedLocales()
        self.lc = WeChatLocale(locale)
        self.last_message_monitoring = False
        self.last_captured_text = ""
        self.last_message_callback = None
        self.last_monitor_status = ""

    def check_target_window(self, target_name: str) -> AutomationResult:
        if not target_name:
            return fail("目标窗口名为空")
        try:
            win = self.get_independent_window(target_name)
        except Exception as exc:
            return fail(f"检查目标窗口失败: {exc}")
        if win:
            return ok(f"已找到独立聊天窗口: {target_name}")
        return fail(f"未找到独立聊天窗口: {target_name}")

    def press_enter(self):
        auto.SendKeys("{enter}")

    def paste_text(self, text: str) -> AutomationResult:
        try:
            pyperclip.copy(text)
            time.sleep(0.3)
            auto.SendKeys("{Ctrl}v")
            return ok("文本已粘贴到输入框")
        except Exception as exc:
            return fail(f"文本粘贴失败: {exc}")

    def get_independent_window(self, target_name):
        win = auto.WindowControl(Name=target_name, searchDepth=1)
        if win.Exists(0.2, 0):
            return win
        return None

    def _activate_window(self, chat_win) -> AutomationResult:
        try:
            hwnd = chat_win.NativeWindowHandle
            user32 = windll.user32
            if user32.IsIconic(hwnd):
                user32.OpenIconicWindow(hwnd)
                time.sleep(0.5)
            user32.ShowWindow(hwnd, 9)
            user32.SetForegroundWindow(hwnd)
            chat_win.SetFocus()
            time.sleep(0.3)
            return ok("聊天窗口已激活")
        except Exception as exc:
            try:
                chat_win.SetFocus()
                time.sleep(0.3)
                return ok("聊天窗口已通过 SetFocus 激活")
            except Exception:
                return fail(f"无法激活聊天窗口: {exc}")

    def _find_chat_input(self, chat_win):
        candidates = [
            chat_win.EditControl(Name="输入"),
            chat_win.EditControl(foundIndex=1),
            chat_win.EditControl(searchDepth=8, foundIndex=1),
        ]
        for edit_input in candidates:
            try:
                if edit_input.Exists(0.3, 0):
                    return edit_input
            except Exception:
                pass
        return None

    def _focus_independent_chat_input(self, target_name: str):
        chat_win = self.get_independent_window(target_name)
        if not chat_win:
            return None, fail(f"找不到名为 '{target_name}' 的独立聊天窗口，请确认聊天窗口已经单独拖出")

        activate_result = self._activate_window(chat_win)
        if not activate_result:
            return None, activate_result

        current_mouse_pos = auto.GetCursorPos()
        try:
            edit_input = self._find_chat_input(chat_win)
            if edit_input:
                move(edit_input)
                click(edit_input)
                time.sleep(0.2)
                return chat_win, ok("已定位并聚焦输入框")

            rect = chat_win.BoundingRectangle
            if not rect:
                return None, fail(f"未能定位 '{target_name}' 的输入区域")

            click_x = rect.left + (rect.right - rect.left) // 2
            click_y = rect.bottom - 60
            auto.Click(click_x, click_y)
            time.sleep(0.2)
            return chat_win, ok("未找到输入控件，已通过窗口底部坐标聚焦输入区")
        finally:
            auto.SetCursorPos(current_mouse_pos[0], current_mouse_pos[1])

    def _get_message_list(self, chat_win):
        try:
            msg_list = chat_win.ListControl(Name=self.lc.message)
            if msg_list.Exists(0.2, 0):
                return msg_list
        except Exception:
            pass

        try:
            msg_list = chat_win.ListControl(foundIndex=1)
            if msg_list.Exists(0.2, 0):
                return msg_list
        except Exception:
            pass
        return None

    def _message_signature(self, item) -> str:
        try:
            if item.Name and str(item.Name).strip():
                return str(item.Name).strip()
            child_names = []
            for child in item.GetChildren():
                name = getattr(child, "Name", "")
                if name and str(name).strip():
                    child_names.append(str(name).strip())
            return "|".join(child_names)
        except Exception:
            return ""

    def _control_text(self, control) -> str:
        try:
            if control.Name and str(control.Name).strip():
                return str(control.Name).strip()
            child_names = []
            for child in control.GetChildren():
                name = getattr(child, "Name", "")
                if name and str(name).strip():
                    child_names.append(str(name).strip())
            return "|".join(child_names)
        except Exception:
            return ""

    def _capture_message_state(self, chat_win) -> Tuple[int, str]:
        msg_list = self._get_message_list(chat_win)
        if not msg_list:
            return 0, ""
        try:
            items = msg_list.GetChildren()
            count = len(items)
            tail = [self._message_signature(item) for item in items[-3:]]
            tail = [text for text in tail if text]
            return count, "||".join(tail)
        except Exception:
            return 0, ""

    def _wait_for_message_change(self, chat_win, before_state, timeout=8.0) -> AutomationResult:
        deadline = time.time() + timeout
        while time.time() < deadline:
            after_state = self._capture_message_state(chat_win)
            if after_state[0] > before_state[0]:
                return ok("发送后消息列表数量已变化")
            if after_state != before_state and after_state[0] > 0:
                return ok("发送后消息列表内容已变化")
            time.sleep(0.4)
        return fail("发送后未检测到消息列表变化，可能未发送成功或微信控件未刷新")

    def _get_toolbar_buttons(self, chat_win):
        buttons = []
        try:
            for control in chat_win.GetChildren():
                try:
                    if getattr(control, "ControlTypeName", "") == "ToolBarControl":
                        for child in control.GetChildren():
                            if getattr(child, "ControlTypeName", "") == "ButtonControl":
                                buttons.append(child)
                            else:
                                try:
                                    for grandchild in child.GetChildren():
                                        if getattr(grandchild, "ControlTypeName", "") == "ButtonControl":
                                            buttons.append(grandchild)
                                except Exception:
                                    pass
                except Exception:
                    pass
        except Exception:
            pass
        return buttons

    def _click_send_button(self, chat_win) -> AutomationResult:
        candidates = self._get_toolbar_buttons(chat_win)
        if not candidates:
            try:
                candidates = chat_win.GetChildren()
            except Exception:
                candidates = []

        exact_match = None
        fallback_match = None

        for control in candidates:
            try:
                if getattr(control, "ControlTypeName", "") != "ButtonControl":
                    continue
                text = self._control_text(control)
                if not text:
                    continue
                if text == "发送":
                    exact_match = control
                    break
                if "发送" in text and all(keyword not in text for keyword in ("表情", "收藏", "文件")):
                    fallback_match = control
            except Exception:
                pass

        button = exact_match or fallback_match
        if button is None:
            return fail("未找到发送按钮")

        click(button)
        return ok("已点击发送按钮")

    def _click_send_file_button(self, chat_win) -> AutomationResult:
        exact_match = None
        fallback_match = None
        for control in self._get_toolbar_buttons(chat_win):
            try:
                text = self._control_text(control)
                if not text:
                    continue
                if text == "发送文件":
                    exact_match = control
                    break
                if "文件" in text and "发送" in text:
                    fallback_match = control
            except Exception:
                pass

        button = exact_match or fallback_match
        if button is None:
            return fail("未找到发送文件按钮")

        click(button)
        return ok("已点击发送文件按钮")

    def _attach_file_via_dialog(self, chat_win, path: str) -> AutomationResult:
        click_file_result = self._click_send_file_button(chat_win)
        if not click_file_result:
            return click_file_result

        dialog = None
        deadline = time.time() + 5.0
        while time.time() < deadline:
            try:
                dialog = auto.WindowControl(ClassName="#32770", searchDepth=1)
                if dialog.Exists(0.2, 0):
                    break
            except Exception:
                dialog = None
            time.sleep(0.2)

        if dialog is None or not dialog.Exists(0.2, 0):
            return fail("点击发送文件后未出现文件选择窗口")

        try:
            dialog.SetFocus()
        except Exception:
            pass

        try:
            pyperclip.copy(path)
            time.sleep(0.2)
            auto.SendKeys("{Ctrl}v")
            time.sleep(0.2)
            auto.SendKeys("{Enter}")
            time.sleep(1.5)
            return ok("已通过文件选择窗口附加文件")
        except Exception as exc:
            return fail(f"文件选择窗口附加文件失败: {exc}")

    def send_msg(self, name, text: str = None) -> AutomationResult:
        chat_win, focus_result = self._focus_independent_chat_input(name)
        if not chat_win:
            return focus_result
        before_state = self._capture_message_state(chat_win)

        if text is not None:
            paste_result = self.paste_text(text)
            if not paste_result:
                return paste_result

        try:
            self.press_enter()
        except Exception as exc:
            return fail(f"按 Enter 发送文本失败: {exc}")

        wait_result = self._wait_for_message_change(chat_win, before_state, timeout=5.0)
        if wait_result:
            return ok("文本发送成功")
        return wait_result

    def send_file(self, name: str, path: str) -> AutomationResult:
        if not path or not os.path.exists(path):
            return fail(f"文件不存在: {path}")

        chat_win, focus_result = self._focus_independent_chat_input(name)
        if not chat_win:
            return focus_result
        before_state = self._capture_message_state(chat_win)

        try:
            setClipboardFiles([path])
        except Exception as exc:
            return fail(f"设置文件剪贴板失败: {exc}")
        time.sleep(0.5)

        self._activate_window(chat_win)
        time.sleep(0.2)
        rect = chat_win.BoundingRectangle
        if rect:
            auto.Click(rect.left + (rect.right - rect.left) // 2, rect.bottom - 60)
            time.sleep(0.2)

        try:
            auto.SendKeys("{Ctrl}v")
            time.sleep(1.5)
            self.press_enter()
        except Exception as exc:
            return fail(f"粘贴并发送文件失败: {exc}")

        wait_result = self._wait_for_message_change(chat_win, before_state, timeout=4.0)
        if wait_result:
            return ok("图片发送成功")

        attach_result = self._attach_file_via_dialog(chat_win, path)
        if attach_result:
            wait_result = self._wait_for_message_change(chat_win, before_state, timeout=5.0)
            if wait_result:
                return ok("图片通过文件选择窗口发送成功")

            click_send_result = self._click_send_button(chat_win)
            if click_send_result:
                return self._wait_for_message_change(chat_win, before_state, timeout=5.0)
            return click_send_result

        click_send_result = self._click_send_button(chat_win)
        if click_send_result:
            return self._wait_for_message_change(chat_win, before_state, timeout=5.0)

        return fail(f"{wait_result.message}; {attach_result.message}; {click_send_result.message}")

    def _emit_monitor_status(self, status_callback: Optional[Callable[[str], None]], message: str) -> None:
        if message == self.last_monitor_status:
            return
        self.last_monitor_status = message
        print(message)
        if status_callback:
            try:
                status_callback(message)
            except Exception:
                pass

    def start_last_message_monitor(self, target_name=None, callback=None, check_interval=1, status_callback=None):
        if self.last_message_monitoring:
            return fail("最后一条消息监控已经在运行中")

        self.last_message_monitoring = True
        self.last_captured_text = ""
        self.last_message_callback = callback
        self.last_monitor_status = ""

        def monitor_loop():
            _uia_init = auto.UIAutomationInitializerInThread()
            self._emit_monitor_status(status_callback, f"已启动监听: {target_name or '未设置目标'}")

            while self.last_message_monitoring:
                try:
                    if not target_name:
                        self._emit_monitor_status(status_callback, "目标窗口名为空")
                        time.sleep(check_interval)
                        continue

                    chat_win = self.get_independent_window(target_name)
                    if not chat_win:
                        self._emit_monitor_status(status_callback, f"未找到目标窗口: {target_name}")
                        time.sleep(check_interval)
                        continue

                    msg_list = self._get_message_list(chat_win)
                    if not msg_list:
                        self._emit_monitor_status(status_callback, "已找到窗口，但未找到消息列表")
                        time.sleep(check_interval)
                        continue

                    self._emit_monitor_status(status_callback, f"监听中: {target_name}")
                    items = msg_list.GetChildren()
                    if items:
                        last_text = ""
                        for item in reversed(items):
                            text = self._message_signature(item)
                            if not text:
                                continue
                            if re.match(r"^(\d{1,2}:\d{2})$", text):
                                continue
                            if re.match(r"^(昨天|前天|星期.)\s+\d{1,2}:\d{2}$", text):
                                continue
                            if re.match(r"^\d{4}年\d{1,2}月\d{1,2}日\s+\d{1,2}:\d{2}$", text):
                                continue
                            last_text = text
                            break

                        if last_text and last_text != self.last_captured_text:
                            self.last_captured_text = last_text
                            current_time = time.strftime("%H:%M:%S")
                            print(f"[监控日志] 捕获到消息: {last_text}")
                            if self.last_message_callback:
                                try:
                                    self.last_message_callback(last_text, current_time)
                                except Exception as exc:
                                    print(f"执行回调报错: {exc}")
                except Exception as exc:
                    self._emit_monitor_status(status_callback, f"监听异常: {exc}")

                time.sleep(check_interval)

            _uia_init = None
            self._emit_monitor_status(status_callback, "监听已停止")

        monitor_thread = threading.Thread(target=monitor_loop, daemon=True)
        monitor_thread.start()
        return ok("已启动最后一条消息监控")

    def stop_last_message_monitor(self):
        self.last_message_monitoring = False
        return ok("已请求停止最后一条消息监控")
