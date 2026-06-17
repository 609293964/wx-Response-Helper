import os
import re
import threading
import time
from ctypes import windll
from dataclasses import dataclass
from typing import Callable, Optional, Tuple

import pyperclip

from windows_dpi import configure_windows_dpi_awareness


configure_windows_dpi_awareness()

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


def _safe_bounding_rect(control):
    try:
        rect = control.BoundingRectangle
        if rect and rect.right > rect.left and rect.bottom > rect.top:
            return rect
    except Exception:
        pass
    return None


def _rect_center(rect):
    return (
        int(round((rect.left + rect.right) / 2)),
        int(round((rect.top + rect.bottom) / 2)),
    )


def _click_point(x, y):
    auto.Click(int(round(x)), int(round(y)))


def _right_click_point(x, y):
    auto.RightClick(int(round(x)), int(round(y)))


def move(element):
    x, y = element.GetPosition()
    auto.SetCursorPos(int(round(x)), int(round(y)))


def click(element):
    try:
        element.Click()
        return True
    except Exception:
        pass

    try:
        x, y = element.GetPosition()
        _click_point(x, y)
        return True
    except Exception:
        return False


def click_control_center(control):
    rect = _safe_bounding_rect(control)
    if not rect:
        return False
    _click_point(*_rect_center(rect))
    return True


def right_click_control_center(control):
    rect = _safe_bounding_rect(control)
    if not rect:
        return False
    _right_click_point(*_rect_center(rect))
    return True


def click_chat_input_fallback(chat_win):
    rect = _safe_bounding_rect(chat_win)
    if not rect:
        return False

    width = rect.right - rect.left
    height = rect.bottom - rect.top
    bottom_offset = max(56, min(140, int(round(height * 0.1))))
    _click_point(rect.left + width / 2, rect.bottom - bottom_offset)
    return True


class WeChat:
    def __init__(self, locale="zh-CN"):
        assert locale in WeChatLocale.getSupportedLocales()
        self.lc = WeChatLocale(locale)
        self.last_message_monitoring = False
        self.last_captured_text = ""
        self.last_captured_direction = ""
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

    def _find_control(self, root, predicate, max_depth=25):
        try:
            for control, _depth in auto.WalkControl(root, includeTop=False, maxDepth=max_depth):
                try:
                    if predicate(control):
                        return control
                except Exception:
                    continue
        except Exception:
            pass
        return None

    def _wait_for_control(self, getter, timeout=5.0, interval=0.2):
        deadline = time.time() + timeout
        while time.time() < deadline:
            control = getter()
            if control is not None:
                return control
            time.sleep(interval)
        return None

    def _get_wechat_top_level_controls(self, process_ids=None):
        controls = []
        process_ids = {int(pid) for pid in (process_ids or []) if pid}
        try:
            root = auto.GetRootControl()
            for control in root.GetChildren():
                class_name = str(getattr(control, "ClassName", "") or "")
                name = str(getattr(control, "Name", "") or "")
                process_id = int(getattr(control, "ProcessId", 0) or 0)
                if (
                    process_id in process_ids
                    or class_name.startswith("mmui::")
                    or class_name in {"WeChatLoginWndForPC", "WeChatMainWndForPC"}
                    or name in {"微信", "WeChat"}
                ):
                    controls.append(control)
        except Exception:
            pass
        return controls

    def _collect_login_controls(self, process_ids=None):
        details = []
        controls = []
        for window in self._get_wechat_top_level_controls(process_ids):
            try:
                iterator = auto.WalkControl(window, includeTop=True, maxDepth=20)
                for control, _depth in iterator:
                    try:
                        detail = {
                            "name": str(getattr(control, "Name", "") or "").strip(),
                            "class_name": str(getattr(control, "ClassName", "") or "").strip(),
                            "automation_id": str(
                                getattr(control, "AutomationId", "") or ""
                            ).strip(),
                            "control_type": str(
                                getattr(control, "ControlTypeName", "") or ""
                            ).strip(),
                        }
                    except Exception:
                        continue
                    if any(detail.values()):
                        details.append(detail)
                        controls.append(control)
                    if len(details) >= 160:
                        return details, controls
            except Exception:
                continue
        return details, controls

    def _find_main_window_with_search(self, process_ids=None):
        for window in self._get_wechat_top_level_controls(process_ids):
            try:
                if str(getattr(window, "ClassName", "") or "") != "mmui::MainWindow":
                    continue
                search_input = window.EditControl(
                    ClassName="mmui::XValidatorTextEdit",
                    searchDepth=20,
                )
                if search_input.Exists(0.3, 0):
                    return window, search_input
            except Exception:
                continue
        return None, None

    @staticmethod
    def _is_login_button_detail(detail):
        login_button_names = {"登录", "进入微信", "重新登录", "登 录", "log in", "login"}
        name = detail.get("name", "").strip().lower()
        class_name = detail.get("class_name", "").lower()
        control_type = detail.get("control_type", "")
        return name in login_button_names and (
            control_type in {"ButtonControl", "HyperlinkControl"}
            or "button" in class_name
        )

    @staticmethod
    def _classify_login_controls(details):
        searchable = [
            " ".join(
                (
                    detail.get("name", ""),
                    detail.get("class_name", ""),
                    detail.get("automation_id", ""),
                )
            ).lower()
            for detail in details
        ]

        mobile_confirmation_hints = (
            "请在手机",
            "需在手机上完成登录",
            "在手机上完成登录",
            "需在手机上确认",
            "手机上确认",
            "手机确认登录",
            "等待手机确认",
            "已发送登录请求",
            "confirm on phone",
        )
        if any(
            hint.lower() in text
            for text in searchable
            for hint in mobile_confirmation_hints
        ):
            return "waiting_mobile", "微信正在等待手机确认登录"

        if any(WeChat._is_login_button_detail(detail) for detail in details):
            return "login_ready", "已找到微信登录按钮"

        qr_hints = (
            "扫描二维码",
            "扫码登录",
            "二维码登录",
            "使用微信扫描",
            "使用手机微信扫描",
            "qrcode",
            "qr_code",
        )
        if any(hint in text for text in searchable for hint in qr_hints):
            return "qr_required", "微信当前需要扫码登录"

        return "unknown", "暂未识别微信登录界面"

    def get_login_state(self, process_ids=None) -> Tuple[str, str]:
        try:
            main_window, search_input = self._find_main_window_with_search(process_ids)
            if main_window is not None and search_input is not None:
                return "logged_in", "已检测到微信主界面，登录完成"

            for window in self._get_wechat_top_level_controls(process_ids):
                class_name = str(getattr(window, "ClassName", "") or "")
                if class_name == "mmui::ChatSingleWindow":
                    return "logged_in", "已检测到微信独立聊天窗口，登录完成"

            details, _controls = self._collect_login_controls(process_ids)
            if not details:
                return "starting", "等待微信登录界面或主界面出现"
            return self._classify_login_controls(details)
        except Exception as exc:
            return "unknown", f"检查微信登录状态失败: {exc}"

    def get_login_ui_summary(self, process_ids=None) -> str:
        details, _controls = self._collect_login_controls(process_ids)
        windows = []
        buttons = []
        texts = []
        for detail in details:
            label = detail["name"] or detail["automation_id"]
            if not label:
                continue
            if detail["control_type"] == "WindowControl":
                windows.append(f"{label}<{detail['class_name']}>")
            elif detail["control_type"] == "ButtonControl":
                buttons.append(label)
            elif detail["name"]:
                texts.append(detail["name"])

        def compact(values, limit):
            unique = []
            for value in values:
                if value not in unique:
                    unique.append(value)
                if len(unique) >= limit:
                    break
            return ", ".join(unique) if unique else "-"

        return (
            f"窗口: {compact(windows, 6)}; "
            f"按钮: {compact(buttons, 10)}; "
            f"文本: {compact(texts, 10)}"
        )

    def _find_login_button(self, process_ids=None):
        details, controls = self._collect_login_controls(process_ids)
        for detail, control in zip(details, controls):
            if self._is_login_button_detail(detail):
                try:
                    parent_window = control.GetTopLevelControl()
                except Exception:
                    parent_window = None
                return parent_window, detail, control
        return None, None, None

    def _wait_for_login_button_effect(self, process_ids=None, timeout=4.0):
        deadline = time.time() + timeout
        last_state = ("login_ready", "登录按钮仍然可见")
        while time.time() < deadline:
            last_state = self.get_login_state(process_ids)
            if last_state[0] != "login_ready":
                return last_state
            time.sleep(0.25)
        return last_state

    def click_login_button(self, process_ids=None) -> AutomationResult:
        methods = ("invoke", "control_click", "coordinate_click")
        errors = []
        for method in methods:
            window, detail, control = self._find_login_button(process_ids)
            if control is None:
                state, message = self.get_login_state(process_ids)
                if state != "login_ready":
                    return ok(f"微信登录界面已变化: {message}")
                errors.append("未找到可点击的微信登录按钮")
                continue

            try:
                if window is not None:
                    try:
                        window.SetActive()
                    except Exception:
                        pass
                    try:
                        hwnd = int(getattr(window, "NativeWindowHandle", 0) or 0)
                        if hwnd:
                            windll.user32.ShowWindow(hwnd, 9)
                            windll.user32.SetForegroundWindow(hwnd)
                    except Exception:
                        pass
                try:
                    control.SetFocus()
                except Exception:
                    pass
                time.sleep(0.2)

                if method == "invoke":
                    control.GetInvokePattern().Invoke()
                elif method == "control_click":
                    control.Click()
                else:
                    if not click_control_center(control):
                        raise RuntimeError("invalid control rectangle")

                state, message = self._wait_for_login_button_effect(
                    process_ids,
                    timeout=4.0,
                )
                if state != "login_ready":
                    return ok(
                        f"已点击微信登录按钮: {detail.get('name')}；"
                        f"界面状态: {message}"
                    )
                errors.append(f"{method} 后登录按钮仍然可见")
            except Exception as exc:
                errors.append(f"{method} 失败: {exc}")

        if errors:
            return fail("微信登录按钮点击未生效: " + "; ".join(errors))
        return fail("未找到可点击的微信登录按钮")

    def _ensure_window_topmost(self, window) -> AutomationResult:
        try:
            hwnd = window.NativeWindowHandle
            if not hwnd:
                return fail("独立聊天窗口没有有效句柄")

            user32 = windll.user32
            if not user32.GetWindowLongW(hwnd, -20) & 0x8:
                pin_button = window.ButtonControl(ClassName="mmui::PinnedButton", searchDepth=8)
                if pin_button.Exists(1, 0):
                    pin_button.Click()
                    deadline = time.time() + 3
                    while time.time() < deadline:
                        if user32.GetWindowLongW(hwnd, -20) & 0x8:
                            break
                        time.sleep(0.2)

            if not user32.GetWindowLongW(hwnd, -20) & 0x8:
                hwnd_topmost = -1
                swp_no_move = 0x0002
                swp_no_size = 0x0001
                swp_show_window = 0x0040
                if not user32.SetWindowPos(
                    hwnd,
                    hwnd_topmost,
                    0,
                    0,
                    0,
                    0,
                    swp_no_move | swp_no_size | swp_show_window,
                ):
                    return fail("独立聊天窗口已打开，但设置置顶失败")

            user32.ShowWindow(hwnd, 9)
            user32.SetForegroundWindow(hwnd)
            window.SetFocus()
            return ok(f"已打开并置顶独立聊天窗口: {window.Name}")
        except Exception as exc:
            return fail(f"设置独立聊天窗口置顶失败: {exc}")

    def open_independent_chat_window(self, target_name: str) -> AutomationResult:
        target_name = str(target_name or "").strip()
        if not target_name:
            return fail("目标联系人名称为空")

        existing = auto.WindowControl(
            Name=target_name,
            ClassName="mmui::ChatSingleWindow",
            searchDepth=1,
        )
        if existing.Exists(0.5, 0):
            return self._ensure_window_topmost(existing)

        main_window = None
        search_input = None
        deadline = time.time() + 2
        while time.time() < deadline:
            main_window, search_input = self._find_main_window_with_search()
            if main_window is not None and search_input is not None:
                break
            time.sleep(0.2)
        if main_window is None or search_input is None:
            return fail("未找到微信主窗口，请先启动微信")

        try:
            main_window.SetActive()
            time.sleep(0.3)
            if not search_input.Exists(1, 0):
                main_window, search_input = self._find_main_window_with_search()
            if main_window is None or search_input is None:
                return fail("未找到微信搜索框，请确认微信 4.1 主窗口已正常显示")

            search_input.Click()
            auto.SendKeys("{Ctrl}a")
            auto.PressKey(auto.Keys.VK_BACK)
            time.sleep(0.2)
            auto.SendKeys(target_name)

            def get_exact_search_result():
                exact_item = main_window.ListItemControl(
                    AutomationId=f"search_item_{target_name}",
                    searchDepth=12,
                )
                if exact_item.Exists(0.2, 0):
                    return exact_item
                return self._find_control(
                    main_window,
                    lambda control: (
                        control.ControlTypeName == "ListItemControl"
                        and control.ClassName == "mmui::SearchContentCellView"
                        and control.Name.strip() == target_name
                    ),
                    max_depth=12,
                )

            result_item = self._wait_for_control(get_exact_search_result, timeout=8)
            if result_item is None:
                auto.SendKeys(" ")
                auto.PressKey(auto.Keys.VK_BACK)
                result_item = self._wait_for_control(get_exact_search_result, timeout=5)

            if result_item is None or not result_item.Exists(0.2, 0):
                try:
                    search_value = search_input.GetValuePattern().Value
                except Exception:
                    search_value = ""
                return fail(
                    f"未找到完全匹配的联系人: {target_name}"
                    f"（搜索框内容: {search_value or '-'}）"
                )

            result_item.Click()
            time.sleep(1)

            def get_session_item():
                exact_item = main_window.ListItemControl(
                    AutomationId=f"session_item_{target_name}",
                    searchDepth=25,
                )
                if exact_item.Exists(0.2, 0):
                    return exact_item
                return self._find_control(
                    main_window,
                    lambda control: (
                        control.ControlTypeName == "ListItemControl"
                        and control.ClassName == "mmui::ChatSessionCell"
                        and control.Name.strip().startswith(target_name)
                    ),
                    max_depth=25,
                )

            session_item = self._wait_for_control(get_session_item, timeout=5)
            if session_item is None:
                return fail(f"已打开联系人 '{target_name}'，但未找到对应会话项")

            if not right_click_control_center(session_item):
                return fail("Unable to right-click the session item because its rectangle is invalid")

            def get_independent_menu_item():
                return self._find_control(
                    main_window,
                    lambda control: (
                        control.ControlTypeName == "MenuItemControl"
                        and control.ClassName == "mmui::XMenuView"
                        and control.Name == "独立窗口显示"
                    ),
                    max_depth=10,
                )

            independent_item = self._wait_for_control(get_independent_menu_item, timeout=3)
            if independent_item is None:
                return fail("未找到“独立窗口显示”菜单项")

            independent_item.Click()

            def get_independent_window():
                chat_window = auto.WindowControl(
                    Name=target_name,
                    ClassName="mmui::ChatSingleWindow",
                    searchDepth=1,
                )
                return chat_window if chat_window.Exists(0.2, 0) else None

            independent_window = self._wait_for_control(get_independent_window, timeout=5)
            if independent_window is None:
                return fail(f"已执行独立窗口命令，但未检测到窗口: {target_name}")
            return self._ensure_window_topmost(independent_window)
        except Exception as exc:
            return fail(f"查找并打开联系人失败: {exc}")

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

            rect = _safe_bounding_rect(chat_win)
            if not rect:
                return None, fail(f"未能定位 '{target_name}' 的输入区域")

            click_chat_input_fallback(chat_win)
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

    def _is_message_time_text(self, text: str) -> bool:
        if re.match(r"^(\d{1,2}:\d{2})$", text):
            return True
        if re.match(r"^(昨天|前天|星期.)\s+\d{1,2}:\d{2}$", text):
            return True
        if re.match(r"^\d{4}年\d{1,2}月\d{1,2}日\s+\d{1,2}:\d{2}$", text):
            return True
        return False

    def _safe_rect(self, control):
        try:
            rect = control.BoundingRectangle
            if rect and rect.right > rect.left and rect.bottom > rect.top:
                return rect
        except Exception:
            pass
        return None

    def _rect_center_x(self, rect) -> float:
        return (rect.left + rect.right) / 2

    def _infer_message_direction(self, item, msg_list) -> str:
        list_rect = self._safe_rect(msg_list)
        if list_rect is None:
            return "unknown"

        list_width = max(1, list_rect.right - list_rect.left)
        list_center_x = self._rect_center_x(list_rect)
        tolerance = max(24, list_width * 0.08)
        candidate_rects = []

        try:
            iterator = auto.WalkControl(item, includeTop=True, maxDepth=8)
            for control, _depth in iterator:
                try:
                    text = self._control_text(control)
                    if not text or self._is_message_time_text(text):
                        continue
                    rect = self._safe_rect(control)
                    if rect is None:
                        continue
                    rect_width = rect.right - rect.left
                    if rect_width >= list_width * 0.9:
                        continue
                    candidate_rects.append(rect)
                except Exception:
                    continue
        except Exception:
            pass

        if candidate_rects:
            rect = max(candidate_rects, key=lambda item_rect: abs(self._rect_center_x(item_rect) - list_center_x))
        else:
            rect = self._safe_rect(item)
            if rect is None or (rect.right - rect.left) >= list_width * 0.9:
                return "unknown"

        center_x = self._rect_center_x(rect)
        if center_x > list_center_x + tolerance:
            return "self"
        if center_x < list_center_x - tolerance:
            return "contact"
        return "unknown"

    def _get_last_message_snapshot(self, msg_list) -> Tuple[str, str]:
        try:
            items = msg_list.GetChildren()
        except Exception:
            return "", "unknown"

        for item in reversed(items):
            text = self._message_signature(item)
            if not text or self._is_message_time_text(text):
                continue
            direction = self._infer_message_direction(item, msg_list)
            return text, direction
        return "", "unknown"

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
        if click_chat_input_fallback(chat_win):
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
        self.last_captured_direction = ""
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
                    last_text, direction = self._get_last_message_snapshot(msg_list)
                    if last_text:
                        changed = (
                            last_text != self.last_captured_text
                            or direction != self.last_captured_direction
                        )
                        if changed:
                            self.last_captured_text = last_text
                            self.last_captured_direction = direction
                            current_time = time.strftime("%H:%M:%S")
                            print(f"[监控日志] 捕获到消息: {last_text} ({direction})")
                            if self.last_message_callback:
                                try:
                                    self.last_message_callback(last_text, current_time, direction)
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
