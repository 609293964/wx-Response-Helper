import time
import uiautomation as auto
import subprocess
import numpy as np
import pandas as pd
import pyperclip
import os
import pyautogui
import threading
import re

from ctypes import *
from PIL import ImageGrab
from clipboard import setClipboardFiles
from PyQt5.QtWidgets import QApplication
from PyQt5.QtCore import QMimeData, QUrl
from typing import List

from wechat_locale import WeChatLocale

def move(element):
    x, y = element.GetPosition()
    auto.SetCursorPos(x, y)

def click(element):
    x, y = element.GetPosition()
    auto.Click(x, y)

def right_click(element):
    x, y = element.GetPosition()
    auto.RightClick(x, y)

def double_click(element):
    x, y = element.GetPosition()
    auto.SetCursorPos(x, y)
    element.DoubleClick()

def wheel_down():
    auto.WheelDown()

class WeChat:
    def __init__(self, path, locale="zh-CN"):
        self.path = path
        self.app = QApplication([])
        self.auto_reply_contacts = []
        self.auto_reply_msg = "[自动回复]您好，我现在正在忙，稍后会主动联系您，感谢理解。"
        self.message_callback = None
        self.last_message_id = 0
        self.monitoring = False

        assert locale in WeChatLocale.getSupportedLocales()
        self.lc = WeChatLocale(locale)
        
    def is_wechat_visible(self):
        try:
            wechat_window = auto.WindowControl(Depth=1, Name=self.lc.weixin, searchDepth=1)
            if wechat_window.Exists(0, 0):
                hwnd = wechat_window.NativeWindowHandle
                user32 = windll.user32
                is_visible = user32.IsWindowVisible(hwnd)
                is_minimized = user32.IsIconic(hwnd)
                return is_visible and not is_minimized
            return False
        except:
            return False
    
    def ensure_wechat_visible(self):
        try:
            wechat_window = auto.WindowControl(Depth=1, Name=self.lc.weixin, searchDepth=1)
            if wechat_window.Exists(0, 0):
                hwnd = wechat_window.NativeWindowHandle
                user32 = windll.user32
                is_minimized = user32.IsIconic(hwnd)
                if is_minimized:
                    user32.OpenIconicWindow(hwnd)
                    time.sleep(1)
                wechat_window.SetFocus()
                return True
        except:
            pass
        return False

    def open_wechat(self):
        if self.is_wechat_visible():
            wechat_window = self.get_wechat()
            wechat_window.SetFocus()
            return
        if self.ensure_wechat_visible():
            time.sleep(1)
            if self.is_wechat_visible():
                return
        auto.SendKeys("{Ctrl}{Alt}w")
        time.sleep(2)
        if not self.is_wechat_visible() and self.path and os.path.exists(self.path):
            subprocess.Popen(self.path)
            time.sleep(5)
    
    def get_wechat(self):
        return auto.WindowControl(Depth=1, Name=self.lc.weixin)

    def get_current_name(self):
        self.open_wechat()
        root = self.get_wechat()
        click(root)
        time.sleep(1)
        window = auto.TextControl(Depth=20)
        return window.Name
    
    def prevent_offline(self):
        self.open_wechat()
        self.get_wechat()
        search_box = auto.EditControl(Depth=8, Name=self.lc.search)
        click(search_box)
    
    def get_contact(self, name):
        self.open_wechat()
        self.get_wechat()
        search_box = auto.EditControl(Depth=13, Name=self.lc.search)
        click(search_box)
        pyperclip.copy(name)
        auto.SendKeys("{Ctrl}v")
        time.sleep(0.3)
        list_control = auto.ListControl(Depth=4)
        for item in list_control.GetChildren():
            if "XTableCell" not in item.ClassName:
                click(item)
                break
        tool_bar = auto.ToolBarControl(Depth=15)
        move(tool_bar)
        click(tool_bar)
    
    def press_enter(self):
        auto.SendKeys("{enter}")

    def paste_text(self, text: str) -> None:
        pyperclip.copy(text)
        time.sleep(0.3)
        auto.SendKeys("{Ctrl}v")

    def get_independent_window(self, target_name):
        win = auto.WindowControl(Name=target_name, searchDepth=1)
        if win.Exists(0.2, 0):
            return win
        return None

    # ==================== 发送纯文本(新增对独立窗口的支持) ====================
    def send_msg(self, name, at_names: List[str] = None, text: str = None, search_user: bool = True) -> bool:
        if search_user:
            self.get_contact(name)
        else:
            # 支持独立窗口的抓取
            chat_win = self.get_independent_window(name)
            if chat_win:
                chat_win.SetFocus()
                time.sleep(0.2)
                
                edit_input = chat_win.EditControl(Name="输入")
                
                # 记住当前鼠标位置，防止乱跳
                current_mouse_pos = auto.GetCursorPos()
                try:
                    if edit_input.Exists(0.5, 0):
                        move(edit_input)
                        click(edit_input)
                        time.sleep(0.2)
                    else:
                        rect = chat_win.BoundingRectangle
                        if rect:
                            click_x = rect.left + (rect.right - rect.left) // 2
                            click_y = rect.bottom - 60
                            pyautogui.click(click_x, click_y)
                            time.sleep(0.2)
                finally:
                    # 恢复鼠标位置
                    auto.SetCursorPos(current_mouse_pos[0], current_mouse_pos[1])
            else:
                print(f"发送失败：找不到名为 '{name}' 的独立窗口。请确认窗口是否已拖出！")
                return False

        if at_names is not None:
            for at_name in at_names:
                if at_name == "所有人":
                    auto.SendKeys("@{UP}{enter}")
                elif at_name != "":
                    auto.SendKeys(f"@{at_name}")
                    auto.SendKeys("{enter}")
                    
        if text is not None:
            self.paste_text(text)
            
        self.press_enter()
        return True

    # ==================== 发送文件 ====================
    def send_file(self, name: str, path: str, search_user: bool = True) -> None:
        if search_user:
            self.get_contact(name)
        else:
            chat_win = self.get_independent_window(name)
            if chat_win:
                chat_win.SetFocus()
                time.sleep(0.2)
                
                edit_input = chat_win.EditControl(Name="输入")
                
                # 记住当前鼠标位置，防止乱跳
                current_mouse_pos = auto.GetCursorPos()
                try:
                    if edit_input.Exists(0.5, 0):
                        move(edit_input)
                        click(edit_input)
                        time.sleep(0.2)
                    else:
                        rect = chat_win.BoundingRectangle
                        if rect:
                            click_x = rect.left + (rect.right - rect.left) // 2
                            click_y = rect.bottom - 60
                            pyautogui.click(click_x, click_y)
                            time.sleep(0.2)
                finally:
                    # 恢复鼠标位置
                    auto.SetCursorPos(current_mouse_pos[0], current_mouse_pos[1])
            else:
                print(f"发送失败：找不到名为 '{name}' 的独立窗口。请确认窗口是否已拖出！")
                return
        
        setClipboardFiles([path])
        time.sleep(0.3)
        auto.SendKeys("{Ctrl}v")
        time.sleep(0.5)
        self.press_enter()
    
    # ... 其他兼容的方法 ...
    def find_all_contacts(self) -> pd.DataFrame: pass
    def find_all_groups(self) -> list: pass
    def check_new_msg(self): pass
    def set_auto_reply(self, contacts): self.auto_reply_contacts = contacts
    def _auto_reply(self, element, text): pass
    def _detect_type(self, list_item_control: auto.ListItemControl) -> int: return 0 
    def _get_chat_frame(self, name: str): pass
    def save_dialog_pictures(self, name: str, num: int, save_dir: str) -> None: pass
    def get_dialogs(self, name: str, n_msg: int, search_user: bool = True) -> List: return []
    def get_dialogs_by_time_blocks(self, name: str, n_time_blocks: int, search_user: bool = True) -> List[List]: return []
    def set_message_callback(self, callback): self.message_callback = callback
    def get_chat_list(self): return []
    def get_current_chat_messages(self, max_count=20): return []
    def start_monitor(self, check_interval=2): pass
    def stop_monitor(self): pass

    def start_last_message_monitor(self, target_name=None, callback=None, check_interval=1):
        if hasattr(self, 'last_message_monitoring') and self.last_message_monitoring:
            print("最后一条消息监控已经在运行中")
            return
        
        self.last_message_monitoring = True
        self.last_captured_text = ""
        self.last_message_callback = callback
        
        def monitor_loop():
            _com_init = auto.UIAutomationInitializerInThread()
            print(f"✅ 【纯独立窗口模式】已启动！目标对象锁定为: [{target_name}]")
            
            while self.last_message_monitoring:
                try:
                    if not target_name:
                        time.sleep(check_interval)
                        continue

                    chat_win = self.get_independent_window(target_name)
                    if not chat_win:
                        time.sleep(check_interval)
                        continue
                    
                    msg_list = chat_win.ListControl(Name=self.lc.message)
                    if not msg_list.Exists(0.1, 0):
                        time.sleep(check_interval)
                        continue
                    
                    items = msg_list.GetChildren()
                    if items:
                        last_text = ""
                        # 倒序遍历，跳过空白占位符、时间戳或格式异常的节点，找到真正的最后一条消息
                        for i in range(len(items)-1, -1, -1):
                            item = items[i]
                            text = item.Name
                            # 尝试对结构解析优化：如果Name为空，但内部有TextControl，尝试获取TextControl的内容
                            if not text or len(text.strip()) == 0:
                                try:
                                    texts = item.GetChildren()
                                    for t in texts:
                                        if getattr(t, 'ControlTypeName', '') == 'TextControl' and t.Name:
                                            text = t.Name
                                            break
                                except:
                                    pass
                            
                            if not text or not str(text).strip():
                                continue
                            
                            text = str(text).strip()
                            
                            # 跳过时间戳类型的项（如 "19:12", "12:30", "昨天 12:30" 等）
                            # 微信消息列表中时间标签和消息混在一起，必须过滤
                            if re.match(r'^(\d{1,2}:\d{2})$', text):
                                continue
                            if re.match(r'^(昨天|前天|星期.)\s+\d{1,2}:\d{2}$', text):
                                continue
                            if re.match(r'^\d{4}年\d{1,2}月\d{1,2}日\s+\d{1,2}:\d{2}$', text):
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
                                except Exception as e:
                                    print(f"执行回调报错: {e}")
                except Exception as e:
                    # 获取频繁时避免打印太多，实际可记录下来
                    # print(f"监控抓取异常: {e}")
                    pass
                
                time.sleep(check_interval)
        
        monitor_thread = threading.Thread(target=monitor_loop, daemon=True)
        monitor_thread.start()
    
    def stop_last_message_monitor(self):
        if hasattr(self, 'last_message_monitoring'):
            self.last_message_monitoring = False
            print("精准最后一条消息监控已停止")

if __name__ == '__main__':
    pass