import time
import uiautomation as auto
import subprocess
import pyperclip
import os

from ctypes import *
from PyQt5.QtWidgets import QApplication
from typing import List

from wechat_locale import WeChatLocale

# 鼠标移动到控件上
def move(element):
    x, y = element.GetPosition()
    auto.SetCursorPos(x, y)

# 鼠标快速点击控件
def click(element):
    x, y = element.GetPosition()
    auto.Click(x, y)

# 鼠标滚轮向下滑动
def wheel_down():
    auto.WheelDown()

class WeChat:
    def __init__(self, path="", locale="zh-CN"):
        self.path = path
        
        # 【修复】防止 QApplication 重复实例化导致程序崩溃
        self.app = QApplication.instance() or QApplication([])
        
        # 自动回复的联系人列表与内容
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
                if user32.IsIconic(hwnd):
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
        search_box = auto.EditControl(Depth=8, Name=self.lc.search)
        click(search_box)
    
    def get_contact(self, name):
        self.open_wechat()
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

    def send_msg(self, name, at_names: List[str] = None, text: str = None, search_user: bool = True) -> bool:
        if search_user:
            self.get_contact(name)
        
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

    def send_file(self, name: str, path: str, search_user: bool = True) -> None:
        if search_user:
            self.get_contact(name)
        else:
            self.open_wechat()
            wechat_window = self.get_wechat()
            wechat_window.SetFocus()
            time.sleep(0.2)
            
            clicked = False
            try:
                edit_input = auto.EditControl(Name="输入", searchDepth=20)
                if edit_input.Exists(0.5, 0):
                    move(edit_input)
                    click(edit_input)
                    clicked = True
                    time.sleep(0.2)
            except Exception:
                pass
            
            if not clicked:
                try:
                    tool_bar = auto.ToolBarControl(searchDepth=20)
                    if tool_bar.Exists(0.5, 0):
                        move(tool_bar)
                        click(tool_bar)
                        clicked = True
                        time.sleep(0.2)
                except Exception:
                    pass
            
            if not clicked:
                import pyautogui
                screen_width, screen_height = pyautogui.size()
                pyautogui.click(screen_width // 2, int(screen_height * 0.85))
                time.sleep(0.2)
        
        from clipboard import setClipboardFiles
        setClipboardFiles([path])
        time.sleep(0.3)
        
        auto.SendKeys("{Ctrl}v")
        time.sleep(0.5)
        self.press_enter()
    
    # ======== 废弃方法精简（直接抛出异常，清除无用的死代码） ========
    def check_new_msg(self):
        raise NotImplementedError("该方法尚未适配新版微信")

    def save_dialog_pictures(self, name: str, num: int, save_dir: str) -> None:
        raise NotImplementedError("该方法尚未适配新版微信")
            
    def get_dialogs(self, name: str, n_msg: int, search_user: bool = True) -> List:
        raise NotImplementedError("该方法尚未适配新版微信")

    def get_dialogs_by_time_blocks(self, name: str, n_time_blocks: int, search_user: bool = True) -> List[List]:
        raise NotImplementedError("该方法尚未适配新版微信")
    
    # ==================== 精准最后一条消息监控（已优化不抢焦点） ====================
    def start_last_message_monitor(self, callback=None, check_interval=1):
        if hasattr(self, 'last_message_monitoring') and self.last_message_monitoring:
            print("最后一条消息监控已经在运行中")
            return
        
        self.last_message_monitoring = True
        self.last_captured_text = ""
        self.last_message_callback = callback
        
        import threading
        
        def monitor_loop():
            print("✅ 精准后台静默监控已启动（不再抢占你的鼠标和窗口焦点）")
            
            while self.last_message_monitoring:
                try:
                    # 【核心修复】静默查找微信窗口，不存在就等待，坚决不调用 open_wechat 或 SetFocus
                    wechat_window = auto.WindowControl(Depth=1, Name=self.lc.weixin)
                    if not wechat_window.Exists(0, 0):
                        time.sleep(check_interval)
                        continue
                    
                    # 仅在后台读取控件树，不影响用户前台操作
                    msg_list = wechat_window.ListControl(Name=self.lc.message)
                    
                    if not msg_list.Exists(1, 0.5):
                        time.sleep(check_interval)
                        continue
                    
                    items = msg_list.GetChildren()
                    if items:
                        last_item = items[-1]
                        last_text = last_item.Name
                        
                        if last_text != self.last_captured_text:
                            self.last_captured_text = last_text
                            current_time = time.strftime("%H:%M:%S")
                            print(f"\n[{current_time}] 最后一条消息已更新: '{last_text}'")
                            
                            if self.last_message_callback:
                                try:
                                    self.last_message_callback(last_text, current_time)
                                except Exception as e:
                                    print(f"回调执行出错: {e}")
                except Exception:
                    # 忽略界面刷新时抓不到数据的偶发错误
                    pass
                    
                time.sleep(check_interval)
        
        monitor_thread = threading.Thread(target=monitor_loop, daemon=True)
        monitor_thread.start()
    
    def stop_last_message_monitor(self):
        if hasattr(self, 'last_message_monitoring'):
            self.last_message_monitoring = False
            print("精准最后一条消息监控已停止")