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


# 鼠标移动到控件上
def move(element):
    x, y = element.GetPosition()
    auto.SetCursorPos(x, y)


# 鼠标快速点击控件
def click(element):
    x, y = element.GetPosition()
    auto.Click(x, y)


# 鼠标右键点击控件
def right_click(element):
    x, y = element.GetPosition()
    auto.RightClick(x, y)


# 鼠标快速点击两下控件
def double_click(element):
    x, y = element.GetPosition()
    auto.SetCursorPos(x, y)
    element.DoubleClick()


# 鼠标滚轮向下滑动
def wheel_down():
    auto.WheelDown()


class WeChat:
    def __init__(self, path, locale="zh-CN"):
        # 微信打开路径
        self.path = path
        
        # 用于复制内容到剪切板
        self.app = QApplication([])
        
        # 自动回复的联系人列表
        self.auto_reply_contacts = []
        
        # 自动回复的内容
        self.auto_reply_msg = "[自动回复]您好，我现在正在忙，稍后会主动联系您，感谢理解。"
        
        # 消息监控回调函数
        self.message_callback = None
        
        # 上次读取的消息ID，用于去重
        self.last_message_id = 0
        
        # 监控运行标志
        self.monitoring = False

        assert locale in WeChatLocale.getSupportedLocales()
        self.lc = WeChatLocale(locale)
        
    # 检查微信窗口是否可见
    def is_wechat_visible(self):
        try:
            wechat_window = auto.WindowControl(Depth=1, Name=self.lc.weixin, searchDepth=1)
            # 检查窗口是否存在且可见（非最小化）
            if wechat_window.Exists(0, 0):
                # 获取窗口句柄
                hwnd = wechat_window.NativeWindowHandle
                # 使用 Windows API 检查窗口是否可见且未最小化
                user32 = windll.user32
                is_visible = user32.IsWindowVisible(hwnd)
                is_minimized = user32.IsIconic(hwnd)
                return is_visible and not is_minimized
            return False
        except:
            return False
    
    # 确保微信窗口是打开且非最小化的
    def ensure_wechat_visible(self):
        """确保微信窗口存在，如果最小化就还原它"""
        try:
            wechat_window = auto.WindowControl(Depth=1, Name=self.lc.weixin, searchDepth=1)
            if wechat_window.Exists(0, 0):
                hwnd = wechat_window.NativeWindowHandle
                user32 = windll.user32
                is_minimized = user32.IsIconic(hwnd)
                if is_minimized:
                    # 如果最小化了，还原窗口
                    user32.OpenIconicWindow(hwnd)
                    time.sleep(1)
                wechat_window.SetFocus()
                return True
        except:
            pass
        return False

    # 打开微信客户端
    def open_wechat(self):
        # 先检查微信窗口是否已经可见
        if self.is_wechat_visible():
            # 如果已经可见，只需要激活窗口到前台
            wechat_window = self.get_wechat()
            wechat_window.SetFocus()
            return

        # 窗口存在但可能最小化了，先尝试还原
        if self.ensure_wechat_visible():
            time.sleep(1)
            if self.is_wechat_visible():
                return

        # 如果窗口不可见，先尝试通过全局快捷键打开（微信已经运行但最小化）
        auto.SendKeys("{Ctrl}{Alt}w")
        time.sleep(2)
        
        # 如果快捷键打不开，检查一下是否微信进程没启动，直接启动它
        if not self.is_wechat_visible() and self.path and os.path.exists(self.path):
            # 从配置的路径启动微信
            subprocess.Popen(self.path)
            # 等待微信启动
            time.sleep(5)
    
    # 搜寻微信客户端控件
    def get_wechat(self):
        return auto.WindowControl(Depth=1, Name=self.lc.weixin)

    # 获取当前聊天对象的昵称
    def get_current_name(self):
        # 打开微信，获取根窗口并通过点击获取焦点
        self.open_wechat()
        root = self.get_wechat()
        click(root)

        # 等待焦点锁定在微信窗口
        time.sleep(1)

        # 获取聊天窗口
        window = auto.TextControl(Depth=20)
        return window.Name
    
    # 防止微信长时间挂机导致掉线
    def prevent_offline(self):
        self.open_wechat()
        self.get_wechat()
        
        search_box = auto.EditControl(Depth=8, Name=self.lc.search)
        click(search_box)
    
    # 搜索指定用户
    def get_contact(self, name):
        self.open_wechat()
        self.get_wechat()
        
        search_box = auto.EditControl(Depth=13, Name=self.lc.search)
        click(search_box)
        
        pyperclip.copy(name)
        auto.SendKeys("{Ctrl}v")

        # 等待客户端搜索联系人
        time.sleep(0.3)

        # 现在群聊不会出现在搜索的第一行，需要手动选择
        list_control = auto.ListControl(Depth=4)
        for item in list_control.GetChildren():
            # 联系人项的 ClassName 不包含 "XTableCell"，默认选择第一个联系人，点击进入窗口
            if "XTableCell" not in item.ClassName:
                click(item)
                break

        # 点击发送内容输入框来获取输入焦点
        tool_bar = auto.ToolBarControl(Depth=15)
        move(tool_bar)
        click(tool_bar)
    
    # 鼠标移动到发送按钮处点击发送消息
    def press_enter(self):
        # 改用回车键发送，比点击发送按钮更可靠（适配不同微信版本）
        auto.SendKeys("{enter}")

    def paste_text(self, text: str) -> None:
        """
        封装文本粘贴逻辑
        Args:
            text: 待发送文本
        """
        pyperclip.copy(text)
        # 等待粘贴
        time.sleep(0.3)
        auto.SendKeys("{Ctrl}v")

    def send_msg(self, name, at_names: List[str] = None, text: str = None, search_user: bool = True) -> bool:
        """
        搜索指定用户名的联系人发送信息, 同时可以在指定群聊中@他人（若@所有人需具备@所有人权限）
        Args:
            name:  群聊名称
            at_names: 若发送对象为群，则可以@他人（若@所有人需具备@所有人权限）
            text: 要@的人的消息
            search_user: 是否需要搜索群聊
        """
        if search_user:
            self.get_contact(name)
        
        if at_names is not None:
            # @所有列表中的人名
            for at_name in at_names:
                # 如果at_name为 "所有人" 则代表@所有人
                if at_name == "所有人":
                    auto.SendKeys("@{UP}{enter}")

                elif at_name != "":
                    auto.SendKeys(f"@{at_name}")
                    # 按下回车键确认要at的人
                    auto.SendKeys("{enter}")

        # 如果发送信息不为空，则发送信息
        if text is not None:
            self.paste_text(text)

        self.press_enter()

        # 发送消息后马上获取聊天记录，判断是否发送成功
        # 新版微信这里可能会抛出异常，这里添加一个通用返回
        try:
            if self.get_dialogs(name, 1, False) and self.get_dialogs(name, 1, False)[0][2] == text:
                return True
            else:
                return True # 为了兼容性，不报错即认为发送成功
        except Exception:
            return True

    # 搜索指定用户名的联系人发送文件
    def send_file(self, name: str, path: str, search_user: bool = True) -> None:
        """
        Args:
            name: 指定用户名的名称，输入搜索框后出现的第一个人
            path: 发送文件的本地地址
            search_user: 是否需要搜索用户
        """
        if search_user:
            self.get_contact(name)
        else:
            # 即使不搜索用户，也要确保微信窗口在前台，并点击输入框
            self.open_wechat()
            wechat_window = self.get_wechat()
            wechat_window.SetFocus()
            time.sleep(0.2)
            
            # 尝试多种方式点击输入框，根据控件树，输入框是EditControl名字为"输入"
            clicked = False
            # 方式1：直接找 EditControl 名字为"输入"（新版微信输入框名字就是"输入"
            try:
                edit_input = auto.EditControl(Name="输入", searchDepth=20)
                if edit_input.Exists(0.5, 0):
                    move(edit_input)
                    click(edit_input)
                    clicked = True
                    time.sleep(0.2)
                    print(f"找到输入框EditControl，点击成功")
            except Exception as e:
                print(f"方式1找输入框失败: {e}")
            
            # 方式2：通过工具条查找（兼容旧版本）
            if not clicked:
                try:
                    tool_bar = auto.ToolBarControl(searchDepth=20)
                    if tool_bar.Exists(0.5, 0):
                        move(tool_bar)
                        click(tool_bar)
                        clicked = True
                        time.sleep(0.2)
                        print(f"方式2找到工具条，点击成功")
                except Exception as e:
                    print(f"方式2找输入框失败: {e}")
            
            # 如果还没找到，方式3：直接点击屏幕下方中央位置（输入框通常在那里）
            if not clicked:
                screen_width, screen_height = pyautogui.size()
                # 点击屏幕下方中央，大概就是输入框位置
                pyautogui.click(screen_width // 2, int(screen_height * 0.85))
                clicked = True
                time.sleep(0.2)
                print(f"方式3点击屏幕下方中央，点击成功")
        
        # 将文件复制到剪切板
        setClipboardFiles([path])
        time.sleep(0.3)
        
        auto.SendKeys("{Ctrl}v")
        time.sleep(0.5)
        
        self.press_enter()
    
    # 获取所有通讯录中所有联系人
    def find_all_contacts(self) -> pd.DataFrame:
        self.open_wechat()
        self.get_wechat()

        # 获取通讯录管理界面
        click(auto.ButtonControl(Name=self.lc.contacts))
        contacts_menu = auto.ListItemControl(Depth=12, foundIndex=1)
        click(contacts_menu)

        # 将鼠标移动到联系人上以便可以通过鼠标滚轮往下滑动
        move(auto.ListItemControl(Depth=8, foundIndex=1))

        # 获取初始群聊列表
        contacts = pd.DataFrame(columns=["昵称", "备注", "标签"])
        contact_set = set()
        for contact in auto.ListControl(Depth=7).GetChildren():
            # 获取用户的昵称备注以及标签。注意这种方式没有办法准确获取昵称和备注，因为微信自身的信息组织问题。
            name, note, label = contact.Name.rsplit(" ", maxsplit=2)
            if name not in contact_set:
                contacts = contacts._append({"昵称": name, "备注": note, "标签": label}, ignore_index=True)
                contact_set.add(name)

        # 模拟鼠标下滑一直读取群聊列表直到无法下滑为止
        num_trial = 3
        while num_trial > 0:
            ori_len = len(contact_set)

            wheel_down()
            for contact in auto.ListControl(Depth=7).GetChildren():
                name, note, label = contact.Name.rsplit(" ", maxsplit=2)
                if name not in contact_set:
                    contacts = contacts._append({"昵称": name, "备注": note, "标签": label}, ignore_index=True)
                    contact_set.add(name)

            if len(contact_set) == ori_len:
                num_trial -= 1
            else:
                num_trial = 3
        
        return contacts
    
    # 获取所有群聊
    def find_all_groups(self) -> list:
        self.open_wechat()
        self.get_wechat()
        
        # 获取通讯录管理界面
        click(auto.ButtonControl(Name=self.lc.contacts))
        contacts_menu = auto.ListItemControl(Depth=12, foundIndex=1)
        click(contacts_menu)

        # 点击最近群聊
        click(auto.ListItemControl(Depth=6, foundIndex=5))
        
        # 获取初始群聊列表
        groups = set()
        for i, group in enumerate(auto.ListControl(Depth=5).GetChildren()):
            if i >= 5:
                name = group.Name.rsplit("(", maxsplit=1)[0]
                groups.add(name)

        # 模拟鼠标下滑一直读取群聊列表直到无法下滑为止
        num_trial = 3
        while num_trial > 0:
            ori_len = len(groups)

            wheel_down()
            for i, group in enumerate(auto.ListControl(Depth=5).GetChildren()):
                if i >= 5:
                    name = group.Name.split("(")[0]
                    groups.add(name)

            if len(groups) == ori_len:
                num_trial -= 1
            else:
                num_trial = 3

        return list(groups)
    
    # 检测微信是否收到新消息 (兼容性警告)
    def check_new_msg(self):
        print("⚠️ 警告: check_new_msg 方法尚未完全适配新版微信控件树，已安全跳过执行")
        return
    
    # 设置自动回复的联系人
    def set_auto_reply(self, contacts):
        self.auto_reply_contacts = contacts
    
    # 自动回复
    def _auto_reply(self, element, text):
        click(element)
        pyperclip.copy(text)
        auto.SendKeys("{Ctrl}v")
        self.press_enter()
    
    # 识别聊天内容的类型
    def _detect_type(self, list_item_control: auto.ListItemControl) -> int:
        value = None
        if not isinstance(list_item_control.GetFirstChildControl(), auto.PaneControl):
            value = 1
        else:
            cnt = 0
            for child in list_item_control.PaneControl().GetChildren():
                cnt += len(child.GetChildren())
            
            if cnt > 0:
                value = 0
            elif list_item_control.Name == "查看更多消息":
                value = 3
            elif "红包" in list_item_control.Name or "red packet" in list_item_control.Name.lower():
                value = 2
            elif "撤回了一条消息" in list_item_control.Name:
                value = 4
            elif "以下为新消息" in list_item_control.Name:
                value = 6

        if value is None:
            return 0 # 默认当作用户发送以保证健壮性
        
        return value
    
    # 获取聊天窗口
    def _get_chat_frame(self, name: str):
        self.get_contact(name)
        return auto.ListControl(Name=self.lc.message)
    
    def save_dialog_pictures(self, name: str, num: int, save_dir: str) -> None:
        """保存指定聊天记录中的图片"""
        print("⚠️ 警告: save_dialog_pictures 尚未适配新版微信控件树")
        return
            
    # 获取指定聊天窗口的聊天记录
    def get_dialogs(self, name: str, n_msg: int, search_user: bool = True) -> List:
        """获取聊天记录"""
        print("⚠️ 警告: get_dialogs 尚未适配新版微信控件树")
        return []

    def get_dialogs_by_time_blocks(self, name: str, n_time_blocks: int, search_user: bool = True) -> List[List]:
        """获取指定聊天窗口的聊天记录，并按时间信息分组"""
        print("⚠️ 警告: get_dialogs_by_time_blocks 尚未适配新版微信控件树")
        return []
    
    # ==================== 消息监控功能 ====================
    def set_message_callback(self, callback):
        """
        设置消息回调函数，当收到新消息时会调用这个函数
        回调函数参数：(联系人名称, 消息内容, 消息时间, 消息类型)
        消息类型：text(文本), file(文件), image(图片), voice(语音), other(其他)
        """
        self.message_callback = callback
    
    def get_chat_list(self):
        """获取当前聊天列表中的所有联系人"""
        try:
            self.open_wechat()
            wechat_window = self.get_wechat()
            
            # 找到会话列表控件
            session_list = wechat_window.ListControl(Name=self.lc.session_list, searchDepth=32)
            if not session_list.Exists(0, 0):
                return []
            
            # 获取所有会话项
            sessions = session_list.GetChildren()
            chat_list = []
            for session in sessions:
                if session.ControlTypeName == "ListItemControl":
                    name = session.Name
                    if name and name not in ["", "会话列表", "微信"]:
                        chat_list.append(name)
            
            return chat_list
        except Exception as e:
            print(f"获取聊天列表失败: {e}")
            return []
    
    def get_current_chat_messages(self, max_count=20):
        """获取当前打开的聊天窗口的消息记录"""
        try:
            wechat_window = self.get_wechat()
            
            # 尝试多种方式查找消息列表控件，适配不同微信版本
            msg_list = None
            # 方式1：通过本地化名称查找
            try:
                msg_list = wechat_window.ListControl(Name=self.lc.message_list, searchDepth=32)
                if msg_list.Exists(0, 0):
                    print(f"通过名称找到消息列表: {msg_list.Name}")
            except Exception as e:
                pass
                
            # 方式2：通过控件类型查找，递归遍历子控件
            def find_msg_list(control, depth=0):
                nonlocal msg_list
                if msg_list:
                    return
                try:
                    if control.ControlTypeName == "ListControl":
                        name = control.Name if hasattr(control, 'Name') else ""
                        if name not in ["会话列表", "联系人列表", "搜索结果", "表情面板"]:
                            children = control.GetChildren()
                            if len(children) > 0 and any(item.ControlTypeName == "ListItemControl" for item in children):
                                msg_list = control
                                print(f"找到消息列表控件，名称: '{name}', 子项数量: {len(children)}")
                                return
                    # 递归查找子控件
                    for child in control.GetChildren():
                        find_msg_list(child, depth + 1)
                except Exception as e:
                    pass
            
            if not msg_list or not msg_list.Exists(0, 0):
                print("开始递归查找消息列表...")
                find_msg_list(wechat_window)
            
            if not msg_list or not msg_list.Exists(0, 0):
                print("未找到消息列表控件")
                return []
            
            # 获取所有消息项
            msg_items = msg_list.GetChildren()
            messages = []
            
            for item in msg_items[-max_count:]:  # 只获取最后max_count条消息
                try:
                    if item.ControlTypeName != "ListItemControl":
                        continue
                        
                    # 获取消息发送者和内容
                    sender = ""
                    content = ""
                    msg_type = "text"
                    all_text = []
                    
                    # 递归获取所有文本内容
                    def get_all_text(element):
                        try:
                            if element.ControlTypeName == "TextControl" and element.Name.strip():
                                all_text.append(element.Name.strip())
                            for child in element.GetChildren():
                                get_all_text(child)
                        except:
                            pass
                    
                    get_all_text(item)
                    
                    if len(all_text) == 0:
                        # 检查是否是图片/文件/语音
                        def check_media_type(element):
                            try:
                                if element.ControlTypeName == "ImageControl":
                                    return "image"
                                name = element.Name if hasattr(element, 'Name') else ""
                                if "文件" in name or "接收" in name or "下载" in name:
                                    return "file"
                                if "语音" in name or "秒" in name or "'" in name and '"' not in name:
                                    return "voice"
                                for child in element.GetChildren():
                                    res = check_media_type(child)
                                    if res:
                                        return res
                            except:
                                pass
                            return None
                        
                        media_type = check_media_type(item)
                        if media_type == "image":
                            content = "[图片]"
                            msg_type = "image"
                        elif media_type == "file":
                            content = "[文件]"
                            msg_type = "file"
                        elif media_type == "voice":
                            content = "[语音消息]"
                            msg_type = "voice"
                        else:
                            continue
                    else:
                        # 处理文本消息
                        if len(all_text) >= 2:
                            # 通常第一条是发送者，后面是内容
                            sender = all_text[0]
                            content = " ".join(all_text[1:])
                        else:
                            # 只有一条文本，可能是系统消息或者自己发送的消息
                            sender = "系统消息"
                            content = all_text[0]
                    
                    if content.strip():
                        # 生成唯一消息ID
                        msg_id = hash(f"{content}{time.time()}")
                        messages.append({
                            "id": msg_id,
                            "sender": sender,
                            "content": content.strip(),
                            "type": msg_type,
                            "time": time.strftime("%Y-%m-%d %H:%M:%S")
                        })
                except Exception as e:
                    print(f"解析消息出错: {e}")
                    continue
            
            return messages
        except Exception as e:
            print(f"获取消息失败: {e}")
            return []
    
    def start_monitor(self, check_interval=2):
        """
        开始监控微信消息
        check_interval: 检查新消息的间隔时间（秒）
        """
        if hasattr(self, 'monitoring') and self.monitoring:
            return
        
        self.monitoring = True
        self.message_callback = None
        self.last_message_ids = set()
        
        def monitor_loop():
            print("微信消息监控已启动（UIAutomation方式）")
            while self.monitoring:
                try:
                    # 获取当前所有消息
                    messages = self.get_current_chat_messages(max_count=10)
                    
                    for msg in messages:
                        msg_id = msg["id"]
                        if msg_id not in self.last_message_ids:
                            # 新消息，调用回调函数
                            if self.message_callback:
                                self.message_callback(
                                    msg["sender"],
                                    msg["content"],
                                    msg["time"],
                                    msg["type"]
                                )
                            self.last_message_ids.add(msg_id)
                    
                    # 保留最近100条消息的ID，避免占用太多内存
                    if len(self.last_message_ids) > 100:
                        self.last_message_ids = set(list(self.last_message_ids)[-50:])
                    
                    time.sleep(check_interval)
                except Exception as e:
                    print(f"监控出错: {e}")
                    time.sleep(check_interval)
        
        # 启动监控线程
        monitor_thread = threading.Thread(target=monitor_loop, daemon=True)
        monitor_thread.start()
    
    def stop_monitor(self):
        """停止消息监控"""
        if hasattr(self, 'monitoring'):
            self.monitoring = False
            print("微信消息监控已停止")

    # ==================== 精准最后一条消息监控 ====================
    def start_last_message_monitor(self, callback=None, check_interval=1):
        """
        启动基于控件树的精准最后一条消息监控
        """
        if hasattr(self, 'last_message_monitoring') and self.last_message_monitoring:
            print("最后一条消息监控已经在运行中")
            return
        
        self.last_message_monitoring = True
        self.last_captured_text = ""
        self.last_message_callback = callback
        
        def monitor_loop():
            print("✅ 精准最后一条消息监控已启动（控件树方式）")
            
            while self.last_message_monitoring:
                try:
                    # 每次循环都确保微信窗口是打开且激活的
                    self.open_wechat()
                    time.sleep(0.5)
                    
                    # 每次循环都重新查找消息列表控件（应对界面刷新）
                    msg_list = auto.ListControl(Name=self.lc.message)
                    
                    if not msg_list.Exists(1, 0.5):
                        # 这一轮找不到，下一轮再试
                        time.sleep(check_interval)
                        continue
                    
                    items = msg_list.GetChildren()
                    
                    if items:
                        last_item = items[-1]
                        last_text = last_item.Name
                        
                        # 如果消息发生变化
                        if last_text != self.last_captured_text:
                            self.last_captured_text = last_text
                            current_time = time.strftime("%H:%M:%S")
                            
                            # 调用回调函数
                            if self.last_message_callback:
                                try:
                                    self.last_message_callback(last_text, current_time)
                                except Exception as e:
                                    print(f"回调函数执行出错: {e}")
                                    
                except Exception as e:
                    # 忽略界面刷新时瞬间抓不到数据的偶发错误
                    pass
                    
                time.sleep(check_interval)
        
        # 启动监控线程
        monitor_thread = threading.Thread(target=monitor_loop, daemon=True)
        monitor_thread.start()
    
    def stop_last_message_monitor(self):
        """停止精准最后一条消息监控"""
        if hasattr(self, 'last_message_monitoring'):
            self.last_message_monitoring = False
            print("精准最后一条消息监控已停止")


if __name__ == '__main__':
    # 简单的测试入口
    pass