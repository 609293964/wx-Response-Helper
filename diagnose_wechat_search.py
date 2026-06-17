import csv
import ctypes
import io
import json
import subprocess
import sys
import time
from pathlib import Path

from windows_dpi import configure_windows_dpi_awareness


configure_windows_dpi_awareness()

import uiautomation as auto
from ui_auto_wechat import WeChat


OUTPUT_PATH = Path(__file__).resolve().parent / "diagnose_wechat_search.json"


def get_weixin_process_ids():
    result = subprocess.run(
        ["tasklist", "/FI", "IMAGENAME eq Weixin.exe", "/FO", "CSV", "/NH"],
        capture_output=True,
        text=True,
        errors="replace",
        timeout=10,
    )
    process_ids = set()
    for row in csv.reader(io.StringIO(result.stdout)):
        if len(row) >= 2 and row[0].strip().lower() == "weixin.exe":
            process_ids.add(int(row[1].replace(",", "").strip()))
    return process_ids


def control_detail(control, depth=0):
    rect = control.BoundingRectangle
    return {
        "depth": depth,
        "process_id": int(getattr(control, "ProcessId", 0) or 0),
        "type": str(getattr(control, "ControlTypeName", "") or ""),
        "name": str(getattr(control, "Name", "") or ""),
        "class_name": str(getattr(control, "ClassName", "") or ""),
        "automation_id": str(getattr(control, "AutomationId", "") or ""),
        "rect": [rect.left, rect.top, rect.right, rect.bottom],
    }


def relevant(detail):
    text = " ".join(
        (
            detail["name"],
            detail["class_name"],
            detail["automation_id"],
        )
    ).lower()
    return any(
        keyword in text
        for keyword in (
            "momo",
            "search",
            "搜索",
            "session_item",
            "chat",
        )
    )


def main():
    if "--close-app" in sys.argv:
        hwnd = ctypes.windll.user32.FindWindowW(
            None,
            "EasyChat Momo - 微信自动回复助手",
        )
        closed = bool(hwnd and ctypes.windll.user32.PostMessageW(hwnd, 0x0010, 0, 0))
        OUTPUT_PATH.write_text(
            json.dumps({"hwnd": hwnd, "close_requested": closed}, indent=2),
            encoding="utf-8",
        )
        return

    if "--click-one-key" in sys.argv:
        window = auto.WindowControl(
            Name="EasyChat Momo - 微信自动回复助手",
            searchDepth=1,
        )
        button = window.ButtonControl(Name="一键启动监控", searchDepth=20)
        clicked = False
        error = ""
        try:
            if not window.Exists(5, 0):
                raise RuntimeError("未找到 EasyChat Momo 主窗口")
            if not button.Exists(5, 0):
                raise RuntimeError("未找到一键启动监控按钮")
            button.Click()
            clicked = True
        except Exception as exc:
            error = f"{type(exc).__name__}: {exc}"
        OUTPUT_PATH.write_text(
            json.dumps(
                {"clicked": clicked, "error": error},
                ensure_ascii=False,
                indent=2,
            ),
            encoding="utf-8",
        )
        return

    if "--open-target" in sys.argv:
        wechat = WeChat(locale="zh-CN")
        result = wechat.open_independent_chat_window("momo")
        statuses = []
        monitor_result = None
        if result and "--verify-monitor" in sys.argv:
            monitor_result = wechat.start_last_message_monitor(
                target_name="momo",
                callback=lambda *_args: None,
                check_interval=0.5,
                status_callback=statuses.append,
            )
            time.sleep(3)
            wechat.stop_last_message_monitor()
            time.sleep(0.5)
        OUTPUT_PATH.write_text(
            json.dumps(
                {
                    "ok": bool(result),
                    "message": result.message,
                    "monitor_ok": bool(monitor_result) if monitor_result else None,
                    "monitor_message": (
                        monitor_result.message if monitor_result else ""
                    ),
                    "statuses": statuses,
                },
                ensure_ascii=False,
                indent=2,
            ),
            encoding="utf-8",
        )
        return

    process_ids = get_weixin_process_ids()
    report = {
        "process_ids": sorted(process_ids),
        "before": [],
        "after": [],
        "search_value_before": "",
        "search_value_after": "",
        "error": "",
    }
    try:
        root = auto.GetRootControl()
        main_window = None
        search_input = None
        for window in root.GetChildren():
            if int(getattr(window, "ProcessId", 0) or 0) not in process_ids:
                continue
            report["before"].append(control_detail(window))
            if str(getattr(window, "ClassName", "") or "") == "mmui::MainWindow":
                candidate = window.EditControl(
                    ClassName="mmui::XValidatorTextEdit",
                    searchDepth=25,
                )
                if candidate.Exists(0.5, 0):
                    main_window = window
                    search_input = candidate

        if main_window is None or search_input is None:
            raise RuntimeError("未找到包含搜索框的微信主窗口")

        try:
            report["search_value_before"] = search_input.GetValuePattern().Value
        except Exception:
            pass

        main_window.SetActive()
        search_input.Click()
        auto.SendKeys("{Ctrl}a")
        auto.PressKey(auto.Keys.VK_BACK)
        auto.SendKeys("momo")
        time.sleep(6)

        try:
            report["search_value_after"] = search_input.GetValuePattern().Value
        except Exception:
            pass

        for window in root.GetChildren():
            if int(getattr(window, "ProcessId", 0) or 0) not in process_ids:
                continue
            top_detail = control_detail(window)
            report["after"].append(top_detail)
            try:
                for control, depth in auto.WalkControl(
                    window,
                    includeTop=False,
                    maxDepth=30,
                ):
                    detail = control_detail(control, depth)
                    if relevant(detail):
                        report["after"].append(detail)
            except Exception:
                continue
    except Exception as exc:
        report["error"] = f"{type(exc).__name__}: {exc}"

    OUTPUT_PATH.write_text(
        json.dumps(report, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )


if __name__ == "__main__":
    main()
