import time
import subprocess
import pyautogui
import win32api
import win32gui
import win32con
import threading
import sys
import pystray
import tkinter as tk
from tkinter import simpledialog
from PIL import Image, ImageDraw

# 默认空闲时间设置为 3 分钟（3 分钟 * 60 秒）
default_idle_time = 3  # 默认 3 分钟
idle_time = default_idle_time  # 初始化为空闲时间为默认值（3 分钟）


def get_idle_time():
    """获取键鼠空闲时间（秒）"""
    lii = win32api.GetLastInputInfo()
    millis = win32api.GetTickCount() - lii
    return millis / 1000.0


def find_wechat_window():
    """查找微信主窗口句柄"""
    return win32gui.FindWindow("WeChatMainWndForPC", None)


def open_wechat():
    """启动或激活微信"""
    hwnd = find_wechat_window()
    if hwnd:  # 如果微信已运行
        win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
        win32gui.SetForegroundWindow(hwnd)
        return True
    else:  # 启动微信
        try:
            subprocess.Popen(r"C:\Program Files\Tencent\WeChat\WeChat.exe")
            # 最多等待 5 秒直到窗口出现
            for _ in range(50):
                time.sleep(0.1)
                hwnd = find_wechat_window()
                if hwnd:
                    win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
                    win32gui.SetForegroundWindow(hwnd)
                    return True
        except Exception as e:
            print("无法启动微信:", e)
    return False


def lock_wechat():
    """发送 Ctrl+L 锁定微信"""
    pyautogui.hotkey("ctrl", "l")
    print(">>> 已执行微信锁定")


def monitor_idle():
    """监控键鼠空闲"""
    global idle_time  # 使用全局变量
    locked = False
    while True:
        idle = get_idle_time()

        # 空闲时间判断：将空闲时间转为秒进行比较（空闲时间是分钟单位，所以乘以 60）
        if idle >= idle_time * 60 and not locked:  # 根据用户设定的空闲时间判断
            print(f"检测到空闲 {idle_time} 分钟，锁定微信...")
            if open_wechat():
                time.sleep(0.2)  # 给系统切换窗口一点缓冲
                lock_wechat()
                locked = True

        if idle < 1 and locked:
            locked = False
            print("检测到用户活动，重置锁定状态")

        time.sleep(1)


def create_image():
    """托盘图标"""
    img = Image.new("RGB", (64, 64), "white")
    draw = ImageDraw.Draw(img)
    draw.ellipse((16, 16, 48, 48), fill="green")
    return img


def on_quit(icon, item):
    icon.stop()
    sys.exit()


def prompt_for_idle_time():
    """弹出输入框，用户输入空闲时间（分钟）"""
    global idle_time
    # 创建 Tkinter 输入框对话框
    root = tk.Tk()
    root.withdraw()  # 不显示主窗口

    # 弹出输入框，限制输入时间为 10 秒
    user_input = simpledialog.askstring("空闲时间设置", f"请输入空闲时间（分钟，默认 {default_idle_time} 分钟）：")
    
    if user_input:
        if user_input.isdigit():
            idle_time = int(user_input)
            print(f"已设置空闲时间为 {idle_time} 分钟")
        else:
            print("输入无效，使用默认空闲时间")
    else:
        print("未输入任何内容，使用默认空闲时间")
    
    # 进入后台最小化并退出 Tkinter
    root.iconify()  # 最小化窗口
    root.after(100, root.quit)  # 关闭输入框
    root.destroy()  # 销毁窗口
    root.mainloop()


def run_tray():
    icon = pystray.Icon(
        "WeChat Locker",
        create_image(),
        menu=pystray.Menu(pystray.MenuItem("退出", on_quit)),
    )

    # 提示用户输入空闲时间
    prompt_for_idle_time()

    t1 = threading.Thread(target=monitor_idle, daemon=True)
    t1.start()

    icon.run()


if __name__ == "__main__":
    print("程序已启动：请输入空闲时间（单位：分钟），如果没有输入则默认空闲时间为 3 分钟")
    run_tray()
