"""
文件保存工具，使用pywinauto处理Windows打印/保存对话框
"""
import time
import pyperclip
from pywinauto import keyboard


def save_file_with_pywinauto(save_dir, order_number, first_save=False, max_retries=3):
    """
    使用pywinauto处理Windows打印/保存对话框

    Args:
        save_dir: 保存目录
        order_number: 订单号（文件名）
        first_save: 是否首次保存（首次需要输入路径）
        max_retries: 最大重试次数

    Returns:
        bool: 成功返回True，失败返回False
    """
    for attempt in range(max_retries):
        try:
            # 使用pywinauto的键盘输入（比pyautogui更稳定）
            keyboard.send_keys('{ENTER}')
            time.sleep(2)

            # 仅第一次保存时需要输入路径
            if first_save:
                # Ctrl+L 打开地址栏
                keyboard.send_keys('^l')
                time.sleep(1)
                pyperclip.copy(save_dir)  # 复制到剪贴板
                keyboard.send_keys('^v')
                keyboard.send_keys('{ENTER}')
                time.sleep(0.5)
                keyboard.send_keys('{TAB}')
            else:
                # 后续保存使用默认路径，只需短暂等待
                time.sleep(0.1)

            # Alt+N 聚焦到文件名输入框
            keyboard.send_keys('%n')
            time.sleep(0.5)

            # 输入文件名（订单号）
            keyboard.send_keys(order_number)
            time.sleep(0.5)

            # 按Enter确认保存
            keyboard.send_keys('{ENTER}')
            time.sleep(1)

            return True

        except Exception as e:
            print(f"      尝试 {attempt + 1} 失败: {e}")
            if attempt < max_retries - 1:
                time.sleep(2)
            else:
                raise
