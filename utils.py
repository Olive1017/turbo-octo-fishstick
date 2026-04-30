import os
import time
from pywinauto import keyboard
import pyperclip


def save_file_with_pywinauto(save_dir, order_number, first_save=False, max_retries=3):
    """
    使用pywinauto处理Windows打印/保存对话框
    """
    for attempt in range(max_retries):
        try:
            # 使用pywinauto的键盘输入（比pyautogui更稳定）
            keyboard.send_keys('{ENTER}')
            time.sleep(5)

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


def get_local_orders(save_dir):
    """
    获取本地已下载的订单号（从PDF文件名提取）
    
    Args:
        save_dir: 保存目录路径
        
    Returns:
        set: 本地已存在的订单号集合
    """
    local_orders = set()
    if not os.path.exists(save_dir):
        return local_orders
    
    for filename in os.listdir(save_dir):
        if filename.endswith('.pdf'):
            # 提取订单号（格式：SAP订单号-行号.pdf）
            order_number = filename[:-4]  # 去掉.pdf
            local_orders.add(order_number)
    
    return local_orders


def get_system_orders(page):
    """
    从系统获取所有订单号（遍历所有页面，但不下载）
    
    Args:
        page: playwright page对象
        
    Returns:
        set: 系统中所有订单号集合
    """
    all_orders = set()
    page_num = 1
    
    while True:
        print(f"正在扫描第 {page_num} 页的订单...")
        
        # 获取当前页面的所有TO
        to_cells = page.locator('[role="gridcell"][title^="TO"]')
        to_count = to_cells.count()
        
        if to_count == 0:
            break
        
        # 遍历当前页面的所有TO
        for i in range(to_count):
            try:
                # 双击TO单元格
                to_cells.nth(i).dblclick()
                time.sleep(5)
                
                # 获取该TO下的所有OM
                om_links = page.locator('[role="gridcell"][title^="OM"]')
                om_count = om_links.count()
                
                # 遍历每个OM提取订单号
                for om_index in range(om_count):
                    try:
                        om_row = om_links.nth(om_index).locator('..')
                        cells = om_row.locator('td[role="gridcell"]')
                        
                        # 提取SAP订单号和行号
                        sap_order = cells.nth(3).inner_text().strip() if cells.count() > 1 else ""
                        sap_line = cells.nth(4).inner_text().strip() if cells.count() > 2 else ""
                        
                        order_number = f"{sap_order}-{sap_line}"
                        all_orders.add(order_number)
                        
                    except Exception as e:
                        print(f"    提取OM订单号时出错: {e}")
                        continue
                
                # 关闭TO窗口
                page.get_by_text("关闭").click()
                time.sleep(1)
                
            except Exception as e:
                print(f"  处理第 {i + 1} 个TO时出错: {e}")
                try:
                    page.get_by_text("关闭").click()
                    time.sleep(1)
                except:
                    pass
                continue
        
        # 翻页逻辑
        if to_count < 50:
            print(f"扫描完成，共 {len(all_orders)} 个订单")
            break
        else:
            # 检查下一页按钮是否可用
            next_button = page.locator(".ui-icon-seek-next")
            try:
                is_disabled = next_button.evaluate("el => el.classList.contains('ui-state-disabled')")
                if is_disabled:
                    print(f"扫描完成，共 {len(all_orders)} 个订单")
                    break
                else:
                    next_button.click()
                    page_num += 1
                    time.sleep(3)
            except:
                print(f"扫描完成，共 {len(all_orders)} 个订单")
                break
    
    return all_orders


def get_new_orders(save_dir, page):
    """
    获取需要下载的新订单（对比本地订单和系统订单）
    
    Args:
        save_dir: 保存目录路径
        page: playwright page对象
        
    Returns:
        list: 需要下载的新订单列表
    """
    local_orders = get_local_orders(save_dir)
    system_orders = get_system_orders(page)
    
    new_orders = system_orders - local_orders
    
    print(f"\n本地已存在订单: {len(local_orders)} 个")
    print(f"系统共有订单: {len(system_orders)} 个")
    print(f"需要下载的新订单: {len(new_orders)} 个")
    
    return list(new_orders)
