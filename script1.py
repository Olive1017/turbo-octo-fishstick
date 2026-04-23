import os
import re
import time
from playwright.sync_api import Playwright, sync_playwright, expect
from pywinauto import Application, keyboard


def save_file_with_pywinauto(save_dir, order_number, first_save=False, max_retries=3):
    """
    使用pywinauto处理Windows打印/保存对话框

    Args:
        save_dir: 保存目录
        order_number: 订单号（作为文件名）
        first_save: 是否为首次保存（首次需要输入路径，后续使用默认路径）
        max_retries: 最大重试次数

    Returns:
        bool: 是否成功保存
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
                time.sleep(0.5)
                keyboard.send_keys(save_dir)
                keyboard.send_keys('{ENTER}')
                time.sleep(0.5)
                keyboard.send_keys('{TAB}')


            else:
                # 后续保存使用默认路径，只需短暂等待
                time.sleep(0.5)

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


def process_single_order(page, index, save_dir, first_save=False):

    #处理单个TO下的所有OM的完整流程


    try:
        # 1. 双击TO单元格
        print(f"正在处理第 {index + 1} 个订单...")
        page.locator('[role="gridcell"][title^="TO"]').nth(index).dblclick()
        time.sleep(3)


        try:
            om_links = page.locator('[role="gridcell"][title^="OM"]')
            om_count = om_links.count()
            print(f"  该订单包含 {om_count} 个OM")
        except:
            om_count = 0
            print(f"  该订单包含 0 个OM")

        # 3. 遍历每个OM
        for om_index in range(om_count):
            print(f"  正在处理第 {om_index + 1} 个OM...")
            try:
                # 双击OM链接
                page.locator('[role="gridcell"][title^="OM"]').nth(om_index).dblclick()
                time.sleep(1)

                # 获取OM链接所在的行并提取数据
                om_row = page.locator('[role="gridcell"][title^="OM"]').nth(om_index).locator('..')
                cells = om_row.locator('td[role="gridcell"]')

                # 提取SAP订单号和行号（根据实际列索引调整）
                sap_order = cells.nth(3).inner_text().strip() if cells.count() > 1 else ""
                sap_line = cells.nth(4).inner_text().strip() if cells.count() > 2 else ""

                print(f"    SAP订单号: {sap_order}, 行号: {sap_line}")

                order_number = f"{sap_order}-{sap_line}"

                # 4. 点击"SO CN单打印"并处理打印对话框
                page.locator("a").filter(has_text="SO CN单打印").click()
                time.sleep(5)

                # 使用pywinauto处理打印对话框
                try:
                    save_file_with_pywinauto(save_dir, order_number, first_save=first_save)
                    print(f"    文件已保存到: {save_dir}\\{order_number}.pdf")
                except Exception as e:
                    print(f"    保存文件失败: {e}")
                    raise

                # 首次保存后，后续保存都使用默认路径
                first_save = False
                time.sleep(0.5)

                # 5. 关闭OM窗口
                page.locator("label:has-text('物流订单') + span.ui-icon-close").click()

                print(f"    第 {om_index + 1} 个OM处理完成！")

            except Exception as e:
                print(f"    处理第 {om_index + 1} 个OM时出错: {e}")
                # 尝试关闭可能打开的窗口
                try:
                    page.get_by_text("关闭").click()
                    time.sleep(1)
                except:
                    pass
                continue

        # 6. 关闭TO窗口
        page.get_by_text("关闭").click()
        time.sleep(1)

        print(f"  第 {index + 1} 个订单处理完成！")
        return True

    except Exception as e:
        print(f"  处理第 {index + 1} 个订单时出错: {e}")
        # 尝试关闭可能打开的窗口
        try:
            page.get_by_text("关闭").click()
            time.sleep(1)
        except:
            pass
        return False


class PrintingManager:
    """打印管理器类，负责管理浏览器和打印流程，供GUI调用"""

    def __init__(self):
        self.playwright = None
        self.browser = None
        self.context = None
        self.page = None

    def start_browser_and_login(self, username, password,browser_type="chromium"):

        try:
            self.playwright = sync_playwright().start()
            # 按优先级尝试浏览器
            browsers_to_try = ["msedge", "chrome", "chromium"]
            self.browser = None

            for browser_type in browsers_to_try:
                try:
                    if browser_type == "chromium":
                        self.browser = self.playwright.chromium.launch(headless=False)
                    else:
                        self.browser = self.playwright.chromium.launch(channel=browser_type, headless=False)
                    print(f"成功启动浏览器: {browser_type}")
                    break
                except Exception as e:
                    print(f"{browser_type} 不可用，尝试下一个... ({e})")
                    continue

            if self.browser is None:
                raise Exception("未找到可用的浏览器，请安装 Edge、Chrome 或 Chromium")
            self.context = self.browser.new_context()
            self.page = self.context.new_page()

            # 登录和导航
            self.page.goto("https://lms.cnoocshell.com/")
            self.page.get_by_role("textbox", name="请输入您的账号").fill(username)
            self.page.get_by_role("textbox", name="请输入您的账号").press("CapsLock")
            self.page.get_by_role("textbox", name="请输入您的账号").press("Tab")
            self.page.get_by_role("textbox", name="请输入您的密码").fill(password)
            self.page.get_by_text("登录", exact=True).click()
            self.page.locator("#VehicleReservationCustomBill1").click()
            self.page.locator("#Booking0").click()
            self.page.locator("a").filter(has_text="装运单").click()

            print("已导航到装运单页面，请筛选运输区域...")
            return True

        except Exception as e:
            print(f"启动浏览器失败: {e}")
            self.close()
            return False

    def start_printing(self, save_dir):
        """
        开始打印流程

        Args:
            save_dir: 保存目录

        Returns:
            dict: 包含打印结果的字典
        """
        try:
            print("等待查询结果加载...")
            time.sleep(2)

            print("开始自动下载流程...")

            # 主循环：遍历所有页面
            page_num = 1
            total_orders = 0
            success_orders = 0
            first_save = True

            while True:
                print(f"\n========== 开始处理第 {page_num} 页 ==========")

                # 获取当前页面的所有TO
                to_cells = self.page.locator('[role="gridcell"][title^="TO"]')
                to_count = to_cells.count()

                print(f"当前页面共有 {to_count} 个TO")

                if to_count == 0:
                    print("当前页面没有TO，结束处理")
                    break

                # 遍历当前页面的所有TO
                for i in range(to_count):
                    total_orders += 1

                    success = process_single_order(
                        self.page, i, save_dir, first_save=first_save
                    )
                    if success:
                        success_orders += 1
                        first_save = False

                print(f"\n第 {page_num} 页处理完成！本页共 {to_count} 个TO")

                # 翻页逻辑：如果TO数量=50，说明可能有下一页；如果<50，说明是最后一页
                if to_count < 50:
                    print(f"本页有 {to_count} 个TO（<50），已是最后一页，处理完成！")
                    break
                else:
                    print(f"本页有 {to_count} 个TO（=50），可能有下一页，尝试翻页...")

                    # 检查下一页按钮是否可用
                    next_button = self.page.locator(".ui-icon-seek-next")

                    try:
                        is_disabled = next_button.evaluate("el => el.classList.contains('ui-state-disabled')")
                        if is_disabled:
                            print("下一页按钮已被禁用，确实是最后一页，处理完成！")
                            break
                        else:
                            print("点击下一页...")
                            next_button.click()
                            page_num += 1
                            time.sleep(3)
                    except:
                        print("检查下一页按钮失败，假设已是最后一页")
                        break

            result = {
                "total_pages": page_num,
                "total_orders": total_orders,
                "success_orders": success_orders,
                "failed_orders": total_orders - success_orders
            }

            print(f"\n========== 全部处理完成 ==========")
            print(f"成功处理: {success_orders}")
            print(f"失败订单: {total_orders - success_orders}")

            return result

        except Exception as e:
            print(f"打印过程出错: {e}")
            raise

    def close(self):
        """关闭浏览器"""
        try:
            if self.context:
                self.context.close()
            if self.browser:
                self.browser.close()
            if self.playwright:
                self.playwright.stop()
        except:
            pass


def run(playwright: Playwright) -> None:
    """原始命令行运行函数"""
    # 设置保存目录
    save_dir = r"C:\Users\xinan\Desktop\ces"

    # 使用管理器启动浏览器
    manager = PrintingManager()

    if not manager.start_browser_and_login("LUOJIAN", "Aa123456@@@"):
        return

    print("\n" + "=" * 50)
    print("         筛选运输区域")
    print("=" * 50)
    print()
    print("请在浏览器中：")
    print("  1. 选择运输区域和到厂日期")
    print("  2. 点击'查询'")
    print()
    print("完成后，请返回此窗口按 Enter 键继续...")
    print("=" * 50)
    input("")  # 等待用户按Enter

    # 开始打印
    result = manager.start_printing(save_dir)
    manager.close()


if __name__ == "__main__":
    with sync_playwright() as playwright:
        run(playwright)
