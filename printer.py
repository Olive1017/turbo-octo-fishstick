"""
打印管理器，负责浏览器控制和订单下载流程
"""
import time
from playwright.sync_api import sync_playwright
from record_manager import RecordManager
from file_saver import save_file_with_pywinauto

# 固定的列索引（从0开始）
SAP_COLUMN_IDX = 6  # 第5列，内容格式为 "SAP订单号-行号"


def get_to_sap_info_from_list(page, index):
    """
    从列表页获取TO对应的SAP订单信息

    Args:
        page: Playwright page对象
        index: TO的索引

    Returns:
        dict: {'to_number': 'TOxxx', 'sap_orders': ['SAP001-001', 'SAP002-001']} 或 None
    """
    try:
        # 获取TO单元格
        to_cell = page.locator('[role="gridcell"][title^="TO"]').nth(index)
        to_number = to_cell.inner_text().strip()

        # 获取TO所在行的所有单元格
        to_row = to_cell.locator('..')
        cells = to_row.locator('td[role="gridcell"]')
        cell_count = cells.count()

        # 提取第5列的SAP订单号-行号（可能有多个，逗号分隔）
        if SAP_COLUMN_IDX < cell_count:
            sap_text = cells.nth(SAP_COLUMN_IDX).inner_text().strip()

            if sap_text:
                # 按逗号分割，去除空格
                sap_orders = [s.strip() for s in sap_text.split(',') if s.strip()]

                if sap_orders:
                    return {
                        'to_number': to_number,
                        'sap_orders': sap_orders
                    }

        return None
    except Exception as e:
        print(f"从列表页获取SAP信息失败: {e}")
        return None


def process_single_order(page, index, save_dir, first_save=False, record_manager=None, incremental_mode=False):
    """
    处理单个TO下的所有OM的完整流程

    Args:
        page: Playwright page对象
        index: TO索引
        save_dir: 保存目录
        first_save: 是否首次保存
        record_manager: 记录管理器
        incremental_mode: 是否增量模式

    Returns:
        bool/str: True表示成功，False表示失败，'skipped'表示跳过
    """

    try:
        # 1. 从列表页获取TO和SAP信息
        print(f"正在处理第 {index + 1} 个订单...")
        to_info = get_to_sap_info_from_list(page, index)

        if not to_info:
            print(f"  无法获取TO信息，跳过")
            return False

        to_number = to_info['to_number']
        list_sap_orders = to_info.get('sap_orders', [])
        print(f"  TO单号: {to_number}, 列表SAP: {list_sap_orders}")

        # 增量模式：检查是否需要下载
        if record_manager and incremental_mode:
            if to_number in record_manager.records:
                # TO已记录，检查SAP列表是否完全匹配（内容和顺序都要一致）
                recorded_saps = record_manager.records[to_number]  # 保持列表顺序

                # 精确比较：内容相同且顺序相同
                if list_sap_orders == recorded_saps:
                    print(f"  ✓ 该TO的SAP列表完全匹配，跳过")
                    return 'skipped'  # 跳过，不点开TO
                else:
                    # SAP列表发生变化（新增、删除、顺序变化），需要重新下载
                    print(f"  ⚠ 该TO的SAP列表已变化:")
                    print(f"    已记录: {recorded_saps}")
                    print(f"    当前列表: {list_sap_orders}")
                    print(f"  将重新下载整个TO")

                    # 清除旧记录，重新记录
                    record_manager.records[to_number] = []
            else:
                print(f"  该TO未下载，需要处理")

        page.locator('[role="gridcell"][title^="TO"]').nth(index).dblclick()
        time.sleep(3)

        # 获取OM数量
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

                order_number = f"{sap_order}-{sap_line}"
                print(f"    SAP订单号: {sap_order}, 行号: {sap_line}")

                # 检查记录管理器（如果提供了record_manager）
                if record_manager:
                    if not record_manager.should_download_om(to_number, order_number):
                        continue  # 跳过已下载的OM

                # 4. 点击"SO CN单打印"并处理打印对话框
                page.locator("a").filter(has_text="SO CN单打印").click()
                time.sleep(5)

                # 使用pywinauto处理打印对话框
                try:
                    save_file_with_pywinauto(save_dir, order_number, first_save=first_save)
                    print(f"    文件已保存到: {save_dir}\\{order_number}.pdf")

                    # 记录已下载的OM（如果提供了record_manager）
                    if record_manager:
                        record_manager.record_om(to_number, order_number)

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
                # 尝试关闭可能打开的窗口（使用更精确的选择器）
                try:
                    page.locator("label:has-text('物流订单') + span.ui-icon-close").click()
                    time.sleep(1)
                except:
                    try:
                        page.get_by_role("button", name="关闭").first.click()
                    except:
                        pass
                continue

        # 6. 关闭TO窗口（使用更精确的选择器）
        try:
            # 尝试找到TO窗口的关闭按钮
            to_close_button = page.locator("[id*='_ToolBar1_toolbar'] span.txt").filter(has_text="关闭").first
            to_close_button.click()
        except:
            # 如果失败，尝试通用关闭按钮
            try:
                page.get_by_role("button", name="关闭").first.click()
            except:
                pass
        time.sleep(1)

        print(f"  第 {index + 1} 个订单处理完成！")
        return True

    except Exception as e:
        print(f"  处理第 {index + 1} 个订单时出错: {e}")
        # 尝试关闭可能打开的窗口
        try:
            page.locator("label:has-text('物流订单') + span.ui-icon-close").click()
            time.sleep(1)
        except:
            try:
                page.get_by_role("button", name="关闭").first.click()
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

    def start_browser_and_login(self, username, password, browser_type="chromium"):
        """启动浏览器并登录"""
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
            self.page.get_by_text("运输管理").click()
            self.page.locator("#Booking0").click()
            self.page.locator("a").filter(has_text="装运单").click()

            print("已导航到装运单页面，请筛选运输区域...")
            return True

        except Exception as e:
            print(f"启动浏览器失败: {e}")
            self.close()
            return False

    def start_printing(self, save_dir, incremental_mode=False):
        """开始打印流程"""
        try:
            print("等待查询结果加载...")
            time.sleep(2)

            print("开始自动下载流程...")

            # 初始化记录管理器（无论什么模式都需要，用于记录下载历史）
            record_manager = RecordManager(save_dir)
            if incremental_mode:
                print(f"已启用增量模式，记录文件: {record_manager.record_file}")
            else:
                print("已启用全量下载模式（仍会记录下载历史）")

            # 主循环：遍历所有页面
            page_num = 1
            total_orders = 0
            success_orders = 0
            failed_orders = 0
            skipped_orders = 0
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

                    result = process_single_order(
                        self.page, i, save_dir, first_save=first_save,
                        record_manager=record_manager, incremental_mode=incremental_mode
                    )

                    if result == 'skipped':
                        skipped_orders += 1
                    elif result:
                        success_orders += 1
                        first_save = False
                    else:
                        failed_orders += 1

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
                "failed_orders": failed_orders,
                "skipped_orders": skipped_orders
            }

            print(f"\n========== 全部处理完成 ==========")
            print(f"总订单数: {total_orders}")
            print(f"成功处理: {success_orders}")
            print(f"失败订单: {failed_orders}")
            if incremental_mode:
                print(f"跳过订单: {skipped_orders}")
            print(f"记录文件: {record_manager.record_file}")

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
