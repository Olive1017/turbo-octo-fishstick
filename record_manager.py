"""
记录管理器，负责读取、写入和比对订单记录
"""
import json
import os


class RecordManager:
    """记录管理器，负责读取、写入和比对订单记录"""

    def __init__(self, save_dir):
        """
        初始化记录管理器

        Args:
            save_dir: 保存目录，JSON文件将保存在该目录下
        """
        self.save_dir = save_dir
        self.record_file = os.path.join(save_dir, "downloaded_orders.json")
        self.records = self._load_records()

    def _load_records(self):
        """加载记录文件"""
        if os.path.exists(self.record_file):
            try:
                with open(self.record_file, 'r', encoding='utf-8') as f:
                    return json.load(f)
            except Exception as e:
                print(f"读取记录文件失败: {e}")
                return {}
        return {}

    def _save_records(self):
        """保存记录到文件"""
        try:
            with open(self.record_file, 'w', encoding='utf-8') as f:
                json.dump(self.records, f, ensure_ascii=False, indent=2)
            print(f"记录已保存到: {self.record_file}")
        except Exception as e:
            print(f"保存记录文件失败: {e}")

    def should_download_om(self, to_number, sap_order_line):
        """
        判断是否需要下载该OM

        Args:
            to_number: TO单号
            sap_order_line: SAP订单号-行号

        Returns:
            bool: True表示需要下载，False表示跳过
        """
        if to_number not in self.records:
            # TO号不存在，需要下载
            return True

        if sap_order_line not in self.records[to_number]:
            # SAP号不存在，需要下载
            return True

        print(f"    跳过已下载的OM: {to_number} -> {sap_order_line}")
        return False

    def record_om(self, to_number, sap_order_line):
        """
        记录已下载的OM

        Args:
            to_number: TO单号
            sap_order_line: SAP订单号-行号
        """
        if to_number not in self.records:
            self.records[to_number] = []

        if sap_order_line not in self.records[to_number]:
            self.records[to_number].append(sap_order_line)
            # 保存记录
            self._save_records()
