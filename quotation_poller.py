"""
报价完成检测模块
轮询飞书多维表格报价明细表，检测报价结果是否生成且稳定（连续 N 秒无变化）
"""

import time
import hashlib
import json
import logging
from typing import Optional

logger = logging.getLogger("quotation_poller")


class QuotationPoller:
    """报价完成检测器 - 轮询多维表格"""

    def __init__(self, config: dict):
        polling_config = config.get("polling", {})
        self.stable_seconds = polling_config.get("quotation_stable_time", 60)
        self.poll_interval = polling_config.get("quotation_poll_interval", 30)
        self.timeout = 1800  # 最大等待 30 分钟

    def poll(self, writer, inquiry_record_id: str, req_number: str = "", skip_wait: bool = False) -> Optional[list[dict]]:
        """
        轮询等待报价完成

        Args:
            writer: BaseWriter 实例
            inquiry_record_id: 询盘记录 ID
            req_number: 需求编号（用于过滤报价明细）
            skip_wait: 是否跳过等待（测试模式，数据已存在）

        Returns:
            Optional[list[dict]]: 报价明细记录列表，超时返回 None
        """
        logger.info(f"开始轮询报价结果: inquiry_record={inquiry_record_id}, "
                     f"需求编号={req_number}, 稳定时间={self.stable_seconds}s, 超时={self.timeout}s"
                     f"{', 跳过等待' if skip_wait else ''}")

        start_time = time.time()
        last_hash = ""
        stable_since = None

        while True:
            elapsed = time.time() - start_time
            if elapsed > self.timeout:
                logger.warning(f"轮询超时 ({self.timeout}s)，放弃等待")
                return None

            # 查询报价明细表中关联到该询盘的记录（按需求编号过滤）
            records = writer.list_quotation_records(req_number=req_number)

            # 过滤：只看有报价数据的记录（产品名称不为空）
            quotation_items = [
                r for r in records
                if r.get("fields") and len(r["fields"]) > 0
                and r["fields"][0]  # 产品名称不为空
            ]

            if quotation_items:
                current_hash = self._compute_hash(quotation_items)

                if skip_wait:
                    # 测试模式：数据存在就直接返回
                    logger.info(f"✅ 找到 {len(quotation_items)} 条报价记录（测试模式，跳过等待）")
                    return quotation_items

                if current_hash != last_hash:
                    # 数据有变化，重置稳定计时
                    last_hash = current_hash
                    stable_since = time.time()
                    logger.info(f"报价数据有变化（共 {len(quotation_items)} 条记录），重置稳定计时")
                else:
                    # 数据无变化，检查是否达到稳定时间
                    stable_duration = time.time() - stable_since if stable_since else 0
                    logger.debug(f"报价数据稳定 {stable_duration:.0f}s / {self.stable_seconds}s")

                    if stable_duration >= self.stable_seconds:
                        logger.info(f"✅ 报价结果已稳定 {self.stable_seconds}s，共 {len(quotation_items)} 条记录")
                        return quotation_items
            else:
                logger.debug(f"暂无报价数据，继续等待... ({elapsed:.0f}s)")
                last_hash = ""
                stable_since = None

            time.sleep(self.poll_interval)

    def _compute_hash(self, records: list[dict]) -> str:
        """
        计算记录列表的 hash（用于判断是否变化）

        Args:
            records: 报价明细记录列表

        Returns:
            str: MD5 hash
        """
        # 提取关键字段组成可哈希的字符串
        key_data = []
        for r in records:
            fields = r.get("fields", [])
            # 产品名称 + 型号 + 数量 + 单价 + 总价
            if len(fields) >= 5:
                key_data.append(str(fields[0]))   # 产品名称
                key_data.append(str(fields[1]))   # 产品型号
                key_data.append(str(fields[2]))   # 数量
                key_data.append(str(fields[3]))   # 单价
                key_data.append(str(fields[4]))   # 总价

        raw = "|".join(key_data)
        return hashlib.md5(raw.encode()).hexdigest()
