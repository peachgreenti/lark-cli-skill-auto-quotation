"""
飞书多维表格写入模块
通过 lark-cli 命令行操作飞书多维表格
功能：创建询盘记录、上传附件、写入客户、查询报价结果
"""

import os
import json
import subprocess
import logging
from datetime import datetime
from typing import Optional

logger = logging.getLogger("base_writer")


class BaseWriter:
    """飞书多维表格写入器 - 通过 lark-cli 命令行"""

    def __init__(self, config: dict):
        base_config = config.get("base", {})
        self.app_token = base_config.get("app_token")
        self.inquiry_table_id = base_config.get("inquiry_table_id")
        self.quotation_table_id = base_config.get("quotation_table_id")
        self.customer_table_id = base_config.get("customer_table_id")
        self.inquiry_fields = base_config.get("inquiry_fields", {})
        self.quotation_fields = base_config.get("quotation_fields", {})

    def _run_cli(self, args: list[str], cwd: str = None) -> dict:
        """
        执行 lark-cli 命令并解析 JSON 输出

        Args:
            args: 完整的命令参数列表（从 lark-cli 开始）
            cwd: 工作目录（可选）

        Returns:
            dict: {"ok": bool, "data": ...}

        Raises:
            RuntimeError: 命令执行失败
        """
        cmd = ["lark-cli"] + args
        logger.debug(f"执行命令: {' '.join(cmd)}")

        try:
            result = subprocess.run(
                cmd,
                capture_output=True,
                text=True,
                timeout=60,
                cwd=cwd,
            )
            if result.returncode != 0:
                logger.error(f"lark-cli 命令失败: {result.stderr}")
                raise RuntimeError(f"lark-cli error: {result.stderr}")

            output = json.loads(result.stdout)
            # lark-cli api 命令返回 {"code": 0, ...} 而不是 {"ok": true, ...}
            if not output.get("ok") and output.get("code", 0) != 0:
                logger.error(f"lark-cli 返回错误: {output}")
                raise RuntimeError(f"lark-cli API error: {output}")

            return output

        except json.JSONDecodeError as e:
            logger.error(f"lark-cli 输出 JSON 解析失败: {result.stdout[:500]}")
            raise RuntimeError(f"JSON parse error: {e}")
        except subprocess.TimeoutExpired:
            logger.error("lark-cli 命令超时")
            raise RuntimeError("lark-cli timeout")

    # ============================================================
    # 询盘需求表操作
    # ============================================================

    def create_inquiry_record(
        self,
        pdf_path: str,
        requirement_text: str = "",
        query_date: Optional[str] = None,
    ) -> dict:
        """
        创建询盘需求记录 — 上传询价函附件 + 填入报价需求文字

        Args:
            pdf_path: PDF 询价函本地路径
            requirement_text: AI 提取的报价需求文字（如 "3个CP001、2个SP002"）
            query_date: 查询日期 YYYY-MM-DD（可选，默认今天）

        Returns:
            dict: {"record_id": str}
        """
        if not query_date:
            query_date = datetime.now().strftime("%Y-%m-%d")

        # 创建记录（同时填入报价需求）
        row_values = [
            requirement_text or "",                   # 报价需求
            "调整/准备中",                           # 询价状态
            query_date,                             # 查询日期
            datetime.now().strftime("%Y-%m-%d %H:%M"),  # 查询时间
        ]

        record_data = self._run_cli([
            "base", "+record-batch-create",
            "--base-token", self.app_token,
            "--table-id", self.inquiry_table_id,
            "--json", json.dumps({
                "fields": [
                    "报价需求",
                    "询价状态",
                    "查询日期",
                    "查询时间",
                ],
                "rows": [row_values],
            }),
        ])

        data = record_data.get("data", {})
        record_ids = data.get("record_id_list", [])
        if not record_ids:
            raise RuntimeError("创建询盘记录失败：未返回 record_id")

        record_id = record_ids[0]
        logger.info(f"✅ 询盘记录已创建: record_id={record_id}, 报价需求={requirement_text[:50]}")

        # 读取需求编号（自动编号字段，创建后才有值）
        req_number = self._get_req_number(record_id)

        # 上传 PDF 附件（如果有）
        if pdf_path and os.path.exists(pdf_path):
            self.upload_attachment(pdf_path, record_id)

        return {"record_id": record_id, "req_number": req_number}

    def _get_req_number(self, record_id: str) -> str:
        """读取询盘记录的需求编号"""
        try:
            req_field_id = self.inquiry_fields.get("req_number", "")
            if not req_field_id:
                return ""
            result = self._run_cli([
                "base", "+record-get",
                "--base-token", self.app_token,
                "--table-id", self.inquiry_table_id,
                "--record-id", record_id,
            ])
            record = result.get("data", {}).get("record", {})
            req_number = record.get("需求编号", "")
            if req_number:
                logger.info(f"需求编号: {req_number}")
            return req_number
        except Exception as e:
            logger.warning(f"读取需求编号失败: {e}")
            return ""

    def upload_attachment(self, pdf_path: str, record_id: str) -> bool:
        """
        上传 PDF 附件到询盘记录（两步流程，绕过 lark-cli +record-upload-attachment 的 0B bug）

        步骤1: drive +upload 上传文件到云空间（size 正确）
        步骤2: api PATCH 写入附件字段（带 deprecated_set_attachment 标记绕过 READONLY）

        Args:
            pdf_path: PDF 文件本地路径
            record_id: 询盘记录 ID

        Returns:
            bool: 是否成功
        """
        try:
            file_dir = os.path.dirname(pdf_path) or "."
            file_name = os.path.basename(pdf_path)
            attachment_field_name = "询价函"

            # 步骤1: 用 drive +upload 上传文件到云空间（size 正确）
            upload_result = self._run_cli([
                "drive", "+upload",
                "--file", file_name,
                "--name", file_name,
            ], cwd=file_dir)

            file_token = upload_result.get("data", {}).get("file_token", "")
            file_size = upload_result.get("data", {}).get("size", 0)
            if not file_token:
                logger.error(f"文件上传到云空间失败: {upload_result}")
                return False
            logger.info(f"文件已上传到云空间: token={file_token}, size={file_size}")

            # 步骤2: 读取现有附件（保留已有附件）
            existing_attachments = self._get_existing_attachments(record_id, attachment_field_name)

            # 合并：已有附件 + 新上传的附件
            all_attachments = existing_attachments + [{
                "deprecated_set_attachment": True,
                "file_token": file_token,
                "name": file_name,
            }]

            # 步骤3: 用 api PATCH 写入附件字段（带 deprecated_set_attachment 绕过 READONLY）
            patch_body = {attachment_field_name: all_attachments}
            patch_result = self._run_cli([
                "api", "PATCH",
                f"/open-apis/base/v3/bases/{self.app_token}/tables/{self.inquiry_table_id}/records/{record_id}",
                "--data", json.dumps(patch_body),
            ])

            ignored = patch_result.get("data", {}).get("ignored_fields", [])
            if ignored:
                logger.warning(f"附件字段写入被忽略: {ignored}")
                return False

            logger.info(f"✅ 附件上传成功: {file_name} ({file_size}B) → record {record_id}")
            return True
        except Exception as e:
            logger.error(f"附件上传失败: {e}")
            return False

    def _get_existing_attachments(self, record_id: str, field_name: str) -> list[dict]:
        """
        读取记录中已有的附件（保留已有附件，避免覆盖）

        Returns:
            list[dict]: 已有附件列表（带 deprecated_set_attachment 标记）
        """
        try:
            result = self._run_cli([
                "base", "+record-get",
                "--base-token", self.app_token,
                "--table-id", self.inquiry_table_id,
                "--record-id", record_id,
            ])
            record = result.get("data", {}).get("record", {})
            existing = record.get(field_name, [])
            if not existing:
                return []
            # 为已有附件添加 deprecated_set_attachment 标记
            return [
                {
                    "deprecated_set_attachment": True,
                    "file_token": a.get("file_token", ""),
                    "name": a.get("name", ""),
                }
                for a in existing
            ]
        except Exception as e:
            logger.warning(f"读取已有附件失败: {e}")
            return []

    # ============================================================
    # 客户管理表操作
    # ============================================================

    def write_customer(
        self,
        customer_name: str,
        customer_email: str,
        contact_person: str = "",
        contact_phone: str = "",
    ) -> dict:
        """
        写入客户管理表（如已存在则跳过）
        注意：客户名称是关联选择字段，需要先在表格中手动添加客户选项

        Args:
            customer_name: 客户名称
            customer_email: 客户邮箱
            contact_person: 联系人
            contact_phone: 联系电话

        Returns:
            dict: {"record_id": str}
        """
        # 检查客户是否已存在
        existing = self._search_customer(customer_name)
        if existing:
            logger.info(f"客户已存在，跳过创建: {customer_name}")
            return {"record_id": existing}

        # 客户名称是关联选择字段，不支持 API 写入
        # 仅记录日志，提示用户手动添加
        logger.warning(f"客户 '{customer_name}' 不在客户管理表中，请手动在多维表格中添加该客户选项")
        return {"record_id": ""}

    def _search_customer(self, customer_name: str) -> Optional[str]:
        """
        搜索客户是否已存在

        Args:
            customer_name: 客户名称

        Returns:
            Optional[str]: 已存在的 record_id，不存在返回 None
        """
        try:
            result = self._run_cli([
                "base", "+record-list",
                "--base-token", self.app_token,
                "--table-id", self.customer_table_id,
                "--field-id", "客户名称",
                "--limit", "100",
            ])
            data = result.get("data", {})
            raw_records = data.get("data", [])
            record_ids = data.get("record_id_list", [])
            for i, fields in enumerate(raw_records):
                if fields and fields[0] == customer_name:
                    return record_ids[i] if i < len(record_ids) else None
        except Exception as e:
            logger.warning(f"搜索客户失败: {e}")
        return None

    # ============================================================
    # 报价明细表操作（轮询用）
    # ============================================================

    def list_quotation_records(self, req_number: str = "") -> list[dict]:
        """
        查询报价明细表记录

        Args:
            req_number: 需求编号（可选，用于过滤）

        Returns:
            list[dict]: 记录列表
        """
        field_names = [
            "产品名称", "产品型号", "数量", "产品售价/元", "总价/元",
            "客户名称", "报价状态", "需求编号", "主报价需求", "子报价需求",
        ]
        args = [
            "base", "+record-list",
            "--base-token", self.app_token,
            "--table-id", self.quotation_table_id,
            "--limit", "500",
        ]
        for f in field_names:
            args.extend(["--field-id", f])

        result = self._run_cli(args)
        data = result.get("data", {})
        raw_records = data.get("data", [])
        record_ids = data.get("record_id_list", [])

        # 组装记录
        records = []
        for i, fields in enumerate(raw_records):
            record = {
                "record_id": record_ids[i] if i < len(record_ids) else "",
                "fields": fields,
            }
            records.append(record)

        # 按需求编号过滤
        if req_number:
            req_field_idx = field_names.index("需求编号") if "需求编号" in field_names else None
            if req_field_idx is not None:
                records = [
                    r for r in records
                    if len(r.get("fields", [])) > req_field_idx
                    and r["fields"][req_field_idx] == req_number
                ]

        return records

    def get_quotation_record(self, record_id: str) -> dict:
        """
        获取单条报价明细记录

        Args:
            record_id: 记录 ID

        Returns:
            dict: 记录数据
        """
        result = self._run_cli([
            "base", "+record-get",
            "--base-token", self.app_token,
            "--table-id", self.quotation_table_id,
            "--record-id", record_id,
        ])
        return result.get("data", {}).get("record", {})

    def update_inquiry_status(self, record_id: str, status: str) -> bool:
        """
        更新询盘状态

        Args:
            record_id: 记录 ID
            status: 新状态（调整/准备中、待客户确认、客户已接受、客户已拒绝）

        Returns:
            bool: 是否成功
        """
        try:
            self._run_cli([
                "base", "+record-batch-update",
                "--base-token", self.app_token,
                "--table-id", self.inquiry_table_id,
                "--json", json.dumps({
                    "records": [
                        {
                            "record_id": record_id,
                            "fields": {"询价状态": status},
                        }
                    ]
                }),
            ])
            logger.info(f"✅ 询盘状态已更新: record={record_id}, status={status}")
            return True
        except Exception as e:
            logger.error(f"更新询盘状态失败: {e}")
            return False
