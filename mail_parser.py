"""
邮件解析模块
解析邮件原始内容，提取正文、PDF 附件、发件人信息等
"""

import email
import logging
from email.header import decode_header
from typing import Optional

logger = logging.getLogger("mail_parser")


class MailParser:
    """邮件解析器"""

    def parse(self, raw_email: bytes) -> dict:
        """
        解析邮件原始内容

        Args:
            raw_email: 邮件原始字节（来自 mail_watcher）

        Returns:
            dict: {
                subject: str,              # 邮件主题
                from_addr: str,            # 发件人邮箱
                from_name: str,            # 发件人名称
                date: str,                 # 邮件日期
                body: str,                 # 纯文本正文
                html_body: str,            # HTML 正文
                pdf_attachments: list,     # PDF 附件列表 [{"filename": str, "content": bytes}]
                all_attachments: list,     # 所有附件列表 [{"filename": str, "content": bytes, "content_type": str}]
                message_id_header: str,    # Message-ID（用于回复时保持线程）
                references: str,           # References（用于回复时保持线程）
            }
        """
        msg = email.message_from_bytes(raw_email)

        # 提取头部信息
        headers = self._extract_headers(msg)

        # 提取正文
        body, html_body = self._extract_body(msg)

        # 提取附件
        all_attachments = self._extract_attachments(msg)

        # 过滤 PDF 附件
        pdf_attachments = [
            {"filename": a["filename"], "content": a["content"]}
            for a in all_attachments
            if a["content_type"] == "application/pdf"
        ]

        logger.info(
            f"邮件解析完成: 主题=[{headers['subject']}], "
            f"正文长度={len(body)}, HTML长度={len(html_body)}, "
            f"附件数={len(all_attachments)}, PDF数={len(pdf_attachments)}"
        )

        return {
            **headers,
            "body": body,
            "html_body": html_body,
            "pdf_attachments": pdf_attachments,
            "all_attachments": all_attachments,
        }

    def _extract_body(self, msg: email.message.Message) -> tuple[str, str]:
        """
        提取邮件正文（纯文本 + HTML）

        处理多种情况：
          - 简单文本邮件：直接返回 text/plain
          - multipart/alternative：优先取 text/plain，备选 text/html
          - multipart/mixed：递归遍历所有 part

        Returns:
            tuple[str, str]: (plain_text, html_text)
        """
        plain_text = ""
        html_text = ""

        if msg.is_multipart():
            for part in msg.walk():
                content_type = part.get_content_type()
                content_disposition = str(part.get("Content-Disposition", ""))

                # 跳过附件
                if "attachment" in content_disposition:
                    continue

                if content_type == "text/plain" and not plain_text:
                    payload = part.get_payload(decode=True)
                    if payload:
                        charset = part.get_content_charset() or "utf-8"
                        try:
                            plain_text = payload.decode(charset, errors="replace")
                        except (LookupError, UnicodeDecodeError):
                            plain_text = payload.decode("utf-8", errors="replace")

                elif content_type == "text/html" and not html_text:
                    payload = part.get_payload(decode=True)
                    if payload:
                        charset = part.get_content_charset() or "utf-8"
                        try:
                            html_text = payload.decode(charset, errors="replace")
                        except (LookupError, UnicodeDecodeError):
                            html_text = payload.decode("utf-8", errors="replace")
        else:
            # 非 multipart 邮件
            payload = msg.get_payload(decode=True)
            if payload:
                charset = msg.get_content_charset() or "utf-8"
                try:
                    text = payload.decode(charset, errors="replace")
                except (LookupError, UnicodeDecodeError):
                    text = payload.decode("utf-8", errors="replace")

                if msg.get_content_type() == "text/html":
                    html_text = text
                else:
                    plain_text = text

        return plain_text, html_text

    def _extract_attachments(self, msg: email.message.Message) -> list[dict]:
        """
        提取所有附件

        Returns:
            list[dict]: [{"filename": str, "content": bytes, "content_type": str}]
        """
        attachments = []

        if not msg.is_multipart():
            return attachments

        for part in msg.walk():
            content_disposition = str(part.get("Content-Disposition", ""))

            # 只处理附件（跳过内联内容）
            if "attachment" not in content_disposition:
                continue

            filename = part.get_filename()
            if not filename:
                continue

            filename = self._decode_header_value(filename)
            content_type = part.get_content_type()
            payload = part.get_payload(decode=True)

            if payload:
                attachments.append({
                    "filename": filename,
                    "content": payload,
                    "content_type": content_type,
                })
                logger.debug(f"提取附件: {filename} ({content_type}, {len(payload)} bytes)")

        return attachments

    def _extract_headers(self, msg: email.message.Message) -> dict:
        """
        提取邮件头部信息

        Returns:
            dict: {subject, from_addr, from_name, date, message_id_header, references}
        """
        subject = self._decode_header_value(msg.get("Subject", ""))
        from_header = msg.get("From", "")
        from_name, from_addr = self._parse_from_header(from_header)
        date = msg.get("Date", "")
        message_id_header = msg.get("Message-ID", "")
        references = msg.get("References", "")
        # 真实发件人优先级：X-Sender > Return-Path > From
        # 163/QQ 等邮箱会把真实发件人放在 X-Sender
        x_sender = msg.get("X-Sender", "").strip().strip("<>")
        return_path = msg.get("Return-Path", "").strip().strip("<>")
        real_sender = x_sender or return_path or from_addr

        return {
            "subject": subject,
            "from_addr": from_addr,       # From 头部（可能是伪装的）
            "from_name": from_name,
            "reply_to": real_sender,      # 回复地址用真实发件人
            "date": date,
            "message_id_header": message_id_header,
            "references": references,
        }

    def _decode_header_value(self, value: str) -> str:
        """
        解码邮件头部字段（处理 =?charset?encoding?text?= 编码）

        Args:
            value: 原始头部值

        Returns:
            str: 解码后的 UTF-8 字符串
        """
        if not value:
            return ""
        decoded_parts = decode_header(value)
        result = []
        for part, charset in decoded_parts:
            if isinstance(part, bytes):
                charset = charset or "utf-8"
                try:
                    result.append(part.decode(charset, errors="replace"))
                except (LookupError, UnicodeDecodeError):
                    result.append(part.decode("utf-8", errors="replace"))
            else:
                result.append(part)
        return "".join(result)

    def _parse_from_header(self, from_header: str) -> tuple[str, str]:
        """
        解析 From 头部，提取名称和邮箱

        Args:
            from_header: 原始 From 值（如 '"张三" <zhangsan@example.com>'）

        Returns:
            tuple[str, str]: (name, email_addr)
        """
        from_header = self._decode_header_value(from_header)
        if "<" in from_header and ">" in from_header:
            name = from_header[:from_header.index("<")].strip().strip('"')
            addr = from_header[from_header.index("<") + 1:from_header.index(">")].strip()
            return name, addr
        return "", from_header.strip()
