"""
邮件发送模块
通过 SMTP 回复询价邮件，附上报价单 PDF 附件
"""

import os
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
import logging
from datetime import datetime

logger = logging.getLogger("mail_sender")


class MailSender:
    """邮件发送器 - SMTP 回复 + PDF 附件"""

    def __init__(self, config: dict):
        self.smtp_server = os.getenv("SMTP_SERVER")
        self.smtp_port = int(os.getenv("SMTP_PORT", 465))
        self.smtp_user = os.getenv("SMTP_USER")
        self.smtp_password = os.getenv("SMTP_PASSWORD")
        q = config.get("quotation", {})
        self.company_name_en = q.get("company_name_en", "")
        self.company_name_cn = q.get("company_name_cn", "")
        self.company_email = q.get("company_email", "")
        self.company_phone = q.get("company_phone", "")
        self.validity_days = q.get("validity_days", 30)

    def generate_reply_content(self, customer_name: str, quotation_number: str, language: str = "zh") -> str:
        """
        生成回复邮件内容（HTML 格式）

        Returns:
            str: HTML 邮件正文
        """
        if language == "en":
            return self._get_reply_body_en(customer_name, quotation_number)
        return self._get_reply_body_zh(customer_name, quotation_number)

    def send_reply(
        self,
        to_addr: str,
        subject: str,
        reply_body: str,
        pdf_path: str,
        original_message_id: str = None,
        references: str = None,
    ) -> bool:
        """
        回复原邮件 + 附上 PDF 附件

        Returns:
            bool: 是否发送成功
        """
        msg = MIMEMultipart("mixed")
        msg["From"] = self.smtp_user
        msg["To"] = to_addr
        msg["Subject"] = subject
        msg["Date"] = datetime.now().strftime("%a, %d %b %Y %H:%M:%S +0800")

        # 保持邮件线程
        if original_message_id:
            msg["In-Reply-To"] = original_message_id
        if references:
            msg["References"] = references
        elif original_message_id:
            msg["References"] = original_message_id

        # HTML 正文
        msg.attach(MIMEText(reply_body, "html", "utf-8"))

        # PDF 附件
        if pdf_path and os.path.exists(pdf_path):
            with open(pdf_path, "rb") as f:
                pdf_attachment = MIMEApplication(f.read(), _subtype="pdf")
                filename = os.path.basename(pdf_path)
                pdf_attachment.add_header(
                    "Content-Disposition", "attachment", filename=("utf-8", "", filename)
                )
                msg.attach(pdf_attachment)
            logger.info(f"已添加 PDF 附件: {filename}")
        else:
            logger.warning(f"PDF 文件不存在: {pdf_path}")

        return self._send_smtp(msg)

    def _send_smtp(self, msg: MIMEMultipart) -> bool:
        """通过 SMTP SSL 发送邮件"""
        try:
            with smtplib.SMTP_SSL(self.smtp_server, self.smtp_port) as server:
                server.login(self.smtp_user, self.smtp_password)
                server.sendmail(self.smtp_user, msg["To"], msg.as_string())
            logger.info(f"✅ 邮件已发送至 {msg['To']}")
            return True
        except Exception as e:
            logger.error(f"邮件发送失败: {e}")
            return False

    def _get_reply_body_zh(self, customer_name: str, quotation_number: str) -> str:
        """中文回复邮件模板"""
        return f"""<!DOCTYPE html>
<html><head><meta charset="utf-8"></head><body style="font-family: 'Microsoft YaHei', sans-serif; margin: 20px; color: #333;">
<p>尊敬的 {customer_name}，</p>
<p>感谢贵司的询价！附件为我司的正式报价单（编号：{quotation_number}），请查收。</p>
<p>报价有效期为本邮件发出之日起 <strong>{self.validity_days} 天</strong>。如有任何疑问，欢迎随时联系我们。</p>
<p>期待与贵司的合作！</p>
<br>
<p>此致<br>敬礼</p>
<p><strong>{self.company_name_cn}</strong><br>
联系方式：{self.company_phone}<br>
邮箱：{self.company_email}</p>
</body></html>"""

    def _get_reply_body_en(self, customer_name: str, quotation_number: str) -> str:
        """英文回复邮件模板"""
        return f"""<!DOCTYPE html>
<html><head><meta charset="utf-8"></head><body style="font-family: Arial, sans-serif; margin: 20px; color: #333;">
<p>Dear {customer_name},</p>
<p>Thank you for your inquiry! Please find attached our formal quotation (No. {quotation_number}) for your reference.</p>
<p>This quotation is valid for <strong>{self.validity_days} days</strong> from the date of this email. Should you have any questions, please feel free to contact us.</p>
<p>We look forward to the opportunity of working with you!</p>
<br>
<p>Best regards,</p>
<p><strong>{self.company_name_en}</strong><br>
Tel: {self.company_phone}<br>
Email: {self.company_email}</p>
</body></html>"""
