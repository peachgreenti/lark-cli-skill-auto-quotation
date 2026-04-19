"""
飞书通知模块
通过 lark-cli 发送飞书消息通知
"""

import json
import subprocess
import logging

logger = logging.getLogger("notifier")


class FeishuNotifier:
    """飞书通知器 - 通过 lark-cli im +messages-send"""

    def __init__(self, config: dict):
        notify_config = config.get("notify", {})
        self.chat_id = notify_config.get("chat_id")
        self.salesperson_id = notify_config.get("salesperson_id")

    def _run_cli(self, args: list[str]) -> dict:
        """执行 lark-cli 命令"""
        cmd = ["lark-cli"] + args
        logger.debug(f"执行命令: {' '.join(cmd)}")
        try:
            result = subprocess.run(cmd, capture_output=True, text=True, timeout=30)
            if result.returncode != 0:
                logger.error(f"lark-cli 命令失败: {result.stderr}")
                return {"ok": False, "error": result.stderr}
            return json.loads(result.stdout)
        except Exception as e:
            logger.error(f"lark-cli 命令异常: {e}")
            return {"ok": False, "error": str(e)}

    def send_text(self, text: str, receive_id: str = None) -> bool:
        """
        发送文本消息

        Args:
            text: 消息文本
            receive_id: 接收者 ID（默认使用 chat_id）

        Returns:
            bool: 是否发送成功
        """
        target_id = receive_id or self.chat_id
        if not target_id:
            logger.warning("未配置通知目标 ID，跳过发送")
            return False

        result = self._run_cli([
            "im", "+messages-send",
            "--chat-id", target_id,
            "--text", text,
        ])
        if result.get("ok"):
            logger.info(f"✅ 飞书消息已发送: {text[:50]}...")
            return True
        else:
            logger.error(f"飞书消息发送失败: {result}")
            return False

    def send_markdown(self, title: str, content: str, receive_id: str = None) -> bool:
        """
        发送 Markdown 消息

        Args:
            title: 消息标题
            content: Markdown 内容
            receive_id: 接收者 ID

        Returns:
            bool: 是否发送成功
        """
        target_id = receive_id or self.chat_id
        if not target_id:
            return False

        # 构造 post 格式的 content JSON
        content_json = json.dumps({
            "zh_cn": {
                "title": title,
                "content": [
                    [{"tag": "md", "text": content}]
                ]
            }
        }, ensure_ascii=False)

        result = self._run_cli([
            "im", "+messages-send",
            "--chat-id", target_id,
            "--msg-type", "post",
            "--content", content_json,
        ])
        if result.get("ok"):
            logger.info(f"✅ 飞书 Markdown 消息已发送: {title}")
            return True
        else:
            logger.error(f"飞书 Markdown 消息发送失败: {result}")
            return False

    def send_inquiry_received(self, customer_name: str, subject: str, record_url: str = "") -> bool:
        """发送询价收到通知"""
        content = f"**客户：** {customer_name}\n**主题：** {subject}\n**状态：** 已录入多维表格，等待报价..."
        if record_url:
            content += f"\n\n[查看多维表格记录]({record_url})"
        return self.send_markdown(title="📥 收到新询价", content=content)

    def send_quotation_ready(
        self,
        customer_name: str,
        doc_url: str,
        total_amount: float = 0,
        record_url: str = "",
    ) -> bool:
        """
        发送报价单已生成通知
        包含：多维表格链接 + 报价单云文档链接，提示用户去云文档调整确认
        """
        content = (
            f"**客户：** {customer_name}\n"
            f"**总金额：** {total_amount:,.2f}\n\n"
            f"报价单已生成，请前往云文档查看并调整内容：\n"
            f"📄 [编辑报价单云文档]({doc_url})\n"
        )
        if record_url:
            content += f"📊 [查看多维表格记录]({record_url})\n"
        content += (
            f"\n"
            f"⚠️ **请先在云文档中调整报价内容，确认无误后在群内回复「确认」**"
        )
        return self.send_markdown(title="📊 报价单已生成，请调整确认", content=content)

    def send_reply_sent(self, customer_name: str, quotation_number: str) -> bool:
        """发送回复已发送通知"""
        return self.send_markdown(
            title="✅ 报价邮件已发送",
            content=f"**客户：** {customer_name}\n**报价编号：** {quotation_number}\n报价邮件已发送，请关注客户回复。",
        )

    def send_error(self, error_msg: str) -> bool:
        """发送错误告警"""
        return self.send_markdown(
            title="⚠️ 系统错误",
            content=error_msg,
        )
