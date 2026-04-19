"""
外贸报价全链路自动化系统 - 主入口

完整流程：
  ① IMAP 监听收件箱 → 检查新邮件
  ② AI 识别询价邮件 + 提取客户信息
  ③ 下载 PDF 附件 + AI 提取报价需求
  ④ lark-cli 写入多维表格（询盘需求 + 客户信息 + 附件上传）
  ⑤ 多维表格自动处理（已有功能）
  ⑥ 轮询检测报价完成（60 秒稳定）
  ⑦ 生成报价单（飞书云文档 + PDF）
  ⑧ 飞书通知：报价单已生成，请确认
  ⑨ 用户确认 → 生成回复邮件 → 飞书通知：请确认邮件内容
  ⑩ 用户确认 → 发送回复邮件 + PDF 附件
"""

import os
import sys
import time
import logging
import subprocess
import yaml
from dotenv import load_dotenv

from mail_watcher import MailWatcher
from mail_parser import MailParser
from inquiry_detector import InquiryDetector
from base_writer import BaseWriter
from quotation_poller import QuotationPoller
from quotation_generator import QuotationGenerator
from mail_sender import MailSender
from notifier import FeishuNotifier

# 配置日志
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
)
logger = logging.getLogger("main")

# 全局状态：记录正在处理的询价
# key: inquiry_record_id, value: 上下文数据
processing_inquiries: dict = {}


def load_config(config_path: str = "config.yaml") -> dict:
    """加载 YAML 配置文件"""
    base_dir = os.path.dirname(os.path.abspath(__file__))
    config_file = os.path.join(base_dir, config_path)
    with open(config_file, "r", encoding="utf-8") as f:
        return yaml.safe_load(f)


def wait_for_feishu_confirm(chat_id: str, prompt: str, timeout: int = 1800) -> bool:
    """
    在飞书群中发送提示，等待用户回复「确认」或「取消」

    Args:
        chat_id: 飞书群聊 ID
        prompt: 提示消息
        timeout: 超时秒数（默认 30 分钟）

    Returns:
        bool: 用户是否确认
    """
    logger = logging.getLogger("main")
    notifier = FeishuNotifier({"notify": {"chat_id": chat_id}})

    # 发送提示消息
    notifier.send_markdown(title="⏳ 等待确认", content=prompt)

    # 轮询群消息，等待用户回复
    start_time = time.time()
    poll_interval = 10  # 每 10 秒检查一次
    start_time_str = time.strftime("%Y-%m-%d %H:%M", time.localtime(start_time))

    while time.time() - start_time < timeout:
        time.sleep(poll_interval)
        try:
            r = subprocess.run(
                ["lark-cli", "im", "+chat-messages-list",
                 "--chat-id", chat_id,
                 "--page-size", "5",
                 "--sort", "desc"],
                capture_output=True, text=True, timeout=15,
            )
            if r.returncode != 0:
                continue

            import json
            data = json.loads(r.stdout)
            messages = data.get("data", {}).get("messages", [])

            for msg in messages:
                # 只检查开始时间之后的消息
                msg_time = msg.get("create_time", "")
                if msg_time < start_time_str:
                    continue
                # 只检查文本消息（忽略 post 类型）
                if msg.get("msg_type") != "text":
                    continue
                # 提取纯文本
                content = msg.get("content", "")
                try:
                    msg_text = json.loads(content).get("text", "")
                except (json.JSONDecodeError, TypeError):
                    msg_text = str(content)

                msg_text = msg_text.strip().lower()
                if msg_text in ("确认", "yes", "y", "ok", "confirm"):
                    logger.info("用户在飞书群中确认")
                    return True
                elif msg_text in ("取消", "cancel", "no", "n"):
                    logger.info("用户在飞书群中取消")
                    return False

        except Exception as e:
            logger.warning(f"轮询飞书群消息异常: {e}")
            continue

    logger.warning(f"等待飞书群确认超时（{timeout}秒）")
    return False


def process_inquiry(email_data: dict, parsed: dict, detected: dict, config: dict, modules: dict) -> bool:
    """
    处理单封询价邮件的完整流程（阶段 ③-⑩）

    Args:
        email_data: 原始邮件数据（来自 mail_watcher）
        parsed: 解析后的邮件数据（来自 mail_parser）
        detected: AI 识别结果（来自 inquiry_detector）
        config: 全局配置
        modules: 各模块实例字典 {"writer", "poller", "generator", "sender", "notifier"}

    Returns:
        bool: 处理是否成功
    """
    writer = modules["writer"]
    poller = modules["poller"]
    generator = modules["generator"]
    sender = modules["sender"]
    notifier = modules["notifier"]
    detector = modules["detector"]

    customer_name = detected.get("customer_name", parsed.get("from_name", ""))
    language = detected.get("language", "zh")
    logger.info(f"{'='*60}")
    logger.info(f"开始处理询价: 客户={customer_name}, 语言={language}")

    # ---- 阶段 ④：写入多维表格 ----
    logger.info("【阶段④】写入多维表格...")
    try:
        # 保存 PDF 附件到临时文件
        pdf_path = None
        pdf_text = ""
        if parsed.get("pdf_attachments"):
            pdf = parsed["pdf_attachments"][0]
            os.makedirs("data/temp", exist_ok=True)
            pdf_path = os.path.join("data/temp", pdf["filename"])
            with open(pdf_path, "wb") as f:
                f.write(pdf["content"])
            logger.info(f"PDF 附件已保存: {pdf_path}")

            # 从 PDF 提取文本
            try:
                import fitz
                doc = fitz.open(pdf_path)
                pdf_text = "\n".join(page.get_text() for page in doc)
                doc.close()
                logger.info(f"PDF 文本已提取: {len(pdf_text)} 字符")
            except Exception as e:
                logger.warning(f"PDF 文本提取失败: {e}")

        if not pdf_path:
            logger.error("未找到 PDF 附件，无法创建询盘记录")
            notifier.send_error(f"未找到 PDF 附件: {parsed.get('subject', '')}")
            return False

        # AI 提取报价需求文字（从 PDF 文本中提取）
        requirement_text = ""
        if pdf_text:
            requirement_text = detector.extract_requirement_text(pdf_text)
            logger.info(f"AI 提取报价需求: {requirement_text}")
        elif parsed.get("body"):
            # 如果没有 PDF，从邮件正文提取
            requirement_text = detector.extract_requirement_text(parsed["body"])

        # 创建询盘记录（上传附件 + 填入报价需求文字）
        result = writer.create_inquiry_record(
            pdf_path=pdf_path,
            requirement_text=requirement_text,
        )
        inquiry_record_id = result.get("record_id", "")
        req_number = result.get("req_number", "")
        logger.info(f"询盘记录已创建: {inquiry_record_id}, 需求编号: {req_number}")

    except Exception as e:
        logger.error(f"写入多维表格失败: {e}")
        notifier.send_error(f"写入多维表格失败: {e}")
        return False

    # 通知：询价已收到（含多维表格链接）
    record_url = f"https://my.feishu.cn/base/VgB8bGECraIO4Ush1rJctNOhncb?table=tblWsxK0Tl6ProSp&view=rec{inquiry_record_id}"
    notifier.send_inquiry_received(customer_name, parsed.get("subject", ""), record_url)

    # ---- 阶段 ⑤⑥：等待报价完成 ----
    logger.info("【阶段⑤⑥】等待多维表格生成报价结果...")
    quotation_records = poller.poll(writer, inquiry_record_id, req_number=req_number)
    if not quotation_records:
        logger.warning("报价等待超时，跳过此询价")
        notifier.send_error(f"报价等待超时: 客户={customer_name}")
        return False

    # ---- 阶段 ⑦：生成报价单（飞书云文档 + PDF） ----
    logger.info("【阶段⑦】生成报价单...")
    try:
        gen_result = generator.generate(
            quotation_records=quotation_records,
            customer_name=customer_name,
            language=language,
        )
        doc_url = gen_result["doc_url"]
        pdf_path = gen_result["pdf_path"]
        quotation_number = gen_result["quotation_number"]
        total_amount = gen_result["total_amount"]
    except Exception as e:
        logger.error(f"生成报价单失败: {e}")
        notifier.send_error(f"生成报价单失败: {e}")
        return False

    # ---- 阶段 ⑧：通知用户去云文档调整报价 ----
    logger.info("【阶段⑧】通知用户调整报价单...")
    notifier.send_quotation_ready(customer_name, doc_url, total_amount, record_url)

    # 等待用户在群里确认（已调整完报价单）
    chat_id = config.get("notify", {}).get("chat_id", "")
    confirmed = wait_for_feishu_confirm(
        chat_id,
        f"**客户：** {customer_name}\n**报价金额：** {total_amount:,.2f}\n\n"
        f"请先前往云文档调整报价内容，调整完毕后在群内回复「**确认**」继续，回复「**取消**」放弃。\n"
        f"📄 [编辑报价单]({doc_url})",
    )
    if not confirmed:
        logger.info("用户未确认或取消，跳过此询价")
        return False

    # ---- 阶段 ⑨：生成邮件预览云文档，让用户调整确认 ----
    logger.info("【阶段⑨】生成邮件预览云文档...")
    reply_subject = f"Re: {parsed.get('subject', '')}"
    try:
        reply_doc = generator.generate_reply_doc(
            quotation_doc_url=doc_url,
            customer_name=customer_name,
            to_addr=parsed.get("reply_to", ""),
            language=language,
        )
        reply_doc_url = reply_doc["doc_url"]
    except Exception as e:
        logger.error(f"生成邮件预览云文档失败: {e}")
        notifier.send_error(f"生成邮件预览失败: {e}")
        return False

    # 通知用户去云文档调整邮件内容
    confirmed = wait_for_feishu_confirm(
        chat_id,
        f"**客户：** {customer_name}\n"
        f"**收件人：** {parsed.get('reply_to')}\n"
        f"**主题：** {reply_subject}\n"
        f"**附件：** {os.path.basename(pdf_path) if pdf_path else '无'}\n\n"
        f"邮件预览已生成，请前往云文档查看并调整邮件内容：\n"
        f"📄 [编辑邮件预览]({reply_doc_url})\n\n"
        f"调整完毕后在群内回复「**确认**」发送邮件，回复「**取消**」放弃。",
    )
    if not confirmed:
        logger.info("用户取消发送")
        return False

    # ---- 阶段 ⑩：从邮件预览云文档读取最终内容，发送邮件 ----
    logger.info("【阶段⑩】从邮件预览云文档读取最终内容，发送邮件...")
    reply_body = generator.generate_reply_from_doc(reply_doc_url, customer_name, language)

    success = sender.send_reply(
        to_addr=parsed.get("reply_to", ""),
        subject=reply_subject,
        reply_body=reply_body,
        pdf_path=pdf_path,
        original_message_id=parsed.get("message_id_header"),
        references=parsed.get("references"),
    )

    if success:
        notifier.send_reply_sent(customer_name, quotation_number)
        logger.info(f"✅ 询价处理完成: {customer_name}, 报价编号: {quotation_number}")
    else:
        notifier.send_error(f"邮件发送失败: 客户={customer_name}")

    return success


def run_test_mode(req_number: str, config: dict, modules: dict) -> bool:
    """
    测试模式：直接从已有的需求编号开始处理（跳过邮件监听）

    Args:
        req_number: 需求编号
        config: 配置
        modules: 模块字典
    """
    writer = modules["writer"]
    poller = modules["poller"]
    generator = modules["generator"]
    sender = modules["sender"]
    notifier = modules["notifier"]

    # 模拟 parsed 和 detected
    parsed = {
        "subject": f"询价测试 - 需求编号 {req_number}",
        "from_addr": "test@example.com",
        "from_name": "测试客户",
        "reply_to": "565041990@qq.com",  # 测试用真实发件人
        "message_id_header": "",
        "references": "",
    }
    detected = {
        "is_inquiry": True,
        "customer_name": "苹果贸易有限公司",
        "language": "zh",
    }

    logger.info(f"{'='*60}")
    logger.info(f"测试模式: 需求编号={req_number}, 客户={detected['customer_name']}")

    # 直接从阶段⑤⑥开始
    logger.info("【阶段⑤⑥】等待报价结果...")
    quotation_records = poller.poll(writer, "", req_number=req_number, skip_wait=True)
    if not quotation_records:
        logger.error(f"未找到需求编号 {req_number} 的报价记录")
        return False

    logger.info(f"找到 {len(quotation_records)} 条报价记录")
    # 调用后续流程（⑦⑧⑨⑩）
    return _process_quotation(parsed, detected, config, modules, quotation_records, "")


def _process_quotation(parsed, detected, config, modules, quotation_records, inquiry_record_id):
    """
    阶段⑦⑧⑨⑩的通用处理逻辑（正常模式和测试模式共用）
    """
    generator = modules["generator"]
    sender = modules["sender"]
    notifier = modules["notifier"]
    customer_name = detected.get("customer_name", parsed.get("from_name", ""))
    language = detected.get("language", "zh")
    chat_id = config.get("notify", {}).get("chat_id", "")

    # ---- 阶段 ⑦：生成报价单云文档（不含 PDF） ----
    logger.info("【阶段⑦】生成报价单云文档...")
    try:
        gen_result = generator.generate_doc(
            quotation_records=quotation_records,
            customer_name=customer_name,
            language=language,
        )
        doc_url = gen_result["doc_url"]
        quotation_number = gen_result["quotation_number"]
        total_amount = gen_result["total_amount"]
    except Exception as e:
        logger.error(f"生成报价单云文档失败: {e}")
        notifier.send_error(f"生成报价单失败: {e}")
        return False

    # ---- 阶段 ⑧：通知用户调整报价单云文档 ----
    record_url = ""
    if inquiry_record_id:
        record_url = f"https://my.feishu.cn/base/VgB8bGECraIO4Ush1rJctNOhncb?table=tblWsxK0Tl6ProSp&view=rec{inquiry_record_id}"
    notifier.send_quotation_ready(customer_name, doc_url, total_amount, record_url)

    confirmed = wait_for_feishu_confirm(
        chat_id,
        f"**客户：** {customer_name}\n**报价金额：** {total_amount:,.2f}\n\n"
        f"请先前往云文档调整报价内容，调整完毕后在群内回复「**确认**」继续，回复「**取消**」放弃。\n"
        f"📄 [编辑报价单]({doc_url})",
    )
    if not confirmed:
        logger.info("用户未确认或取消，跳过此询价")
        return False

    # ---- 阶段 ⑦②：从用户确认后的云文档生成 PDF ----
    logger.info("【阶段⑦②】从确认后的云文档生成 PDF...")
    try:
        pdf_result = generator.generate_pdf_from_doc(doc_url, customer_name, language)
        pdf_path = pdf_result["pdf_path"]
        quotation_number = pdf_result["quotation_number"]
        total_amount = pdf_result["total_amount"]
        logger.info(f"✅ PDF 已生成: {pdf_path}, 总价={total_amount}")
    except Exception as e:
        logger.error(f"从云文档生成 PDF 失败: {e}")
        notifier.send_error(f"生成 PDF 失败: {e}")
        return False

    # ---- 阶段 ⑨：生成邮件预览云文档 ----
    logger.info("【阶段⑨】生成邮件预览云文档...")
    reply_subject = f"Re: {parsed.get('subject', '')}"
    try:
        reply_doc = generator.generate_reply_doc(
            quotation_doc_url=doc_url,
            customer_name=customer_name,
            to_addr=parsed.get("reply_to", ""),
            language=language,
        )
        reply_doc_url = reply_doc["doc_url"]
    except Exception as e:
        logger.error(f"生成邮件预览云文档失败: {e}")
        notifier.send_error(f"生成邮件预览失败: {e}")
        return False

    confirmed = wait_for_feishu_confirm(
        chat_id,
        f"**客户：** {customer_name}\n"
        f"**收件人：** {parsed.get('reply_to')}\n"
        f"**主题：** {reply_subject}\n"
        f"**附件：** {os.path.basename(pdf_path) if pdf_path else '无'}\n\n"
        f"邮件预览已生成，请前往云文档查看并调整邮件内容：\n"
        f"📄 [编辑邮件预览]({reply_doc_url})\n\n"
        f"调整完毕后在群内回复「**确认**」发送邮件，回复「**取消**」放弃。",
    )
    if not confirmed:
        logger.info("用户取消发送")
        return False

    # ---- 阶段 ⑩：发送邮件 ----
    logger.info("【阶段⑩】从邮件预览云文档读取最终内容，发送邮件...")
    reply_body = generator.generate_reply_from_doc(reply_doc_url, customer_name, language)

    success = sender.send_reply(
        to_addr=parsed.get("reply_to", ""),
        subject=reply_subject,
        reply_body=reply_body,
        pdf_path=pdf_path,
        original_message_id=parsed.get("message_id_header"),
        references=parsed.get("references"),
    )

    if success:
        notifier.send_reply_sent(customer_name, quotation_number)
        logger.info(f"✅ 询价处理完成: {customer_name}, 报价编号: {quotation_number}")
    else:
        notifier.send_error(f"邮件发送失败: 客户={customer_name}")

    return success


def main():
    """主函数 - 启动邮件监听循环"""
    load_dotenv()
    config = load_config()

    # 检查是否为测试模式
    if len(sys.argv) >= 3 and sys.argv[1] == "--test":
        req_number = sys.argv[2]
        logger.info(f"测试模式: 需求编号={req_number}")

        detector = InquiryDetector(config)
        writer = BaseWriter(config)
        poller = QuotationPoller(config)
        generator = QuotationGenerator(config)
        sender = MailSender(config)
        notifier = FeishuNotifier(config)

        modules = {
            "writer": writer,
            "poller": poller,
            "generator": generator,
            "sender": sender,
            "notifier": notifier,
            "detector": detector,
        }

        run_test_mode(req_number, config, modules)
        return

    # 初始化各模块
    watcher = MailWatcher(config)
    parser = MailParser()
    detector = InquiryDetector(config)
    writer = BaseWriter(config)
    poller = QuotationPoller(config)
    generator = QuotationGenerator(config)
    sender = MailSender(config)
    notifier = FeishuNotifier(config)

    modules = {
        "writer": writer,
        "poller": poller,
        "generator": generator,
        "sender": sender,
        "notifier": notifier,
        "detector": detector,
    }

    check_interval = config.get("polling", {}).get("check_interval", 60)

    logger.info("=" * 60)
    logger.info("外贸报价全链路自动化系统启动")
    logger.info(f"邮件检查间隔: {check_interval}s")
    logger.info("=" * 60)

    # 连接邮箱
    if not watcher.connect():
        logger.error("无法连接邮箱，退出")
        sys.exit(1)

    try:
        while True:
            logger.info("-" * 40)
            logger.info("检查新邮件...")

            # 阶段 ①②：检查新邮件 + 解析 + 识别
            emails = watcher.check_new_emails()
            for email_data in emails:
                # 解析邮件
                parsed = parser.parse(email_data["raw_email"])

                # AI 识别是否为询价
                text_for_detect = parsed.get("body", "") or parsed.get("html_body", "")[:1000]
                detected = detector.detect(parsed["subject"], text_for_detect)

                if detected["is_inquiry"]:
                    logger.info(f"✅ 发现询价邮件: [{parsed['subject']}]")
                    # 标记已读
                    watcher.mark_as_seen(email_data["uid"])
                    # 处理询价
                    process_inquiry(email_data, parsed, detected, config, modules)
                else:
                    logger.info(f"⏭️ 非询价邮件，跳过: [{parsed['subject']}]")

            # 等待下次检查
            logger.info(f"等待 {check_interval}s 后再次检查...")
            time.sleep(check_interval)

    except KeyboardInterrupt:
        logger.info("收到中断信号，系统停止")
    finally:
        watcher.disconnect()


if __name__ == "__main__":
    main()
