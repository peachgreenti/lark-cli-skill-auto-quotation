"""
邮件监听模块
使用 Python imaplib 监听网易邮箱收件箱
功能：连接 IMAP、搜索未读邮件、提取邮件原始数据、标记已读、自动重连
"""

import os
import imaplib
import email
import logging
from typing import Optional
from email.header import decode_header

logger = logging.getLogger("mail_watcher")


class MailWatcher:
    """邮件监听器 - IMAP 收件箱监听"""

    def __init__(self, config: dict):
        self.imap_server = os.getenv("IMAP_SERVER")
        self.imap_port = int(os.getenv("IMAP_PORT", 993))
        self.imap_user = os.getenv("IMAP_USER")
        self.imap_password = os.getenv("IMAP_PASSWORD")
        self.watch_folders = config.get("mail", {}).get("watch_folders", ["INBOX"])
        self.connection: Optional[imaplib.IMAP4_SSL] = None
        self.processed_uids: set[str] = set()

    def connect(self) -> bool:
        """
        建立 IMAP SSL 连接

        Returns:
            bool: 连接是否成功
        """
        try:
            logger.info(f"正在连接 IMAP 服务器 {self.imap_server}:{self.imap_port}...")
            self.connection = imaplib.IMAP4_SSL(self.imap_server, self.imap_port)
            # 网易邮箱要求先发送 ID 命令，否则 SELECT 会被拒绝
            self._send_id_command()
            self.connection.login(self.imap_user, self.imap_password)
            logger.info("IMAP 连接成功")
            return True
        except imaplib.IMAP4.error as e:
            logger.error(f"IMAP 登录失败: {e}")
            return False
        except Exception as e:
            logger.error(f"IMAP 连接失败: {e}")
            return False

    def disconnect(self) -> None:
        """断开 IMAP 连接"""
        if self.connection:
            try:
                self.connection.logout()
            except Exception:
                try:
                    self.connection.close()
                except Exception:
                    pass
            self.connection = None
            logger.info("IMAP 连接已断开")

    def _send_id_command(self) -> None:
        """
        发送 IMAP ID 命令（网易邮箱要求，否则 SELECT 会被拒绝为 Unsafe Login）
        """
        try:
            tag = self.connection._new_tag()
            if isinstance(tag, bytes):
                tag_str = tag.decode()
            else:
                tag_str = tag
            self.connection.send(
                f'{tag_str} ID ("name" "auto-quotation" "version" "1.0.0")\r\n'.encode()
            )
            # 读取响应直到看到 tagged OK
            while True:
                line = self.connection.readline()
                if tag_str.encode() in line:
                    break
            logger.debug("IMAP ID 命令发送成功")
        except Exception as e:
            logger.warning(f"IMAP ID 命令发送失败（非致命）: {e}")

    def _ensure_connected(self) -> bool:
        """
        确保连接可用，断开则自动重连

        Returns:
            bool: 连接是否可用
        """
        if self.connection is None:
            return self.connect()
        try:
            # 发送 NOOP 检测连接是否存活
            self.connection.noop()
            return True
        except Exception:
            logger.warning("IMAP 连接已断开，尝试重新连接...")
            self.connection = None
            return self.connect()

    def check_new_emails(self) -> list[dict]:
        """
        检查指定文件夹中的未读邮件

        Returns:
            list[dict]: 新邮件列表，每封包含:
                - uid: str           # 邮件 UID（唯一标识）
                - subject: str       # 邮件主题
                - from_addr: str     # 发件人邮箱
                - from_name: str     # 发件人名称
                - date: str          # 邮件日期
                - raw_email: bytes   # 邮件原始字节（供 mail_parser 解析）
        """
        if not self._ensure_connected():
            return []

        results = []
        for folder in self.watch_folders:
            folder_emails = self._check_folder(folder)
            results.extend(folder_emails)

        if results:
            logger.info(f"发现 {len(results)} 封新邮件")
        return results

    def _check_folder(self, folder: str) -> list[dict]:
        """
        检查单个文件夹的未读邮件

        Args:
            folder: 文件夹名称（如 "INBOX"）

        Returns:
            list[dict]: 新邮件列表
        """
        try:
            status, _ = self.connection.select(folder, readonly=True)
            if status != "OK":
                logger.warning(f"无法选择文件夹: {folder}")
                return []

            # 搜索未读邮件
            status, data = self.connection.uid("search", None, "UNSEEN")
            if status != "OK" or not data[0]:
                return []

            uid_list = data[0].split()
            new_uids = [uid.decode() for uid in uid_list if uid.decode() not in self.processed_uids]

            if not new_uids:
                return []

            results = []
            for uid in new_uids:
                try:
                    email_data = self._fetch_email(uid)
                    if email_data:
                        results.append(email_data)
                        self.processed_uids.add(uid)
                except Exception as e:
                    logger.error(f"获取邮件 UID={uid} 失败: {e}")

            return results

        except Exception as e:
            logger.error(f"检查文件夹 {folder} 失败: {e}")
            return []

    def _fetch_email(self, uid: str) -> Optional[dict]:
        """
        获取单封邮件的原始数据

        Args:
            uid: 邮件 UID

        Returns:
            Optional[dict]: 邮件数据字典
        """
        status, data = self.connection.uid("fetch", uid, "(RFC822)")
        if status != "OK" or not data or not data[0]:
            return None

        raw_email = data[0][1]
        msg = email.message_from_bytes(raw_email)

        # 提取头部信息
        subject = self._decode_header_value(msg.get("Subject", ""))
        from_header = msg.get("From", "")
        from_name, from_addr = self._parse_from_header(from_header)
        # X-Sender 是 163/QQ 等邮箱的真实发件人（From 可能被伪装）
        x_sender = msg.get("X-Sender", "").strip().strip("<>")
        if x_sender:
            from_addr = x_sender
            logger.info(f"  📧 使用真实发件人 X-Sender: {from_addr}")
        date = msg.get("Date", "")

        logger.info(f"  📧 新邮件: [{subject}] 来自 {from_addr}")

        return {
            "uid": uid,
            "subject": subject,
            "from_addr": from_addr,
            "from_name": from_name,
            "date": date,
            "raw_email": raw_email,
        }

    def mark_as_seen(self, uid: str) -> bool:
        """
        标记邮件为已读

        Args:
            uid: 邮件 UID

        Returns:
            bool: 是否成功
        """
        try:
            status, _ = self.connection.uid("store", uid, "+FLAGS", "\\Seen")
            if status == "OK":
                logger.info(f"已标记邮件 UID={uid} 为已读")
                return True
            return False
        except Exception as e:
            logger.error(f"标记已读失败 UID={uid}: {e}")
            return False

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
