"""
询价识别模块
使用 AI 识别邮件是否为询价邮件，并提取客户信息
流程：关键词预筛 → AI 二次确认 + 信息提取
"""

import os
import json
import re
import logging
from openai import OpenAI

logger = logging.getLogger("inquiry_detector")


class InquiryDetector:
    """询价识别器 - AI 分类 + 客户信息提取"""

    def __init__(self, config: dict):
        ai_config = config.get("ai", {})
        self.model = ai_config.get("model", "doubao-1-5-pro-32k-250115")
        api_key_env = ai_config.get("api_key_env", "OPENAI_API_KEY")
        api_key = os.getenv(api_key_env, "")
        base_url = os.getenv(ai_config.get("base_url_env", "OPENAI_BASE_URL"), "")
        self.client = OpenAI(api_key=api_key, base_url=base_url or None)
        self.keywords = config.get("mail", {}).get("inquiry_keywords", [])

    def detect(self, subject: str, body: str) -> dict:
        """
        识别邮件是否为询价邮件，并提取客户信息

        Args:
            subject: 邮件主题
            body: 邮件正文

        Returns:
            dict: {
                is_inquiry: bool,          # 是否为询价邮件
                confidence: float,         # 置信度 0-1
                customer_name: str,        # 客户公司名称
                contact_person: str,       # 联系人
                contact_email: str,        # 联系邮箱
                contact_phone: str,        # 联系电话
                language: str,             # "zh" / "en"
                products: list,            # [{"name": str, "model": str, "quantity": int}]
                reason: str,               # 判断理由（调试用）
            }
        """
        # 1. 关键词预筛
        if not self._keyword_pre_filter(subject, body):
            logger.info(f"关键词预筛未通过，判定为非询价邮件: [{subject}]")
            return self._build_negative_result("关键词预筛未通过")

        # 2. AI 二次确认 + 信息提取
        logger.info(f"关键词命中，调用 AI 二次确认: [{subject}]")
        result = self._ai_detect(subject, body)

        if result["is_inquiry"]:
            logger.info(
                f"✅ AI 确认为询价邮件 (置信度={result['confidence']}), "
                f"客户={result['customer_name']}, 语言={result['language']}"
            )
        else:
            logger.info(f"❌ AI 判定为非询价邮件 (置信度={result['confidence']}), 理由={result['reason']}")

        return result

    def _keyword_pre_filter(self, subject: str, body: str) -> bool:
        """
        关键词预筛（不区分大小写）

        Returns:
            bool: 是否包含询价关键词
        """
        text = (subject + " " + body).lower()
        for keyword in self.keywords:
            if keyword.lower() in text:
                logger.debug(f"命中关键词: {keyword}")
                return True
        return False

    def _ai_detect(self, subject: str, body: str) -> dict:
        """
        AI 询价识别 + 客户信息提取

        Returns:
            dict: 同 detect() 返回格式
        """
        prompt = self._build_prompt(subject, body)

        try:
            response = self.client.chat.completions.create(
                model=self.model,
                messages=[
                    {"role": "system", "content": "你是一个外贸询价邮件识别助手，请严格按照 JSON 格式输出。"},
                    {"role": "user", "content": prompt},
                ],
                temperature=0.1,
                max_tokens=2000,
            )

            content = response.choices[0].message.content.strip()
            # 提取 JSON（可能被 ```json 包裹）
            json_match = re.search(r'\{[\s\S]*\}', content)
            if not json_match:
                logger.error(f"AI 返回内容无法解析为 JSON: {content[:200]}")
                return self._build_negative_result("AI 返回格式异常")

            result = json.loads(json_match.group())
            return self._normalize_result(result)

        except json.JSONDecodeError as e:
            logger.error(f"AI 返回 JSON 解析失败: {e}")
            return self._build_negative_result(f"JSON 解析失败: {e}")
        except Exception as e:
            logger.error(f"AI 调用失败: {e}")
            return self._build_negative_result(f"AI 调用失败: {e}")

    def _build_prompt(self, subject: str, body: str) -> str:
        """构造 AI prompt"""
        # 截取正文前 3000 字符，避免 token 过长
        body_preview = body[:3000] if body else "(无正文内容)"
        return f"""请分析以下邮件，判断是否为**外贸询价邮件**，并提取关键信息。

邮件主题：{subject}

邮件正文：
{body_preview}

请严格按照以下 JSON 格式输出（不要输出其他内容）：
{{
  "is_inquiry": true/false,
  "confidence": 0.0-1.0,
  "reason": "判断理由（一句话）",
  "customer_name": "客户公司名称（如无法确定则填空字符串）",
  "contact_person": "联系人姓名（如无法确定则填空字符串）",
  "contact_email": "联系邮箱（如无法确定则填空字符串）",
  "contact_phone": "联系电话（如无法确定则填空字符串）",
  "language": "zh 或 en（邮件主要使用的语言）",
  "products": [
    {{"name": "产品名称", "model": "型号规格", "quantity": 数量}}
  ]
}}

判断标准：
- 包含产品询价、报价请求、RFQ、价格咨询等意图 → is_inquiry=true
- 营销邮件、通知邮件、验证码等 → is_inquiry=false
- 如果邮件只是询问产品信息但没有明确询价意图，confidence 设为 0.5 以下"""

    def _normalize_result(self, raw: dict) -> dict:
        """
        标准化 AI 返回结果，填充默认值

        Args:
            raw: AI 返回的原始 dict

        Returns:
            dict: 标准化后的结果
        """
        return {
            "is_inquiry": bool(raw.get("is_inquiry", False)),
            "confidence": float(raw.get("confidence", 0.0)),
            "reason": str(raw.get("reason", "")),
            "customer_name": str(raw.get("customer_name", "")),
            "contact_person": str(raw.get("contact_person", "")),
            "contact_email": str(raw.get("contact_email", "")),
            "contact_phone": str(raw.get("contact_phone", "")),
            "language": self._detect_language(
                raw.get("language", "") + " " + raw.get("customer_name", "")
            ),
            "products": raw.get("products", []),
        }

    def _detect_language(self, text: str) -> str:
        """
        检测文本语言（简单规则：是否包含中文字符）

        Returns:
            str: "zh" 或 "en"
        """
        if not text:
            return "en"
        # 统计中文字符占比
        chinese_chars = len(re.findall(r'[\u4e00-\u9fff]', text))
        if chinese_chars > len(text) * 0.1:
            return "zh"
        return "en"

    def _build_negative_result(self, reason: str) -> dict:
        """构造非询价结果"""
        return {
            "is_inquiry": False,
            "confidence": 0.0,
            "reason": reason,
            "customer_name": "",
            "contact_person": "",
            "contact_email": "",
            "contact_phone": "",
            "language": "zh",
            "products": [],
        }

    def extract_requirement_text(self, pdf_text: str) -> str:
        """
        从询价函文本中提取报价需求文字

        按照用户指定的格式输出，如："3个CP001、2个SP002、8个SM902"

        Args:
            pdf_text: PDF 询价函的文本内容

        Returns:
            str: 提取的报价需求文字
        """
        prompt = f"""你是一个专业的文本信息提取员，能够准确读取询价函内容，提取出询盘需求，包括产品型号和数量，并按照特定格式输出。

## 技能
1. 仔细阅读询价函的文本内容。
2. 识别并提取其中的询盘需求，明确产品型号和数量。
3. 将提取的信息按照"C个产品型号"的格式进行整理，多个需求之间用中文顿号隔开。

## 限制
- 严格按照"C个产品型号"的格式输出提取的询盘需求。
- 多个需求之间必须用中文顿号隔开。
- 只输出产品型号，不输出中文的产品名称。
- 输出内容仅包含提取的询盘需求，不做其他无关说明。如3个CP001、2个SP002、8个SM902

## 用户要求
请读取以下询价函的内容，提取出询盘需求，包括产品型号和数量，按照"C个产品型号"的格式输出，多个需求之间用中文顿号隔开。

询价函内容：
{pdf_text[:4000]}"""

        try:
            response = self.client.chat.completions.create(
                model=self.model,
                messages=[
                    {"role": "system", "content": "你是一个专业的文本信息提取员，只输出提取结果，不做任何其他说明。"},
                    {"role": "user", "content": prompt},
                ],
                temperature=0.1,
                max_tokens=500,
            )
            result = response.choices[0].message.content.strip()
            logger.info(f"AI 提取报价需求: {result}")
            return result
        except Exception as e:
            logger.error(f"AI 提取报价需求失败: {e}")
            return ""
