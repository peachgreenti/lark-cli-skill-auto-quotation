"""
报价单生成模块
功能：
  1. 将报价明细数据转换为结构化列表
  2. 生成飞书云文档（Markdown 格式）
  3. 生成 PDF 报价单（HTML → weasyprint）
"""

import os
import json
import subprocess
import logging
from datetime import datetime, timedelta
from typing import Optional

logger = logging.getLogger("quotation_generator")

# 项目根目录
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(BASE_DIR, "data", "temp")


class QuotationGenerator:
    """报价单生成器 - 飞书文档 + PDF"""

    def __init__(self, config: dict):
        q = config.get("quotation", {})
        self.company_name_cn = q.get("company_name_cn")
        self.company_name_en = q.get("company_name_en")
        self.company_address_cn = q.get("company_address_cn")
        self.company_address_en = q.get("company_address_en")
        self.company_phone = q.get("company_phone")
        self.company_email = q.get("company_email")
        self.bank_name = q.get("bank_name")
        self.bank_account = q.get("bank_account")
        self.bank_swift = q.get("bank_swift")
        self.validity_days = q.get("validity_days", 30)
        self.delivery_days = q.get("delivery_days", 15)
        self.payment_terms_en = q.get("payment_terms_en")
        self.payment_terms_cn = q.get("payment_terms_cn")
        self.packaging_en = q.get("packaging_en")
        self.packaging_cn = q.get("packaging_cn")
        self.moq = q.get("moq", 100)
        self.currency = q.get("currency", "USD")

        # 确保输出目录存在
        os.makedirs(DATA_DIR, exist_ok=True)

    def generate_doc(
        self,
        quotation_records: list[dict],
        customer_name: str,
        language: str = "zh",
    ) -> dict:
        """
        生成报价单飞书云文档（不含 PDF，等用户确认后再生成）

        Returns:
            dict: {"doc_url": str, "quotation_number": str, "total_amount": float, "items": list}
        """
        items = self._parse_quotation_records(quotation_records)
        total_amount = self._calculate_total(items)
        quotation_number = self._generate_quotation_number()
        validity_date = self._calculate_validity_date()

        logger.info(f"生成报价单云文档: {quotation_number}, 客户={customer_name}, "
                     f"产品数={len(items)}, 总价={total_amount}")

        doc_url = self._create_quotation_doc(
            quotation_number=quotation_number,
            customer_name=customer_name,
            items=items,
            total_amount=total_amount,
            validity_date=validity_date,
            language=language,
        )
        logger.info(f"✅ 飞书云文档已创建: {doc_url}")

        return {
            "doc_url": doc_url,
            "quotation_number": quotation_number,
            "total_amount": total_amount,
            "items": items,
            "validity_date": validity_date,
        }

    def generate_pdf_from_doc(self, doc_url: str, customer_name: str, language: str = "zh") -> dict:
        """
        从用户确认后的飞书云文档直接生成 PDF（完全保留用户修改的内容）

        Args:
            doc_url: 用户确认后的报价单云文档 URL
            customer_name: 客户名称
            language: 语言

        Returns:
            dict: {"pdf_path": str, "quotation_number": str, "total_amount": float}
        """
        # 从云文档读取 Markdown 内容
        doc_markdown = self.fetch_doc_markdown(doc_url)
        if not doc_markdown:
            raise RuntimeError("无法读取云文档内容")

        # 提取报价编号和金额
        quotation_number = self._extract_quotation_number(doc_markdown) or self._generate_quotation_number()
        total_amount = self._extract_total_amount(doc_markdown) or 0

        logger.info(f"从云文档生成 PDF: {quotation_number}, 客户={customer_name}, 总价={total_amount}")

        # 将云文档 Markdown 转为 PDF（直接用文档内容，不重新解析）
        pdf_path = self._generate_pdf_from_markdown(doc_markdown, customer_name, quotation_number)
        if not pdf_path:
            raise RuntimeError("PDF 生成失败")

        logger.info(f"✅ PDF 报价单已生成: {pdf_path}")
        return {
            "pdf_path": pdf_path,
            "quotation_number": quotation_number,
            "total_amount": total_amount,
        }

    def _generate_pdf_from_markdown(self, markdown: str, customer_name: str, quotation_number: str) -> str:
        """
        将飞书云文档 Markdown 直接转为 PDF（保留用户所有修改）

        Args:
            markdown: 云文档 Markdown 内容
            customer_name: 客户名称
            quotation_number: 报价编号

        Returns:
            str: PDF 文件路径
        """
        # 将飞书 Markdown（含 lark-table 等标签）转为干净 HTML
        html_content = self._doc_markdown_to_pdf_html(markdown, customer_name, quotation_number)

        filename = f"{quotation_number}_{customer_name}.pdf"
        output_path = os.path.join(DATA_DIR, filename)

        try:
            from weasyprint import HTML
            HTML(string=html_content).write_pdf(output_path)
            return output_path
        except ImportError:
            logger.warning("weasyprint 未安装，跳过 PDF 生成")
            return ""
        except Exception as e:
            logger.error(f"PDF 生成失败: {e}")
            return ""

    def _doc_markdown_to_pdf_html(self, markdown: str, customer_name: str, quotation_number: str) -> str:
        """
        将飞书云文档 Markdown 转为带样式的 HTML（用于 PDF 生成）
        """
        import re

        # 第一步：提取 lark-table 为占位符
        tables = {}
        counter = [0]

        def extract_table(match):
            counter[0] += 1
            key = f"__TABLE_{counter[0]}__"
            table_html = match.group(0)
            cells = re.findall(r'<lark-td[^>]*>(.*?)</lark-td>', table_html, re.DOTALL)
            cells = [re.sub(r'<[^>]+>', '', c).strip() for c in cells]
            tables[key] = cells
            return key

        text = re.sub(r'<lark-table[^>]*>.*?</lark-table>', extract_table, markdown, flags=re.DOTALL)

        # 移除剩余飞书标签
        text = re.sub(r'<lark-[a-z]+[^>]*>', '', text)
        text = re.sub(r'</lark-[a-z]+>', '', text)
        text = re.sub(r'\n{3,}', '\n\n', text)

        # 第二步：Markdown → HTML
        lines = text.strip().split("\n")
        parts = []

        for line in lines:
            line = line.strip()
            if not line:
                continue
            # 表格占位符 → HTML table
            if line.startswith("__TABLE_"):
                cells = tables.get(line, [])
                if cells:
                    cols = 6
                    rows = []
                    for i in range(0, len(cells), cols):
                        row = cells[i:i+cols]
                        if len(row) < cols:
                            row.extend([""] * (cols - len(row)))
                        rows.append(row)
                    if rows:
                        t = "<table class='quotation-table'>"
                        t += "<thead><tr>" + "".join(f"<th>{c}</th>" for c in rows[0]) + "</tr></thead>"
                        for r in rows[1:]:
                            t += "<tr>" + "".join(f"<td>{c}</td>" for c in r) + "</tr>"
                        t += "</table>"
                        parts.append(t)
                continue
            if line.startswith("# "):
                parts.append(f"<h1 class='title'>{line[2:]}</h1>")
            elif line.startswith("## "):
                parts.append(f"<h2 class='section'>{line[3:]}</h2>")
            elif line.startswith("- "):
                parts.append(f"<li>{line[2:]}</li>")
            elif line.startswith("---"):
                parts.append("<hr class='divider'>")
            else:
                clean = re.sub(r'\*\*(.+?)\*\*', r'<strong>\1</strong>', line)
                parts.append(f"<p>{clean}</p>")

        body = "\n".join(parts)

        # 第三步：包装为完整的带样式 HTML
        html = f"""<!DOCTYPE html>
<html lang="zh">
<head>
<meta charset="utf-8">
<style>
@page {{
    size: A4;
    margin: 20mm 15mm;
}}
body {{
    font-family: "Microsoft YaHei", "SimHei", "PingFang SC", sans-serif;
    font-size: 11pt;
    color: #333;
    line-height: 1.6;
}}
h1.title {{
    text-align: center;
    font-size: 22pt;
    color: #1a1a2e;
    border-bottom: 3px solid #e94560;
    padding-bottom: 10px;
    margin-bottom: 20px;
}}
h2.section {{
    font-size: 13pt;
    color: #1a1a2e;
    margin-top: 20px;
    border-left: 4px solid #e94560;
    padding-left: 10px;
}}
table.quotation-table {{
    width: 100%;
    border-collapse: collapse;
    margin: 10px 0 20px 0;
    font-size: 10pt;
}}
table.quotation-table th {{
    background-color: #1a1a2e;
    color: white;
    padding: 8px 6px;
    text-align: center;
    font-weight: bold;
}}
table.quotation-table td {{
    border: 1px solid #ddd;
    padding: 6px;
    text-align: center;
}}
table.quotation-table tr:nth-child(even) {{
    background-color: #f9f9f9;
}}
p {{
    margin: 4px 0;
}}
li {{
    margin-left: 20px;
}}
hr.divider {{
    border: none;
    border-top: 1px solid #ddd;
    margin: 20px 0;
}}
strong {{
    color: #1a1a2e;
}}
</style>
</head>
<body>
{body}
</body>
</html>"""
        return html

    def _extract_total_amount(self, markdown: str) -> float:
        """从云文档中提取总金额"""
        import re
        # 匹配 "合计" 行中的金额
        m = re.search(r'合计.*?([\d,]+\.?\d*)', markdown)
        if m:
            try:
                return float(m.group(1).replace(',', ''))
            except ValueError:
                pass
        return 0

    def _parse_doc_to_items(self, markdown: str) -> list[dict]:
        """从云文档 Markdown 中解析报价明细"""
        import re
        items = []
        # 匹配 lark-table 中的单元格
        cells = re.findall(r'<lark-td[^>]*>(.*?)</lark-td>', markdown, re.DOTALL)
        cells = [re.sub(r'<[^>]+>', '', c).strip() for c in cells]

        if not cells:
            return []

        # 每 6 个单元格一行：序号、产品名称、型号规格、数量、单价、金额
        cols = 6
        for i in range(0, len(cells), cols):
            row = cells[i:i+cols]
            if len(row) < cols:
                continue
            # 跳过表头行
            if row[0] == "序号":
                continue
            try:
                qty = int(re.sub(r'[^\d]', '', row[3])) if row[3] else 0
                price = float(re.sub(r'[^\d.]', '', row[4])) if row[4] else 0
                subtotal = float(re.sub(r'[^\d.]', '', row[5])) if row[5] else 0
                items.append({
                    "name": row[1],
                    "model": row[2],
                    "quantity": qty,
                    "unit_price": price,
                    "amount": subtotal,
                })
            except (ValueError, IndexError):
                continue

        return items

    def _extract_quotation_number(self, markdown: str) -> str:
        """从云文档中提取报价编号"""
        import re
        m = re.search(r'报价编号[：:]\s*(\S+)', markdown)
        return m.group(1) if m else ""

    def _extract_validity_date(self, markdown: str) -> str:
        """从云文档中提取有效期"""
        import re
        m = re.search(r'有效期至[：:]\s*(\S+)', markdown)
        return m.group(1) if m else ""

    def generate(
        self,
        quotation_records: list[dict],
        customer_name: str,
        language: str = "zh",
    ) -> dict:
        """
        生成报价单（飞书云文档 + PDF）— 一次性生成，用于简单场景

        Returns:
            dict: {"doc_url": str, "pdf_path": str, "quotation_number": str, "total_amount": float}
        """
        doc_result = self.generate_doc(quotation_records, customer_name, language)
        pdf_result = self.generate_pdf_from_doc(doc_result["doc_url"], customer_name, language)
        return {
            "doc_url": doc_result["doc_url"],
            "pdf_path": pdf_result["pdf_path"],
            "quotation_number": pdf_result["quotation_number"],
            "total_amount": pdf_result["total_amount"],
        }

    def _parse_quotation_records(self, records: list[dict]) -> list[dict]:
        """
        将报价明细记录解析为标准产品列表

        Args:
            records: 原始记录列表

        Returns:
            list[dict]: [{"name": str, "model": str, "quantity": int, "unit_price": float, "amount": float}]
        """
        items = []
        for r in records:
            fields = r.get("fields", [])
            if len(fields) < 5:
                continue
            try:
                items.append({
                    "name": str(fields[0]),           # 产品名称
                    "model": str(fields[1]),          # 产品型号
                    "quantity": int(fields[2]),        # 数量
                    "unit_price": float(str(fields[3]).replace(",", "")),  # 单价
                    "amount": float(str(fields[4]).replace(",", "")),      # 总价
                })
            except (ValueError, IndexError) as e:
                logger.warning(f"解析报价记录失败: {fields}, 错误: {e}")
        return items

    def _calculate_total(self, items: list[dict]) -> float:
        """计算总金额"""
        return sum(item.get("amount", 0) for item in items)

    def _generate_quotation_number(self) -> str:
        """生成报价单编号（QT-YYYYMMDD-NNN）"""
        today = datetime.now().strftime("%Y%m%d")
        # 简单序号：用时间戳后 3 位
        seq = str(int(datetime.now().timestamp()))[-3:]
        return f"QT-{today}-{seq}"

    def _calculate_validity_date(self) -> str:
        """计算有效期截止日期"""
        valid_until = datetime.now() + timedelta(days=self.validity_days)
        return valid_until.strftime("%Y-%m-%d")

    # ============================================================
    # 飞书云文档
    # ============================================================

    def _create_quotation_doc(
        self,
        quotation_number: str,
        customer_name: str,
        items: list[dict],
        total_amount: float,
        validity_date: str,
        language: str,
    ) -> str:
        """
        创建飞书云文档

        Returns:
            str: 文档 URL
        """
        if language == "en":
            title = f"Quotation {quotation_number} - {customer_name}"
            markdown = self._build_markdown_en(
                quotation_number, customer_name, items, total_amount, validity_date
            )
        else:
            title = f"报价单 {quotation_number} - {customer_name}"
            markdown = self._build_markdown_cn(
                quotation_number, customer_name, items, total_amount, validity_date
            )

        try:
            result = subprocess.run(
                [
                    "lark-cli", "docs", "+create",
                    "--title", title,
                    "--markdown", markdown,
                ],
                capture_output=True,
                text=True,
                timeout=30,
            )

            if result.returncode != 0:
                logger.error(f"创建飞书文档失败: {result.stderr}")
                return ""

            output = json.loads(result.stdout)
            if output.get("ok"):
                data = output.get("data", {})
                return data.get("doc_url", data.get("url", ""))
            else:
                logger.error(f"飞书文档 API 错误: {output}")
                return ""

        except Exception as e:
            logger.error(f"创建飞书文档异常: {e}")
            return ""

    def _build_markdown_cn(
        self, quotation_number, customer_name, items, total_amount, validity_date
    ) -> str:
        """构造中文报价单 Markdown"""
        lines = [
            f"# 报价单",
            f"",
            f"**报价编号：** {quotation_number}",
            f"**客户名称：** {customer_name}",
            f"**报价日期：** {datetime.now().strftime('%Y-%m-%d')}",
            f"**有效期至：** {validity_date}",
            f"",
            f"## 产品明细",
            f"",
            f"| 序号 | 产品名称 | 型号规格 | 数量 | 单价（{self.currency}） | 金额（{self.currency}） |",
            f"| --- | --- | --- | --- | --- | --- |",
        ]
        for i, item in enumerate(items, 1):
            lines.append(
                f"| {i} | {item['name']} | {item['model']} | {item['quantity']} "
                f"| {item['unit_price']:,.2f} | {item['amount']:,.2f} |"
            )
        lines.extend([
            f"| | | | **合计** | | **{total_amount:,.2f}** |",
            f"",
            f"## 付款方式",
            f"{self.payment_terms_cn}",
            f"",
            f"## 交货期",
            f"收到订单后 {self.delivery_days} 个工作日内交货",
            f"",
            f"## 包装方式",
            f"{self.packaging_cn}",
            f"",
            f"## 银行信息",
            f"- **开户行：** {self.bank_name}",
            f"- **账号：** {self.bank_account}",
            f"- **SWIFT：** {self.bank_swift}",
            f"",
            f"---",
            f"*{self.company_name_cn}*",
            f"*{self.company_address_cn}*",
            f"*Tel: {self.company_phone} | Email: {self.company_email}*",
        ])
        return "\n".join(lines)

    def _build_markdown_en(
        self, quotation_number, customer_name, items, total_amount, validity_date
    ) -> str:
        """构造英文报价单 Markdown"""
        lines = [
            f"# Quotation",
            f"",
            f"**Quotation No.:** {quotation_number}",
            f"**Customer:** {customer_name}",
            f"**Date:** {datetime.now().strftime('%Y-%m-%d')}",
            f"**Valid Until:** {validity_date}",
            f"",
            f"## Product Details",
            f"",
            f"| No. | Product Name | Model | Qty | Unit Price ({self.currency}) | Amount ({self.currency}) |",
            f"| --- | --- | --- | --- | --- | --- |",
        ]
        for i, item in enumerate(items, 1):
            lines.append(
                f"| {i} | {item['name']} | {item['model']} | {item['quantity']} "
                f"| {item['unit_price']:,.2f} | {item['amount']:,.2f} |"
            )
        lines.extend([
            f"| | | | **Total** | | **{total_amount:,.2f}** |",
            f"",
            f"## Payment Terms",
            f"{self.payment_terms_en}",
            f"",
            f"## Delivery Time",
            f"Within {self.delivery_days} working days after order confirmation",
            f"",
            f"## Packaging",
            f"{self.packaging_en}",
            f"",
            f"## Bank Information",
            f"- **Bank:** {self.bank_name}",
            f"- **Account:** {self.bank_account}",
            f"- **SWIFT:** {self.bank_swift}",
            f"",
            f"---",
            f"*{self.company_name_en}*",
            f"*{self.company_address_en}*",
            f"*Tel: {self.company_phone} | Email: {self.company_email}*",
        ])
        return "\n".join(lines)

    # ============================================================
    # PDF 生成
    # ============================================================

    def _generate_pdf(
        self,
        quotation_number: str,
        customer_name: str,
        items: list[dict],
        total_amount: float,
        validity_date: str,
        language: str,
    ) -> str:
        """
        生成 PDF 报价单

        Returns:
            str: PDF 文件路径
        """
        if language == "en":
            html_content = self._build_html_en(
                quotation_number, customer_name, items, total_amount, validity_date
            )
        else:
            html_content = self._build_html_cn(
                quotation_number, customer_name, items, total_amount, validity_date
            )

        filename = f"{quotation_number}_{customer_name}.pdf"
        output_path = os.path.join(DATA_DIR, filename)

        try:
            from weasyprint import HTML
            HTML(string=html_content).write_pdf(output_path)
            return output_path
        except ImportError:
            logger.warning("weasyprint 未安装，跳过 PDF 生成")
            return ""
        except Exception as e:
            logger.error(f"PDF 生成失败: {e}")
            return ""

    def _build_html_cn(self, quotation_number, customer_name, items, total_amount, validity_date):
        """构造中文报价单 HTML"""
        rows = ""
        for i, item in enumerate(items, 1):
            rows += f"""
            <tr>
                <td>{i}</td>
                <td>{item['name']}</td>
                <td>{item['model']}</td>
                <td>{item['quantity']}</td>
                <td>{self.currency} {item['unit_price']:,.2f}</td>
                <td>{self.currency} {item['amount']:,.2f}</td>
            </tr>"""

        return f"""<!DOCTYPE html>
<html><head><meta charset="utf-8">
<style>
body {{ font-family: "Microsoft YaHei", sans-serif; margin: 40px; color: #333; }}
h1 {{ text-align: center; color: #1a1a2e; border-bottom: 3px solid #e94560; padding-bottom: 10px; }}
.info {{ display: flex; justify-content: space-between; margin: 20px 0; }}
table {{ width: 100%; border-collapse: collapse; margin: 20px 0; }}
th {{ background: #1a1a2e; color: white; padding: 10px; text-align: center; }}
td {{ border: 1px solid #ddd; padding: 8px; text-align: center; }}
.total {{ font-weight: bold; font-size: 1.1em; text-align: right; margin: 10px 0; }}
.section {{ margin: 20px 0; }}
.section h3 {{ color: #1a1a2e; }}
.footer {{ margin-top: 40px; padding-top: 20px; border-top: 1px solid #ddd; text-align: center; color: #666; font-size: 0.9em; }}
</style></head><body>
<h1>报 价 单</h1>
<div class="info">
    <div><strong>报价编号：</strong>{quotation_number}</div>
    <div><strong>报价日期：</strong>{datetime.now().strftime('%Y-%m-%d')}</div>
</div>
<div class="info">
    <div><strong>客户名称：</strong>{customer_name}</div>
    <div><strong>有效期至：</strong>{validity_date}</div>
</div>
<table>
    <thead><tr><th>序号</th><th>产品名称</th><th>型号规格</th><th>数量</th><th>单价</th><th>金额</th></tr></thead>
    <tbody>{rows}
    <tr><td colspan="5" style="text-align:right;"><strong>合计</strong></td><td><strong>{self.currency} {total_amount:,.2f}</strong></td></tr>
    </tbody>
</table>
<div class="section"><h3>付款方式</h3><p>{self.payment_terms_cn}</p></div>
<div class="section"><h3>交货期</h3><p>收到订单后 {self.delivery_days} 个工作日内交货</p></div>
<div class="section"><h3>包装方式</h3><p>{self.packaging_cn}</p></div>
<div class="section"><h3>银行信息</h3>
    <p>开户行：{self.bank_name}<br>账号：{self.bank_account}<br>SWIFT：{self.bank_swift}</p>
</div>
<div class="footer">
    <p><strong>{self.company_name_cn}</strong></p>
    <p>{self.company_address_cn}</p>
    <p>Tel: {self.company_phone} | Email: {self.company_email}</p>
</div>
</body></html>"""

    # ============================================================
    # 从云文档读取修改后的内容
    # ============================================================

    def fetch_doc_markdown(self, doc_url: str) -> str:
        """
        读取飞书云文档的 Markdown 内容（用户修改后的版本）

        Args:
            doc_url: 飞书云文档 URL

        Returns:
            str: Markdown 文本
        """
        try:
            result = subprocess.run(
                ["lark-cli", "docs", "+fetch", "--doc", doc_url],
                capture_output=True, text=True, timeout=30,
            )
            if result.returncode != 0:
                logger.error(f"读取云文档失败: {result.stderr}")
                return ""
            data = json.loads(result.stdout)
            # lark-cli +fetch 返回格式可能是 {"ok": true, "data": {"markdown": "..."}}
            # 或者直接是 markdown 文本
            if data.get("ok"):
                md = data.get("data", {}).get("markdown", "")
                if not md:
                    md = data.get("data", {}).get("content", "")
                return md
            return ""
        except Exception as e:
            logger.error(f"读取云文档异常: {e}")
            return ""

    def generate_reply_from_doc(self, doc_url: str, customer_name: str, language: str = "zh") -> str:
        """
        根据用户修改后的云文档内容生成回复邮件正文

        Args:
            doc_url: 飞书云文档 URL
            customer_name: 客户名称
            language: 语言

        Returns:
            str: HTML 邮件正文
        """
        doc_markdown = self.fetch_doc_markdown(doc_url)
        if doc_markdown:
            logger.info(f"已读取云文档内容: {len(doc_markdown)} 字符")
            body_html = self._doc_to_email_html(doc_markdown)
        else:
            logger.warning("云文档内容为空，使用默认邮件模板")
            body_html = self._get_default_reply_html(customer_name, language)

        return body_html

    def generate_reply_doc(self, quotation_doc_url: str, customer_name: str, to_addr: str, language: str = "zh") -> dict:
        """
        从报价单云文档读取内容，生成「邮件预览」飞书云文档，供用户调整确认

        Args:
            quotation_doc_url: 报价单云文档 URL
            customer_name: 客户名称
            to_addr: 回复收件人
            language: 语言

        Returns:
            dict: {"doc_url": str, "doc_token": str}
        """
        # 读取报价单云文档内容
        doc_markdown = self.fetch_doc_markdown(quotation_doc_url)
        if not doc_markdown:
            logger.warning("报价单云文档内容为空")
            doc_markdown = "(报价单内容读取失败，请手动填写)"

        # 构造邮件预览文档的 Markdown
        if language == "en":
            email_md = f"""# Reply Email Preview

**To:** {to_addr}
**Subject:** Quotation from {self.company_name_en}

---

{doc_markdown}

---

*Please review and adjust the content above. After confirmation, this will be sent as an email to the customer.*
*Generated at {datetime.now().strftime('%Y-%m-%d %H:%M')}*
"""
        else:
            email_md = f"""# 回复邮件预览

**收件人：** {to_addr}
**主题：** {self.company_name_cn} 报价单

---

{doc_markdown}

---

*请审阅并调整以上邮件内容，确认后将作为邮件发送给客户。*
*生成时间：{datetime.now().strftime('%Y-%m-%d %H:%M')}*
"""

        # 创建飞书云文档
        doc_token = self._create_feishu_doc(email_md, f"邮件预览 - {customer_name}")
        doc_url = f"https://www.feishu.cn/docx/{doc_token}"
        logger.info(f"邮件预览云文档已创建: {doc_url}")
        return {"doc_url": doc_url, "doc_token": doc_token}

    def _create_feishu_doc(self, markdown: str, title: str) -> str:
        """
        创建飞书云文档并写入 Markdown 内容

        Args:
            markdown: Markdown 内容
            title: 文档标题

        Returns:
            str: doc_token
        """
        result = subprocess.run(
            ["lark-cli", "docs", "+create",
             "--title", title,
             "--markdown", markdown],
            capture_output=True, text=True, timeout=30,
        )
        if result.returncode != 0:
            logger.error(f"创建云文档失败: {result.stderr}")
            raise RuntimeError(f"创建云文档失败: {result.stderr}")

        data = json.loads(result.stdout)
        doc_token = data.get("data", {}).get("doc_id", "") or data.get("data", {}).get("doc_token", "")
        if not doc_token:
            logger.error(f"创建云文档未返回 doc_token: {data}")
            raise RuntimeError(f"创建云文档未返回 doc_token")
        return doc_token

    def _doc_to_email_html(self, markdown: str) -> str:
        """
        将飞书云文档 Markdown 转为邮件 HTML
        处理飞书特殊标签（lark-table 等），生成干净的邮件 HTML
        """
        import re

        # 第一步：提取 lark-table 为占位符，保存表格数据
        tables = {}
        counter = [0]

        def extract_table(match):
            counter[0] += 1
            key = f"__TABLE_{counter[0]}__"
            table_html = match.group(0)
            cells = re.findall(r'<lark-td[^>]*>(.*?)</lark-td>', table_html, re.DOTALL)
            cells = [re.sub(r'<[^>]+>', '', c).strip() for c in cells]
            tables[key] = cells
            return key

        text = re.sub(r'<lark-table[^>]*>.*?</lark-table>', extract_table, markdown, flags=re.DOTALL)

        # 第二步：移除剩余飞书标签
        text = re.sub(r'<lark-[a-z]+[^>]*>', '', text)
        text = re.sub(r'</lark-[a-z]+>', '', text)
        text = re.sub(r'\n{3,}', '\n\n', text)

        # 第三步：Markdown → HTML
        lines = text.strip().split("\n")
        parts = []

        for line in lines:
            line = line.strip()
            if not line:
                continue
            # 表格占位符 → HTML table
            if line.startswith("__TABLE_"):
                cells = tables.get(line, [])
                if cells:
                    cols = 6
                    rows = []
                    for i in range(0, len(cells), cols):
                        row = cells[i:i+cols]
                        if len(row) < cols:
                            row.extend([""] * (cols - len(row)))
                        rows.append(row)
                    if rows:
                        t = "<table border='1' cellpadding='8' cellspacing='0' style='border-collapse:collapse;width:100%;margin:10px 0;'>"
                        t += "<thead><tr>" + "".join(f"<th style='background:#1a1a2e;color:white;padding:8px;'>{c}</th>" for c in rows[0]) + "</tr></thead>"
                        for r in rows[1:]:
                            t += "<tr>" + "".join(f"<td style='border:1px solid #ddd;padding:6px;text-align:center;'>{c}</td>" for c in r) + "</tr>"
                        t += "</table>"
                        parts.append(t)
                continue
            if line.startswith("# "):
                parts.append(f"<h2 style='color:#1a1a2e;border-bottom:2px solid #e94560;padding-bottom:8px;'>{line[2:]}</h2>")
            elif line.startswith("## "):
                parts.append(f"<h3 style='color:#1a1a2e;'>{line[3:]}</h3>")
            elif line.startswith("- "):
                parts.append(f"<li style='margin-left:20px;'>{line[2:]}</li>")
            elif line.startswith("---"):
                parts.append("<hr style='border:none;border-top:1px solid #ddd;margin:20px 0;'>")
            else:
                clean = re.sub(r'\*\*(.+?)\*\*', r'<strong>\1</strong>', line)
                parts.append(f"<p>{clean}</p>")

        return "\n".join(parts)

    def _get_default_reply_html(self, customer_name: str, language: str) -> str:
        """默认邮件模板（云文档读取失败时使用）"""
        if language == "en":
            return (
                f"<p>Dear {customer_name},</p>"
                f"<p>Please find attached our formal quotation for your reference.</p>"
                f"<p>This quotation is valid for <strong>{self.validity_days} days</strong>.</p>"
                f"<p>We look forward to working with you!</p>"
                f"<p>Best regards,<br><strong>{self.company_name_en}</strong><br>"
                f"Tel: {self.company_phone} | Email: {self.company_email}</p>"
            )
        return (
            f"<p>尊敬的 {customer_name}，</p>"
            f"<p>感谢贵司的询价！附件为我司的正式报价单，请查收。</p>"
            f"<p>报价有效期为本邮件发出之日起 <strong>{self.validity_days} 天</strong>。</p>"
            f"<p>期待与贵司的合作！</p>"
            f"<p>此致<br>敬礼<br><strong>{self.company_name_cn}</strong><br>"
            f"Tel: {self.company_phone} | Email: {self.company_email}</p>"
        )

    def _build_html_en(self, quotation_number, customer_name, items, total_amount, validity_date):
        """构造英文报价单 HTML"""
        rows = ""
        for i, item in enumerate(items, 1):
            rows += f"""
            <tr>
                <td>{i}</td>
                <td>{item['name']}</td>
                <td>{item['model']}</td>
                <td>{item['quantity']}</td>
                <td>{self.currency} {item['unit_price']:,.2f}</td>
                <td>{self.currency} {item['amount']:,.2f}</td>
            </tr>"""

        return f"""<!DOCTYPE html>
<html><head><meta charset="utf-8">
<style>
body {{ font-family: Arial, sans-serif; margin: 40px; color: #333; }}
h1 {{ text-align: center; color: #1a1a2e; border-bottom: 3px solid #e94560; padding-bottom: 10px; }}
.info {{ display: flex; justify-content: space-between; margin: 20px 0; }}
table {{ width: 100%; border-collapse: collapse; margin: 20px 0; }}
th {{ background: #1a1a2e; color: white; padding: 10px; text-align: center; }}
td {{ border: 1px solid #ddd; padding: 8px; text-align: center; }}
.total {{ font-weight: bold; font-size: 1.1em; text-align: right; margin: 10px 0; }}
.section {{ margin: 20px 0; }}
.section h3 {{ color: #1a1a2e; }}
.footer {{ margin-top: 40px; padding-top: 20px; border-top: 1px solid #ddd; text-align: center; color: #666; font-size: 0.9em; }}
</style></head><body>
<h1>QUOTATION</h1>
<div class="info">
    <div><strong>Quotation No.:</strong> {quotation_number}</div>
    <div><strong>Date:</strong> {datetime.now().strftime('%Y-%m-%d')}</div>
</div>
<div class="info">
    <div><strong>Customer:</strong> {customer_name}</div>
    <div><strong>Valid Until:</strong> {validity_date}</div>
</div>
<table>
    <thead><tr><th>No.</th><th>Product Name</th><th>Model</th><th>Qty</th><th>Unit Price</th><th>Amount</th></tr></thead>
    <tbody>{rows}
    <tr><td colspan="5" style="text-align:right;"><strong>Total</strong></td><td><strong>{self.currency} {total_amount:,.2f}</strong></td></tr>
    </tbody>
</table>
<div class="section"><h3>Payment Terms</h3><p>{self.payment_terms_en}</p></div>
<div class="section"><h3>Delivery Time</h3><p>Within {self.delivery_days} working days after order confirmation</p></div>
<div class="section"><h3>Packaging</h3><p>{self.packaging_en}</p></div>
<div class="section"><h3>Bank Information</h3>
    <p>Bank: {self.bank_name}<br>Account: {self.bank_account}<br>SWIFT: {self.bank_swift}</p>
</div>
<div class="footer">
    <p><strong>{self.company_name_en}</strong></p>
    <p>{self.company_address_en}</p>
    <p>Tel: {self.company_phone} | Email: {self.company_email}</p>
</div>
</body></html>"""
