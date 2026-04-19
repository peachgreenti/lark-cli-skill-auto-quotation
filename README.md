# 外贸报价全链路自动化系统

监听邮箱询价邮件，自动完成从询价识别到报价单发送的全流程。

## 项目目录结构

```
lark-cli-skill-auto-quotation/
├── main.py                  # 入口：流程编排、飞书群消息轮询
├── mail_watcher.py          # IMAP 邮件监听
├── mail_parser.py           # 邮件解析（头、正文、附件）
├── inquiry_detector.py      # AI 询价识别 + 需求提取
├── base_writer.py           # 飞书多维表格读写
├── quotation_poller.py      # 报价结果轮询
├── quotation_generator.py   # 报价单生成（云文档 + PDF）
├── mail_sender.py           # SMTP 邮件发送
├── notifier.py              # 飞书群通知
├── config.yaml              # 配置文件
├── .env                     # 环境变量（密码、密钥）
├── requirements.txt         # Python 依赖
├── SKILL.md                 # AI Agent 调用指令
└── data/
    └── temp/                # 临时文件（PDF、附件）
```

## 系统架构

```
┌─────────────┐     ┌──────────────┐     ┌─────────────────┐
│  IMAP 邮箱   │────▶│  AI 询价识别  │────▶│  飞书多维表格    │
│  (mail_      │     │  (inquiry_   │     │  (base_writer)  │
│   watcher)   │     │   detector)  │     │                 │
└─────────────┘     └──────────────┘     └────────┬────────┘
                                                   │
                                                   ▼
┌─────────────┐     ┌──────────────┐     ┌─────────────────┐
│  SMTP 发送   │◀────│  邮件预览文档  │◀────│  报价结果轮询    │
│  (mail_      │     │  (quotation_  │     │  (quotation_    │
│   sender)    │     │   generator)  │     │   poller)       │
└─────────────┘     └──────────────┘     └─────────────────┘
        ▲                   ▲
        │                   │
        └───────┌───────────┘
                │  飞书群通知
                │  (notifier)
                │
        ┌───────┴───────────┐
        │  用户在飞书群确认   │
        └───────────────────┘
```

## 完整流程

系统运行后自动执行以下 10 个阶段：

| 阶段 | 说明 | 用户操作 |
|------|------|---------|
| ① | IMAP 监听邮箱，检测新邮件 | 无 |
| ② | AI 识别是否为询价邮件（关键词预筛 + AI 二次确认） | 无 |
| ③ | 解析邮件：提取正文、PDF 附件文本 | 无 |
| ④ | AI 提取报价需求 → 写入飞书多维表格 + 上传附件 | 无 |
| ⑤⑥ | 轮询等待报价完成（多维表格 AI 字段自动生成，60 秒稳定） | 无 |
| ⑦ | 生成报价单飞书云文档 | 无 |
| ⑧ | 飞书群通知 → 用户在云文档中调整报价 | **用户编辑云文档，群内回复「确认」** |
| ⑦② | 从用户确认后的云文档生成 PDF | 无 |
| ⑨ | 生成邮件预览云文档 → 飞书群通知 | **用户编辑邮件内容，群内回复「确认」** |
| ⑩ | 从邮件预览云文档读取最终内容 → 发送邮件 + PDF 附件 | 无 |

## 快速开始

### 1. 安装依赖

```bash
pip install -r requirements.txt
```

主要依赖：
- `PyMuPDF` — PDF 文本提取
- `weasyprint` — HTML 转 PDF
- `httpx` — OpenAI 兼容 API 客户端
- `pyyaml` — YAML 配置解析
- `python-dotenv` — 环境变量加载

### 2. 安装 lark-cli

```bash
# 安装飞书 CLI 工具（用于飞书 API 调用）
npm install -g @aspect/lark-cli
# 或参考飞书官方文档安装
```

### 3. 配置环境变量

复制并编辑 `.env` 文件：

```bash
cp .env.example .env
```

必填环境变量：

| 变量名 | 说明 | 示例 |
|--------|------|------|
| `IMAP_SERVER` | IMAP 服务器 | `imap.163.com` |
| `IMAP_PORT` | IMAP 端口 | `993` |
| `IMAP_USER` | 邮箱地址 | `your@163.com` |
| `IMAP_PASSWORD` | IMAP 授权码 | `xxxxxxxx` |
| `SMTP_SERVER` | SMTP 服务器 | `smtp.163.com` |
| `SMTP_PORT` | SMTP 端口 | `465` |
| `SMTP_USER` | SMTP 用户名 | `your@163.com` |
| `SMTP_PASSWORD` | SMTP 授权码 | `xxxxxxxx` |
| `OPENAI_API_KEY` | AI API Key | `xxxxxxxx` |
| `OPENAI_BASE_URL` | AI API 地址 | `https://ark.cn-beijing.volces.com/api/v3` |
| `FEISHU_APP_ID` | 飞书应用 ID | `cli_xxxxxxxx` |
| `FEISHU_APP_SECRET` | 飞书应用 Secret | `xxxxxxxx` |

### 4. 配置 config.yaml

编辑 `config.yaml`，主要配置项：

```yaml
# 邮件监听
mail:
  watch_folders: ["INBOX"]
  inquiry_keywords: ["询价", "报价", "quotation", "inquiry", "RFQ"]

# 飞书多维表格
base:
  app_token: "your_base_app_token"
  inquiry_table_id: "your_inquiry_table_id"
  quotation_table_id: "your_quotation_table_id"

# 飞书通知
notify:
  chat_id: "your_feishu_group_chat_id"

# 报价单模板
quotation:
  company_name_cn: "公司中文名"
  company_name_en: "Company Name"
  company_phone: "+86 xxx xxxx xxxx"
  company_email: "info@company.com"
  validity_days: 30
  currency: "USD"

# 轮询配置
polling:
  check_interval: 60
  quotation_stable_time: 60
```

### 5. 启动

```bash
python3 main.py
```

系统启动后进入无限循环，每 60 秒检查一次新邮件。

## 测试模式

跳过邮件监听，直接从已有的需求编号开始处理后续流程（⑦⑧⑨⑩）：

```bash
# 基本用法
python3 main.py --test 20260419090

# 指定回复收件人
python3 main.py --test 20260419090 --to customer@example.com
```

适用于：
- 报价数据已在多维表格中生成，只需测试后续流程
- 调试报价单生成、PDF 生成、邮件发送等环节

## 模块说明

### main.py

入口模块，流程编排。

- `main()` — 启动邮件监听循环或测试模式
- `process_inquiry()` — 处理单封询价邮件的完整流程（阶段③④⑤⑥⑦⑧⑨⑩）
- `run_test_mode()` — 测试模式入口
- `wait_for_feishu_confirm()` — 在飞书群中发送提示并轮询等待用户回复

### mail_watcher.py — MailWatcher

IMAP 邮件监听，支持断线重连。

- `connect()` — 连接 IMAP 服务器（30 秒超时）
- `check_new_emails()` — 检查新邮件，返回未处理的邮件列表
- `mark_as_seen()` — 标记邮件为已读

### mail_parser.py — MailParser

解析邮件头部、正文、附件。

- `parse(raw_email)` — 返回结构化数据

关键字段：
- `reply_to` — 真实发件人（优先 X-Sender > Return-Path > From），用于回复
- `pdf_attachments` — PDF 附件列表（含文件名和内容）

### inquiry_detector.py — InquiryDetector

AI 询价识别（火山方舟豆包大模型）。

- `detect(subject, body)` — 返回 `{is_inquiry, customer_name, language, confidence, reason}`
- `extract_requirement_text(pdf_text)` — 从 PDF 文本中提取报价需求

识别策略：关键词预筛（询价/报价/quotation 等） → AI 二次确认（降低误判）。

### base_writer.py — BaseWriter

飞书多维表格读写（通过 lark-cli）。

- `create_inquiry_record(pdf_path, requirement_text)` — 创建询盘记录，返回 `{record_id, req_number}`
- `upload_attachment(pdf_path, record_id)` — 上传 PDF 附件
- `list_quotation_records(req_number)` — 查询报价明细（按需求编号过滤）
- `update_inquiry_status(record_id, status)` — 更新询盘状态

### quotation_poller.py — QuotationPoller

轮询等待多维表格 AI 字段生成报价数据。

- `poll(writer, inquiry_record_id, req_number, skip_wait)` — 数据稳定后返回记录列表

稳定策略：连续 60 秒数据无变化视为稳定。超时 30 分钟。

### quotation_generator.py — QuotationGenerator

报价单生成（飞书云文档 + PDF）。

- `generate_doc()` — 生成报价单飞书云文档（用户可编辑）
- `generate_pdf_from_doc(doc_url)` — 从用户确认后的云文档生成 PDF
- `generate_reply_doc()` — 生成邮件预览飞书云文档（用户可编辑）
- `generate_reply_from_doc(doc_url)` — 从邮件预览云文档读取内容，生成 HTML 邮件正文
- `fetch_doc_markdown(doc_url)` — 读取飞书云文档的 Markdown 内容

### mail_sender.py — MailSender

SMTP 邮件发送，支持重试。

- `send_reply(to_addr, subject, reply_body, pdf_path)` — 发送回复邮件（含 PDF 附件）

重试策略：最多 3 次，间隔 5 秒。

### notifier.py — FeishuNotifier

飞书群消息通知。

- `send_inquiry_received()` — 通知收到新询价
- `send_quotation_ready()` — 通知报价单已生成（含云文档链接）
- `send_reply_sent()` — 通知邮件已发送
- `send_error()` — 发送错误告警

## 飞书多维表格结构

系统依赖以下三张表：

### 询盘需求表

| 字段 | 说明 |
|------|------|
| 客户名称 | 客户公司名 |
| 报价需求 | AI 提取的需求描述（如 "50个CP001、30个CP003"） |
| 状态 | pending / quoted / sent / cancelled |
| 附件 | 询价函 PDF |
| 需求编号 | 自动编号（如 20260419090） |

### 报价明细表

| 字段 | 说明 |
|------|------|
| 产品名称 | 产品名 |
| 产品型号 | 型号规格 |
| 数量 | 客户需求数量 |
| 单价 | AI 自动生成 |
| 总价 | AI 自动生成 |
| 客户名称 | 关联客户 |
| 需求编号 | 关联询盘需求 |
| 报价状态 | pending / confirmed |

### 客户管理表

| 字段 | 说明 |
|------|------|
| 客户名称 | 公司名 |
| 联系邮箱 | 邮箱地址 |
| 联系人 | 联系人姓名 |
| 联系电话 | 电话号码 |

## 异常处理

系统对以下异常场景有保护措施：

| 场景 | 处理方式 |
|------|---------|
| IMAP 连接断开 | 自动重连 |
| SMTP 发送失败 | 重试 3 次（间隔 5 秒） |
| AI API 调用失败 | 重试 2 次（间隔 3 秒） |
| AI 返回空结果 | 记录日志，跳过该邮件 |
| lark-cli 命令超时 | base 60s、docs/im 30s（轮询场景 15s） |
| 邮件解析失败 | 记录日志，跳过该邮件 |
| PDF 附件损坏 | 记录日志，跳过提取 |
| 配置文件缺失 | 启动时校验，明确报错 |
| 环境变量缺失 | 启动时校验，明确报错 |
| 用户超时未确认 | 30 分钟超时，自动取消 |
| 单封邮件处理异常 | 记录错误，继续处理下一封 |

## 日志

系统使用 Python `logging` 模块输出日志：

```
2026-04-19 08:18:24 [INFO] main: ✅ 发现询价邮件: [询价函 - 苹果贸易有限公司]
2026-04-19 08:18:30 [INFO] base_writer: ✅ 询盘记录已创建: record_id=recvhbMXMXzx77
2026-04-19 08:19:00 [INFO] quotation_poller: ✅ 报价结果已稳定 60s，共 4 条记录
2026-04-19 08:19:05 [INFO] quotation_generator: ✅ 飞书云文档已创建: https://www.feishu.cn/docx/xxx
2026-04-19 08:20:00 [INFO] main: 用户在飞书群中确认
2026-04-19 08:20:10 [INFO] quotation_generator: ✅ PDF 报价单已生成: data/temp/QT-xxx.pdf
2026-04-19 08:21:00 [INFO] mail_sender: ✅ 回复邮件已发送: customer@example.com
```

## 常见问题

### Q: 如何获取 IMAP/SMTP 授权码？

163 邮箱：设置 → POP3/SMTP/IMAP → 开启服务 → 获取授权码
QQ 邮箱：设置 → 账户 → POP3/IMAP 服务 → 生成授权码
Gmail：需开启 App Password（两步验证）

### Q: 如何获取飞书多维表格的 app_token 和 table_id？

1. 打开多维表格
2. URL 格式为 `https://xxx.feishu.cn/base/{app_token}?table={table_id}&view=...`
3. `app_token` 和 `table_id` 直接从 URL 中提取

### Q: 如何获取飞书群聊 ID？

1. 在飞书群中发送任意消息
2. 使用 `lark-cli im +chat-messages-list --chat-id <群ID>` 测试
3. 或通过飞书开放平台 API 获取

### Q: AI 识别不准确怎么办？

1. 调整 `config.yaml` 中的 `inquiry_keywords` 关键词列表
2. 修改 `inquiry_detector.py` 中的 AI prompt
3. 系统采用「关键词预筛 + AI 二次确认」双重策略，误判率较低

### Q: 报价数据没有自动生成？

1. 检查多维表格中报价明细表的 AI 字段是否已配置
2. 确认需求编号已正确关联
3. 查看轮询日志，确认是否在等待数据稳定

### Q: 如何清理临时文件？

```bash
# 删除 7 天前的临时文件
find data/temp -name "*.pdf" -mtime +7 -delete
```

## AI Agent 集成 (SKILL.md)

本项目提供 `SKILL.md` 文件，让 AI Agent（如 SOLO、Cursor 等）也能调用项目的飞书工具，实现半自动化的报价处理。

### 设计思路

将项目中的操作分为两类：

| 类型 | 说明 | 示例 |
|------|------|------|
| **适合 Skill 指令** | 无状态、单次执行的 lark-cli 命令 | 创建记录、查询数据、创建云文档、发送通知 |
| **维持 Python 代码** | 有状态、长连接、复杂依赖 | IMAP 监听、AI 识别、PDF 生成、SMTP 发送、轮询 |

### AI Agent 可直接调用的 lark-cli 指令

**多维表格操作 (base)：**
- `lark-cli base +record-batch-create` — 创建询盘记录
- `lark-cli base +record-get` — 读取记录详情
- `lark-cli base +record-list` — 查询报价明细
- `lark-cli base +record-batch-update` — 更新询盘状态
- `lark-cli drive +upload` + `lark-cli api PATCH` — 上传附件

**云文档操作 (docs)：**
- `lark-cli docs +create` — 创建报价单/邮件预览云文档
- `lark-cli docs +fetch` — 读取用户修改后的云文档内容

**消息操作 (im)：**
- `lark-cli im +messages-send` — 发送文本/富文本通知
- `lark-cli im +chat-messages-list` — 查询群消息

### AI Agent 典型操作流程

```
流程 A：手动创建询盘记录
  → +record-batch-create → +record-get → drive +upload → api PATCH → im +messages-send

流程 B：查询报价结果
  → +record-list（按需求编号过滤）→ 汇总数据

流程 C：生成报价单云文档
  → 构造 Markdown → docs +create → 发送链接给用户 → docs +fetch 读取确认后内容

流程 D：发送飞书通知
  → im +messages-send --msg-type post
```

> 完整的指令文档、参数说明、返回值格式、注意事项请参阅 [SKILL.md](SKILL.md)。

## 技术栈

| 组件 | 技术 |
|------|------|
| 语言 | Python 3.10+ |
| 邮件监听 | IMAP (imaplib) |
| 邮件发送 | SMTP (smtplib) |
| PDF 生成 | WeasyPrint (HTML → PDF) |
| PDF 读取 | PyMuPDF (fitz) |
| AI 模型 | 火山方舟 豆包大模型 (OpenAI 兼容接口) |
| 飞书集成 | lark-cli (命令行工具) |
| 配置管理 | YAML + python-dotenv |

## License

MIT
