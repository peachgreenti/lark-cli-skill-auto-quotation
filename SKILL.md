---
name: auto-quotation
description: 外贸报价全链路自动化。监听邮箱询价邮件，自动录入飞书多维表格，生成报价单并发送给客户。
---

# Auto Quotation — AI Agent Skill 指令

外贸报价全链路自动化系统。监听邮箱收到询价邮件后，自动完成：AI 识别询价 → 提取报价需求 → 写入飞书多维表格 → 等待报价完成 → 生成报价单云文档 → 用户确认 → 生成 PDF → 生成邮件预览 → 用户确认 → 发送回复邮件。

## 架构概览

```
┌─────────────────────────────────────────────────────────────┐
│                    Python 自动化流水线                        │
│  main.py (编排)                                              │
│  ├── mail_watcher.py  ← IMAP 长连接监听                      │
│  ├── mail_parser.py   ← 邮件解析 (X-Sender/Return-Path)      │
│  ├── inquiry_detector.py ← AI 询价识别 (OpenAI API)          │
│  ├── base_writer.py   ← 飞书多维表格 CRUD (lark-cli base)    │
│  ├── quotation_poller.py ← 轮询报价结果 (lark-cli base)      │
│  ├── quotation_generator.py ← 报价单/PDF生成 (lark-cli docs) │
│  ├── mail_sender.py   ← SMTP 回复邮件                        │
│  └── notifier.py      ← 飞书群通知 (lark-cli im)             │
└─────────────────────────────────────────────────────────────┘
```

---

## 第一部分：适合 AI Agent 直接调用的 lark-cli 工具指令

以下指令基于 `lark-cli` 命令行工具，AI Agent 可以直接在终端中执行这些命令来操作飞书。每条指令都是**无状态的、单次执行**的操作。

### 前置依赖

执行以下任何指令前，需确保：
1. 已安装 `lark-cli` 并完成认证（`lark-cli auth login`）
2. 已配置 `config.yaml` 中的 `base.app_token`、各 `table_id`、`notify.chat_id`

### 1.1 多维表格操作 (lark-cli base)

#### 创建询盘记录

```bash
lark-cli base +record-batch-create \
  --base-token <app_token> \
  --table-id <inquiry_table_id> \
  --json '{"fields":["报价需求","询价状态","查询日期","查询时间"],"rows":[["3个CP001、2个SP002","调整/准备中","2026-04-19","2026-04-19 09:00"]]}'
```

**说明：** 在询盘需求表中创建一条新记录。`rows` 中的字段顺序必须与 `fields` 一一对应。

**返回：** `{"ok": true, "data": {"record_id_list": ["recXXX"]}}`

**关键字段说明：**
| 字段名 | 类型 | 说明 |
|--------|------|------|
| 报价需求 | 文本 | AI 从询价函中提取的产品需求（如 "3个CP001、2个SP002"） |
| 询价状态 | 单选 | 固定值："调整/准备中" |
| 查询日期 | 日期 | YYYY-MM-DD |
| 查询时间 | 日期时间 | YYYY-MM-DD HH:MM |

#### 读取记录详情

```bash
lark-cli base +record-get \
  --base-token <app_token> \
  --table-id <inquiry_table_id> \
  --record-id <record_id>
```

**说明：** 读取单条记录的完整字段值。常用于：
- 创建记录后读取自动编号字段（需求编号）
- 读取已有附件字段（上传附件前需先读取以保留已有附件）

**返回：** `{"ok": true, "data": {"record": {"报价需求": "...", "需求编号": "20260419090", ...}}}`

#### 查询报价明细列表

```bash
lark-cli base +record-list \
  --base-token <app_token> \
  --table-id <quotation_table_id> \
  --limit 500 \
  --field-id "产品名称" \
  --field-id "产品型号" \
  --field-id "数量" \
  --field-id "产品售价/元" \
  --field-id "总价/元" \
  --field-id "客户名称" \
  --field-id "报价状态" \
  --field-id "需求编号" \
  --field-id "主报价需求" \
  --field-id "子报价需求"
```

**说明：** 查询报价明细表中的记录。返回的 `data.data` 是二维数组，每行对应一条记录，列顺序与 `--field-id` 参数顺序一致。

**返回：**
```json
{
  "ok": true,
  "data": {
    "data": [
      ["产品A", "Model-X", 100, 10.50, 1050.00, "客户A", "调整/准备中", "20260419090", "...", "..."],
      ...
    ],
    "record_id_list": ["recYYY", ...]
  }
}
```

**注意：** 需在客户端按 `需求编号` 字段过滤，只保留目标需求的记录。

#### 更新询盘状态

```bash
lark-cli base +record-batch-update \
  --base-token <app_token> \
  --table-id <inquiry_table_id> \
  --json '{"records":[{"record_id":"<record_id>","fields":{"询价状态":"待客户确认"}}]}'
```

**说明：** 更新询盘记录的状态字段。

**可选状态值：** `调整/准备中`、`待客户确认`、`客户已接受`、`客户已拒绝`

#### 上传文件到云空间

```bash
lark-cli drive +upload --file <filename> --name <filename>
```

**说明：** 将本地文件上传到飞书云空间。需在文件所在目录下执行（`cwd` 参数），或使用完整路径。

**返回：** `{"ok": true, "data": {"file_token": "boxcnXXX", "size": 12345}}`

#### 通过 API 写入附件字段

```bash
lark-cli api PATCH \
  "/open-apis/base/v3/bases/<app_token>/tables/<inquiry_table_id>/records/<record_id>" \
  --data '{"询价函":[{"deprecated_set_attachment":true,"file_token":"<file_token>","name":"<filename>"}]}'
```

**说明：** 将已上传的文件关联到询盘记录的附件字段。`deprecated_set_attachment: true` 标记用于绕过附件字段的 READONLY 限制。

**注意：** 如需保留已有附件，需先用 `+record-get` 读取现有附件列表，合并后一起写入。

#### 搜索客户

```bash
lark-cli base +record-list \
  --base-token <app_token> \
  --table-id <customer_table_id> \
  --field-id "客户名称" \
  --limit 100
```

**说明：** 列出客户管理表中的所有客户，用于判断客户是否已存在。

---

### 1.2 云文档操作 (lark-cli docs)

#### 创建云文档

```bash
lark-cli docs +create \
  --title "报价单 QT-20260419-123 - 客户名称" \
  --markdown "# 报价单\n\n**报价编号：** QT-20260419-123\n..."
```

**说明：** 创建飞书云文档并写入 Markdown 内容。支持标准 Markdown 语法，飞书会自动渲染为富文本。

**返回：** `{"ok": true, "data": {"doc_id": "Q2oFdC1eYop5kcxlYWlcqJ4fnvg", "doc_url": "https://..."}}`

**关键点：**
- `doc_id` 和 `doc_token` 是同一个值，不同版本 lark-cli 返回字段名可能不同
- Markdown 中的表格会自动转为飞书表格（`<lark-table>` 标签）
- 文档 URL 格式：`https://www.feishu.cn/docx/<doc_id>`

#### 读取云文档内容

```bash
lark-cli docs +fetch --doc "https://www.feishu.cn/docx/<doc_id>"
```

**说明：** 读取云文档的 Markdown 内容（包含用户修改后的最新版本）。

**返回：**
```json
{
  "ok": true,
  "data": {
    "markdown": "# 报价单\n\n<lark-table>...</lark-table>\n..."
  }
}
```

**关键点：**
- 返回的 Markdown 中，表格以 `<lark-table>` / `<lark-td>` 等 HTML 标签呈现
- 需解析这些标签来提取表格数据
- `data.markdown` 和 `data.content` 两个字段都可能包含内容，优先取 `markdown`

---

### 1.3 消息操作 (lark-cli im)

#### 发送文本消息

```bash
lark-cli im +messages-send \
  --chat-id <chat_id> \
  --text "报价单已生成，请查看"
```

**说明：** 向飞书群发送纯文本消息。

#### 发送富文本消息 (Post)

```bash
lark-cli im +messages-send \
  --chat-id <chat_id> \
  --msg-type post \
  --content '{"zh_cn":{"title":"📊 报价单已生成","content":[[{"tag":"md","text":"**客户：** XX公司\n**总金额：** 1,234.56\n\n📄 [编辑报价单](https://www.feishu.cn/docx/xxx)"}]]}}'
```

**说明：** 向飞书群发送富文本消息，支持 Markdown 格式、链接等。

**content JSON 结构：**
```json
{
  "zh_cn": {
    "title": "消息标题",
    "content": [
      [{"tag": "md", "text": "Markdown 格式的消息正文"}]
    ]
  }
}
```

#### 查询群消息列表

```bash
lark-cli im +chat-messages-list \
  --chat-id <chat_id> \
  --page-size 5 \
  --sort desc
```

**说明：** 获取群聊最近的消息列表，用于轮询用户确认。

**返回：**
```json
{
  "ok": true,
  "data": {
    "messages": [
      {
        "create_time": "2026-04-19 01:17",
        "msg_type": "text",
        "content": "{\"text\":\"确认\"}"
      }
    ]
  }
}
```

**关键点：**
- `--sort desc` 表示按时间倒序（最新在前）
- `create_time` 是字符串格式 `"YYYY-MM-DD HH:MM"`，不是时间戳
- `content` 是 JSON 字符串，需二次解析：`JSON.parse(content).text`
- 只关注 `msg_type: "text"` 类型的消息

---

## 第二部分：不适合编写为 Skill 指令的模块（维持 Python 代码）

以下模块涉及**长连接、有状态操作、复杂依赖库、或需要持续运行的进程**，不适合拆分为独立的 Skill 指令，应维持现有的 Python 代码实现。

### 2.1 IMAP 邮件监听 (mail_watcher.py)

**不适合原因：**
- 需要维护 IMAP 长连接状态（`imaplib.IMAP4_SSL`）
- 包含断线重连逻辑（`_ensure_connected`）
- 需要持续运行的无限循环（每 60 秒检查一次）
- 涉及邮件 UID 追踪、已读标记等有状态操作

**AI Agent 替代方案：** 如果需要手动触发，可以告知用户手动将询价邮件转发或上传，然后跳过邮件监听直接从阶段④开始。

### 2.2 邮件解析 (mail_parser.py)

**不适合原因：**
- 依赖 Python `email` 标准库的复杂 MIME 解析
- 需要处理多种邮件格式（纯文本、HTML、多部分混合）
- X-Sender / Return-Path / From 头部的优先级链逻辑
- PDF 附件的二进制提取

**AI Agent 替代方案：** AI Agent 可以直接提取邮件中的关键信息（客户名、产品需求），跳过自动解析步骤。

### 2.3 AI 询价识别 (inquiry_detector.py)

**不适合原因：**
- 依赖 OpenAI API（火山方舟 Doubao 模型）
- 包含关键词预筛 + AI 二次确认的两阶段逻辑
- 需要处理 API 调用失败、返回格式异常等边界情况
- `extract_requirement_text` 需要精确的 Prompt 工程

**AI Agent 替代方案：** AI Agent 本身就具备理解邮件内容的能力，可以直接判断是否为询价邮件并提取需求信息，无需调用外部 AI。

### 2.4 PDF 生成 (quotation_generator.py 中的 WeasyPrint 部分)

**不适合原因：**
- 依赖 `weasyprint` Python 库（底层为 C 渲染引擎）
- 需要复杂的 HTML → PDF 转换逻辑
- 包含飞书 Markdown 特殊标签（`<lark-table>` 等）的解析和转换
- PDF 样式（CSS）与内容紧密耦合

**AI Agent 替代方案：** AI Agent 可以生成报价单云文档后，由用户手动导出 PDF，或调用 Python 脚本执行 PDF 生成。

### 2.5 SMTP 邮件发送 (mail_sender.py)

**不适合原因：**
- 依赖 SMTP SSL 长连接（`smtplib.SMTP_SSL`）
- 需要处理邮件线程（In-Reply-To、References 头部）
- PDF 附件的二进制编码（MIMEApplication）
- 重试机制（3 次重试，5 秒间隔）

**AI Agent 替代方案：** AI Agent 可以生成邮件内容后，由用户手动发送，或调用 Python 脚本执行发送。

### 2.6 报价轮询 (quotation_poller.py)

**不适合原因：**
- 需要持续运行的轮询循环（每 30 秒检查一次，最长 30 分钟）
- 包含数据稳定性检测（MD5 哈希比对，连续 60 秒无变化）
- 有状态操作（追踪 `last_hash`、`stable_since`）

**AI Agent 替代方案：** AI Agent 可以主动查询报价明细表（使用 1.1 中的 `+record-list` 指令），判断数据是否已生成，无需被动轮询。

### 2.7 飞书群消息轮询确认 (main.py 中的 wait_for_feishu_confirm)

**不适合原因：**
- 需要持续轮询群消息（每 10 秒检查一次，最长 30 分钟）
- 包含时间窗口过滤（只检查开始时间之后的消息）
- 需要解析多种消息格式（text、post 等）

**AI Agent 替代方案：** AI Agent 可以直接询问用户是否确认，无需通过飞书群消息轮询。

---

## 第三部分：AI Agent 可执行的典型操作流程

AI Agent 可以利用第一部分的 lark-cli 指令，手动执行以下操作流程（替代 Python 自动化流水线的部分环节）：

### 流程 A：手动创建询盘记录

当用户告知有新的询价需求时：

```
1. 执行 lark-cli base +record-batch-create 创建询盘记录
2. 执行 lark-cli base +record-get 读取需求编号
3. 如有询价函 PDF，执行 lark-cli drive +upload 上传文件
4. 执行 lark-cli api PATCH 关联附件到记录
5. 执行 lark-cli im +messages-send 通知用户
```

### 流程 B：查询报价结果

当需要检查报价是否已生成时：

```
1. 执行 lark-cli base +record-list 查询报价明细
2. 在返回结果中按需求编号过滤
3. 汇总报价数据（产品名称、数量、单价、总价）
```

### 流程 C：生成报价单云文档

当报价数据已就绪时：

```
1. 根据报价数据构造 Markdown 内容（参考 _build_markdown_cn 格式）
2. 执行 lark-cli docs +create 创建云文档
3. 将文档链接发送给用户确认
4. 用户确认后，执行 lark-cli docs +fetch 读取最新内容
```

### 流程 D：发送飞书通知

在任何需要通知用户的场景：

```
1. 执行 lark-cli im +messages-send --msg-type post 发送富文本通知
2. 通知内容中可包含多维表格链接和云文档链接
```

---

## 第四部分：完整自动化流水线（Python 程序）

当需要端到端自动执行时，运行 Python 程序：

### 正常模式（监听邮箱）

```bash
cd lark-cli-skill-auto-quotation
python3 main.py
```

### 测试模式（跳过邮件监听）

```bash
python3 main.py --test <需求编号>
```

### 完整流程（10 个阶段）

```
① IMAP 监听邮箱，检测新邮件
② AI 识别是否为询价邮件（关键词预筛 + AI 二次确认）
③ 解析邮件：提取正文、PDF 附件文本
④ AI 提取报价需求 → 写入飞书多维表格（询盘需求表）+ 上传附件
⑤⑥ 轮询等待报价完成（多维表格 AI 字段自动生成报价明细，60 秒稳定）
⑦  生成报价单飞书云文档（用户可编辑）
⑧  飞书群通知（含云文档链接）→ 用户在云文档中调整报价 → 群内回复「确认」
⑦② 从用户确认后的云文档生成 PDF 报价单
⑨  生成邮件预览飞书云文档 → 飞书群通知 → 用户调整邮件内容 → 群内回复「确认」
⑩  从邮件预览云文档读取最终内容 → 发送回复邮件 + PDF 附件给客户
```

### 用户交互点

系统通过飞书群消息与用户交互，共 2 个确认节点：

1. **阶段⑧**：报价单云文档确认 — 用户在云文档中调整报价后，在群内回复「确认」或「取消」
2. **阶段⑨**：邮件预览确认 — 用户在邮件预览云文档中调整内容后，在群内回复「确认」或「取消」

---

## 第五部分：配置参考

### 环境变量（.env）

| 变量名 | 说明 |
|--------|------|
| `IMAP_SERVER` | IMAP 服务器地址（如 `imap.163.com`） |
| `IMAP_PORT` | IMAP 端口（如 `993`） |
| `IMAP_USER` | IMAP 用户名（邮箱地址） |
| `IMAP_PASSWORD` | IMAP 授权码 |
| `SMTP_SERVER` | SMTP 服务器地址（如 `smtp.163.com`） |
| `SMTP_PORT` | SMTP 端口（如 `465`） |
| `SMTP_USER` | SMTP 用户名 |
| `SMTP_PASSWORD` | SMTP 授权码 |
| `OPENAI_API_KEY` | AI API Key（火山方舟） |
| `OPENAI_BASE_URL` | AI API Base URL |
| `FEISHU_APP_ID` | 飞书应用 ID |
| `FEISHU_APP_SECRET` | 飞书应用 Secret |

### config.yaml 关键配置

```yaml
base:
  app_token: "多维表格 app token"
  inquiry_table_id: "询盘需求表 ID"
  quotation_table_id: "报价明细表 ID"
  customer_table_id: "客户管理表 ID"

notify:
  chat_id: "飞书群聊 ID"

quotation:
  company_name_cn: "公司中文名"
  company_name_en: "公司英文名"
  validity_days: 30
  currency: "USD"

polling:
  check_interval: 60
  quotation_stable_time: 60
  quotation_poll_interval: 30
```

### 多维表格字段映射

**询盘需求表 (inquiry_table_id)：**

| 字段名 | 字段 ID | 类型 | 说明 |
|--------|---------|------|------|
| 客户名称 | fldgrNbTnQ | 关联选择 | 需在客户表中预先添加 |
| 报价需求 | fldLYihHyq | 文本 | AI 提取的产品需求 |
| 询价状态 | flds9zibsi | 单选 | 调整/准备中、待客户确认、客户已接受、客户已拒绝 |
| 询价函 | fld8kbT6o1 | 附件 | 询价函 PDF |
| 查询日期 | fld1QCOXYT | 日期 | YYYY-MM-DD |
| 查询时间 | fldvDiKseN | 日期时间 | YYYY-MM-DD HH:MM |
| 需求编号 | fldFvm5QQe | 自动编号 | 只读，创建后自动生成 |

**报价明细表 (quotation_table_id)：**

| 字段名 | 字段 ID | 类型 | 说明 |
|--------|---------|------|------|
| 产品名称 | fldhSwvJ3z | 查找自产品管理 | |
| 产品型号 | fldyHWDpBj | 文本 | |
| 数量 | fldCAUopQE | 数字 | |
| 产品售价/元 | fldQpjwnHO | 查找自产品管理 | |
| 总价/元 | fld8GjmjQV | 公式 | 只读 |
| 客户名称 | fldZQkyIIZ | 查找自询盘需求表 | |
| 报价状态 | fld4CII0Cl | 查找自询盘需求表 | |
| 需求编号 | fldHfNxmRA | 文本 | 用于过滤 |

---

## 第六部分：注意事项

1. **邮件发件人识别**：系统优先使用 `X-Sender` 头部（163/QQ 邮箱的真实发件人），其次 `Return-Path`，最后 `From`。回复邮件会发送到真实发件人地址。
2. **报价数据过滤**：轮询报价明细时按需求编号过滤，确保只获取当前询价的报价数据。
3. **云文档确认机制**：报价单和邮件预览均以飞书云文档形式供用户编辑确认，PDF 和邮件内容均从用户确认后的云文档生成。
4. **附件上传两步流程**：由于 `lark-cli +record-upload-attachment` 存在 0B bug，采用 drive +upload → api PATCH 两步流程。
5. **lark-cli 超时设置**：base 操作 60 秒超时，docs 操作 30 秒超时，im 操作 30 秒超时（轮询场景下 `+chat-messages-list` 使用 15 秒超时以加快错误恢复）。
