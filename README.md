# 外贸报价全链路自动化系统

基于飞书多维表格 + lark-cli + AI 的外贸询价自动处理系统。自动监听邮箱询价邮件，写入飞书多维表格触发报价流程，生成报价单云文档和 PDF，经人工确认后自动回复客户。

## 系统架构

```
客户询价邮件
    ↓
① IMAP 监听收件箱（只处理新邮件）
    ↓
② AI 识别询价邮件 + 提取客户信息
    ↓
③ 解析邮件正文/PDF，AI 提取报价需求
    ↓
④ 写入飞书多维表格（询盘需求 + 附件上传）
    ↓
⑤ 多维表格自动处理（AI 字段捷径生成报价明细）
    ↓
⑥ 轮询检测报价完成（按需求编号过滤，60s 稳定）
    ↓
⑦ 生成报价单飞书云文档
    ↓
⑧ 飞书群通知 → 用户在云文档中调整报价 → 群内回复「确认」
    ↓
⑦② 从用户确认后的云文档生成 PDF
    ↓
⑨ 生成邮件预览云文档 → 用户调整 → 群内回复「确认」
    ↓
⑩ 从邮件预览云文档读取内容 → 发送回复邮件 + PDF 附件
```

## 功能特性

- **自动邮件监听**：IMAP 协议监听收件箱，只处理启动后的新邮件
- **AI 询价识别**：关键词预筛 + AI 二次确认，自动提取客户名称和语言
- **智能需求提取**：从 PDF 附件或邮件正文中 AI 提取产品报价需求
- **飞书多维表格集成**：自动创建询盘记录、上传附件、读取报价结果
- **云文档协作**：报价单和邮件预览均生成飞书云文档，支持人工调整
- **PDF 报价单生成**：基于用户确认后的云文档内容生成 PDF
- **自动邮件回复**：从邮件预览云文档读取内容，回复真实发件人
- **中英双语支持**：根据客户语言自动切换报价单语言
- **测试模式**：支持 `--test <需求编号>` 直接从已有数据开始处理

## 前置条件

- Python 3.10+
- [lark-cli](https://github.com/nicepkg/lark-cli) 已安装并完成登录授权
- 飞书多维表格（需创建三个表：询盘需求表、报价明细表、客户管理表）
- 飞书机器人（已加入目标群聊）
- 支持 IMAP/SMTP 的邮箱（如网易邮箱、QQ 邮箱）

## 安装

```bash
# 1. 克隆项目
git clone https://github.com/your-username/lark-cli-skill-auto-quotation.git
cd lark-cli-skill-auto-quotation

# 2. 安装依赖
pip install -r requirements.txt

# 3. 配置环境变量
cp .env.example .env
# 编辑 .env 填入真实的邮箱密码、AI API Key 等

# 4. 配置飞书多维表格
cp config.yaml.example config.yaml
# 编辑 config.yaml 填入多维表格的 app_token、table_id、字段 ID 等
```

## 配置说明

### 1. 环境变量（`.env`）

| 变量 | 说明 | 示例 |
|------|------|------|
| `IMAP_SERVER` | IMAP 服务器地址 | `imap.163.com` |
| `IMAP_PORT` | IMAP 端口 | `993` |
| `IMAP_USER` | 邮箱账号 | `your@163.com` |
| `IMAP_PASSWORD` | IMAP 授权码（非登录密码） | `xxxxxxxx` |
| `SMTP_SERVER` | SMTP 服务器地址 | `smtp.163.com` |
| `SMTP_PORT` | SMTP 端口 | `465` |
| `SMTP_USER` | SMTP 发件账号 | `your@163.com` |
| `SMTP_PASSWORD` | SMTP 授权码 | `xxxxxxxx` |
| `OPENAI_API_KEY` | AI API Key（火山方舟 Ark） | `ark-xxx` |
| `OPENAI_BASE_URL` | AI API 地址 | `https://ark.cn-beijing.volces.com/api/v3` |

### 2. 飞书多维表格配置（`config.yaml`）

需要创建以下三个表：

**询盘需求表**（`inquiry_table_id`）：
- 报价需求（文本）：AI 提取的产品需求描述
- 询价状态（单选）：调整/准备中、待客户确认、客户已接受、客户已拒绝
- 查询日期（日期）
- 查询时间（日期时间）
- 询价函（附件）
- 需求编号（自动编号，只读）

**报价明细表**（`quotation_table_id`）：
- 产品名称、产品型号、数量、产品售价/元、总价/元
- 客户名称（关联询盘需求表）
- 报价状态（关联询盘需求表）
- 需求编号（文本）
- 主报价需求、子报价需求

**客户管理表**（`customer_table_id`）：
- 客户名称、邮箱、联系人、联系电话

> 各表的字段 ID 需要在多维表格中查看并填入 `config.yaml`。

### 3. 飞书群聊配置

- 创建一个飞书群聊
- 将飞书机器人加入群聊
- 获取群聊 ID（`chat_id`）填入 `config.yaml`
- 系统会在群中发送通知，用户在群内回复「确认」或「取消」来控制流程

## 使用方法

### 正常模式（监听邮箱）

```bash
python3 main.py
```

系统启动后会：
1. 连接 IMAP 服务器，跳过所有历史未读邮件
2. 每 60 秒检查一次新邮件
3. 发现询价邮件后自动进入处理流程
4. 在飞书群中通知用户确认

### 测试模式（跳过邮件监听）

```bash
# 直接从已有的需求编号开始处理（跳过阶段①-④）
python3 main.py --test 20260419090
```

## 项目结构

```
lark-cli-skill-auto-quotation/
├── main.py                  # 主入口（邮件监听循环 + 流程编排）
├── mail_watcher.py          # IMAP 邮件监听（新邮件检测 + 去重）
├── mail_parser.py           # 邮件解析（正文/附件/发件人提取）
├── inquiry_detector.py      # AI 询价识别（关键词预筛 + AI 确认）
├── base_writer.py           # 飞书多维表格操作（创建记录/上传附件/查询报价）
├── quotation_poller.py      # 报价完成检测（轮询 + 稳定判断）
├── quotation_generator.py   # 报价单生成（云文档 + PDF + 邮件预览）
├── mail_sender.py           # 邮件发送（SMTP 回复 + PDF 附件）
├── notifier.py              # 飞书通知（群消息发送）
├── config.yaml              # 配置文件（不提交到 Git）
├── config.yaml.example      # 配置模板
├── .env                     # 环境变量（不提交到 Git）
├── .env.example             # 环境变量模板
├── .gitignore               # Git 忽略规则
├── requirements.txt         # Python 依赖
└── data/temp/               # 临时文件目录（PDF 等）
```

## 依赖

```
pyyaml
python-dotenv
PyMuPDF
weasyprint
httpx
```

## 注意事项

1. **邮箱授权码**：网易邮箱需要在设置中开启 IMAP/SMTP 并获取授权码（非登录密码）
2. **lark-cli 授权**：使用前需运行 `lark-cli login` 完成飞书授权
3. **AI 字段捷径**：多维表格中需配置 AI 字段捷径，自动从询盘需求生成报价明细
4. **首次启动**：系统会跳过所有历史未读邮件，只处理启动后收到的新邮件
5. **真实发件人**：系统优先从 `X-Sender` / `Return-Path` 头部提取真实发件人作为回复地址

## License

MIT
