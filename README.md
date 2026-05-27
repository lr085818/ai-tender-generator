# AI智能招投标文档生成系统

基于 DeepSeek V4 Flash 大语言模型的智能投标文件生成平台。自动分析招标文件，生成符合格式要求的标准 Word 投标响应文件。

---

## 核心功能

- **招标文件智能分析** — 自动提取招标要求、评分标准、技术参数
- **章节结构设计** — AI 根据招标文件自动设计投标书目录结构
- **动态 Word 文档生成** — 完全按照 AI 设计的章节结构生成投标响应文件
- **设备清单提取** — 自动从招标文件中提取设备/货物清单填入报价表
- **AI 技术方案撰写** — 自动生成项目实施方案、培训方案、售后服务方案
- **在线文档预览** — 类似 Word 的只读预览，支持直接下载

## 技术栈

| 组件 | 技术 |
|------|------|
| 后端框架 | ![Flask](https://img.shields.io/badge/Flask-000000?style=flat-square&logo=flask&logoColor=white) (Python 3.9+) |
| 大语言模型 | ![DeepSeek](https://img.shields.io/badge/DeepSeek-4F6BED?style=flat-square&logo=deepseek&logoColor=white) V4 Flash |
| 关系型数据库 | ![SQLite](https://img.shields.io/badge/SQLite-003B57?style=flat-square&logo=sqlite&logoColor=white) |
| 文档生成 | ![python-docx](https://img.shields.io/badge/python--docx-3776AB?style=flat-square&logo=python&logoColor=white) |
| 前端 | ![Vue.js](https://img.shields.io/badge/Vue_3-4FC08D?style=flat-square&logo=vuedotjs&logoColor=white) + ![Element Plus](https://img.shields.io/badge/Element_Plus-409EFF?style=flat-square&logo=element&logoColor=white) |

## 项目结构

```
├── main.py                    # 服务入口
├── routes.py                  # API 路由（上传、预分析、章节分析等）
├── generate_route.py          # 投标文件生成路由（AI 内容生成 + 设备清单提取）
├── bid_document_generator.py  # Word 文档动态渲染引擎（核心）
├── qwen_client.py             # DeepSeek API 封装
├── users.py                   # 用户管理
├── templates/index.html       # 前端页面（Vue 3 + Element Plus）
├── .env                       # 环境变量配置
├── outputs/                   # 生成文件输出目录
├── uploads/                   # 上传文件存储
└── start.bat                  # 一键启动脚本
```

## 业务流程

```
上传招标文件 (.docx/.pdf)
       │
       ▼
 AI 预分析（提取需求/摘要/评分标准）
       │
       ▼
 AI 章节分析（提取响应文件格式）
       │
       ▼
 AI 章节设计（生成投标书目录结构）
       │
       ▼
 ┌─────────────────────────────────────┐
 │  Word 动态渲染引擎                   │
 │                                     │
 │  根据 AI 设计的章节结构逐章渲染：     │
 │  ├─ 封面 + 目录                      │
 │  ├─ 各章节 sections/subsections      │
 │  ├─ 分项报价表（含设备清单表格）      │
 │  ├─ AI 生成实施方案/培训/售后内容     │
 │  └─ 签名区 / 彩页占位                │
 └─────────────────────────────────────┘
       │
       ▼
 在线预览 / 下载 .docx 投标文件
```

## API 接口

| 方法 | 路径 | 说明 |
|------|------|------|
| POST | `/api/bidding/upload` | 上传招标文件 |
| POST | `/api/bidding/pre-analysis_bid` | AI 预分析 |
| POST | `/api/bidding/chapter-analysis_bid` | 章节分析 |
| POST | `/api/bidding/chapter-design` | 章节设计 |
| POST | `/api/bidding/generate-bid-document` | 生成投标文件 |
| GET  | `/api/bidding/document-preview/<id>` | 文档预览 |
| POST | `/api/users/identify` | 用户识别 |

## 快速开始

### 1. 克隆并安装依赖

```bash
git clone https://github.com/lr085818/ai-tender-generator.git
cd ai-tender-generator
pip install flask flask-cors python-docx mammoth PyPDF2 python-dotenv requests
```

### 2. 配置 API Key

```bash
# 复制环境变量模板
cp .env.example .env
```

编辑 `.env` 文件，填入你的 DeepSeek API Key：

```ini
DEEPSEEK_API_KEY=sk-你的真实key
```

> 获取 API Key：[DeepSeek Platform](https://platform.deepseek.com)（注册免费，送额度）

### 3. 启动服务

**方式一：一键启动**
```
双击 start.bat
```

**方式二：命令行**
```bash
python main.py
```

访问 [http://localhost:3012](http://localhost:3012)

## 文档格式规范

- **页面**：A4，上边距 2.6cm，下边距 2.2cm，左右边距 2.5cm
- **封面**：主标题 48pt 加粗，副标题 36pt 加粗，项目信息 14pt 加粗
- **正文**：宋体 12pt，首行缩进 2 字符，1.5 倍行距
- **标题编号**：一级 一、二、三… | 二级 1. 2. 3.… | 三级 1.1 1.2 1.3…
- **表格**：完整边框，表头黑体加粗

## 模型切换

系统使用 OpenAI 兼容协议，切换模型只需修改 `.env`：

```ini
# 切换为通义千问
DEEPSEEK_API_KEY=your_dashscope_key
DEEPSEEK_BASE_URL=https://dashscope.aliyuncs.com/compatible-mode/v1
DEEPSEEK_CHAT_MODEL=qwen-plus
```

## 许可证

MIT License
