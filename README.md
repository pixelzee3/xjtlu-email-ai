# 📧 XJTLU 邮件智能助手

> 基于 Playwright + LLM 的 XJTLU 邮箱自动总结工具。关键词搜索、智能提取正文、一键生成中文摘要。

## ✨ 功能亮点

- 🔍 **关键词检索** — 搜索收件箱，或直接获取最新邮件
- 📄 **正文提取** — 自动点击并提取多封邮件全文（支持深度扫描模式）
- 🤖 **AI 总结** — 并行调用 LLM（支持 OpenAI / DeepSeek 等）生成智能摘要
- 🌐 **Web 界面** — 基于 FastAPI 的浏览器操作界面，零门槛可视化配置
- 🍪 **Cookie 登录** — 适配 OWA 双因素认证环境
- 📝 **Prompt 预设** — 内置多种总结风格（综合摘要、求职、活动、Deadline 等），一键切换
- ⏰ **定时摘要** — 每日/每周自动生成摘要，服务端后台排队执行
- 🎯 **优先级排序** — 深度扫描模式下按紧急程度自动评分排序
- 👥 **多用户** — 本地多账号、独立配置与历史

## 🚀 快速开始（3 步上手）

### 第 1 步：安装

**Windows 新手推荐** — 双击 **`run_helper.bat`**，在图形界面中一键安装依赖：

![helper](https://img.shields.io/badge/双击-run__helper.bat-brightgreen?style=for-the-badge)

或者命令行：

```bash
pip install -r src/requirements.txt
playwright install msedge
```

### 第 2 步：启动

**双击 `run_app.bat`** 即可启动 Web 服务并自动打开浏览器。

或命令行：

```bash
cd src
python app.py
# 访问 http://localhost:8001
```

首次启动会提示**注册账号**（邮箱 + 密码），注册后即可进入主界面。

### 第 3 步：配置 & 使用

1. 点击页面右上角 **⚙️ 设置**
2. 填入你的 **AI API Key**（OpenAI / DeepSeek 等）和 **Base URL**
3. 粘贴学校邮箱的 **Cookie**（用 Cookie-Editor 插件导出 JSON）
4. 保存 → 回到主页 → 输入关键词或留空 → 点击 **生成摘要**

> **获取 Cookie**：在 Edge/Chrome 中安装 [Cookie-Editor](https://cookie-editor.com/) → 登录学校邮箱 → Export as JSON → 粘贴到设置中。

---

## 🖥️ 界面预览

启动后你会看到：

- **搜索栏** — 输入关键词或留空获取最新邮件
- **Prompt 预设** — 一键选择总结风格（综合摘要、求职招聘、Deadline 提取等）
- **摘要卡片** — AI 生成的结构化摘要，支持卡片/原文视图切换
- **侧栏** — 搜索历史、定时摘要设置、安全说明

---

## ⚙️ 技术栈

| 模块         | 技术                            |
| ------------ | ------------------------------- |
| 浏览器自动化 | Playwright (Microsoft Edge)     |
| Web 服务     | FastAPI + Uvicorn               |
| AI 接口      | OpenAI 兼容 API（并行批量调用） |
| 正文解析     | BeautifulSoup4                  |
| 本地账号     | SQLite + bcrypt                 |

---

## 📂 项目结构

```
├── run_app.bat              # 一键启动 Web 服务
├── run_helper.bat           # 图形化安装与环境诊断
├── src/
│   ├── app.py               # FastAPI 服务入口
│   ├── main.py              # 邮件抓取与 LLM 核心
│   ├── auth_db.py           # 用户账号与配置存储
│   ├── digest_schedule.py   # 定时摘要逻辑
│   ├── deep_priority.py     # 深度扫描优先级评分
│   ├── startup_helper_*.py  # 启动助手（环境检查）
│   ├── templates/           # 前端页面
│   └── config.example.json  # 配置模板
└── docs/                    # 详细文档、维护指南、计划等
```

---

## 📝 配置说明

| 字段                | 说明                             |
| ------------------- | -------------------------------- |
| `ai.base_url`       | LLM API 地址                     |
| `ai.api_key`        | API Key                          |
| `ai.model`          | 模型名称（如 `gpt-4o-mini`）     |
| `email.url`         | 邮箱 URL（默认 XJTLU OWA）       |
| `email.cookies`     | Cookie JSON 数组                 |
| `email.cookie_file` | 或 Netscape 格式 Cookie 文件路径 |

> 推荐通过 **Web 界面设置**，无需手动编辑文件。

---

## ⚠️ 注意事项

- **Cookie 有效期**：通常几天到数周后失效，需重新导出
- **隐私安全**：`config.json`、`cookies.txt`、`user.db` 均已在 `.gitignore` 中，**请勿提交到公开仓库**
- **选择器兼容**：若 OWA 更新 UI，可在设置中调整 CSS 选择器

---

## 📚 更多文档

| 文档                                                                         | 说明                                                     |
| ---------------------------------------------------------------------------- | -------------------------------------------------------- |
| [`docs/README.md`](docs/README.md)                                           | 完整技术文档与部署说明                                   |
| [`docs/i18n-translation-workflow.md`](docs/i18n-translation-workflow.md)     | 前端中英文本地化（i18n）开发与补译流程                   |
| [`AGENTS.md`](AGENTS.md)                                                     | 面向 Cursor / Claude Code / Codex 等代理的 i18n 执行规范 |
| [`docs/维护指南.txt`](docs/维护指南.txt)                                     | Cookie 过期、OWA 调整等运维指南                          |
| [`docs/项目完整介绍与文件说明.txt`](docs/项目完整介绍与文件说明.txt)         | 每个文件的作用详解                                       |
| [`docs/校友产品优先级计划.md`](docs/校友产品优先级计划.md)                   | 产品路线图                                               |
| [`docs/在线部署与本地抓取方案教程.txt`](docs/在线部署与本地抓取方案教程.txt) | 上线部署教程                                             |

---

## 📄 License

MIT License
