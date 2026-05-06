# 营销数据报表自动化工具

## 如何使用

### 1. 克隆项目

```bash
git clone https://github.com/troywjz/MKT_Data_Report_Auto.git
cd MKT_Data_Report_Auto
```

### 2. 创建虚拟环境并安装依赖

Windows PowerShell:

```powershell
python -m venv venv
.\venv\Scripts\pip.exe install -r requirements.txt
```

### 3. 配置本地环境变量

复制 `.env.example` 为 `.env`，并填写本地真实配置：

```powershell
Copy-Item .env.example .env
```

需要重点填写：

- 各平台账号和密码：`BAIDU_USERNAME`、`BAIDU_PASSWORD`、`BING_USERNAME`、`BING_PASSWORD`、`CRM_USERNAME`、`CRM_PASSWORD`、`SLL_USERNAME`、`SLL_PASSWORD`
- 各平台登录页、目标页、URL 关键字
- Chromium 用户数据目录和调试端口
- DingTalk 机器人相关配置

`.env` 不会被 Git 提交。

### 4. 准备 Excel 模板

仓库提供脱敏示例模板：

- `examples/data_template.example.xlsx`
- `examples/config.example.xlsx`
- `examples/daily_output.example.xlsx`

首次运行前，将示例模板复制为真实运行文件，并按业务情况补充真实账户、渠道、返点、线索识别规则等数据：

```powershell
Copy-Item .\examples\data_template.example.xlsx .\数据模版.xlsx
```

如果仍想使用旧版 Excel 账号配置回退，可以复制：

```powershell
Copy-Item .\examples\config.example.xlsx .\config.xlsx
```

推荐优先使用 `.env` 管理账号密码。

### 5. 运行时报

```powershell
.\venv\Scripts\python.exe main.py shibao
```

### 6. 运行日报

自动补最近缺失日期：

```powershell
.\venv\Scripts\python.exe main.py ribao
```

指定日期：

```powershell
.\venv\Scripts\python.exe main.py ribao 2026-04-01
```

### 7. 浏览器登录说明

项目使用 DrissionPage 驱动 Chromium。首次运行时可能需要手动完成登录、短信验证或安全验证。登录态会保存在 `User_Data_Chrome/`，该目录包含账号状态和浏览器数据，不会提交 Git。

## 项目说明

本项目用于自动化生成市场投放日报和时报。它会从多个广告平台采集展现、点击、消费等数据，从 CRM 获取线索数据，再通过 Excel 模板进行汇总、匹配和报表生成。

主要能力：

- 抓取多平台广告投放数据。
- 导出并读取 CRM 线索数据。
- 基于 Excel 模板生成日报和时报。
- 生成时报发送模板和 Markdown 表格。
- 支持 DingTalk 机器人发送流程。
- 使用 `.env` 隔离账号、密码、URL、Webhook Token 等敏感配置。

## 安全注意事项

- 不要提交 `.env`、`config.xlsx`、真实业务 Excel、下载数据、浏览器用户数据目录。
- 如果真实 token 或密码曾经提交到远程仓库，应立即吊销并重新生成。
- 开源前建议运行一次敏感信息扫描。

---

# Marketing Data Report Automation

## How To Use

### 1. Clone The Repository

```bash
git clone https://github.com/troywjz/MKT_Data_Report_Auto.git
cd MKT_Data_Report_Auto
```

### 2. Create A Virtual Environment And Install Dependencies

Windows PowerShell:

```powershell
python -m venv venv
.\venv\Scripts\pip.exe install -r requirements.txt
```

### 3. Configure Local Environment Variables

Copy `.env.example` to `.env` and fill in your local private values:

```powershell
Copy-Item .env.example .env
```

Important values include platform credentials, login URLs, target URLs, Chromium user data path, local debugging port, and DingTalk robot settings.

`.env` is ignored by Git and must not be committed.

### 4. Prepare Excel Templates

Sanitized example templates are included:

- `examples/data_template.example.xlsx`
- `examples/config.example.xlsx`
- `examples/daily_output.example.xlsx`

Copy the main template before running:

```powershell
Copy-Item .\examples\data_template.example.xlsx .\数据模版.xlsx
```

If you need the legacy Excel credential fallback:

```powershell
Copy-Item .\examples\config.example.xlsx .\config.xlsx
```

Using `.env` for credentials is recommended.

### 5. Run The Hourly Report

```powershell
.\venv\Scripts\python.exe main.py shibao
```

### 6. Run The Daily Report

Automatically process missing dates:

```powershell
.\venv\Scripts\python.exe main.py ribao
```

Run a specific date:

```powershell
.\venv\Scripts\python.exe main.py ribao 2026-04-01
```

### 7. Browser Login

The project uses DrissionPage with Chromium. On the first run, manual login or verification may be required. Browser state is stored in `User_Data_Chrome/`, which is ignored by Git.

## Project Description

This project automates marketing daily and hourly reports. It collects ad metrics such as impressions, clicks, and cost from advertising platforms, reads lead data from CRM, then generates reports through Excel templates.

Main features:

- Multi-platform ad data collection.
- CRM lead export and parsing.
- Excel-based daily and hourly report generation.
- Markdown table generation for hourly report delivery.
- DingTalk robot delivery support.
- Sensitive configuration isolation through `.env`.

## Security Notes

- Never commit `.env`, `config.xlsx`, real business Excel files, downloaded data, or browser profile data.
- Rotate secrets immediately if they were ever committed to a remote repository.
- Run a secret scan before publishing the repository.
