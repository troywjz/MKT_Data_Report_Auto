---
name: mkt-report-shibao
description: Run the hourly marketing report workflow, generate a Markdown table, and optionally send it through DingTalk using environment variables.
triggers:
  - 做时报
  - 发时报
  - hourly report
---

# 市场时报 Skill

## 用途

当用户要求生成或发送市场时报时，运行项目脚本抓取当天多平台广告数据、汇总 CRM 线索，并更新时报发送模板。Webhook、账号、密码、URL 等敏感信息必须全部来自 `.env` 或系统环境变量。

## 前置条件

1. 已创建 `.env`，并按 `.env.example` 填写本地配置。
2. 已准备真实业务模板文件 `数据模版.xlsx`，该文件默认不提交 Git。
3. 已安装依赖：`pip install -r requirements.txt`。
4. 浏览器用户数据目录 `User_Data_Chrome/` 只保留在本地，不提交 Git。

## 运行时报

Windows PowerShell:

```powershell
.\venv\Scripts\python.exe main.py shibao
```

WSL 示例:

```bash
/mnt/d/path/to/project/venv/Scripts/python.exe main.py shibao
```

## Markdown 表格要求

读取 `数据模版.xlsx` 的 `时报发送模板` sheet，按原始数据生成 Markdown 表格。

要求：

1. 保留所有需要展示的列：时间、渠道、账户、线索、实际成本、名片成本、展现、点击、消费、实际消费、CPC、CTR、CVR、CPM。
2. 账户名称必须原样输出，不做替换、缩写或泛化。
3. 数据以 Excel 中的实际结果为准，不在 Skill 中手工重算。

## 发送到 DingTalk

发送逻辑应使用环境变量：

- `DINGTALK_SHIBAO_ROBOT_ACCESS_TOKEN`
- `DINGTALK_SHIBAO_WEBHOOK_URL`

不要在 Skill 文档中写真实 Webhook URL 或 access token。发送失败时，把接口返回值反馈给用户。
