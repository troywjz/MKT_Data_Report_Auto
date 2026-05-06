---
name: mkt-report-ribao
description: Run the daily marketing report workflow and optionally send the generated Excel file through DingTalk using environment variables.
triggers:
  - 做日报
  - 发日报
  - daily report
---

# 市场日报 Skill

## 用途

当用户要求生成或发送市场日报时，运行项目脚本生成日报 Excel。所有账号、密码、DingTalk AppKey/AppSecret、Webhook Token、业务后台 URL 都必须从本地 `.env` 或系统环境变量读取，不能写在 Skill 文档或命令里。

## 前置条件

1. 已创建 `.env`，并按 `.env.example` 填写本地配置。
2. 已准备真实业务模板文件 `数据模版.xlsx`，该文件默认不提交 Git。
3. 已安装依赖：`pip install -r requirements.txt`。
4. Excel 文件没有被桌面 Excel 程序打开，否则写入可能失败。

## 运行日报

Windows PowerShell:

```powershell
.\venv\Scripts\python.exe main.py ribao
```

指定日期:

```powershell
.\venv\Scripts\python.exe main.py ribao 2026-04-01
```

WSL 示例:

```bash
/mnt/d/path/to/project/venv/Scripts/python.exe main.py ribao
```

## 发送到 DingTalk

发送逻辑应使用环境变量：

- `DINGTALK_APP_KEY`
- `DINGTALK_APP_SECRET`
- `DINGTALK_RIBAO_ROBOT_ACCESS_TOKEN` 或 `DINGTALK_RIBAO_WEBHOOK_URL`

不要在命令、文档、日志中暴露真实 token。执行前先确认生成文件存在，例如 `01日报新增.xlsx`。

## 错误处理

- 脚本失败：把错误信息反馈给用户，不发送文件。
- 登录失败：提示用户检查 `.env` 中平台账号、密码、URL 配置，或在浏览器中完成验证。
- DingTalk 发送失败：反馈接口返回值，要求用户检查 token、机器人权限和网络。
