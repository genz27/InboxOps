# InboxOps

一个基于 Flask 的管理员后台，支持：

- 管理员登录
- 添加、编辑、删除多个微软邮箱档案
- 针对指定邮箱切换三种接入方式
  - 旧版 IMAP
  - 新版 IMAP
  - Microsoft Graph API
- 查看收件箱概览、邮件列表、邮件详情
- 切换邮件已读 / 未读状态

## 启动方式

```bash
python -m venv venv
venv\Scripts\activate
pip install -r requirements.txt
python app.py
```

默认访问地址：`http://127.0.0.1:5000`

## 默认管理员账号

如果没有配置环境变量，默认管理员登录信息为：

- 用户名：`admin`
- 密码：`admin123456`

建议启动前通过环境变量覆盖：

```powershell
$env:MAIL_ADMIN_USERNAME="admin"
$env:MAIL_ADMIN_PASSWORD="your-strong-password"
$env:MAIL_ADMIN_SECRET_KEY="replace-this-secret"
python app.py
```

## 数据存储

- 邮箱档案使用 SQLite，默认文件：`data/mailboxes.db`
- 每个档案保存：
  - 备注名称
  - 邮箱账号
  - Client ID
  - Refresh Token
  - 代理
  - 默认方法
  - 备注

## 已实现接口

- `GET /api/auth/me`
- `POST /api/auth/login`
- `POST /api/auth/logout`
- `GET /api/mailboxes?q=&page=&page_size=`
- `POST /api/mailboxes`
- `GET /api/mailboxes/<id>`
- `PUT /api/mailboxes/<id>`
- `DELETE /api/mailboxes/<id>`
- `POST /api/mailboxes/delete/batch`
- `POST /api/mailboxes/test-connection`
- `POST /api/mailboxes/test-connection/batch`
- `POST /api/mailboxes/preferred-method/batch`
- `POST /api/mailbox/overview`
- `POST /api/mailbox/messages`
- `POST /api/mailbox/message`
- `POST /api/mailbox/message/read-state`

## 列表接口说明

- `GET /api/mailboxes` 现在返回邮箱摘要列表，不再包含 `client_id`、`refresh_token`、`proxy`
- 支持查询参数：
  - `q`：按备注名称、邮箱账号、备注信息搜索
  - `page`：页码，默认 `1`
  - `page_size`：每页数量，默认 `20`
- 编辑邮箱时，请继续使用 `GET /api/mailboxes/<id>` 获取完整详情
- 侧边栏支持跨页多选，批量测试连接，以及批量切换默认接入方式
- 侧边栏搜索词、页码、跨页已选邮箱、批量方法选择会保存在当前浏览器会话中

## 注意事项

- `refresh_token` 和 `client_id` 需要具备对应 IMAP 或 Graph API 权限。
- IMAP 连接本身仍是直连，当前代理主要覆盖 token 和 Graph 的 HTTP 请求。
- 当前多邮箱后台适合内部管理使用；若用于生产环境，建议再补密码哈希、CSRF、防爆破和敏感信息加密存储。
