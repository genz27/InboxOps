# InboxOps

一个基于 Flask 的多邮箱管理员后台，支持：

- 管理员登录
- 添加、编辑、删除多个微软邮箱档案
- 在 `graph_api`、`imap_new`、`imap_old` 三种方式之间切换
- 查看收件箱概览、邮件列表、邮件详情
- 切换邮件已读 / 未读状态
- 侧边栏分页、搜索、跨页多选
- 批量测试连接、批量切换默认方式、批量删除
- 通过 `邮箱 + API Key` 直接读取邮件列表和邮件详情

## 启动方式

```bash
python -m venv venv
venv\Scripts\activate
pip install -r requirements.txt
python app.py
```

默认访问地址：`http://127.0.0.1:5000`

## 环境变量

如果没有额外配置，默认管理员账号为：

- 用户名：`admin`
- 密码：`admin123456`

推荐启动前显式配置：

```powershell
$env:MAIL_ADMIN_USERNAME="admin"
$env:MAIL_ADMIN_PASSWORD="your-strong-password"
$env:MAIL_ADMIN_SECRET_KEY="replace-this-secret"
$env:INBOXOPS_API_KEY="replace-this-project-key"
python app.py
```

说明：

- `MAIL_ADMIN_USERNAME`：管理员用户名
- `MAIL_ADMIN_PASSWORD`：管理员密码
- `MAIL_ADMIN_SECRET_KEY`：Flask Session 密钥
- `INBOXOPS_API_KEY`：项目级访问 Key，用于 `邮箱 + Key` 接口

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

## GitHub Packages / GHCR 自动发布

项目已增加 GitHub Actions 自动构建并发布镜像到 GitHub Container Registry。

- 工作流文件：`.github/workflows/publish-ghcr.yml`
- 镜像地址：`ghcr.io/genz27/inboxops`
- 触发条件：
  - 推送到 `main`
  - 推送版本标签，例如 `v1.0.0`
  - 在 Actions 页面手动执行 `workflow_dispatch`
- `main` 分支推送时会发布：
  - `latest`
  - `main`
  - `sha-<commit>`
- 推送语义化版本标签时会额外发布：
  - `v1.2.3`
  - `1.2`
  - `1`

拉取示例：

```bash
docker pull ghcr.io/genz27/inboxops:latest
docker pull ghcr.io/genz27/inboxops:v1.0.0
```

## 已实现接口

管理员接口：

- `GET /api/auth/me`
- `POST /api/auth/login`
- `POST /api/auth/logout`
- `GET /api/mailboxes?q=&page=&page_size=`
- `POST /api/mailboxes`
- `POST /api/mailboxes/import`
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

项目级 Key 接口：

- `POST /api/key/mailbox/messages`
- `POST /api/key/mailbox/message`

## 列表与侧边栏说明

- `GET /api/mailboxes` 返回邮箱摘要列表，不包含 `client_id`、`refresh_token`、`proxy`
- 支持查询参数：
  - `q`：按备注名称、邮箱账号、备注信息搜索
  - `page`：页码，默认 `1`
  - `page_size`：每页数量，默认 `20`
- 编辑邮箱时，继续通过 `GET /api/mailboxes/<id>` 获取完整详情
- 侧边栏支持跨页多选
- 支持批量测试连接、批量切换默认接入方式、批量删除
- 搜索词、页码、跨页已选邮箱、批量方法选择会保存在当前浏览器会话中

## 邮箱 + Key 接口

为了让外部服务不依赖管理员登录态直接读取邮件，项目新增了项目级访问 Key。

- 服务端配置：`INBOXOPS_API_KEY`
- 推荐传递方式：请求头 `X-InboxOps-Key`
- 兼容请求体字段：`api_key` 或 `key`
- 通过邮箱账号定位档案，不暴露 `client_id`、`refresh_token`、`proxy`
- 当服务端未配置 `INBOXOPS_API_KEY` 时，接口返回 `503`
- 当 Key 缺失或错误时，接口返回 `401`

### 1. 获取邮件列表

接口：`POST /api/key/mailbox/messages`

请求头示例：

```http
X-InboxOps-Key: replace-this-project-key
Content-Type: application/json
```

请求体示例：

```json
{
  "email": "demo@example.com",
  "method": "graph_api",
  "top": 10,
  "unread_only": false,
  "keyword": "invoice"
}
```

字段说明：

- `email`：必填，邮箱账号
- `method`：可选，`graph_api`、`imap_new`、`imap_old`
- `top`：可选，返回数量，范围 `1-50`
- `unread_only`：可选，是否只返回未读邮件
- `keyword`：可选，按主题 / 发件人 / 摘要关键字过滤

返回内容包含：

- `mailbox`：邮箱公开信息
- `method`：本次使用的读取方式
- `messages`：邮件摘要列表
- `count`：本次返回条数

`curl` 示例：

```bash
curl -X POST http://127.0.0.1:5000/api/key/mailbox/messages \
  -H "Content-Type: application/json" \
  -H "X-InboxOps-Key: replace-this-project-key" \
  -d "{\"email\":\"demo@example.com\",\"top\":5}"
```

### 2. 获取邮件详情

接口：`POST /api/key/mailbox/message`

请求体示例：

```json
{
  "email": "demo@example.com",
  "message_id": "AAMkAG...",
  "method": "graph_api"
}
```

字段说明：

- `email`：必填，邮箱账号
- `message_id`：必填，邮件唯一标识
- `method`：可选，不传时使用该邮箱的默认读取方式

返回内容包含：

- `mailbox`：邮箱公开信息
- `message`：完整邮件详情，包括正文、收件人、抄送、附件状态、会话信息等

## 注意事项

- `refresh_token` 和 `client_id` 需要具备对应 IMAP 或 Graph API 权限
- IMAP 连接本身仍是直连，当前代理主要覆盖 token 和 Graph 的 HTTP 请求
- 当前多邮箱后台适合内部管理使用；若用于生产环境，建议再补密码哈希、CSRF、防爆破和敏感信息加密存储
