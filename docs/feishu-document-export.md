# 飞书文档导出（第一期）

入口：生成文件右侧下拉 → 导出为飞书文档。

- Word / TXT 在线模板 / layout / JSON 在线模板；多条记录合并为一份。
- 用户首次导出时在新窗口完成 OAuth。普通 PDF、Word、Excel 下载不需要授权。
- 使用 OAuth 用户令牌导入个人云盘根目录，不需要 folder_token，不回退到应用身份。
- 默认标题沿用文件名规则；合并时使用第一条记录的规则并加记录数量。
- 飞书原生文档的分页、字体等可能与 Word 不完全一致，以导入结果为准。
- 当记录造成页眉/页脚等共享包内容不同，合并会报错，请单条导出，避免丢失内容。
- Excel/PDF 模板本期不转换为飞书文档，原有导出保持不变。

## 服务端配置（不能放入 Git）

在 `feishu-layout-print.service` 的 systemd drop-in 配置：

```ini
[Service]
Environment="FEIYE_FEISHU_APP_ID=<应用ID>"
Environment="FEIYE_FEISHU_APP_SECRET=<应用密钥>"
Environment="FEIYE_FEISHU_OAUTH_REDIRECT_URI=https://gzwy.online/feishu/api/oauth/feishu/callback"
```

开放平台登记完全一致的重定向地址。申请并发布用户授权范围：

```text
offline_access docs:document:import docs:document.media:upload
```

如开放平台为当前应用提供不同的等效权限组合，可设置 `FEIYE_FEISHU_OAUTH_SCOPES`（空格分隔），但必须保留 `offline_access` 才能取得刷新令牌。以应用控制台实际可用权限为准。运行环境 Node.js 18.17+。

## 凭据与数据边界

`data/oauth/key` 为随机 32 字节本机加密密钥，`tokens.enc` 为 AES-256-GCM 加密的用户令牌和会话。两者均以 0600 权限创建，已加入 .gitignore，不能打入部署包。备份需同时保留这两项，并限制访问。

访问令牌、刷新令牌不返回浏览器。浏览器仅在 sessionStorage 存放随机应用会话，不凭用户 ID 直接获取授权。关闭会话或换设备后，可能需要重新通过飞书认证；服务器保存令牌不等于新设备可以跳过身份验证。OAuth 返回的实际用户是文档归属人。

按接口返回的 `expires_in`、`refresh_token_expires_in` 保存有效期，不写死“30天”。每次刷新保存新令牌，单进程按用户合并刷新请求；多进程部署需改为共享事务存储和分布式锁。

授权 state 与浏览器 HttpOnly cookie 绑定，领取结果还需私有 challenge/verifier。state 十分钟有效，一次使用。账户绑定以 OAuth user_info 的 app / tenant / open_id 为准，不信任插件上报的 ID。

点击结果页的“解除本插件授权”会删除服务端保存的令牌、使所有本插件会话失效；不删除云文档，也不代表撤销飞书平台侧授权，平台授权可在飞书中管理。

## 导出任务

导出接口返回任务 ID，前端每两秒查询，不保持长时间 HTTP 连接。相同 requestId 不重复创建，任务只允许所属 OAuth 用户查询。重启后已提交飞书的导入任务可继续查询；上传前中断会明确报错。超时不能确认是否已导入成功时，应先检查个人云盘。

`data/oauth/jobs` 只保存所属用户、标题、任务状态和结果链接，不保存业务记录内容与访问令牌。飞书 `ccm_import_open` 临时素材导入后由飞书清理。本地无需产生临时 DOCX 文件。

## 验证

`npm run build`、`npm test` 包含令牌加密、刷新并发、state 防重放、默认根目录导入、重复提交、跨用户/跨租户访问拦截以及完整 mock HTTP 流程。发布后仍需真实用户授权并导出一次确认应用权限及文档布局；自动测试不等价于线上 OAuth 实测。

官方依据：[用户授权](https://open.feishu.cn/document/common-capabilities/sso/web-application-end-user-consent/guide)、[导入素材](https://open.feishu.cn/document/server-docs/docs/drive-v1/import_task/import-user-guide)、[创建导入任务](https://open.feishu.cn/document/server-docs/docs/drive-v1/import_task/create)、[查询导入结果](https://open.feishu.cn/document/server-docs/docs/drive-v1/import_task/get)。
