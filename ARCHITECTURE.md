# 飞页技术架构与程序逻辑

## 1. 边界和目标

飞页是一个独立的飞书多维表侧边插件。它和工作区中的“印序”是两个项目：目录、端口、依赖、模板存储和接口均不共享。首版链路为：当前多维表选中记录 → 绑定模板 → DOCX/XLSX/XLS/PDF 或飞书在线布局输出。

## 2. 总体架构

```text
飞书多维表侧边栏 / 浏览器
    │  官方 Base JS SDK（只读当前上下文）
    │  POST 选中记录 JSON
    ▼
Node.js server.mjs :4318
    ├─ 静态白名单：public/、官方 SDK dist/
    ├─ 模板 API：上传、列表、删除
    ├─ 渲染 API：模板 + records → 原格式文件或 ZIP
    └─ 下载 API：随机 output id → 文件流
          │
          ├─ data/templates/*.docx
          ├─ data/templates.json
          └─ data/outputs/*.docx
```

插件规范约束：发布到飞书“扩展脚本”时使用可公网访问的 HTTPS 页面 URL（开发时可用本地映射/隧道），直接指向 `/feishu`；该路径不做登录跳转。页面以官方 SDK `@lark-base-open/js-sdk@0.3.8` 的 `bitable` 为唯一宿主入口，优先调用 `base.getSelection()` + `base.getTableById(tableId)`，无 tableId 时才调用 `base.getActiveTable()`，并通过 `base.onSelectionChange()` 和 `ui.showToast()` 提供宿主内反馈。SDK 由本服务同源提供，避免在飞书宿主中依赖外部 CDN。

前端不发送飞书 App Secret，也不把 `PersonalBaseToken` 放入服务端。字段读取发生在飞书宿主环境，服务端只接收用户点击生成时提交的已选记录值。

## 3. 程序逻辑

1. 页面加载后并行请求 `/api/templates`，并尝试动态加载 `/vendor/lark-base/index.mjs`。
   左侧菜单默认收起；展开后只渲染 `/api/templates` 返回的真实模板，不写死模板名称，也没有登录用户面板。
   模板设置按钮只操作本地说明面板；调试面板把当前模板字段渲染为字段依赖，不改变模板或数据。
2. SDK 获取 `bitable.base.getSelection()`；有 tableId 时使用 `base.getTableById(tableId)`，单条读取始终优先使用当前 `selection.recordId`，只有宿主没有返回 recordId 时才回退到视图选择列表或活动视图第一页首条记录；批量模式独立读取当前视图记录。
3. 飞书 1.0.2 的分页结果是 `IRecordList` 记录代理，不能直接读取 `record.fields[fieldId]`。页面按每个选中 `recordId` 调用字段的 `getCellString(recordId)`，组装出稳定的 `{字段名:值}` 对象。
   同时用 `getValue` 检查关联字段的 `recordIds/tableId`，按模板循环根读取关联表字段；宿主触发 `base.onSelectionChange` 时自动重跑该步骤。若宿主返回 `not registered`，事件注册异常会被捕获，700ms 轮询 `base.getSelection()` 作为降级方案。
4. 上传 DOCX/XLSX/XLS/PDF 或飞书导出的 TXT/LAYOUT/JSON 在线模板时，浏览器以原始二进制 POST `/api/templates`，文件名和当前 Base/Table/视图上下文放在编码请求头；服务端检查格式和大小，提取 Office 或布局变量元数据后写入独立目录。模板查询按 Base/Table 过滤，切换到另一张表不会看到上一张表的模板。
5. 点击生成时，前端 POST `{ templateId, records:[{字段名:值}] }` 到 `/api/generate-docx`。
6. 服务端用 `docxtemplater + pizzip` 渲染 DOCX；XLSX 在 OpenXML 工作表中替换普通变量；在线布局根据导出的 document/pages/rows/columns/blocks 生成带表格样式的 HTML；XLS/PDF 首版保留原始格式直接导出。单记录输出原格式，多记录逐条渲染后写入 ZIP。
7. 输出以随机 UUID 文件名写入 `data/outputs`，响应输出 ID；浏览器通过 `/api/outputs/:id/download` 下载 DOCX/XLSX/XLS/PDF 或 ZIP。

模板预览通过 `GET /api/templates/:id/preview` 返回字段、结构化布局/可见文本和原文件 URL；`GET /api/templates/:id/file` 以 `inline` 文件流返回。PDF 使用 iframe 原生预览，DOCX/XLSX 使用服务端提取的内容和变量清单，在线布局在浏览器内重建页面和表格，XLS 提供原始文件流，所有路径均由模板 ID 解析，不能访问任意本地路径。

在线模板编辑器是独立的基础编辑层：组件按钮使用 HTML5 `draggable/drop` 和 `contenteditable` 画布插入实际块；数据源面板使用 `[字段名]` 作为排版变量，复制按钮写入剪贴板，系统 Tab 提供 `[表格名]` 和 `[打印时间]`；文本格式、清空、撤销/重做、页面尺寸和字体设置均绑定真实 DOM 操作；“保存模板”将标题和画布 HTML 按当前 `tableId` 保存到浏览器本地草稿。

## 4. 安全和稳定性

- 文件名只保留字母、数字、中文、空格、点、下划线和连字符；文件大小上限 20 MB。
- 所有 JSON API `cache-control: no-store`；静态路由只允许 `public/` 与 SDK `dist/` 内的路径。
- 资源 ID 使用 UUID，服务端不会把用户传入的路径拼接为文件路径。
- 记录值只在请求内存中处理，不写入 JSON 或日志。
- 首版没有服务端飞书 OAuth 和 App Secret，适合个人侧边栏使用；正式多人部署前应增加登录、租户边界、文件清理和限流。

## 5. 演进路线

| 阶段 | 增量 |
| --- | --- |
| MVP | DOCX/XLSX 基础变量、XLS/PDF 原格式导出、选中记录自动读取 |
| V1 | 多选记录批量 ZIP、模板变量预览、字段 ID 映射 |
| V2 | XLSX 复杂循环、PDF 版式填充、打印队列 |
| V3 | 企业 OAuth、模板共享权限、审计和对象存储 |
