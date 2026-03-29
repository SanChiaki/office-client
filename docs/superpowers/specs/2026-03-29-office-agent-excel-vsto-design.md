# OfficeAgent Excel VSTO 设计说明

日期：2026-03-29

状态：已切换为 VSTO 主线方案，待进入实施

## 1. 目标

将当前 `Office Add-in + Office.js` 技术方向切换为 `VSTO Excel Add-in`，并保留已经确认过的产品边界：

- 仅支持 `Windows 桌面版 Excel 2019 及之后版本`
- 主界面仍然是 AI 对话形态
- 需要实时展示当前选中的 `Sheet / 地址 / 行列 / 样例值`
- 需要支持 Excel 深度操作，不再受 Office.js 边界限制
- 需要支持外部 API 调用
- 需要支持 `upload_data` 这类可扩展 skill
- 会话历史按插件级、本地机器级持久化
- 优先满足内网分发和企业安装简化

## 2. 主方案结论

主方案改为：

`VSTO Excel Add-in + Ribbon + CustomTaskPane + WebView2 + React/TypeScript + .NET Framework 4.8 + MSI 分发`

这条路线的核心取舍是：

- 宿主、Excel 自动化、安装分发，交给 `VSTO`
- 聊天 UI、会话列表、设置面板、确认卡片，继续使用 `Web 技术栈`
- Excel 和本地能力，不再直接从前端调用，而是通过 `WebView2 JS/.NET Bridge` 调宿主服务

## 3. 方案对比

### 方案 A：VSTO + WebView2 + React

说明：

- VSTO 负责 Excel 宿主、Ribbon、自定义任务窗格、Excel Interop、本地存储、安装分发
- WebView2 负责承载聊天前端
- React/TypeScript 负责 UI 和前端交互
- 前端通过消息桥调用本地 .NET 服务

优点：

- 最适合聊天类产品界面
- 能复用当前 MVP 的前端交互思路
- UI 开发速度和可维护性明显优于纯 WPF
- 仍然保留 VSTO 的分发和 Excel 深度操作优势

缺点：

- 需要设计一层 Web 与 .NET 的桥接协议
- 需要处理 WebView2 Runtime 的检测与分发
- 前端不能继续直接依赖 Office.js

结论：

- `推荐，作为主方案`

### 方案 B：VSTO + 纯 WPF

说明：

- 任务窗格 UI 完全使用 WPF

优点：

- 宿主集成最直接
- 不需要浏览器容器和消息桥

缺点：

- 对聊天产品 UI 来说开发效率偏低
- 会话列表、富文本消息、复杂状态交互的迭代成本更高
- 难以复用当前 Web 前端经验

结论：

- `作为备选，不作为主路线`

### 方案 C：继续 Office Add-in

说明：

- 保持 Office Add-in 路线，继续走 manifest、任务窗格网页、Office.js

缺点：

- 分发链路复杂
- 仍受 Office.js 能力边界约束
- 与你当前“内网分发简单、Excel 操作更强”的目标冲突

结论：

- `不再作为主路线`

## 4. 官方边界

这条方案的几个关键前提来自 Microsoft 官方文档：

- WebView2 可以嵌入 `WinForms` 宿主，支持 `.NET Framework` 桌面项目：[Get started with WebView2 in WinForms apps](https://learn.microsoft.com/en-us/microsoft-edge/webview2/get-started/winforms)
- Web 与宿主之间可以通过 `window.chrome.webview.postMessage` 和宿主侧 `PostWebMessageAsJson / WebMessageReceived` 通讯：[Interop of native and web code](https://learn.microsoft.com/en-us/microsoft-edge/webview2/how-to/communicate-btwn-web-native)
- WebView2 本地静态资源可以通过 `virtual host name mapping` 承载，这样前端页面拥有 `http/https` origin，支持 `localStorage`、`indexedDB` 和安全上下文 API：[Using local content in WebView2 apps](https://learn.microsoft.com/en-us/microsoft-edge/webview2/concepts/working-with-local-content)
- WebView2 Runtime 支持在线 bootstrapper 和离线 standalone installer，适合企业安装包集成：[Distribute your app and the WebView2 Runtime](https://learn.microsoft.com/en-us/microsoft-edge/webview2/concepts/distribution)
- VSTO 的正式发布支持 `ClickOnce` 或 `Windows Installer`，其中 Windows Installer 更适合内网机器级安装：[Deploy an Office solution](https://learn.microsoft.com/en-us/visualstudio/vsto/deploying-an-office-solution?view=vs-2022)
- 自定义任务窗格没有默认显示入口，因此需要自行提供 Ribbon 按钮：[CustomTaskPane.Visible Property](https://learn.microsoft.com/en-us/dotnet/api/microsoft.office.tools.customtaskpane.visible?view=vsto-2022)

## 5. 总体架构

```text
Excel
  -> VSTO Add-in Host
      -> Ribbon Controller
      -> TaskPane Controller
      -> WinForms TaskPaneHostControl
          -> WebView2
              -> React Frontend
      -> Native Application Services
          -> Excel Event Bridge
          -> Excel Interop Adapter
          -> Agent Orchestrator
          -> Skill Registry
          -> Confirmation Service
          -> Session Store
          -> Settings Store
          -> Secret Protector
          -> HTTP API Clients
```

宿主层和前端层的职责边界：

- `VSTO 宿主`
  负责 Excel 生命周期、Ribbon、任务窗格、Interop、本地文件、密钥保护、安装分发
- `React 前端`
  负责聊天界面、会话列表、消息渲染、设置面板、确认卡片、选区展示
- `Bridge`
  负责前后端之间的结构化消息交互

## 6. 项目拆分

推荐拆成 5 个项目：

- `src/OfficeAgent.ExcelAddIn`
  说明：VSTO 宿主项目，包含 `ThisAddIn`、Ribbon、TaskPane 控制器、Excel 事件注册
- `src/OfficeAgent.Core`
  说明：领域模型、命令模型、Agent 编排、skill 路由、确认策略接口
- `src/OfficeAgent.Infrastructure`
  说明：Excel Interop、HTTP 客户端、本地文件存储、DPAPI、运行日志
- `src/OfficeAgent.Frontend`
  说明：React/Vite 前端工程，产出任务窗格静态资源
- `installer/OfficeAgent.Setup`
  说明：MSI 安装包工程，负责部署 VSTO add-in、前端资源和 WebView2 Runtime

## 7. UI 形态

### 7.1 入口

通过 Excel Ribbon 上的 `OfficeAgent` 按钮显示或隐藏任务窗格。

### 7.2 任务窗格结构

前端界面继续沿用你已经确认过的产品形态：

- 左侧：会话列表
- 右侧：聊天主区
- 顶部：标题与设置入口
- 中部：消息线程
- 底部：输入框
- 输入框附近：当前选区摘要
- 中部卡片：写操作和外部提交的确认卡片

### 7.3 设置项

首版设置面板继续保留：

- `API Key`
- `Base URL`
- `Model`

后续正式版再切到 OAuth。

## 8. WebView2 承载策略

### 8.1 前端资源来源

不走远端 HTTPS 页面托管。前端构建产物直接随安装包落地到本机安装目录。

推荐目录：

- `%ProgramFiles%\OfficeAgent\app\`
  或
- `%LocalAppData%\OfficeAgent\app\`

### 8.2 本地页面加载方式

推荐用 `virtual host name mapping`，例如：

- `https://appassets.officeagent.local/index.html`

对应映射到本地静态目录。

这样比直接 `file:///` 更合适，因为：

- 有明确 origin
- 支持相对资源路径
- 支持安全上下文相关 Web API
- 更贴近正常 Web 前端运行环境

### 8.3 前端状态

前端自己的临时 UI 状态可以保存在内存。

业务级持久化不要依赖浏览器存储作为唯一数据源，应由宿主统一落盘，再回填给前端。

## 9. JS/.NET Bridge 设计

### 9.1 通讯方式

Web 到宿主：

- `window.chrome.webview.postMessage(...)`

宿主到 Web：

- `CoreWebView2.PostWebMessageAsJson(...)`

### 9.2 消息协议

建议使用结构化 envelope，不走自由文本协议。

示例：

```json
{
  "type": "excel.readSelection",
  "requestId": "req-001",
  "payload": {}
}
```

返回：

```json
{
  "type": "excel.readSelection.result",
  "requestId": "req-001",
  "ok": true,
  "payload": {
    "sheetName": "Sheet1",
    "address": "A1:D5"
  }
}
```

### 9.3 Bridge 原则

- 只允许白名单消息类型
- 每条消息必须有 `requestId`
- 宿主层统一做参数校验
- 错误必须结构化返回给前端
- 不把 Excel COM 对象、路径句柄等原生对象直接暴露给前端

## 10. Excel 交互架构

### 10.1 事件桥

宿主监听 Excel 事件，至少包括：

- `SheetSelectionChange`
- `WorkbookOpen`
- `WorkbookBeforeClose`
- 必要时的 `SheetActivate`

### 10.2 选区上下文服务

统一输出 `SelectionContext`：

```json
{
  "workbookName": "Budget.xlsx",
  "sheetName": "Sheet1",
  "address": "A1:D20",
  "rowCount": 20,
  "columnCount": 4,
  "isContiguous": true,
  "headerPreview": ["Name", "Owner", "StartDate", "Budget"],
  "sampleValues": [["Project A", "张三"]],
  "capturedAt": "2026-03-29T14:00:00+08:00"
}
```

### 10.3 Excel 命令适配器

所有 Excel 操作统一通过 `ExcelInteropAdapter` 暴露。

建议命令族：

- `ReadSelectionTable`
- `ReadRange`
- `WriteRange`
- `AddWorksheet`
- `DeleteWorksheet`
- `RenameWorksheet`
- `InsertRows`
- `DeleteRows`
- `AutoFit`

## 11. Agent 与 Skill 架构

继续保留已经确认的业务规则：

- 自然语言优先
- slash 命令作为强制 skill 入口
- 读操作直接执行
- 写操作必须确认
- 会改动外部系统状态的 API 提交也必须确认

Agent 输出统一为结构化命令，而不是任意脚本。

示例：

```json
{
  "assistantMessage": "我先读取当前选区并生成上传预览。",
  "mode": "skill",
  "skillName": "upload_data",
  "requiresConfirmation": true,
  "actions": [
    {
      "type": "excel.readSelectionTable",
      "args": {}
    },
    {
      "type": "skill.upload_data.preview",
      "args": {
        "project": "项目A"
      }
    }
  ]
}
```

### upload_data Skill

流程不变：

```text
用户输入
  -> 识别 upload_data skill
  -> 读取当前选区
  -> 根据首行/首列推断字段
  -> 生成上传 payload 预览
  -> 显示确认卡片
  -> 用户确认
  -> 调用 upload_data_api
  -> 在聊天线程显示结果
```

## 12. 本地存储与安全

### 12.1 会话与设置

首版使用本地文件存储：

- `%LocalAppData%\OfficeAgent\sessions\index.json`
- `%LocalAppData%\OfficeAgent\sessions\<sessionId>.json`
- `%LocalAppData%\OfficeAgent\settings.json`

### 12.2 敏感信息

`API Key` 不明文存入 settings JSON。

推荐：

- 普通设置和会话：JSON
- API Key：Windows `DPAPI` 用户级加密

### 12.3 安全边界

- Bridge 消息白名单
- 不允许前端直接执行任意本地命令
- 不默认把整块大选区发送给外部 API
- 写 Excel 和外部提交统一确认

## 13. 部署与分发

正式分发建议：

`MSI + WebView2 Runtime 检测/安装`

推荐策略：

- 开发调试：Visual Studio F5 + 前端本地 dev server 可选
- UAT：MSI 测试包
- 正式发布：MSI

WebView2 Runtime 策略：

- 内网/离线环境优先：随安装包集成 `Evergreen Standalone Installer`
- 在线环境可选：Bootstrapper

因为 Microsoft 官方文档明确支持：

- 在线 bootstrapper 安装
- 离线 standalone installer 安装
- per-machine 和 per-user 两种安装模式

来源：
- [Distribute your app and the WebView2 Runtime](https://learn.microsoft.com/en-us/microsoft-edge/webview2/concepts/distribution)

## 14. 兼容性约束

- Windows 桌面版 Excel 2019+
- 需要验证 Office x86/x64
- 不支持 Mac
- 不支持 Excel Web
- VSTO 主体仍应基于 `.NET Framework 4.8`

## 15. 测试策略

### 15.1 单元测试

覆盖：

- skill 路由
- 命令 envelope 解析
- payload 组装
- Base URL 解析
- 会话持久化
- 确认策略
- Bridge 消息校验

### 15.2 集成测试

覆盖：

- WebView2 bridge handler
- Excel Interop 适配器
- 本地存储
- HTTP 客户端

### 15.3 手工验证

至少覆盖：

- Excel 2019 32 位
- Excel 2019 64 位
- 一个更高版本 Excel 桌面端
- 安装包安装/升级/卸载
- 首次打开任务窗格
- 选区实时刷新
- upload_data 预览/确认/提交

## 16. 结论

对于你当前的真实目标：

- 企业内网
- Windows Excel
- 安装分发简单
- 聊天 UI 开发效率高
- Excel 操作能力强

最合适的路线不是纯 WPF，也不是继续 Office Add-in，而是：

- `VSTO Excel Add-in`
- `Ribbon + CustomTaskPane`
- `WinForms Host + WebView2`
- `React/TypeScript UI`
- `Excel Interop Adapter`
- `本地 JSON + DPAPI`
- `MSI 分发`

