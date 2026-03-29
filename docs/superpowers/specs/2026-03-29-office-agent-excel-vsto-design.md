# OfficeAgent Excel VSTO 设计说明

日期：2026-03-29

状态：架构改版草案，待评审

## 1. 目标

将当前 `Office Add-in + Office.js` 路线切换为 `VSTO Excel Add-in` 路线，满足以下核心目标：

- 仅支持 `Windows 桌面版 Excel 2019 及之后版本`
- 以 `Excel 插件` 形态承载 AI 聊天任务窗格
- 能稳定获取当前选中 `Sheet / Range / 行列 / 表头 / 样例值`
- 能执行 Excel 深度操作，包括读写单元格、增删改 Sheet、结构修改
- 支持外部 API 调用
- 支持 `upload_data` 等可扩展 skill
- 会话历史按插件级、本地机器级持久化
- 优先满足企业内网分发与安装简化

## 2. 方向结论

推荐技术路线：

`VSTO Excel Add-in + Ribbon + CustomTaskPane + WPF Chat UI + .NET Framework 4.8 + MSI 分发`

原因：

- 更贴合你现在的单平台目标：只做 Windows Excel
- 分发方式更像传统桌面插件，适合内网统一安装
- Excel 自动化能力更强，不再受 `Office.js` 能力边界限制
- 不再依赖远端 HTTPS 托管页面、manifest 侧载、证书信任链等 Office Add-in 额外复杂度

## 3. 方案对比

### 方案 A：VSTO + 原生 WPF 聊天窗格

说明：
- VSTO 负责 Excel 生命周期、Ribbon、自定义任务窗格
- 任务窗格 UI 采用 `WPF`
- 通过 `Microsoft.Office.Interop.Excel` 直接操作 Excel

优点：
- 原生桌面体验
- 不依赖浏览器容器
- 对 Excel 交互最直接
- 打包部署路径最清晰

缺点：
- 需要重写当前 Web UI
- 前端迭代效率低于 Web 技术栈

结论：
- `推荐`

### 方案 B：VSTO + WebView2 承载现有 React UI

说明：
- VSTO 仍负责 Ribbon、TaskPane、Excel 自动化
- 聊天界面继续用现有 Web 前端，通过 `WebView2` 承载
- 通过 JS/.NET 桥接访问本地 Excel 能力

优点：
- 可复用当前 Office Add-in MVP 的部分 UI/交互逻辑
- 聊天界面开发效率更高

缺点：
- 需要额外处理 JS/.NET 通讯
- 部署时要处理 `WebView2 Runtime`
- 技术栈混合度更高

结论：
- 适合作为“迁移成本优先”备选

### 方案 C：VSTO + 纯 WinForms 窗格

说明：
- 全部使用 WinForms 实现任务窗格 UI

优点：
- 与 VSTO 集成最直接
- 工程最简单

缺点：
- 聊天界面体验和可维护性明显较差

结论：
- 不推荐作为长期方案

## 4. 推荐架构

### 4.1 宿主结构

```text
Excel
  -> VSTO Add-in Host
      -> Ribbon Controller
      -> Custom Task Pane Host
      -> Excel Event Bridge
      -> Application Services
          -> Agent Orchestrator
          -> Skill Registry
          -> Confirmation Service
          -> Session Service
          -> Settings Service
      -> Infrastructure
          -> Excel Interop Adapter
          -> HTTP API Clients
          -> Local File Storage
          -> Secret Protection
      -> WPF Chat UI
```

### 4.2 项目拆分

推荐按 4 个项目组织：

- `src/OfficeAgent.ExcelAddIn`
  说明：VSTO 宿主项目，包含 `ThisAddIn`、Ribbon、TaskPane 注册、Excel 事件绑定
- `src/OfficeAgent.DesktopUI`
  说明：WPF 视图、ViewModel、命令绑定、确认卡片、会话列表、设置面板
- `src/OfficeAgent.Core`
  说明：领域模型、Agent 编排、skill 路由、确认策略、会话与设置服务接口
- `src/OfficeAgent.Infrastructure`
  说明：Excel Interop 适配器、HTTP 客户端、本地持久化、凭据保护实现

## 5. UI 形态

### 5.1 主入口

通过 Ribbon 上的 `OfficeAgent` 按钮显示或隐藏任务窗格。

Microsoft 官方文档明确指出，自定义任务窗格没有默认显示 UI，因此插件应提供一个按钮让用户显式打开或关闭任务窗格。这个按钮就放在 Ribbon 上。来源：
- [CustomTaskPane.Visible Property](https://learn.microsoft.com/en-us/dotnet/api/microsoft.office.tools.customtaskpane.visible?view=vsto-2022)

### 5.2 任务窗格布局

任务窗格布局保持你已经确认过的产品形态：

- 左侧会话列表
- 右侧聊天主区
- 顶部标题栏和设置入口
- 中部消息线程
- 底部输入框
- 输入框上方或下方显示当前选区摘要
- 写操作和外部提交显示确认卡片

### 5.3 设置项

首版设置面板保留：

- `API Key`
- `Base URL`
- `Model`

正式版再演进到 OAuth。

## 6. Excel 交互架构

### 6.1 Excel 事件来源

VSTO 宿主监听 Excel 应用程序事件，至少包括：

- `SheetSelectionChange`
- `WorkbookOpen`
- `WorkbookBeforeClose`
- 必要时的 `SheetActivate`

### 6.2 选区上下文服务

统一产出 `SelectionContext`：

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

### 6.3 Excel 操作适配器

所有 Excel 读写通过 `ExcelInteropAdapter` 统一暴露，避免 UI 或 skill 直接接触 Interop 对象。

建议对外提供的命令族：

- `ReadSelectionTable`
- `ReadRange`
- `WriteRange`
- `AddWorksheet`
- `DeleteWorksheet`
- `RenameWorksheet`
- `InsertRows`
- `DeleteRows`
- `AutoFit`

## 7. Agent 与 Skill 架构

### 7.1 Agent Orchestrator

保留你当前已确认的产品规则：

- 自然语言优先
- slash 命令作为强制 skill 入口
- 读操作可直接执行
- 写操作和会改动外部系统状态的动作必须确认

Agent 输出不直接生成任意文本脚本，而是生成结构化命令：

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

### 7.2 Skill Registry

首版至少支持：

- `upload_data`

保留扩展位：

- `format_selection`
- `generate_sheet`
- `sync_to_system`

### 7.3 upload_data Skill 流程

```text
用户输入
  -> 识别 upload_data skill
  -> 读取当前选区
  -> 根据首行/首列推断字段
  -> 生成上传 payload 预览
  -> 弹出确认卡片
  -> 用户确认
  -> 调用 upload_data_api
  -> 在聊天线程显示结果
```

## 8. 本地存储与安全

### 8.1 会话持久化

首版不再使用浏览器 `localStorage`，改用本地文件存储：

- 路径建议：`%LocalAppData%\OfficeAgent\`
- `sessions\index.json`
- `sessions\<sessionId>.json`
- `settings.json`

### 8.2 敏感信息存储

`API Key` 不建议明文写入 `settings.json`。

推荐：

- 会话与普通设置：JSON 文件
- API Key：Windows `DPAPI` 用户级加密

这样在 demo 阶段也比浏览器本地存储更安全。

## 9. 部署与分发

VSTO 的正式分发策略建议直接走：

`MSI 安装包`

原因：

- 可按机器安装
- 更适合企业软件中心、组策略、SCCM、Intune、内网分发
- 不需要每个用户手工侧载或信任目录

Microsoft 官方文档明确指出：

- VSTO 可通过 `ClickOnce` 或 `Windows Installer`
- `Windows Installer` 支持为一台机器上的所有用户安装

来源：
- [Deploy an Office solution](https://learn.microsoft.com/en-us/visualstudio/vsto/deploying-an-office-solution?view=vs-2022)
- [Deploy an Office solution by using ClickOnce](https://learn.microsoft.com/en-us/visualstudio/vsto/deploying-an-office-solution-by-using-clickonce?view=visualstudio)

部署建议：

- 开发调试：Visual Studio F5
- UAT：ClickOnce 可选
- 正式发布：MSI

## 10. 兼容性与环境约束

### 10.1 支持范围

- Windows 桌面版 Excel 2019+
- 32/64 位 Office 都要验证
- 不支持 Mac
- 不支持 Excel 网页版

### 10.2 开发环境

VSTO 开发机需具备：

- Visual Studio Office 开发工具
- .NET Framework 4.x
- Excel

来源：
- [Configure a computer to develop Office solutions](https://learn.microsoft.com/en-us/visualstudio/vsto/how-to-configure-a-computer-to-develop-office-solutions?view=visualstudio)

### 10.3 技术边界

VSTO UI 与文档交互代码不能直接复用到 Office Add-in。

Microsoft 的迁移教程明确把 VSTO 代码分成三类：

- UI code
- Document code
- Logic code

其中能共享的主要是“逻辑代码”，UI 和文档交互代码不能直接平移到 Office Add-in，也反过来说明当前 React/Office.js MVP 不能原样迁到 VSTO。来源：
- [Share code between both a VSTO Add-in and an Office Add-in](https://learn.microsoft.com/en-us/office/dev/add-ins/tutorials/migrate-vsto-to-office-add-in-shared-code-library-tutorial)

## 11. 测试策略

### 11.1 单元测试

覆盖：

- skill 路由
- 命令解析
- payload 组装
- 会话持久化
- 设置与 Base URL 解析
- 确认策略

### 11.2 集成测试

覆盖：

- Excel Interop 适配器
- 选区上下文服务
- HTTP 客户端
- 任务窗格与 ViewModel 交互

### 11.3 手工验证

至少覆盖：

- Excel 2019 32 位
- Excel 2019 64 位
- 一个更新版本的 Excel 桌面端
- 内网 MSI 安装
- 首次打开任务窗格
- 选区实时刷新
- upload_data 预览/确认/提交

## 12. 分阶段实施建议

### Phase 1：VSTO 外壳

- 建立 Excel VSTO 项目
- 加 Ribbon 按钮
- 打通 CustomTaskPane
- 放入空白聊天 UI

### Phase 2：核心桌面能力

- 会话持久化
- 设置面板
- API Key / Base URL
- Excel 选区监听

### Phase 3：Agent 与 Skill

- 命令模型
- upload_data skill
- 写操作确认流
- HTTP 客户端

### Phase 4：分发与交付

- 安装包
- 签名
- UAT 验收
- 内网部署文档

## 13. 结论

如果目标明确是：

- 企业内网
- Windows Excel
- 安装分发简单
- Excel 操作能力强

那么 `VSTO + MSI` 比继续推进纯 Office Add-in 更契合你的实际需求。

在这条路线上，我建议首版采用：

- `VSTO Excel Add-in`
- `Ribbon + CustomTaskPane`
- `WPF Chat UI`
- `Excel Interop Adapter`
- `本地 JSON + DPAPI`
- `MSI 分发`

