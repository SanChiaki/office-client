# OfficeAgent Excel Add-in 设计说明

日期：2026-03-29

状态：已完成设计评审，待进入实现计划

## 1. 项目目标

开发一个以 Excel 任务窗格为承载形态的 OfficeAgent。首期只支持 Windows 桌面版 Excel 2019 及之后版本，核心体验是“在 Excel 内通过 AI 对话完成数据读取、写入、Sheet 操作、外部 API 调用和可复用 skill 执行”。

## 2. 本期范围

本期必须实现：

- Excel 任务窗格聊天主界面
- 上方展示用户与 Agent 对话，下方输入框发送消息
- 实时感知当前选区，并在输入框下方展示选中的 Sheet、地址、行数、列数等信息
- Agent 可根据用户输入执行 Excel 读操作、写操作、Sheet 增删改等动作
- 支持调用外部 API
- 支持 skill 封装，并以 `upload_data` 作为样例 skill
- 聊天记录按会话粒度持久化，支持新建对话

本期明确不做：

- Mac、Excel 网页版、移动端兼容
- 纯本地离线运行
- 用户侧本地 companion 程序
- 正式版 OAuth 账号体系
- 云端会话同步

## 3. 已确认约束

### 3.1 产品和部署约束

- 形态为纯 Office Add-in，不给用户安装额外本地程序
- 开发期使用 `https://localhost:<port>` 调试
- 发布期前端页面部署到正式 HTTPS 站点
- manifest 指向远端 HTTPS 页面，不依赖用户本机启动本地服务

### 3.2 客户端和兼容性约束

- 仅支持 Windows 桌面版 Excel 2019 及之后版本
- 首版能力边界以 Office.js 可覆盖范围为准
- 采用经典 add-in-only XML manifest
- 采用保守的 ExcelApi 基线，并在运行时做能力检查
- 前端页面保持单页任务窗格形态，不依赖浏览器路由能力

### 3.3 Agent 和执行约束

- Agent 为前端内置轻量编排，不引入本地 companion
- 读操作可直接执行
- 写操作一律先预览、再确认
- 自然语言优先，slash 命令作为强制 skill 入口
- demo 阶段通过用户填写 API Key 访问外部 API
- 正式版再演进到用户登录 OAuth

### 3.4 会话和存储约束

- 聊天记录需要持久化
- 按会话粒度保存，不绑定工作簿
- 页面支持新建对话、切换对话、删除对话
- 首版仅当前机器本地保存，不做云同步

## 4. 总体架构

整体采用“单一任务窗格 UI + 前端分层执行”的结构。

```text
Excel Task Pane UI
  -> Chat Shell
  -> Selection Context Service
  -> Agent Orchestrator
  -> Skill Registry
  -> Confirmation Guard
  -> Excel Adapter (Office.js)
  -> API Client (LLM API / Business API)
  -> Session Manager
  -> Storage Adapter
```

系统的产品形态是“纯 Office Add-in”，但内部代码边界按“聊天层 / skill 层 / Excel 执行层 / 存储层”拆分，保证 demo 先跑通，后续再逐步迁移到 OAuth 和更正式的远端架构。

## 5. 模块设计

### 5.1 Chat Shell

负责：

- 渲染用户和 Agent 消息
- 渲染写操作确认卡片
- 渲染 skill 预览卡片和执行结果
- 提供输入框、发送按钮、新建对话入口、会话列表
- 在输入框下方展示当前选区上下文

### 5.2 Selection Context Service

负责：

- 监听 Excel 选区变化
- 读取当前工作表名、选区地址、行数、列数
- 在需要时补充首行、首列、样例值
- 生成统一的 `SelectionContext` 对象，供 UI、Agent 和 skill 复用

建议字段：

```json
{
  "workbookName": "Budget.xlsx",
  "sheetName": "Sheet1",
  "address": "A1:D20",
  "rowCount": 20,
  "columnCount": 4,
  "hasHeaders": true,
  "headerPreview": ["Name", "Owner", "StartDate", "Budget"],
  "firstColumnPreview": ["项目A", "项目B"],
  "isMultiRange": false,
  "sampleValues": [["Name", "Owner"], ["项目A", "张三"]],
  "capturedAt": "2026-03-29T00:40:00+08:00"
}
```

### 5.3 Agent Orchestrator

负责：

- 汇总用户输入、当前会话历史、当前选区上下文
- 判断当前请求是普通对话、skill 调用还是 Excel 动作
- 调用 LLM API 获取结构化执行计划
- 解析 LLM 返回结果并分发到 Skill Registry、Excel Adapter 或普通聊天输出

Agent 不直接返回“任意可执行文本”，而是返回结构化命令包。

### 5.4 Skill Registry

负责：

- 注册所有可复用 skill
- 管理 skill 的元信息、入参规范和执行流程
- 让自然语言与 slash 命令都能映射到 skill

首个样例 skill 为 `upload_data`。

### 5.5 Confirmation Guard

负责：

- 对所有写操作统一做预览和确认
- 对所有会修改外部系统状态的 API 提交统一做确认
- 在确认前渲染影响范围、目标位置、payload 摘要、预期结果

### 5.6 Excel Adapter

负责：

- 封装 Office.js 对工作簿、工作表、单元格、区域的访问
- 屏蔽 UI 层对 Excel API 的直接依赖
- 对外提供稳定命令接口，例如 `readSelection`、`writeRange`、`addSheet`、`renameSheet`

### 5.7 API Client

负责：

- 调用 LLM API
- 调用业务 API，例如 `upload_data_api`
- 统一处理超时、鉴权失败、重试、错误格式化

### 5.8 Session Manager

负责：

- 新建对话、切换对话、删除对话
- 恢复最近一次活动会话
- 为每个会话维护消息流、标题、更新时间

### 5.9 Storage Adapter

负责：

- 本地保存会话索引
- 本地保存会话消息
- 本地保存设置项，例如 API Key
- 为后续云端存储演进保留统一接口

## 6. 交互和执行流

### 6.1 普通对话流

1. 用户输入自然语言
2. Agent Orchestrator 组合输入、会话上下文、选区摘要
3. 调用 LLM API
4. LLM 返回普通聊天消息
5. Chat Shell 展示结果并落盘到当前会话

### 6.2 直接 Excel 动作流

1. 用户输入自然语言，例如“新增一个 Sheet，名字叫汇总”
2. Agent Orchestrator 请求 LLM 输出结构化命令
3. 若命令为读操作，则 Excel Adapter 直接执行
4. 若命令为写操作，则 Confirmation Guard 展示确认卡片
5. 用户确认后由 Excel Adapter 执行
6. 结果回写到聊天流并持久化

### 6.3 Skill 路由流

支持三种路径：

- 普通自然语言被识别为 skill
- 显式 slash 命令强制进入 skill
- 用户先对话，再由 Agent 建议调用 skill

### 6.4 `upload_data` skill 流

1. 用户输入：
   - `把选中数据上传到项目A`
   - 或 `/upload_data 把选中数据上传到项目A`
2. Skill Registry 将请求路由到 `upload_data`
3. Excel Adapter 读取当前连续选区
4. skill 根据首行、首列推断字段和记录
5. 构造待上传 payload
6. UI 渲染预览卡片，展示字段、记录数、样例数据和目标项目
7. 用户确认后调用 `upload_data_api`
8. UI 展示保存结果、失败原因或部分成功摘要

## 7. 结构化命令协议

Agent 返回统一命令包，例如：

```json
{
  "assistant_message": "我将先读取选区并识别字段，再给你预览上传内容。",
  "mode": "chat | excel_action | skill",
  "skill_name": "upload_data",
  "requires_confirmation": true,
  "actions": [
    {
      "type": "excel.readRange",
      "args": {
        "source": "current_selection",
        "includeHeaders": true
      }
    },
    {
      "type": "skill.upload_data.preview",
      "args": {
        "project": "项目A",
        "mappingStrategy": "infer_from_first_row_and_first_column"
      }
    }
  ]
}
```

推荐命令族：

- `excel.read*`
- `excel.write*`
- `skill.*`
- `http.call`

所有命令都必须经过白名单校验，不允许执行未注册命令。

## 8. 选区展示和上下文预算

为了兼顾体验和 token 成本，选区信息分两层处理。

### 8.1 UI 展示层

输入框下方实时展示：

- 当前 Sheet 名
- 选区地址
- 行数
- 列数
- 是否识别出表头

### 8.2 Agent 上下文层

默认只传：

- 工作表名
- 地址
- 行列数
- 表头预览
- 少量样例值

只有在用户问题明确依赖选区数据内容时，才读取更多单元格值并送入 Agent 或 skill 流程。

首版规则：

- 默认只取前 5 行 x 5 列样本
- 大选区先摘要，不直接整块注入 prompt
- 首版只支持连续选区
- 多区域选区先提示用户重新选择

## 9. 会话持久化设计

首版使用本地存储，并通过统一的 Storage Adapter 封装。

建议 key 设计：

- `oa:sessions:index`
- `oa:sessions:<sessionId>`
- `oa:settings`
- `oa:runtime`

推荐存储内容：

- 会话索引：`sessionId`、标题、创建时间、更新时间、最近预览
- 会话消息：消息角色、内容、卡片元信息、确认状态
- 设置项：API Key、模型选择、UI 偏好
- 运行态：当前激活会话 ID

为了避免无限膨胀：

- 单会话消息条数要设上限
- 大 payload 不完整持久化，只保留摘要
- 历史过长时做消息摘要压缩

## 10. 错误处理和安全边界

### 10.1 错误处理

必须覆盖以下场景：

- 未选择区域就触发依赖选区的 skill
- 多区域选区
- 受保护工作表
- 合并单元格导致的解析异常
- API Key 缺失或无效
- LLM 返回格式不合法
- 外部 API 超时、失败、部分成功
- Excel 写操作在确认后执行失败

所有错误都必须在聊天流中以可恢复方式呈现，不能只写控制台。

### 10.2 安全边界

- demo 阶段允许用户手工填写 API Key
- API Key 仅保存在插件本地存储中
- 不把 API Key 写入工作簿
- 不把完整大选区数据默认发送给外部 API
- 所有写操作和外部系统状态变更都必须显式确认

## 11. 兼容性策略

根据当前设计目标，首版兼容策略如下：

- 以 Windows 桌面版 Excel 2019 作为最低验证环境
- 使用经典 XML manifest
- 使用保守的 ExcelApi 能力边界
- 运行时检查能力，不假设最新 Office.js 要求集一定可用
- 前端避免依赖激进的现代浏览器特性

说明：

- Office 2019 已于 2025-10-14 结束官方支持，因此“兼容 Office 2019”是产品目标，不意味着仍可依赖官方当前支持周期

## 12. 测试策略

### 12.1 单元测试

覆盖：

- 命令解析
- skill 路由
- payload 构造
- 会话管理
- 确认逻辑

### 12.2 集成测试

覆盖：

- Excel Adapter 的命令映射
- API Client 的失败和超时处理
- Storage Adapter 的读写一致性

### 12.3 手工验证矩阵

至少覆盖：

- Excel 2019 Windows 桌面版
- 一个更高版本的 Windows 桌面 Excel

重点验证场景：

- 任务窗格加载
- 选区变化实时展示
- 读操作直执
- 写操作确认
- `/upload_data` 成功路径
- `/upload_data` 失败路径
- API Key 缺失和错误
- 大选区降级逻辑
- 会话新建、切换、删除和恢复

## 13. 后续演进方向

正式版可逐步演进：

- API Key -> OAuth
- 本地会话存储 -> 云端会话同步
- 单个样例 skill -> 可扩展 skill 市场
- 前端轻量编排 -> 更正式的远端 Agent 编排

这些演进不应改变当前模块边界，只替换具体实现。

## 14. 结论

首版 OfficeAgent 的最优方案是：

- 纯 Office Add-in
- Excel 任务窗格聊天 UI
- Office.js 作为唯一 Excel 操作通道
- 前端内置轻量 Agent 编排
- 自然语言优先、slash 命令强制 skill
- 读操作直接执行，写操作统一确认
- `upload_data` 作为首个样例 skill
- 会话级本地持久化，插件级历史，不绑定工作簿
- 开发期 localhost，发布期正式 HTTPS 站点

该设计适合进入实现计划阶段。
