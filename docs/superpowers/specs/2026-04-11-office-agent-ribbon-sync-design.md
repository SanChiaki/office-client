# OfficeAgent Ribbon 数据同步设计说明

日期：2026-04-11

状态：设计已确认，待进入实施计划

## 1. 目标

在当前 Excel VSTO 插件中新增一组独立于 Agent 的 Ribbon 数据同步能力，满足以下业务目标：

- 用户从 Ribbon 直接执行下载和上传
- 不通过任务窗格承载任何本功能的交互
- 使用 Office/WinForms 原生弹框完成确认、提示和错误反馈
- 支持全量下载、部分下载、全量上传、部分上传、增量上传
- 首版只接入当前这一个应用系统
- 架构上预留未来接入多个应用系统的扩展位

## 2. 范围

### 2.1 首版范围

- Ribbon 上新增项目下拉框与 5 个按钮
- 工作表绑定项目上下文
- 通过 `/find` 完成全量下载和部分下载
- 通过 `/head` 获取字段映射与表头定义
- 通过 `/batchSave` 提交上传改动
- 支持仅可见单元格优先的非连续选区解析
- 支持同一项目内混合单层表头列和双层活动表头列
- 支持按 `id + apiFieldKey` 的快照差异实现增量上传
- 提供 mock server 测试接口覆盖下载、表头、上传

### 2.2 明确不做

- 不通过任务窗格展示本功能的确认内容或执行结果
- 不通过 Agent、skill、自然语言路由本功能
- 不处理无 `id` 新增行上传
- 不处理删除行同步
- 不实现多应用系统切换 UI

## 3. Ribbon 交互

Ribbon 新增三个分组：

- `项目`
  - 项目下拉框
- `下载`
  - `全量下载`
  - `部分下载`
- `上传`
  - `全量上传`
  - `部分上传`
  - `增量上传`

交互规则：

- 用户先通过 Ribbon 下拉框选择项目
- 当前活动 sheet 已绑定项目信息时，切换到该 sheet 后自动回填下拉框
- 当前活动 sheet 未绑定项目信息时，下拉框显示“先选择项目”
- 所有按钮都以当前活动 sheet 为作用目标
- 所有确认、警告、成功、失败反馈都通过原生弹框完成

首版项目下拉数据来源：

- 由 `CurrentBusinessSystemConnector` 提供本地静态项目清单
- mock server 不新增独立项目列表接口
- 后续若真实系统提供项目接口，只替换 connector 的项目清单实现

## 4. 总体架构

本功能单独走一条 Ribbon 驱动的结构化同步链路，不与现有 Agent/task pane 共享执行入口。

```text
Ribbon
  -> RibbonSyncController
      -> WorksheetSyncService
          -> WorksheetProjectBindingService
          -> WorksheetSchemaLayoutService
          -> WorksheetSelectionResolver
          -> WorksheetChangeTracker
          -> SystemConnectorRegistry
              -> CurrentBusinessSystemConnector
                  -> /head
                  -> /find
                  -> /batchSave
      -> Native Dialogs
      -> Excel Interop
      -> Metadata Worksheet
```

职责划分：

- Ribbon 层：只处理按钮和下拉框事件
- 同步应用层：编排下载、上传、预检、确认预览、快照更新
- Connector 层：封装当前业务系统的接口协议与字段展开规则
- Excel 层：读写表格、解析选区、生成表头、管理元数据 sheet
- Dialog 层：原生确认与结果弹框

## 5. 扩展模型

虽然首版只接当前业务系统，但内部统一使用 `systemKey` 识别外部系统。

### 5.1 `systemKey`

`systemKey` 是内部扩展标识，用于区分工作表绑定的是哪一个外部系统。首版固定为当前系统的一个常量值，例如：

```text
current-business-system
```

### 5.2 Connector 抽象

建议新增以下接口：

- `ISystemConnector`
- `IProjectCatalogProvider`
- `ISchemaProvider`
- `IDataDownloadGateway`
- `IDataUploadGateway`

首版实现：

- `CurrentBusinessSystemConnector`

未来新系统只需要新增 connector 与字段展开规则，不需要改 Ribbon 主体和大部分 Excel 同步逻辑。

## 6. 工作表绑定与元数据

### 6.1 工作表绑定

每个受管理 sheet 绑定以下信息：

- `sheetName`
- `systemKey`
- `projectId`
- `projectName`

当用户在 Ribbon 中选择项目时：

- 将当前活动 sheet 绑定到该项目
- 后续切回该 sheet 时自动回填 Ribbon 下拉框

### 6.2 元数据存储方式

Workbook 内维护一个独立元数据 sheet，例如：

```text
_OfficeAgentMetadata
```

当前阶段该 sheet 保持可见，方便开发调试。后续若需要改为隐藏，只调整显示策略，不改变数据结构。

元数据 sheet 内至少维护三类结构化数据：

- `SheetBindings`
  - 记录 sheet 与 `systemKey/projectId/projectName` 的绑定关系
- `SheetSchemas`
  - 记录当前 sheet 的列布局结果和列绑定结果
- `SheetSnapshots`
  - 记录最近一次成功下载后的 `id + apiFieldKey -> value` 基线

## 7. 表头与列绑定模型

### 7.1 核心原则

当前业务系统返回的是扁平 JSON list。Excel 展示可以是单层或双层表头，但内部字段身份必须始终基于真实接口字段，而不是显示名。

因此每一列都必须绑定一个稳定的：

- `apiFieldKey`

例如：

- `name`
- `owner`
- `start_12345678`
- `end_12345678`

上传、部分刷新、增量比较都只认 `apiFieldKey`。

### 7.2 混合表头

同一项目内允许同时存在：

- 单层表头列
- 双层活动表头列

单层与双层不是互斥模式，而是同一张表里的不同列类型。

### 7.3 单层列

单层列规则：

- 扁平 JSON 中的普通字段直接映射到显示表头
- Excel 中第 1-2 行纵向合并显示该表头
- 示例：

```text
| 负责人 |
|        |
| 张三   |
```

### 7.4 双层活动列

双层列表示“活动实例 + 活动属性”。

例如返回数据中的字段：

```json
{
  "start_12345678": "2026-01-02",
  "end_12345678": "2026-01-02"
}
```

其含义为：

- `12345678` 是活动 ID
- `start` 和 `end` 是活动属性字段

展开规则：

1. 从 `apiFieldKey` 中解析出：
   - `propertyKey`
   - `activityId`
2. 用 `activityId` 查询活动信息，得到活动名
3. 用静态属性字段关系表将 `propertyKey` 映射为实际第二层表头名

例如：

- `activityId = 12345678`
- 活动名 = `测试活动111`
- `start -> 开始时间`
- `end -> 结束时间`

最终 Excel 表头：

```text
|      测试活动111      |
| 开始时间 | 结束时间 |
|2026-01-02|2026-01-02|
```

第 1 行的活动名单元格为横向合并单元格。

### 7.5 列绑定模型

建议每列至少包含：

- `columnIndex`
- `apiFieldKey`
- `columnKind`
  - `Single`
  - `ActivityProperty`
- `parentHeaderText`
- `childHeaderText`
- `activityId`
- `activityName`
- `propertyKey`
- `isIdColumn`

## 8. 当前系统接口模型

当前 mock server 和 connector 按以下三类接口固化。

### 8.1 `/find`

用途：

- 全量下载
- 部分下载

规则：

- 全量下载：按项目查询完整数据
- 部分下载：按项目 + `id` 集合 + 字段集合查询局部数据
- 返回扁平 JSON list

### 8.2 `/head`

用途：

- 获取字段映射和表头配置

返回 JSON 中包含：

- `headList`

`headList` 保存所有表头定义，但不直接展开活动属性列。

其中：

- 普通表头直接定义 `apiFieldKey -> headerText`
- 活动类定义通过 `headType = "activity"` 标识
- 活动类 item 同时提供活动实例身份信息，至少包含：
  - `activityId`
  - `activityName`

活动属性的第二层表头名来自静态属性字段关系表，不直接硬编码在 Excel 逻辑中。

### 8.3 `/batchSave`

用途：

- 全量上传
- 部分上传
- 增量上传

请求体为一个 list，每个 item 表示一个单元格改动，而不是整行改动。

首版按以下结构设计：

```json
{
  "projectId": "p1",
  "id": "row_1001",
  "fieldKey": "start_12345678",
  "value": "2026-01-02"
}
```

内部字段命名：

- `projectId`
- `id`
- `fieldKey`
- `value`

## 9. 选区解析与定位规则

### 9.1 选区口径

部分下载和部分上传默认按：

- `仅可见单元格优先`

如果用户选区包含隐藏行或隐藏列，则默认不处理隐藏单元格。

### 9.2 非连续选区

选区可以是不连续区域。解析时需要：

- 先拆成多个 area
- 过滤出可见单元格
- 合并成待处理的坐标集合

### 9.3 不要求选中表头或 ID 列

用户选区可以不包含：

- 表头行
- ID 列

系统必须能自动回溯：

- 该列对应哪个 `apiFieldKey`
- 该行对应哪个 `id`

前提规则：

- 表头固定在第 1-2 行
- 数据区从第 3 行开始
- `ID` 列字段身份固定，可从列绑定模型中识别

### 9.4 解析结果

选区最终应被解析成：

- `id` 集合
- `apiFieldKey` 集合
- 精确到单元格坐标的写回目标

## 10. 五个动作的执行流

### 10.1 全量下载

流程：

1. 校验当前 sheet 是否已绑定项目
2. 检查当前 sheet 是否存在未上传本地改动
3. 调 `/head` 获取字段与表头定义
4. 调 `/find` 获取全量数据
5. 创建或刷新当前项目对应的专用 sheet
6. 写入混合表头和数据
7. 更新元数据 sheet 中的绑定、schema、snapshot

默认写入位置：

- 项目专用 sheet

若存在未上传改动：

- 强提醒覆盖风险
- 默认取消
- 用户二次确认后才允许继续

### 10.2 部分下载

流程：

1. 校验当前 sheet 为受管理 sheet
2. 解析当前选区得到 `id` 集合与 `apiFieldKey` 集合
3. 检查本次将覆盖的目标单元格是否存在未上传改动
4. 调 `/find` 获取局部数据
5. 按原单元格位置写回
6. 更新本次涉及单元格的快照

### 10.3 全量上传

流程：

1. 校验当前 sheet 为受管理 sheet
2. 从整个数据区读取所有带 `id` 的行
3. 将每个数据单元格展开为 `/batchSave` item list
4. 调 `/batchSave`
5. 成功后刷新对应快照

### 10.4 部分上传

流程：

1. 校验当前 sheet 为受管理 sheet
2. 解析当前仅可见选区
3. 从列绑定模型回溯到 `apiFieldKey`
4. 从对应数据行回溯到 `id`
5. 将选中的每个改单元格展开为 `/batchSave` item list
6. 成功后只更新这些单元格的快照

### 10.5 增量上传

流程：

1. 校验当前 sheet 为受管理 sheet
2. 校验存在最近一次成功下载形成的快照
3. 扫描整个数据区
4. 仅处理已有 `id` 的行
5. 比较当前值与快照值，找出所有变化单元格
6. 展开为 `/batchSave` item list
7. 成功后更新这些单元格的快照

限制：

- 首版只上传已有 `id` 行中的改单元格
- 不处理新增行
- 不处理删除行

## 11. 差异与覆盖规则

### 11.1 本地未上传改动

定义：

- 某个有 `id` 的数据单元格，当前值与最近下载快照中的 `id + apiFieldKey` 基线不同

### 11.2 下载覆盖拦截

- 全量下载：
  - 只要当前 sheet 存在任意未上传改动，即触发强提醒
- 部分下载：
  - 只要本次目标单元格中存在未上传改动，即触发强提醒

### 11.3 上传成功与失败

- 上传成功：
  - 只更新实际成功提交的单元格快照
- 上传失败：
  - 不修改快照
  - 不清理本地脏状态

## 12. 原生弹框策略

本功能不使用任务窗格展示确认内容。

建议使用两类原生交互：

- 简单提示：
  - `MessageBox`
- 复杂确认：
  - WinForms 模态窗

建议新增三类对话框：

- `DownloadConfirmDialog`
  - 展示项目、sheet、记录数、字段数、覆盖风险
- `UploadConfirmDialog`
  - 展示项目、记录数、字段数、改单元格数、diff 示例
- `OperationResultDialog`
  - 展示成功/失败摘要与错误明细

确认强度：

- 下载：轻确认
- 上传：强确认，且带可展开 diff 预览

## 13. 模块拆分建议

### 13.1 Ribbon 与宿主

- `src/OfficeAgent.ExcelAddIn/AgentRibbon.cs`
- `src/OfficeAgent.ExcelAddIn/AgentRibbon.Designer.cs`
- `src/OfficeAgent.ExcelAddIn/ThisAddIn.cs`

### 13.2 Excel Add-in 层

建议新增：

- `src/OfficeAgent.ExcelAddIn/Excel/WorksheetMetadataStore.cs`
- `src/OfficeAgent.ExcelAddIn/Excel/WorksheetProjectBindingService.cs`
- `src/OfficeAgent.ExcelAddIn/Excel/WorksheetSchemaLayoutService.cs`
- `src/OfficeAgent.ExcelAddIn/Excel/WorksheetSelectionResolver.cs`
- `src/OfficeAgent.ExcelAddIn/Excel/WorksheetChangeTracker.cs`
- `src/OfficeAgent.ExcelAddIn/Dialogs/DownloadConfirmDialog.cs`
- `src/OfficeAgent.ExcelAddIn/Dialogs/UploadConfirmDialog.cs`
- `src/OfficeAgent.ExcelAddIn/Dialogs/OperationResultDialog.cs`

### 13.3 Core 层

建议新增：

- `src/OfficeAgent.Core/Sync/WorksheetSyncService.cs`
- `src/OfficeAgent.Core/Sync/SyncOperationPreviewFactory.cs`
- `src/OfficeAgent.Core/Models/ProjectBinding.cs`
- `src/OfficeAgent.Core/Models/WorksheetSchema.cs`
- `src/OfficeAgent.Core/Models/WorksheetColumnBinding.cs`
- `src/OfficeAgent.Core/Models/WorksheetSnapshot.cs`
- `src/OfficeAgent.Core/Models/CellChange.cs`
- `src/OfficeAgent.Core/Services/ISystemConnector.cs`
- `src/OfficeAgent.Core/Services/IProjectCatalogProvider.cs`
- `src/OfficeAgent.Core/Services/ISchemaProvider.cs`
- `src/OfficeAgent.Core/Services/IDataDownloadGateway.cs`
- `src/OfficeAgent.Core/Services/IDataUploadGateway.cs`

### 13.4 Infrastructure 层

建议新增或拆分：

- `src/OfficeAgent.Infrastructure/Http/CurrentBusinessSystemConnector.cs`
- `src/OfficeAgent.Infrastructure/Http/ProjectCatalogClient.cs`
- `src/OfficeAgent.Infrastructure/Http/SchemaClient.cs`
- `src/OfficeAgent.Infrastructure/Http/DataFindClient.cs`
- `src/OfficeAgent.Infrastructure/Http/BatchSaveClient.cs`

现有 `BusinessApiClient` 不再继续承载这条 Ribbon 同步链路的全部职责。

## 14. 测试与 Mock Server

### 14.1 Mock Server 接口

`tests/mock-server/server.js` 需要补齐：

- `/find`
- `/head`
- `/batchSave`

其中：

- `/find` 同时支持全量下载和部分下载
- `/head` 返回 `headList`
- `/batchSave` 按改单元格 list 处理保存

### 14.2 Mock 数据要求

至少包含以下场景：

- 单层字段
- 混合表头项目
- 活动实例数据
- 静态属性字段关系表
- 固定 ID 字段
- 部分字段更新
- 未登录
- 非法项目
- 非法字段
- 找不到记录

### 14.3 单元测试重点

- 单层列与双层活动列展开
- 活动名横向合并和单层表头纵向合并的布局规则
- 非连续且仅可见选区的解析
- 从选区回溯 `id + apiFieldKey`
- 快照差异计算
- 增量上传仅处理已有 `id` 行
- 下载覆盖未上传改动的拦截
- `/batchSave` payload 生成为改单元格 list

### 14.4 集成测试重点

- `/find` 下载结果到 Excel 表布局的完整链路
- `/head` 到列绑定模型的转换
- `/batchSave` 请求正确性
- 元数据 sheet 写入和读取
- Ribbon 动作到同步服务的编排

### 14.5 手工验证

需要补充到 `docs/vsto-manual-test-checklist.md`：

- Ribbon 项目下拉绑定与自动回填
- 全量下载
- 部分下载
- 全量上传
- 部分上传
- 增量上传
- 本地有未上传改动时再次下载的拦截
- 混合表头项目的渲染与回写

## 15. 结论

首版实现采用：

- Ribbon 作为唯一入口
- 原生弹框作为唯一交互承载
- 当前业务系统的三接口模型：`/find`、`/head`、`/batchSave`
- Workbook 内可见元数据 sheet 作为调试期状态载体
- 以 `apiFieldKey` 为真实字段身份
- 以混合表头布局承载单层字段与活动属性字段
- 以 `id + apiFieldKey` 快照实现增量上传

这条方案满足当前业务系统接入需求，同时保留了未来按 `systemKey + connector` 扩展到多系统的架构空间。
