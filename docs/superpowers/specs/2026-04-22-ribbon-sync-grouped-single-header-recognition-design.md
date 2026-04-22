# Ribbon Sync 分组 Single 表头识别设计说明

日期：2026-04-22

状态：设计已确认，待进入实施计划

## 1. 目标

在当前 Ribbon Sync 表头识别能力上，支持一种特殊但合理的人工维护场景：

- 某些业务字段本质上仍然是单字段 `single`
- 用户为了便于 Excel 内分类管理，把这些字段在业务表中改造成两层表头
- 用户同时手工维护 `SheetFieldMappings` 中的 `Excel L1 / Excel L2`
- 插件后续按 metadata 精确识别这些列，而不是要求这类列必须保持插件默认生成的单层样式

本次设计的目标是：

- 支持“分组 single 表头”在现有业务表中的运行时识别
- 保持 `single` 与 `activityProperty` 的字段身份边界清晰
- 不新增新的业务字段类型
- 不改变当前默认表头生成样式
- 不做只改 Excel、不改 metadata 的自动猜测识别

## 2. 范围

### 2.1 本次范围

- 支持 `HeaderType = single` 且 `Excel L2` 非空时，按双层表头识别该列
- 让以下路径支持这种识别方式：
  - `部分下载`
  - `部分上传`
  - `全量下载` 在“当前表头已存在且可识别”时复用现有布局
- 为冲突 metadata 提供显式错误
- 为上述行为补单元测试
- 更新 Ribbon Sync 当前行为文档

### 2.2 本次明确不做

- 不新增新的 `HeaderType`，例如 `groupedSingle`
- 不支持只改 Excel 表头、不改 `SheetFieldMappings` 的自动识别
- 不支持空表头时自动生成这种分组 single 表头
- 不改变当前 `WorksheetSchemaLayoutService` 的默认 single 渲染样式
- 不支持 `HeaderRowCount = 1` 下识别带 `Excel L2` 的 `single`

## 3. 场景定义

当前已支持的两类字段是：

- 普通单层字段 `single`
- 活动双层字段 `activityProperty`

本次新增支持的是第三种展示形态，但不是第三种字段类型：

- 分组 single 字段

它的字段身份仍然是 `single`，只是 Excel 展示从：

- 单层表头：`负责人`

变成：

- 第一行：`联系人信息`
- 第二行：`负责人`

也就是说，字段仍然只有一个 `ApiFieldKey`，没有活动属性展开，也没有额外业务语义，只是表头展示变成“分组名 + 字段名”。

## 4. Metadata 约定

本次继续沿用当前 `SheetFieldMappings` 四列显示模型：

- `ISDP L1`
- `Excel L1`
- `ISDP L2`
- `Excel L2`

以及当前列顺序：

- `HeaderType`
- `ISDP L1`
- `Excel L1`
- `ISDP L2`
- `Excel L2`
- `HeaderId`
- `ApiFieldKey`
- `IsIdColumn`
- `ActivityId`
- `PropertyId`

对 `single` 的解释规则调整为：

- `HeaderType = single` 且 `Excel L2` 为空
  - 表示普通 single
- `HeaderType = single` 且 `Excel L2` 非空
  - 表示分组 single

对 `activityProperty` 的解释保持不变：

- `HeaderType = activityProperty`
  - `Excel L1` 是父表头
  - `Excel L2` 是子表头

示例：

| HeaderType | Excel L1 | Excel L2 | ApiFieldKey | 含义 |
| --- | --- | --- | --- | --- |
| `single` | `负责人` |  | `owner_name` | 普通 single |
| `single` | `联系人信息` | `负责人` | `owner_name` | 分组 single |
| `activityProperty` | `测试活动111` | `开始时间` | `start_12345678` | 活动双层 |

## 5. 运行时匹配规则

### 5.1 索引分类

在 `WorksheetHeaderMatcher` 中，匹配索引拆成三类：

- 普通 single 索引
  - 条件：`HeaderType = single` 且 `Excel L2` 为空
  - key：`Excel L1`
- 分组 single 索引
  - 条件：`HeaderType = single` 且 `Excel L2` 非空
  - key：`Excel L1 + Excel L2`
- activity 索引
  - 条件：`HeaderType = activityProperty`
  - key：`Excel L1 + Excel L2`

### 5.2 HeaderRowCount = 2 时的匹配顺序

当 `HeaderRowCount = 2` 时，每列读取：

- 第一行文本 `topText`
- 第二行文本 `bottomText`
- 当前连续父头 `currentParent`

匹配顺序如下：

1. 先匹配普通 single
   - 条件：
     - `topText = Excel L1`
     - 且 `bottomText` 为空，或 `bottomText == topText`
2. 再匹配分组 single
   - 条件：
     - `currentParent = Excel L1`
     - `bottomText = Excel L2`
3. 最后匹配 `activityProperty`
   - 条件：
     - `currentParent = Excel L1`
     - `bottomText = Excel L2`

这样做的目的：

- 保持原有普通 single 行为不回归
- 仅在明确存在子表头时，才进入双层 single / activity 的识别路径
- 避免把现有纵向合并 single 误认成分组 single

### 5.3 HeaderRowCount = 1 的行为

当 `HeaderRowCount = 1` 时：

- 仅支持普通 single 按 `Excel L1` 匹配
- 若 metadata 中存在 `single` 且 `Excel L2` 非空，则该列在当前布局下不支持识别

本次不提供该场景的自动降级或猜测逻辑。

## 6. 执行链路行为

### 6.1 需要支持的路径

本次能力只覆盖以下路径：

- `部分下载`
- `部分上传`
- `全量下载` 在当前工作表表头已存在且可识别时复用现有布局

这些路径都依赖 `WorksheetHeaderMatcher` 产出的运行时列映射，因此只要 matcher 能识别，后续选区解析、`row_id` 回找、字段回写都可以复用现有链路。

### 6.2 空表头时的行为

当当前表头区为空时：

- `全量下载` 仍按现有默认 single / activity 布局生成表头
- 即使 metadata 中某些 `single` 字段配置了 `Excel L1 + Excel L2`
- 插件也不会主动生成分组 single 父头

换言之：

- 分组 single 只是一种“现有布局识别能力”
- 不是一种“默认生成布局能力”

### 6.3 BuildConfiguredColumns 的要求

在空表头走默认生成路径时：

- 普通 single：继续取 `Excel L1` 作为显示文本
- 分组 single：按普通 single 处理，显示文本取 `Excel L2`

这样可以保证：

- 空表生成路径不扩展新布局
- metadata 中存在 `Excel L1 + Excel L2` 不会阻塞默认 single 布局生成

## 7. 冲突与错误处理

本次不做猜测匹配，metadata 冲突必须显式报错。

### 7.1 需要拦截的冲突

- 两个 `single` 字段拥有相同的 `Excel L1 + Excel L2`
- 一个分组 single 与一个 `activityProperty` 拥有相同的 `Excel L1 + Excel L2`
- `HeaderRowCount = 1`，但 metadata 中存在 `single` 且 `Excel L2` 非空，并要求按当前表头识别

### 7.2 错误策略

建议在 `WorksheetHeaderMatcher` 构建索引阶段直接抛出 `InvalidOperationException`，错误信息指向 `AI_Setting` / `SheetFieldMappings` 配置问题，而不是继续走泛化的“表头无法匹配”。

推荐错误口径：

- `SheetFieldMappings 中存在重复的双层表头键，请先修正 AI_Setting。`
- `当前 HeaderRowCount=1，无法识别带 Excel L2 的 single 表头，请先修正 AI_Setting。`

## 8. 示例

业务字段：

- `owner_name`
- `department_name`
- `phone_number`

用户希望在 Excel 中把它们归到 `联系人信息` 这个一级分组下。

对应 metadata：

| HeaderType | Excel L1 | Excel L2 | ApiFieldKey |
| --- | --- | --- | --- |
| `single` | `联系人信息` | `负责人` | `owner_name` |
| `single` | `联系人信息` | `所属部门` | `department_name` |
| `single` | `联系人信息` | `联系电话` | `phone_number` |

对应 Excel 表头：

| 第一行 | 第二行 |
| --- | --- |
| 联系人信息 | 负责人 |
| 联系人信息 | 所属部门 |
| 联系人信息 | 联系电话 |

在这个场景下：

- `部分下载` 能正确把 `负责人` 识别为 `owner_name`
- `部分上传` 能正确把 `所属部门` 识别为 `department_name`
- `全量下载` 如果表头已存在，则复用当前布局

但如果表头区为空：

- 插件不会自动生成第一行 `联系人信息`
- 只会按当前默认 single 样式生成 `负责人 / 所属部门 / 联系电话`

## 9. 测试要求

至少补以下测试：

- `WorksheetHeaderMatcherTests`
  - `HeaderRowCount = 2` 时，分组 single 按 `L1 + L2` 识别成功
  - 普通 single 识别不回归
  - `activityProperty` 识别不回归
  - 重复分组 single 键时报错
  - 分组 single 与 activity 双层键冲突时报错
  - `HeaderRowCount = 1` 下带 `Excel L2` 的 single 报错

- `WorksheetSyncExecutionServiceTests`
  - 已有分组 single 表头时，`PreparePartialDownload` 识别正确列
  - 已有分组 single 表头时，`PreparePartialUpload` 识别正确列
  - 已有分组 single 表头时，`PrepareFullDownload` 走 `UsesExistingLayout = true`
  - 空表头时，分组 single 仍按默认 single 样式生成，不生成父头

- 文档
  - `docs/modules/ribbon-sync-current-behavior.md`
  - 如有必要，同步手工测试清单

## 10. 风险与取舍

本次设计刻意选择“显式 metadata 驱动”，而不是“宽松猜测匹配”。

优点：

- 行为稳定
- 排错路径清晰
- 不会把同名子表头误识别到错误字段

代价：

- 用户必须同时维护 Excel 和 `SheetFieldMappings`
- 配置不一致时会更早报 metadata 错误

这个取舍符合当前 Ribbon Sync 的总体原则：

- 表头文本是运行时事实
- metadata 是唯一受信的字段映射来源
- 插件不自动猜测用户的手工重构意图
