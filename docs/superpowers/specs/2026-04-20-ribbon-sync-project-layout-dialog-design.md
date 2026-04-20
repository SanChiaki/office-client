# Ribbon Sync 项目切换布局参数弹框设计说明

日期：2026-04-20

状态：设计已确认，待进入实施计划

## 1. 目标

在当前 Ribbon Sync 项目下拉框交互上补一个显式参数输入步骤，使用户在下列场景中可以手工填写：

- `HeaderStartRow`
- `HeaderRowCount`
- `DataStartRow`

本次设计要满足以下目标：

- 在当前 sheet 首次绑定项目时弹框
- 在当前 sheet 切换到其他项目时弹框
- 弹框默认带出当前可用的布局参数
- 用户点击取消时，完全取消本次项目切换
- 只有用户确认且参数合法时，才写入 `SheetBindings`
- 保持现有“选择项目后不自动初始化”的行为

## 2. 范围

### 2.1 本次范围

- 在项目下拉框选中项目后增加一个 WinForms 原生弹框
- 为弹框提供默认值填充、数值校验和取消行为
- 将确认后的参数写入 `AI_Setting` 的 `SheetBindings`
- 让 Ribbon 下拉框在取消时回到原来的选中显示
- 为上述行为补单元测试和配置型测试

### 2.2 本次明确不做

- 不改 `初始化当前表` 的职责
- 不在弹框确认后自动生成或刷新 `SheetFieldMappings`
- 不新增 task pane 交互
- 不新增更多布局参数
- 不改下载、上传、表头匹配的主流程规则

## 3. 交互设计

### 3.1 触发时机

当用户从 Ribbon 项目下拉框选择一个项目时：

- 如果当前 sheet 还没有绑定项目，则弹框
- 如果当前 sheet 已绑定其他项目，则弹框
- 如果当前 sheet 选择的还是同一个项目，则不弹框，也不重复写入 binding

这里“同一个项目”的判断规则是：

- `SystemKey` 相同
- `ProjectId` 相同

### 3.2 弹框内容

弹框包含三个输入项：

- `HeaderStartRow`
- `HeaderRowCount`
- `DataStartRow`

以及两个按钮：

- `确定`
- `取消`

弹框标题和说明文本应明确表明：这些参数将写入当前 sheet 的同步配置，用于后续初始化、下载和上传。

### 3.3 默认值来源

弹框打开时，三个输入项按以下优先级带默认值：

1. 当前 sheet 已保存的 binding 值
2. 当前连接器 `CreateBindingSeed()` 返回的 seed 值

这意味着：

- 首次绑定项目时，默认显示连接器 seed，当前系统下就是 `1 / 2 / 3`
- 已有 binding 且切换项目时，默认显示当前 sheet 已保存的 `HeaderStartRow / HeaderRowCount / DataStartRow`

这样可以保留用户已经维护过的布局习惯。

### 3.4 确认行为

用户点击 `确定` 后：

1. 读取三个输入值
2. 执行数值校验
3. 校验通过后，生成新的 `SheetBinding`
4. 写入 `AI_Setting`
5. 更新 `RibbonSyncController` 当前项目状态
6. 刷新 Ribbon 下拉框文本为新项目

### 3.5 取消行为

用户点击 `取消` 后：

- 本次项目切换完全终止
- 不写 `AI_Setting`
- 不修改当前 sheet 的 `SheetBindings`
- `RibbonSyncController` 当前项目状态保持不变
- Ribbon 下拉框回到取消前的文本

取消必须是强取消，不允许“项目已切换但参数未更新”的半完成状态。

## 4. 校验规则

弹框确认时必须满足以下规则：

- `HeaderStartRow > 0`
- `HeaderRowCount > 0`
- `DataStartRow > 0`
- `DataStartRow >= HeaderStartRow + HeaderRowCount`

校验失败时：

- 不关闭弹框
- 给出原生错误提示
- 保持用户当前输入，允许继续修改

这条 `DataStartRow` 规则用于确保：

- 表头区不会和数据区重叠
- 后续下载、上传、表头匹配逻辑继续可以假设布局合法

## 5. 数据写入规则

### 5.1 新 binding 生成规则

确认通过后，`RibbonSyncController.SelectProject()` 不再直接使用“seed + 旧值保留”的静默合成方式，而是：

1. 先构造一个建议 binding
2. 把建议 binding 交给弹框编辑
3. 使用弹框返回值作为最终 binding
4. 保存到 `metadataStore.SaveBinding()`

最终保存的字段包括：

- `SheetName`
- `SystemKey`
- `ProjectId`
- `ProjectName`
- `HeaderStartRow`
- `HeaderRowCount`
- `DataStartRow`

### 5.2 与现有初始化流程的关系

本次设计保持当前行为：

- 选择项目后只写 `SheetBindings`
- 不自动执行 `InitializeCurrentSheet`
- `SheetFieldMappings` 仍需用户显式点击 `初始化当前表` 才会写入或刷新

这意味着本次变更只影响“项目绑定时如何收集布局参数”，不改变初始化边界。

## 6. 架构调整

### 6.1 RibbonSyncController

`RibbonSyncController` 需要承担新的职责：

- 判断当前选择是否真的发生项目切换
- 构造建议 binding
- 调起布局参数弹框
- 处理取消、确认、校验失败后的结果

建议新增一个内部步骤：

- `PromptForBindingLayout(...)`

它返回三种结果语义：

- 确认并返回 binding
- 用户取消
- 理论上不返回非法 binding

### 6.2 新增对话框

新增一个 WinForms 对话框类型，职责单一：

- 展示三个数值输入项
- 做本地输入校验
- 在确认时返回用户填写的参数

不把项目切换逻辑放进对话框内部。

推荐让对话框只处理：

- 展示
- 输入解析
- 校验提示
- 返回结果

而由 `RibbonSyncController` 负责：

- 何时弹
- 取消后如何恢复状态
- 最终保存什么 binding

### 6.3 AgentRibbon

`AgentRibbon` 现有 `TextChanged -> SelectProject(project)` 链路保持不变，但要满足一个额外结果：

- 如果控制器返回“用户取消”，Ribbon 下拉框必须刷新回原有项目文本

也就是说，控制器和 Ribbon 之间需要继续依赖现有的 `ActiveProjectChanged` 刷新链路，而不是在 Ribbon 层自行猜测取消后的文本。

## 7. 失败与边界处理

### 7.1 非数字输入

如果用户输入非数字：

- 阻止确认
- 提示必须输入正整数

### 7.2 合法整数但违反布局规则

例如：

- `HeaderStartRow = 1`
- `HeaderRowCount = 2`
- `DataStartRow = 2`

这类输入必须阻止确认，并提示数据区起始行不能落在表头区内。

### 7.3 当前活动 sheet 不可用

如果下拉框事件发生时当前活动 sheet 不可用：

- 沿用现有控制器异常行为
- 不显示布局弹框
- 不写 binding

### 7.4 已有 binding 但项目名为空

如果当前 binding 中 `ProjectName` 为空：

- 下拉框回显格式退化为 `ProjectId`
- 弹框默认值仍然优先读取该 binding 中保存的三个行号字段

## 8. 测试策略

需要覆盖以下场景：

- 首次绑定项目时会弹框，并在确认后保存用户输入值
- 切换到其他项目时会弹框，并保留用户输入值
- 重新选择同一个项目时不弹框
- 取消弹框时不保存 binding，并恢复原下拉框显示
- 弹框默认值优先读取已有 binding
- 无已有 binding 时，弹框默认值来自连接器 seed
- 非数字输入无法确认
- `DataStartRow < HeaderStartRow + HeaderRowCount` 无法确认
- 选择项目后仍然不会自动初始化

测试分层建议：

- `RibbonSyncControllerTests`
  - 覆盖项目切换、取消、默认值来源、保存结果
- 新对话框测试
  - 覆盖输入解析和校验规则
- `AgentRibbonConfigurationTests`
  - 覆盖控制器取消后仍通过现有刷新链路恢复下拉框

## 9. 结论

本次采用的最终方案是：

- 仅在“首次绑定项目”或“切换到其他项目”时弹出布局参数框
- 默认值优先取当前 sheet 已保存布局，其次取连接器 seed
- 用户确认后才写 `AI_Setting`
- 用户取消时完全取消项目切换
- 强制校验三个值为正整数，且 `DataStartRow >= HeaderStartRow + HeaderRowCount`
- 保持“选择项目后不自动初始化”的现有行为

这套方案能把布局控制权交给用户，同时继续保持当前 Ribbon Sync 的状态一致性和可预测性。
