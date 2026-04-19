# Matrix Workspace

当前默认入口是 `matrix_calculator.ps1`（薄加载器，按顺序 dot-source 加载 `modules/` 下的模块）。
- 推荐启动方式：双击 `启动矩阵计算器.cmd`
- `matrix_calculator.py` 目前仍是占位提示
- `MatrixWorkspace.cs` / `MatrixWorkspace.exe` 是另一套实现，本 README 以 PowerShell 版本为准

## 项目结构（模块化）

当前主脚本已拆分为模块化结构，每个文件不超过 5000 行：

- `matrix_calculator.ps1` — 主入口（薄加载器，7 行），定义 `` 和 ``，按顺序 dot-source 所有模块
- `modules/matrix_core.ps1` — 核心模块（1718 行）：Add-Type 程序集加载、DPI 感知、UI 字符串字典、张量数据操作、矩阵数学运算、特征值/奇异值/分解、表达式解析器、值运算与函数调度
- `modules/matrix_ui_helpers.ps1` — UI 辅助模块（1091 行）：全局状态变量、按钮样式与窗体缩放、矩阵显示辅助、桌宠集成、历史记录、数据导入/导出/处理、帮助字典、矩阵卡片辅助函数
- `modules/matrix_dialogs.ps1` — 对话框模块（1024 行）：帮助内容显示、特殊矩阵弹窗、历史记录窗口、数据对话框、矩阵卡片管理、表达式快捷按钮
- `modules/matrix_form.ps1` — 主窗体模块（660 行）：主窗体创建、事件处理、应用运行
- `modules/data_profile.ps1` — 数据画像模块（688 行）
- `modules/data_plotting.ps1` — 数据绘图模块（816 行）

### 模块加载顺序

1. `matrix_core.ps1` — 无前置依赖
2. `matrix_ui_helpers.ps1` — 依赖 core 中的函数和 `` 字典
3. `matrix_dialogs.ps1` — 依赖 core 和 ui_helpers
4. `matrix_form.ps1` — 依赖所有前置模块

### 封装注意

若后续重新封装 EXE，需要在 `MatrixWorkspaceLauncher.cs` 和 `exe_build_tools/一键生成exe.cmd` 中同步更新资源列表，确保以上 6 个模块文件均被嵌入。

### 备份文件说明

- `matrix_calculator.original.ps1` — 原始主脚本完整备份（4496 行），确认模块化版本运行正常后可删除
- `corrupted_backup_parts/` — 损坏备份的存档拆分（part_1.ps1 ~ part_7.ps1），仅供存档参考
- `matrix_calculator.corrupted.backup.ps1` — 原始损坏备份文件（33410 行），已拆分存档

## 当前能力

- 支持二维矩阵、批量矩阵和更高维规则数组
- 所有矩阵函数都支持“前置维度批量化”
  规则：把最后两维视为矩阵，前面的维度视为批量维
- 支持中文/英文函数名、变量别名、全角符号输入
- 支持特殊矩阵快捷输入与弹窗配置
- 支持分块矩阵输入与计算
- 支持本地数据文件导入与表格处理窗口
- 支持最近 50 次运算历史记录查看与回填
- 主界面已合并为一个四分区页面
  - 左上：矩阵工作区
  - 左下：结果区
  - 右上：表达式区
  - 右下：函数/运算符帮助区
- 四个区域都可拖动改变大小

## 输入格式

### 1. 普通矩阵

```text
[[1,2],[3,4]]
```

```text
[1 2; 3 4]
```

### 2. 多维数组 / 批量矩阵

```text
[[[1,2],[3,4]],[[5,6],[7,8]]]
```

解释：上例可理解为“两个 2x2 矩阵组成的批量矩阵”。

### 3. 分块矩阵

分块矩阵请写在某个矩阵输入框中，例如：

```text
[A | B; B | A]
```

规则：
- 同一行中的块行数必须一致
- 同一列中的块列数必须一致
- 标量会自动按 `1x1` 块处理
- 块中可以引用其他矩阵变量，或引用特殊矩阵简写对应的矩阵卡片

### 4. 中文与全角符号兼容

以下写法都支持：

```text
[1 2； 3 4]
转置（甲）
行列式（A）
```

## 特殊矩阵

### 文本简写

这些写法可以直接填入矩阵输入框：

- `I3` 或 `单位阵3`
- `zeros(2,3)`
- `ones(2,3)`
- `diag(1,2,3)`

### 弹窗配置

每个矩阵卡片右上角有“特殊”按钮，工具栏也有“特殊矩阵”按钮。
点击后会弹出单独窗口，可配置：
- 单位矩阵
- 零矩阵
- 全 1 矩阵
- 对角矩阵

配置完成后：
- 对应矩阵输入框会自动填入简写
- 矩阵卡片会显示当前为“预设”模式
- 表达式区变量快捷按钮会同步显示预设标签，例如 `A:I3`

## 运算符

### 逐元素运算

- `A + B`
- `A - B`
- `A .* B`
- `A ./ B`

要求：两个数组形状一致，或一侧为标量（仅 `.*` / `./` 支持标量混算）。

示例：

```text
A.*B
A./B
A.*2
2./A
```

### 矩阵运算

- `A * B`：矩阵乘法
- `A / B`：右除，按 `A * inv(B)` 处理
- `A ^ 2`：矩阵幂

对至少二维的输入，程序会：
- 将最后两维视为矩阵
- 将前置维度视为批量维
- 对每个批量切片逐个执行矩阵运算

例如：
- `2x3x3` 的张量可理解为“2 个 3x3 矩阵”
- `A^2` 会对每个 `3x3` 切片分别做矩阵平方

## 函数

以下函数均把“最后两维”视为矩阵，对前置维度做批量处理：

- `T(X)` / `转置(X)`
- `REF(X)`
- `RREF(X)`
- `RANK(X)`
- `TR(X)` / `TRACE(X)`
- `ROWSUMS(X)`
- `COLSUMS(X)`
- `ROWMEANS(X)`
- `COLMEANS(X)`
- `ROWMAXS(X)`
- `COLMAXS(X)`
- `ROWMINS(X)`
- `COLMINS(X)`
- `INV(X)`
- `EIGVALS(X)`
- `EIGVECS(X)`
- `SVALS(X)`
- `CHOL(X)`
- `LU(X)`
- `QR(X)`
- `DET(X)`
- `COF(X)`
- `ADJ(X)`
- `NULL(X)`

其中：
- `RANK(X)`、`TR(X)`、`DET(X)` 在批量输入下会返回标量数组
- `ROWSUMS(X)`、`COLSUMS(X)`、`ROWMEANS(X)`、`COLMEANS(X)`、`ROWMAXS(X)`、`COLMAXS(X)`、`ROWMINS(X)`、`COLMINS(X)`、`EIGVALS(X)`、`SVALS(X)` 在批量输入下会返回批量向量
- `EIGVECS(X)`、`CHOL(X)` 在批量输入下会返回批量矩阵
- `LU(X)`、`QR(X)`、`NULL(X)` 在批量输入下会返回逐批次的文本结果

`NORM(X)` 对任意维规则数组都可用，返回整体 Frobenius 范数。

示例：

```text
ROWSUMS(A)
COLMEANS(A)
ROWMAXS(A)
COLMINS(A)
```

## 表达式区与帮助区

### 表达式区

- 支持变量快捷按钮
- 支持常用运算符与函数快捷按钮
- 回车可直接计算
- “示例”按钮会自动生成一个包含分块矩阵的案例

### 帮助区

右下角帮助区支持：
- 输入函数名查询，例如 `DET`
- 输入运算符查询，例如 `.*`
- 输入关键词查询，例如 `BLOCK`、`diag(1,2,3)`
- 点击快捷查询按钮后，下方立即显示含义与用法

## 数据导入与处理

工具栏中的“数据导入与处理”会打开独立窗口。

当前已支持：
- 导入本地 `CSV / TXT / XLSX / XLS / XLSM`
- 默认先全量导入，再通过“遴选行列”按钮筛选所需范围
- 首行作为列名
- 表格数据概览
  - 基础信息：数据集名称、来源、总行数、总列数
  - 缺失与异常：缺失值数量、缺失率、重复行数量、极值提示
  - 总结评分：整体质量、完整性、一致性、可用性
  - 字段类型分布：数值、文本、日期时间、布尔/分类
  - 按列明细：数值列统计、文本/分类统计、日期列统计
- 缺失比例统计
  - 支持按列查看
  - 支持按行查看
- 缺失值清洗
  - 删除含缺失的行
  - 用 0 填充
  - 用列均值填充
- 遴选行列范围
- 修改指定单元格数据
- 分组统计
  - 支持按列分组
  - 支持按行分组
  - 当前统计方法：`COUNT / SUM / MEAN / MAX / MIN`
- 导出表格为 CSV
- 导出当前表格视图为 PNG 图片

当前状态：
- 已开始拆分到模块文件：
  - `modules/data_profile.ps1`
  - `modules/data_plotting.ps1`
- “数据绘图”已进入原型开发阶段：
  - 单图设计页已接入柱状图、条形图、折线图、饼图、面积图、散点图、样条图、环形图、堆积柱状图、雷达图
  - 已接入标题、副标题、配色、字体、导出图片、组合图、四宫格拼图原型
  - 目前以命令级验证为主，GUI 交互还需要继续联调

## 2026-04-16 当前开发备注

这一轮已经完成并验证：
- `INV` 对普通方阵可用
- 非方阵输入会正确报错
- 奇异矩阵会正确报错
- `INV(X)` 已支持按“最后两维为矩阵、前置维度为批量维”逐批处理
- 数据概览、缺失比例、分组统计、单图预览的命令级回归已通过

这一轮仍在处理中，暂时不要当成已彻底解决：
- `1x1` / `n x 1` 文字矩阵解析边界问题已完成修复，当前以最新验证结果为准
- 主脚本底层张量/集合处理还在继续清理，目标是彻底消除 PowerShell 对单元素集合的自动拆平影响
- 打包 `1.1.3` 前，需要把新增模块文件一起接入 EXE 封装链

## 历史记录

工具栏中的“历史记录”会打开独立窗口。

当前已支持：
- 自动保存最近 50 次表达式运算
- 显示时间、状态、表达式、结果
- 点击后查看详细内容
- 一键把历史表达式回填到主界面继续计算

## 说明与限制

- 当前最多可同时放置 8 个矩阵卡片
- 分块矩阵建议在矩阵输入框中定义后，再在表达式中引用对应变量
- 一维向量不适用于需要矩阵结构的函数；这类函数要求输入至少有两维
- `[[5]]` 可正确识别为 `1x1`，`[[1],[2]]` 可正确识别为 `2x1`，`INV([[5]]) = 0.2`，批量 `INV` 保持正常
- 表达式解析目前仍以变量和函数调用为主，建议先在矩阵卡片中定义矩阵，再在表达式区引用
- `matrix_calculator.py` 仍未同步这些能力
- `MatrixWorkspace.cs` / `MatrixWorkspace.exe` 也未同步这些能力

### 开发续记（2026-04-16）

- 已把单元素矩阵解析从旧的 JSON 反序列化思路切换为递归扫描实现
- 这一步的目标是修复 [[5]]、[[1],[2]]、批量矩阵这类输入的层级丢失问题
- 当前语法检查已通过，但命令级回归还在继续，暂时不要把这一条当成已经完全收口
- 当前命令级验证还受到本机 PowerShell 执行策略影响，若看到测试脚本无法直接点源，这属于环境限制，不等同于主脚本语法错误

### 开发续记（2026-04-16，第二次收口）

- 已修复 1x1 与部分单元素结构的文字矩阵解析边界问题
- 已确认：
  - [[5]] 现在可识别为 1x1 矩阵
  - [[1],[2]] 现在可识别为 2x1 矩阵
  - INV([[5]]) = 0.2
  - 批量 INV 仍可正常工作
  - 奇异矩阵、非方阵仍会正确报错
- 本轮先只收口已完成与已验证部分，剩余更大范围的数据模块和绘图模块扩展后续继续

### 状态修正（2026-04-16）

- 先前文档中关于 1x1 / 单元素矩阵“仍在修复”的描述已过时
- 以当前最新验证结果为准：
  - [[5]] 可正确识别为 1x1
  - [[1],[2]] 可正确识别为 2x1
  - INV([[5]]) = 0.2
  - 批量 INV 保持正常

### 封装链状态（2026-04-16）

- MatrixWorkspaceLauncher.cs 已同步支持释放模块文件
- exe_build_tools/一键生成exe.cmd 已同步支持嵌入：
  - matrix_calculator.ps1
  - modules/data_profile.ps1
  - modules/data_plotting.ps1
- 已完成一次实际构建验证，当前 EXE 封装链已能覆盖主脚本 + 数据模块
- 当前构建脚本的版本输出名仍保持 1.1.2，后续正式打包下一版时再统一调整版本号

### 构建脚本版本控制（2026-04-16）

- exe_build_tools/一键生成exe.cmd 已支持用环境变量 MATRIX_WORKSPACE_VERSION 控制版本输出名
- 未设置时，默认仍输出 1.1.2
- 后续正式打包 1.1.3 时，只需要在运行脚本前设置目标版本号，不必再手改批处理内容
- 已顺手清理一处旧回归记录中的乱码数值，避免后续交接时误读验证结果

## 2026-04-16 进度更新
- 已修复绘图模块的图表程序集加载时机问题，数据绘图 不再因为 Charting.Chart 类型未加载而在入口处直接异常。
- 已重构 数据导入与处理 主窗口的上下布局，顶部标题/状态/按钮与下方数据表格分离，避免导入后顶部数据被遮挡。
- 已增强表达式区变量快捷按钮区域：当矩阵数量增多时，按钮区会按数量扩展并启用滚动，不再被直接裁掉。
- 本轮修改后已完成命令级验证：主脚本 PARSE_OK，图表控件创建验证通过。
- 进行中：数据概览中文排版、首行作为列名参与规则、万/亿 等中文数值归一化、分组统计中文化，以及后续的特征值/奇异值/矩阵分解功能。

## 2026-04-16 数据分析模块更新
- 数据概览 已改为中文分节报告，包含基础信息、缺失与异常、质量评分、字段类型分布和逐列分析。
- 数值识别已支持中文数量级：例如 1万、3.6万、1.2亿 会在统计过程中自动换算为标准数值。
- 分组统计页面与输出结果已开始中文化，缺失比例查看也支持按列/按行中文切换。
- 当前统计口径说明已写入界面输出：默认按当前表头作为字段名，不把列名纳入数据统计。

## 2026-04-16 当前可交接状态
### 已完成
- 修复了数据绘图入口的 Charting.Chart 类型加载异常。
- 修复了数据导入窗口顶部遮挡问题，数据表格与顶部信息区已分离。
- 修复了矩阵数量增多时表达式区变量按钮被遮挡的问题。
- 数据概览已中文化，并支持 万/亿 数值归一化。
- 分组统计界面已开始中文化。
- 新增矩阵函数：EIGVALS、EIGVECS、SVALS、CHOL、LU、QR。

### 当前边界
- EIGVALS / EIGVECS 当前稳定支持实对称方阵。
- CHOL 要求输入为对称正定矩阵。
- LU / QR 当前返回文本结果，不是可继续参与表达式计算的矩阵对象。
- 绘图模块当前主要完成了“能打开、不会因类型缺失直接报错”，后续中文化与图表拼接功能仍待继续。

### 下一步建议
- 先做 GUI 联调与帮助说明补全，再考虑打包 1.1.3。
- 若要继续扩展线性代数功能，优先统一新函数的帮助文案、边界提示和结果展示形式。

## 2026-04-16 交接补充（文档已对齐后的当前状态）

- `matrix_calculator.ps1` 当前语法检查仍通过：`PARSE_OK`
- 新增函数 `EIGVALS / EIGVECS / SVALS / CHOL / LU / QR` 的帮助正文已经存在于主脚本帮助字典中
- README 现已同步：
  - 新函数列表
  - 批量返回说明
  - `1x1 / n x 1` 已修复后的真实状态

下一位 AI 建议优先继续：
- 做一次真实 GUI 联调
- 补界面内对新函数边界的提示一致性
- 再决定是否继续推进 1.1.3 打包



## 2026-04-16 绘图模块中文化进度

- `modules/data_plotting.ps1` 已完成第一轮中文化：
  - 图表中心入口窗口已中文化
  - 单图设计器主要标题、字段标签、按钮、状态提示已中文化
  - 组合图 / 四图拼接窗口的主要标题与导出按钮已中文化
- 当前仍保留少量内部英文项：
  - 部分图表类型下拉项仍显示 `Column / Line / Area ...` 这类内部 key
  - 配色方案名仍显示 `Ocean / Warm / Forest ...`
- 本轮为界面文案调整，未改动绘图数据逻辑；模块语法检查已通过：`PLOTTING_PARSE_OK`

## 2026-04-16 绘图模块中文化补充

- 组合图与四图拼接页面中的图表类型下拉框也已中文化。
- 当前采用“显示标签 + 内部 key”兼容方式，例如 `Column|柱状图`：
  - 好处是无需改动底层图表类型传参逻辑
  - 风险较低，便于继续迭代
- 模块语法检查在本轮后仍通过：`PLOTTING_PARSE_OK`

## 2026-04-16 绘图模块中文化补充（二）

- 绘图模块的图表类型与配色方案下拉框现已支持“纯中文显示、内部自动映射”模式。
- 当前用户界面会直接显示：
  - 图表类型：`柱状图 / 折线图 / 饼图 / 面积图 ...`
  - 配色方案：`海洋蓝 / 暖阳橙 / 森林绿 / 珊瑚红 / 中性灰`
- 内部仍自动映射回原有 key（如 `Column`、`Ocean`），因此不需要改写底层图表绘制逻辑。
- 本轮后 `modules/data_plotting.ps1` 语法检查仍通过：`PLOTTING_PARSE_OK`

## 2026-04-16 GUI 联调进度

- 主界面的表达式快捷按钮与帮助快捷按钮现已补齐新线性代数函数：
  - `EIGVALS`
  - `EIGVECS`
  - `SVALS`
  - `CHOL`
  - `LU`
  - `QR`
- 数据导入窗口顶部按钮区现已支持随按钮换行自动增高，减少窗口变窄时的遮挡风险。
- 本轮后 `matrix_calculator.ps1` 语法检查仍通过：`PARSE_OK`

## 2026-04-16 GUI 联调补充（二）

- 表达式区函数快捷按钮面板现已支持按按钮数量和窗口宽度动态调整显示高度，超出部分继续滚动。
- 帮助区快捷按钮条现已支持鼠标滚轮横向滚动，不再只依赖底部滚动条拖动。
- 本轮后 `matrix_calculator.ps1` 语法检查仍通过：`PARSE_OK`

## 2026-04-16 绘图故障修复补充

- 已定位并修复一处导致“数据绘图”入口仍报错的真实根因：
  - 问题位置：`modules/data_profile.ps1`
  - 问题原因：`Get-ColumnProfile` 中错误调用了 4 参数版本的 `[Math]::Max(...)`
  - 影响范围：绘图入口在分析列类型时会直接异常，导致图表设计页面无法继续打开
- 当前修复后：
  - `modules/data_profile.ps1` 语法检查通过：`PROFILE_PARSE_OK`
  - 命令级验证确认：
    - 绘图前置列分析链路已不再因 `Math.Max` 异常中断
    - 中文数值列如 `1万 / 2.5万 / 1.1亿` 仍可被识别为数值列
- 仍建议下一步做一次真实 GUI 复测，确认图表设计器窗口在实际点击下已恢复可用。

## 2026-04-16 绘图数值字段修复补充

- 已修复图表设计器中“数值字段（可多选）”列表为空的问题。
- 原因是绘图模块此前只接受 `InferredType = Numeric` 的列，判定过于严格；只要列里夹少量异常值或占位符，就会整列被排除。
- 当前绘图候选规则已放宽为：
  - 完全数值列直接入选
  - 多数值可解析的列也可入选，条件为：
    - 非 `Boolean` / `DateTime`
    - 至少 2 个单元格可解析为数值
    - 数值可解析比例 >= 0.6
- 命令级回归已确认：
  - 含 `1万 / 2.5万 / 1.1亿` 的列可继续被识别为数值字段
  - 含少量 `-` 的多数值列也能进入绘图候选列表

## 2026-04-16 绘图设计器补充说明

当前 `modules/data_plotting.ps1` 已补上以下能力：
- 图表设计器支持两种提取方式：
  - `按列提取`：横轴来自分类字段，勾选的数值列分别生成系列
  - `按行提取`：勾选的数据行分别生成系列，横轴可使用数值列名或列序号
- 数值字段候选改为直接扫描真实表格值：
  - 支持中文数值如 `1万 / 2.5万 / 1.1亿`
  - 支持夹少量 `-` 的多数值列
- 图表设计器左侧配置区已改为可滚动：
  - `导出图表` 按钮不会再被底部遮挡
  - 状态提示也会完整可见
- 预览逻辑已补强：
  - 若当前选择下没有可绘制数据，会明确提示 `当前选择下没有可绘制的数据点。`
  - 生成成功时会显示本次预览的数据点数量

当前已做过命令级验证：
- 候选检测：`NUM=播放量|点赞量`
- 行候选检测：`ROWS=第1行|第2行|第3行`
- 按列预览、按行预览均已能正常生成数据点

说明：
- 本轮中途曾用 `MatrixWorkspace.exe` 内嵌资源恢复 `modules/data_plotting.ps1` 的干净基线，然后在该基线上重新合并本轮修复。
- 组合图与四图拼接目前仍主要按“按列提取”设计；若要进一步扩展到按行提取，需要继续补 `Update-ComboChartPreview` 与四图编辑器逻辑。

## 2026-04-16 打包说明

当前项目已完成一次基于现有代码状态的 `1.1.3` 封装，输出文件为：
- `矩阵工作台1.1.3.exe`

本次封装通过 `exe_build_tools\一键生成exe.cmd` 完成，嵌入内容包括：
- `matrix_calculator.ps1`
- `modules\data_profile.ps1`
- `modules\data_plotting.ps1`

## 2026-04-16 数据窗口与绘图入口补充说明

- 已修复绘图设计器滚动区的 `Size` 构造错误，避免点击绘图后直接因 `Size` 重载异常而中断。
- 绘图中心窗口当前统一使用中文文案：`数据绘图`、`选择图表类型`、`组合图 / 拼接`。
- 数据导入窗口顶部按钮区已改为带自适应滚动的容器，`导出数据` 按钮在按钮换行时不应再被遮挡。
- 数据概览中的字段类型判断和数值统计，仍然是先通过 `Try-ParseDoubleValue` 对 `亿 / 万 / 千 / %` 等中文数量单位做标准化，再输出概览结果。
- 命令级验证：`1万 -> 10000`，`1.1亿 -> 110000000`，字段推断类型为 `Numeric`。

## 2026-04-16 桌宠集成与按钮区增强

已完成：
- 主程序工具栏新增 桌宠 与 桌宠设置 入口。
- 新增桌宠设置窗口，可配置：
  - DyberPet-main 路径
  - Python / 启动器路径
  - 启动方式：自动 / 优先EXE / 优先Python
  - 自动启动开关
- 主程序当前会优先尝试从本地 DyberPet-main 启动桌宠；若存在 un_DyberPet.exe 则可直接走 EXE，否则走 un_DyberPet.py。
- DyberPet 上游项目本身已包含扑鼠标、跟随鼠标、拖拽、点击等互动代码；当前主程序侧主要负责启动与设置接入。
- “数据导入与处理”窗口顶部按钮区继续加强了自适应高度，尽量避免 导出数据 按钮在初始打开时被遮挡。

当前验证：
- matrix_calculator.ps1：PARSE_OK
- modules/data_plotting.ps1：PLOTTING_PARSE_OK
- 本地桌宠资源检查：
  - DyberPet-main 存在
  - un_DyberPet.py 存在
  - 当前未发现 un_DyberPet.exe

后续封装注意：
- 如果要把桌宠随主程序一起封装，后续需要继续改打包链，把 DyberPet-main 的运行资源与启动方式一起纳入。
- 若本机缺少 PySide6 / PySide6-Fluent-Widgets / pynput / tendo，则 Python 方式启动桌宠会失败。

## 2026-04-16 绘图解压副本说明

已确认：
- 项目根目录中的 modules/data_plotting.ps1 已是新版中文绘图界面。
- 若仍看到英文 Data Plotting / Choose A Chart Type，通常是因为运行的是旧版 矩阵工作台1.1.3.exe，它会把旧资源重新解压到 %LocalAppData%\\MatrixWorkspaceApp。
- 本轮已将新版绘图模块同步覆盖到 %LocalAppData%\\MatrixWorkspaceApp\\modules\\data_plotting.ps1，用于当前环境的直接修正。

注意：
- 只要继续使用旧的 1.1.3.exe 重启程序，它仍可能再次覆盖成旧资源。
- 如需彻底一致，后续应基于当前代码重新封装新版本 EXE。





## 2026-04-16 界面自适应与高分屏说明

已完成：
- 当前主程序 DPI 感知已升级为：优先 Per-Monitor V2，失败时回退到旧式 DPI 感知。
- 新增统一窗体缩放默认设置：主要窗口和常用弹窗都会按 AutoScaleMode = Dpi 参与系统缩放。
- 主工具栏不再只依赖固定横向像素坐标，而是会根据可用宽度自动换行，并按文本宽度自适应按钮尺寸。
- 说明文字区域和工具栏整体高度也会自动随内容变化，减少不同电脑上出现的遮挡。

当前判断：
- 这套程序原先是“固定像素布局 + 局部自适应”的混合模式，因此在不同分辨率、不同字体缩放、不同 DPI 机器上确实容易出现遮挡。
- 本轮已经把主显示链路往“自动缩放优先”的方向推进了一大步，更适合发给别人直接运行。

验证结果：
- matrix_calculator.ps1：MAIN_PARSE_OK
- modules/data_profile.ps1：PROFILE_PARSE_OK
- modules/data_plotting.ps1：PLOTTING_PARSE_OK

后续建议：
- 若要让别人拿到的 EXE 也完整继承这轮改动，应基于当前代码重新封装新版本，而不是继续分发旧的 1.1.3.exe。

## 2026-04-16 启动器修复说明

已确认本次“启动器打不开”的根因不是 启动矩阵计算器.vbs 或 启动矩阵计算器.cmd 本身，而是主脚本 matrix_calculator.ps1 的主窗体初始化段在前一轮 DPI 改造时被误改坏。

已修复：
- 补回 Set-FormScaleDefaults
- 补回 $form = New-Object System.Windows.Forms.Form
- 启动链恢复后，以下方式均可正常拉起程序：
  - powershell -File matrix_calculator.ps1
  - 启动矩阵计算器.cmd
  - 启动矩阵计算器.vbs

说明：
- 启动矩阵计算器.vbs 采用隐藏窗口方式静默调用 PowerShell；因此只要主脚本报错，用户表面上就会看到“没反应”。
- 当前这条启动链已经恢复正常。

## 2026-04-16 启动器当前状态修正说明

当前以最新复测结果为准：
- `matrix_calculator.ps1`：可直接正常启动
- `启动矩阵计算器.cmd`：当前可用
- `启动矩阵计算器.vbs`：当前仍不稳定，未确认彻底修复

说明：
- 之前 README / 日志中存在“`启动矩阵计算器.vbs` 已恢复正常”的旧结论；本轮重新实测后，确认这条结论不应继续作为当前状态使用。
- 后续若用户急用，优先建议运行：
  - `启动矩阵计算器.cmd`
  - 或直接运行 `matrix_calculator.ps1`
- `vbs` 启动链后续需继续单独排查，重点关注：
  - `wscript.exe` 下的调用行为
  - 中文路径/编码
  - Shell.Run / ShellExecute 调用方式

## 2026-04-16 启动方式更新

当前项目已切换到更稳的启动方案：
- 新增 ASCII 启动器：`launch_matrix_workspace.cmd`
- `启动矩阵计算器.cmd` 现在会转调该 ASCII 启动器
- `启动矩阵计算器.vbs` 也已改为转调该 ASCII 启动器，而不是自己直接拼接 PowerShell 命令

当前推荐启动顺序：
1. `启动矩阵计算器.cmd`
2. `launch_matrix_workspace.cmd`
3. `matrix_calculator.ps1`

说明：
- `launch_matrix_workspace.cmd` 已实测可正常拉起“矩阵工作台”窗口。
- VBS 启动器虽然已经改走新的后端，但 `wscript.exe` 在当前环境下仍未拿到完全稳定的复测结果，因此暂不建议把 VBS 作为唯一入口。

## 2026-04-16 数据页面修复说明

已修复“数据导入与处理”页打不开的问题。

原因：
- 模块拆分后，多个弹窗函数头部漏掉了 `New-Object System.Windows.Forms.Form`，导致对 `Null` 对象设置 `Font / Controls / Show` 时直接报错。

已修复的窗口包括：
- 数据导入与处理
- 历史记录
- 特殊矩阵设置
- 缺失值清洗
- 遴选行列
- 修改数据
- 分组统计
- 导出数据

验证结果：
- `Show-DataWindow` 已可成功构造 `数据导入与处理` 窗口。
- 当前优先推荐从 `launch_matrix_workspace.cmd` 或 `启动矩阵计算器.cmd` 启动。

## 2026-04-17 运行修复补充（数据页 / 绘图 / 桌宠）

### 本轮根因排查
- 继续追查上两位 AI 留下的残留问题后，确认当前真正阻断使用的不是启动器本身，而是模块化后仍有若干窗口函数漏掉了 `New-Object System.Windows.Forms.Form`。
- 具体影响：
  - `modules/matrix_ui_helpers.ps1`
    - `Show-DesktopPetSettingsDialog`
    - `Show-TextViewerDialog`
    - `Show-DataTableDialog`
  - `modules/data_plotting.ps1`
    - `Show-ChartDesignerWindow`
    - `Show-ChartComposerWindow`
    - `Show-PlotCenterWindow`
- 另外还发现 `modules/matrix_ui_helpers.ps1` 中 `Get-DesktopPetRuntimeHint` 的异常返回写成了 `return .Exception.Message`，会导致桌宠设置页在异常分支下再次报错。
- 更深一层的问题是：`modules/data_profile.ps1` 已经被乱码污染，加载时会直接产生 PowerShell 解析错误，导致数据概览、缺失比例、分组统计等函数根本没有成功定义。

### 本轮修复
- `modules/matrix_ui_helpers.ps1`
  - 为桌宠设置、文本查看、数据表格三个窗口补回 Form 实例化。
  - 修复 `Get-DesktopPetRuntimeHint` 的异常返回写法，改为 `$_ .Exception.Message` 对应的正确形式：`$_.Exception.Message`。
- `modules/data_plotting.ps1`
  - 为图表设计器、组合图、数据绘图中心补回 Form 实例化。
- `modules/data_profile.ps1`
  - 放弃继续在乱码文件上做局部修补。
  - 直接从 `%LocalAppData%\MatrixWorkspaceApp\modules\data_profile.ps1` 回填一份可解析、内容完整的干净副本。
  - 再将 `Set-FormScaleDefaults -Form $dialog` 补回 `Show-MissingRatioDialog` 和 `Show-GroupStatsDialogEx`，保证这两个弹窗继续沿用当前项目的 DPI/缩放策略。
- `modules/matrix_dialogs.ps1`
  - 微调“数据导入与处理”页顶部布局：压缩顶部卡片和按钮区高度，移除 `AutoScrollMinSize` 那段多余滚动尺寸逻辑，减少顶部出现“像滚动条一样压住按钮文字”的视觉干扰。

### 桌宠启动问题处理
- 已确认 DyberPet 本地目录存在：`DyberPet-main`
- 已确认当前设置使用的 Python 为：
  - `C:\Users\23258\AppData\Local\Programs\Python\Python313\python.exe`
- 实际运行 `DyberPet-main\run_DyberPet.py` 后，先后补齐了以下依赖：
  - `PySide6`
  - `PySide6-Fluent-Widgets==1.5.4`
  - `pynput==1.7.6`
  - `tendo`
  - `apscheduler`
- 补齐后再次实际拉起桌宠进程，桌宠进程可持续运行至少 8 秒，不再启动即秒退。
- 当前仍会在标准错误里看到上游 DyberPet 自身的 Qt 警告：
  - `Qt::WindowType::Desktop has been deprecated`
  - `Shiboken::Conversions::_pythonToCppCopy ... NoneType`
- 上述两条目前属于上游项目运行告警，不影响“桌宠按钮能启动桌宠”这一目标。

### 本轮验证
- 真实 GUI 弹窗烟雾验证通过：
  - `Show-TextViewerDialog`
  - `Show-MissingRatioDialog`
  - `Show-GroupStatsDialogEx`
  - `Show-PlotCenterWindow`
  - `Show-DesktopPetSettingsDialog`
- 验证结果：`GUI_SMOKE_OK`
- 数据链路弹窗验证结果：`DATA_GUI_SMOKE_OK`
- 桌宠实际运行验证结果：`PET_LAUNCH_OK`（运行 8 秒后由测试脚本主动结束）

### 当前结论
- 现在“数据概览 / 缺失比例 / 分组统计 / 数据绘图打不开”的主因已经排除。
- 现在“桌宠设置打不开 / 桌宠按钮启动失败”的主因也已经排除。
- 启动方式仍优先推荐：
  1. `启动矩阵计算器.cmd`
  2. `launch_matrix_workspace.cmd`
  3. `matrix_calculator.ps1`
- `vbs` 仍不建议作为唯一可靠入口。

### 后续可继续优化的点
- 数据导入页顶部布局虽然已收一轮，但如果后续还要继续兼容更多 DPI/更窄窗口，最好把顶部区域进一步改成 `TableLayoutPanel` / 纯自动布局，减少依赖固定 Point 坐标。
- 桌宠若后续要随主程序一同封装进 EXE，仍需继续改封装链，把 DyberPet 资源目录与 Python/EXE 启动策略一起纳入。

## 2026-04-17 数据预览页布局再收口

### 本轮修改
- `modules/matrix_dialogs.ps1`
  - 继续压缩“数据导入与处理”页顶部区域高度，给下方数据预览区腾出更多可视空间。
  - 将顶部状态区、按钮区和外层卡片高度再收紧一轮，减少不同 DPI / 缩放下表格最后几行被截断的概率。
  - 移除了顶部布局里残留的 `AutoScrollMinSize` 干扰项，避免出现类似滚动条的视觉噪音。

### 本轮验证
- `modules/matrix_dialogs.ps1` 通过 PowerShell 语法解析检查。

### 当前状态
- 本轮只做了页面布局收口，没有继续打包。
- 后续如果要发版，可再基于当前代码直接封装 `矩阵工作台1.1.4`。
- 当前推荐启动方式仍是：
  1. `启动矩阵计算器.cmd`
  2. `launch_matrix_workspace.cmd`
  3. `matrix_calculator.ps1`

## 2026-04-17 数据页比例布局与绘图 BeginInvoke 修复

### 本轮目标
- 继续修复“数据导入与处理”页面里，数据预览区看起来被顶部工作栏遮挡的问题。
- 在不影响其它功能的前提下，修复绘图模块中 `BeginInvoke` 重载报错。

### 本轮修改
- `modules/matrix_dialogs.ps1`
  - 将“数据导入与处理”页根布局由普通 `Panel + Top/Fill Dock` 调整为 `TableLayoutPanel` 两行结构：
    - 第 1 行固定给顶部状态/工具栏区。
    - 第 2 行百分比填充给数据预览区。
  - 顶部工作栏与下方数据预览现在物理分层，不再共享同一块父面板的叠放关系。
  - 在顶部布局刷新时，同步更新 `TableLayoutPanel.RowStyles[0].Height`，让顶部区高度跟按钮换行后的真实高度一致。
- `modules/matrix_ui_helpers.ps1`
  - 在 `Set-CurrentDataTable` 里增加表格视图重置：
    - 绑定数据后清除选中
    - 将 `CurrentCell` 置空
    - 横向滚动位置重置到第 0 列
    - 纵向滚动位置重置到第 0 行
  - 这样导入新数据后，预览页会尽量从表头和第一行开始显示，减少“看起来像被上方挡住”的错觉。
- `modules/data_plotting.ps1`
  - 将绘图设计器里原来的 `BeginInvoke($refreshPreview)` 改为标准 WinForms 委托调用：
    - 新增 `$refreshPreviewInvoker = [System.Windows.Forms.MethodInvoker]{ & $refreshPreview }`
    - `Mode` 切换和 `CheckedListBox.ItemCheck` 触发时，改为 `BeginInvoke($refreshPreviewInvoker)`
  - 这样避免命中 `BeginInvoke` 的错误重载，从而修复用户截图中“找不到 BeginInvoke 的重载，参数计数为 1”的异常。

### 本轮验证
- `modules/matrix_dialogs.ps1`、`modules/matrix_ui_helpers.ps1`、`modules/data_plotting.ps1` 均通过 PowerShell 语法解析检查。
- 数据页结构验证结果：
  - `DATA_LAYOUT=System.Windows.Forms.TableLayoutPanel`
  - `DATA_ROWS=2`
  - `DATA_TOPHEIGHT=188`
  - `GRID_FIRST_ROW=0`
  - `GRID_FIRST_COL=0`
- 上述结果说明：数据页已经切成“上工具栏、下预览区”两行布局，且导入数据后预览会从首行首列开始显示。

### 当前结论
- 数据预览区现在是独立占用工作栏下方空间，不再依赖原先的 Dock 叠放关系。
- 绘图模块里的 `BeginInvoke` 重载错误已从代码层修正为标准委托调用。
- 若后续继续微调，可优先在真实 GUI 中观察：
  - 数据页顶部按钮在极窄窗口或高 DPI 下是否还需要进一步换行压缩。
  - 绘图设计器在切换“按行提取 / 按列提取”与勾选数值字段时，预览是否还有别的运行时异常。

## 桌宠专项研究

### 当前接入状态
- 当前矩阵工作台里的桌宠功能仍是“外部启动 + 本地设置”模式：
  - `桌宠` 按钮负责启动 DyberPet。
  - `桌宠设置` 负责路径、启动方式、自动启动等配置。
- 这一层已经能工作，但还没有把上游 DyberPet 的角色管理、控制面板、仪表盘完整接进来。

### 已确认的上游能力
- 角色选择 / 切换：
  - `DyberPet-main/run_DyberPet.py`
  - `DyberPet-main/DyberPet/DyberSettings/CharCardUI.py`
- 宠物 / 子宠物管理：
  - `DyberPet-main/DyberPet/DyberSettings/PetCardUI.py`
- 物品 / 模块管理：
  - `DyberPet-main/DyberPet/DyberSettings/ItemCardUI.py`
- 控制面板 / 仪表盘：
  - `DyberPet-main/DyberPet/DyberSettings/DyberControlPanel.py`
  - `DyberPet-main/DyberPet/Dashboard/DashboardUI.py`
- 鼠标互动、拖拽、跟随、随机移动、掉落：
  - `DyberPet-main/DyberPet/DyberPet.py`
  - `DyberPet-main/DyberPet/Accessory.py`
  - `DyberPet-main/DyberPet/conf.py`

### 研究结论
- 可以加入矩阵工作台与桌宠的互动逻辑。
- 可以加入角色选择、宠物管理、控制面板、仪表盘等上游已有能力。
- 可以实现桌宠自己在屏幕里移动，这部分 DyberPet 上游已具备。
- 但最稳妥的接法不是把 PySide6 页面直接硬嵌进当前 WinForms 主窗体，而是：
  - 矩阵工作台负责控制入口与状态桥接。
  - DyberPet 继续作为独立桌宠进程运行。

### 推荐实施顺序
1. 先做“桌宠控制中心”页面。
2. 再做矩阵工作台 -> 桌宠的状态桥接。
3. 先接低风险互动：计算成功、计算失败、数据导入成功。
4. 再接角色/宠物选择、控制面板、仪表盘。
5. 最后统一处理打包。

### 专项日志
- 已新增：`PET_WORK_LOG.md`
- 后续桌宠专项进展优先同步到这份日志，避免影响主日志可读性。

## 1.1.4 发布方式（方案 B）

### 发布策略
- 采用“矩阵核心内嵌 EXE，桌宠运行目录外置”的发布方式。
- 这样可以：
  - 尽量不把矩阵工作台核心源码直接发出去。
  - 同时保留桌宠功能。
  - 后续桌宠升级时，只需要替换 `DyberPet-main` 文件夹。

### 本轮封装优化
- `MatrixWorkspaceLauncher.cs`
  - 去掉了最敏感的 `-ExecutionPolicy Bypass` 和 `-WindowStyle Hidden` 启动参数。
  - 改为 `-STA -NoProfile -File`，尽量降低误报风险。
- `matrix_calculator.ps1`
  - 支持通过 `MATRIX_WORKSPACE_ROOT` 识别真实发布目录。
  - 即使矩阵主程序从 EXE 内嵌资源解压运行，桌宠默认目录仍可指向 EXE 同级 `DyberPet-main`。
- `launch_matrix_workspace.cmd`
  - 现在优先启动 `MatrixWorkspace.exe`，只有找不到 EXE 时才回退到脚本直启。

### 当前推荐分发物
- `release\matrix_workspace_1.1.4_bundle.zip`

### 完整包包含内容
- `MatrixWorkspace.exe`
- `矩阵工作台1.1.4.exe`
- `launch_matrix_workspace.cmd`
- `launch_matrix_workspace.vbs`
- `DyberPet-main` 运行目录子集：
  - `run_DyberPet.py`
  - `DyberPet`
  - `res`
  - `data`
  - `LICENSE`

### 不再包含
- 矩阵主程序 `.ps1`
- `modules` 目录
- 主开发日志与说明文档
- 其它项目源文件

### 当前结论
- 这个方案比“直接发源码型完整包”更适合当前需求。
- 桌宠后续更新也更方便，主要替换自己的文件夹即可。

## 2026-04-17 启动兼容性说明

### 为什么先回退启动参数
- 1.1.4 之前为了尽量减少“病毒风险”提示，把启动器参数从：
  - `-ExecutionPolicy Bypass -WindowStyle Hidden`
  - 调整成了更保守的 `-STA -NoProfile -File`
- 实际反馈显示：
  - 在部分设备上，EXE 和启动器会直接无法打开。
  - 根因更接近 PowerShell 执行策略或启动环境差异，而不是主程序逻辑损坏。

### 当前已采取的策略
- 已把以下两条启动链恢复成兼容优先：
  - `MatrixWorkspaceLauncher.cs`
  - `launch_matrix_workspace.cmd`
- 当前实际启动参数恢复为：
  - `powershell.exe -STA -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File ...`

### 当前取舍
- 现在的优先级是：
  - 先保证别的电脑能正常打开。
  - 再考虑怎么降低误报。
- 这意味着：
  - 误报风险可能比“保守参数版”更高。
  - 但跨设备启动成功率会更好。

### 当前建议
- 若对外分发，请优先使用最新重新生成后的：
  - `release\matrix_workspace_1.1.4_bundle.zip`
- 如果后续还要继续降误报，建议改造方向是：
  - 逐步减少对 PowerShell 启动链的依赖。
  - 长期目标改成更原生的 EXE 主程序。

## 桌宠透明框状态

- 已对桌宠外层白色/透明矩形框做过一轮“仅视觉隐藏”的保守修复。
- 当前策略是：
  - 优先缩小桌宠窗体的可见留白。
  - 默认隐藏未使用的顶部状态条占位。
  - 暂不做像素级命中区裁切，避免影响拖拽和鼠标互动。

## 作图功能当前状态

- 已修正单图设计器中一处仍残留的 `BeginInvoke` 句柄时机问题。
- 当前作图模块在切换“按列提取 / 按行提取”或勾选数值字段时，不再依赖窗口句柄已创建后才能调用的旧刷新方式。
- 单图设计器页面已从固定像素排版重构为：
  - 左侧属性栏独立滚动区
  - 右侧预览独立显示区
  - 分区宽度按窗口大小动态调整
- 已补充修复 `SplitContainer` 最小宽度约束在窗口早期创建阶段触发的异常。
- 已补充降低图表预览高频刷新导致的 `GDI+ 一般性错误` 风险。
- 当前建议：
  - 在 GUI 中重新测试图表设计器是否已能正常打开并更新预览。

## 外部作图工具桥接

- “数据绘图”中心已新增两个外部工具按钮：
  - `ScottPlot`
  - `LiveCharts2`
- 当前桥接方式是：
  - 先将矩阵工作台当前 `DataTable` 导出为 CSV/JSON
  - 再尝试打开对应工具的本地 WinForms 示例项目

### 导出位置
- `%LocalAppData%\MatrixWorkspaceApp\ExternalPlotBridge\ScottPlot\current_data.csv`
- `%LocalAppData%\MatrixWorkspaceApp\ExternalPlotBridge\LiveCharts2\current_data.csv`

### 仓库内桥接位置
- `ScottPlot-main\MatrixWorkspaceBridge\current_data.csv`
- `LiveCharts2-master\MatrixWorkspaceBridge\current_data.csv`

### 当前入口定位
- ScottPlot：
  - `ScottPlot-main\src\ScottPlot5\ScottPlot5 Demos\ScottPlot5 WinForms Demo\WinForms Demo.csproj`
- LiveCharts2：
  - `LiveCharts2-master\samples\WinFormsSample\WinFormsSample.csproj`

### 当前限制
- 当前开发机只有 `.NET runtime`，没有 `.NET SDK`
- 也没有 `devenv / msbuild`
- 因此当前按钮已经能完成：
  - 数据导出
  - 工具入口定位
  - 项目/目录唤起
- 但若要在当前机器上直接把上游示例项目编译并弹出最终独立 GUI，还需要：
  - 安装 `.NET SDK`
  - 或进一步补专门的桥接示例项目

## 2026-04-17 更新：外部作图桥接页已落地

- 当前仓库内已经安装本地 `.NET SDK`：
  - `C:\Users\23258\Desktop\矩阵计算器\.dotnet`
- 当前不再只是“定位示例项目”，而是已经补了两个真正可直接打开的桥接页：
  - `external_tools\ScottPlotBridge`
  - `external_tools\LiveCharts2Bridge`

### 当前按钮行为
- 在 `数据绘图` 主页面点击：
  - `ScottPlot`
  - `LiveCharts2`
- 会执行以下流程：
  1. 导出当前导入/处理后的 `DataTable`
  2. 写入桥接 CSV / JSON
  3. 若桥接 EXE 不存在，则用仓库内 `.dotnet` 自动编译
  4. 直接打开各自独立的桥接窗口

### 当前桥接 EXE
- ScottPlot：
  - `external_tools\ScottPlotBridge\bin\Release\net10.0-windows\ScottPlotBridge.exe`
- LiveCharts2：
  - `external_tools\LiveCharts2Bridge\bin\Release\net10.0-windows\LiveCharts2Bridge.exe`

### 桥接页能力
- 选择图表类型
- 选择横轴 / 分类字段
- 多选数值字段
- 实时预览
- 导出 PNG

### 当前说明
- 现在这两个按钮已经指向“可直接用的独立页面”，不再是只打开上游源码项目。
- 本地 SDK 与桥接项目都在仓库内，后续迁移、重建和继续调整更稳定。

## 当前最新分发包

- 最新可直接发给别人的运行包是：
  - `C:\Users\23258\Desktop\矩阵计算器\release\matrix_workspace_1.1.4_bundle.zip`
- 包内当前包含：
  - `MatrixWorkspace.exe`
  - `矩阵工作台1.1.4.exe`
  - `launch_matrix_workspace.cmd`
  - `launch_matrix_workspace.vbs`
  - `DyberPet-main` 运行目录
- 当前建议：
  - 优先发整个 zip，不要只发单个 EXE。

## 2026-04-17 更新：外部作图按钮 and 报错已修复

- 现象：点击 `ScottPlot` / `LiveCharts2` 时会提示：
  - `找不到与参数名称“and”匹配的参数。`
- 根因：
  - 外部桥接编译检查函数里，把 `-and` 写进了 `Test-Path` 调用。
- 修复后：
  - 外部桥接按钮会正确检查 EXE 是否已存在，再决定是否编译和启动。
- 当前验证：
  - `BRIDGE_OK=ScottPlot`
  - `BRIDGE_OK=LiveCharts2`

## 2026-04-17 更新：外部作图桥接已改为自带运行时发布

- 之前 `ScottPlotBridge.exe` / `LiveCharts2Bridge.exe` 仍可能提示安装 `.NET`。
- 现在这两个桥接页已经改成：
  - `win-x64`
  - `self-contained`
  - 可随项目一起打包的发布目录

### 当前发布目录
- ScottPlot：
  - `external_tools\ScottPlotBridge\publish\win-x64`
- LiveCharts2：
  - `external_tools\LiveCharts2Bridge\publish\win-x64`

### 当前按钮逻辑
- `数据绘图` 页面中的：
  - `ScottPlot`
  - `LiveCharts2`
- 现在会优先检查并使用上述发布目录中的 EXE。
- 若发布目录不存在，会自动用仓库内 `.dotnet` 重新发布。

### 当前验证
- 发布版桥接 EXE 已实际拉起成功：
  - `SCOTT_PUBLISHED_LAUNCHED=True`
  - `LIVE_PUBLISHED_LAUNCHED=True`

### 打包建议
- 后续若要一起打包外部作图功能，请把以下目录一并带上：
  - `external_tools\ScottPlotBridge\publish\win-x64`
  - `external_tools\LiveCharts2Bridge\publish\win-x64`

## 2026-04-17 更新：桥接页面布局与官方页面入口

- `ScottPlotBridge` / `LiveCharts2Bridge` 当前都已改成：
  - 左侧属性栏
  - 右侧预览区
  - 左右独立分区
  - 左侧独立滚动
  - 更适合高 DPI / 窄窗口的自适应布局

### 当前桥接页能力
- 左侧支持：
  - `CSV 文件`
  - `图表类型`
  - `横轴 / 分类字段`
  - `数值字段（可多选）`
  - `主标题`
  - `刷新图表`
  - `导出 PNG`
  - `打开官方页面`

### 官方页面入口
- ScottPlot：
  - `ScottPlot-main\src\ScottPlot5\ScottPlot5 Demos\ScottPlot5 WinForms Demo\WinForms Demo.csproj`
- LiveCharts2：
  - `LiveCharts2-master\samples\WinFormsSample\WinFormsSample.csproj`

### 当前说明
- 桥接页仍然是主入口，因为它已经接好了矩阵工作台导出的当前数据。
- 如果需要看官方示例或在官方页面中重新导入数据，可以直接点桥接页里的：
  - `打开官方页面`

## 2026-04-17 更新：官方页面按钮已区分为“运行”和“打开项目”

- 之前点“打开官方页面”会直接打开 `.csproj` 文件。
- 如果系统默认把 `.csproj` 交给 VS Code，就会表现成“打开源码文件”，不是弹出示例窗口。

### 当前桥接页按钮
- `运行官方页面`
  - 先构建官方 sample
  - 再直接启动官方示例窗口
- `打开官方项目`
  - 仅打开对应 `.csproj` / 项目目录

### 已接入的官方入口
- ScottPlot：
  - `ScottPlot-main\src\ScottPlot5\ScottPlot5 Demos\ScottPlot5 WinForms Demo\WinForms Demo.csproj`
- LiveCharts2：
  - `LiveCharts2-master\samples\WinFormsSample\WinFormsSample.csproj`

## 2026-04-17 更新：官方按钮现在区分“运行窗口”和“查看项目”

- 桥接页现在保留两类官方入口：
  - `运行官方页面`
  - `打开官方项目`

### 运行官方页面
- 不再直接启动官方 sample 的 `.exe`。
- 当前行为是：
  1. 使用仓库内本地 `.dotnet\dotnet.exe` 构建官方示例项目
  2. 查找 sample 输出的 `.dll`
  3. 继续使用仓库内本地 `.dotnet\dotnet.exe "<sample.dll>"` 直接运行官方示例窗口
- 这样做的好处是：
  - 不依赖目标电脑是否安装全局 `.NET Desktop Runtime`
  - 不会再因为官方 sample 的框架依赖型 `.exe` 弹出“需要安装或更新 .NET”

### 打开官方项目
- 仍然只用于：
  - 打开 `.csproj`
  - 或在项目文件缺失时打开对应目录
- 适合查看源码结构，不等同于运行官方示例窗口。

### 桥接页按钮布局
- 当前桥接页按钮已改为两排：
  - 第一排：`刷新图表`、`导出 PNG`
  - 第二排：`运行官方页面`、`打开官方项目`
- 第二排按钮已加宽，文本显示更完整。

## 2026-04-18 更新：桥接页设计项增强与词云图补充

### 当前桥接页新增设计项
- `ScottPlotBridge` / `LiveCharts2Bridge` 当前都已补充：
  - `副标题`
  - `X 轴标题`
  - `Y 轴标题`
  - `配色方案`
  - `图例位置`
  - `显示网格`
  - `线宽`
  - `点大小`
  - `圆环内径 %`

### 当前 ScottPlot 桥接页图型
- `柱状图`
- `条形图`
- `折线图`
- `阶梯线`
- `面积图`
- `样条图`
- `样条面积图`
- `散点图`
- `饼图`
- `圆环图`
- `雷达图`

### 当前 LiveCharts2 桥接页图型
- `柱状图`
- `条形图`
- `折线图`
- `面积图`
- `阶梯线`
- `阶梯面积图`
- `堆叠柱状图`
- `堆叠条形图`
- `堆叠面积图`
- `散点图`
- `饼图`
- `圆环图`

### 当前基础作图补充
- 基础作图入口现已额外加载：
  - `modules\data_plotting_ext.ps1`
- 当前新增：
  - 更丰富的基础图型定义
  - 更丰富的调色板
  - `词云图` 独立设计器

### 当前词云图能力
- 可选分词模式：
  - `Auto`
  - `Whole Cell`
  - `Whitespace`
  - `Single Char`
  - `Chinese Bigram`
  - `Chinese Trigram`
- 当前词云形状：
  - `Rectangle`
  - `Circle`
  - `Ellipse`
  - `Diamond`
  - `Hexagon`
  - `Triangle`
  - `Star`
  - `Heart`
  - `Cloud`
  - `Leaf`

### 当前验证
- `modules\data_plotting_ext.ps1` 已可正常加载。
- `ScottPlotBridge` / `LiveCharts2Bridge` 已重新 `build + publish` 到：
  - `external_tools\ScottPlotBridge\publish\win-x64`
  - `external_tools\LiveCharts2Bridge\publish\win-x64`
