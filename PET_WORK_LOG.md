# 桌宠功能专项日志

## 2026-04-17 桌宠能力研究与集成建议

### 本轮范围
- 仅研究 DyberPet 与“矩阵工作台”的可集成能力、风险边界与推荐接法。
- 不改动矩阵计算、数据导入、绘图等既有主功能逻辑，避免引入新的主流程回归。

### 当前项目里已经接入的桌宠能力
- 主程序入口：
  - `modules/matrix_form.ps1`
    - 顶部已有 `桌宠`、`桌宠设置` 按钮。
- 当前桌宠集成层：
  - `modules/matrix_ui_helpers.ps1`
    - 已支持桌宠目录保存/读取。
    - 已支持 Python 或 EXE 两种启动方式。
    - 已支持自动启动、启动/停止、路径配置。
- 当前结论：
  - 现在的矩阵工作台对 DyberPet 的接入还是“外部启动器”模式。
  - 还没有把 DyberPet 上游的角色管理、控制面板、仪表盘、任务系统真正桥接进主程序。

### DyberPet 上游已经存在、可直接复用的功能

#### 1. 角色/宠物选择
- `DyberPet-main/run_DyberPet.py`
  - `self.conp.charCardInterface.change_pet.connect(self.p._change_pet)`
  - 说明：上游已支持角色切换，并且主宠对象已经接好切换信号。
- `DyberPet-main/DyberPet/DyberSettings/CharCardUI.py`
  - `change_pet = Signal(str, name='change_pet')`
  - 已有“角色管理”页面与切换逻辑。
- `DyberPet-main/DyberPet/DyberSettings/PetCardUI.py`
  - 已有“迷你宠物/子宠物”页面。
- `DyberPet-main/DyberPet/DyberSettings/ItemCardUI.py`
  - 已有物品/模块管理页面。

#### 2. 上游控制台与仪表盘
- `DyberPet-main/run_DyberPet.py`
  - 已实例化：
    - `ControlMainWindow()`
    - `DashboardMainWindow()`
- `DyberPet-main/DyberPet/DyberSettings/DyberControlPanel.py`
  - 已有完整系统面板。
- `DyberPet-main/DyberPet/Dashboard/DashboardUI.py`
  - 已有完整仪表盘。

#### 3. 鼠标互动、拖拽、跟随、掉落、随机动作
- `DyberPet-main/DyberPet/DyberPet.py`
  - `mousePressEvent / mouseMoveEvent / mouseReleaseEvent`
  - 已支持点击、拖拽、松手后掉落、拍打互动等。
- `DyberPet-main/DyberPet/Accessory.py`
  - 已包含 `follow_mouse`、`mousedrag`、`Listener(on_move / on_click)` 等逻辑。
- `DyberPet-main/DyberPet/conf.py`
  - 已支持：
    - `random_act`
    - `need_move`
    - `frame_move`
    - `follow_mouse`
    - `mouseDecor`
- 当前结论：
  - “桌宠自己在屏幕里移动”
  - “扑鼠标 / 被鼠标拖拽 / 跟随鼠标”
  - 这些上游本身就已经具备，不需要从零重新写。

### 与矩阵工作台互动的可行方向

#### A. 可安全加入的互动逻辑
- 计算成功时：
  - 桌宠弹出鼓励气泡，如“算出来啦”“这个矩阵搞定了”。
- 计算失败时：
  - 桌宠弹出提醒气泡，如“维度不匹配哦”“这里要再检查一下”。
- 长时间无操作时：
  - 桌宠随机巡游或弹出提醒。
- 数据导入成功时：
  - 桌宠切换成“搬运/整理数据”类提示。
- 进行 `DET / INV / EIGVALS / QR` 等较重计算时：
  - 桌宠可切换为“思考中/工作中”状态。

#### B. 推荐的桥接方式
- 推荐使用“轻量文件/状态桥接”，不要直接把 WinForms 和 PySide6 强绑在一个 UI 事件循环里。
- 推荐新增：
  - `%LocalAppData%\\MatrixWorkspaceApp\\pet_bridge_state.json`
  - `%LocalAppData%\\MatrixWorkspaceApp\\pet_bridge_events.json`
- 由矩阵工作台写入：
  - 当前任务类型
  - 最近一次计算状态
  - 最近一次导入数据状态
  - 可显示的气泡文本
- 由 DyberPet 侧单独读取并转成气泡/动作。

### 关于“把上游页面内置进矩阵工作台”的结论

#### 能做，但不建议直接硬嵌 WinForms 主窗体
- 原因：
  - 当前矩阵工作台主体是 PowerShell + WinForms。
  - DyberPet 上游主体是 Python + PySide6 + qfluentwidgets。
  - 两边 UI 框架、消息循环、资源加载方式都不同。
- 风险：
  - 直接把 PySide6 页面“嵌”进 WinForms 主窗体，容易带来：
    - 焦点问题
    - DPI/缩放问题
    - 窗口句柄管理问题
    - 打包复杂度暴涨
    - 后续维护困难

#### 更稳的方案
- 第一阶段：
  - 在矩阵工作台里增加“桌宠控制中心”页面。
  - 这个页面只负责：
    - 角色/宠物列表展示
    - 启动/停止桌宠
    - 打开上游控制面板
    - 打开上游仪表盘
    - 设置互动模式
    - 设置桥接开关
- 第二阶段：
  - 新增一个独立 Python 桥接脚本，例如：
    - `pet_bridge.py`
  - 由它负责：
    - 监听矩阵工作台状态文件
    - 给 DyberPet 派发气泡、动作、切换提示
- 第三阶段：
  - 再考虑把“角色选择”等能力做成本地镜像页面，或直接打开上游原生窗口。

### 打包层面的结论
- 若后续要随矩阵工作台一起打包，桌宠至少需要一起带上：
  - `DyberPet-main\\res`
  - `DyberPet-main\\DyberPet`
  - `run_DyberPet.py` 或 `run_DyberPet.exe`
- 如果仍走 Python 启动：
  - 还要解决目标机器无 Python 或依赖缺失的问题。
- 更稳的发布思路：
  - 要么给 DyberPet 预编译 EXE 再随主程序一起发布。
  - 要么主安装包里一起带 Python 运行时与依赖。

### 推荐的下一步实施顺序
1. 新建“桌宠控制中心”而不是直接把 PySide6 页面嵌进 WinForms。
2. 先做矩阵工作台 -> 桌宠 的状态桥接文件。
3. 先接 3 类低风险互动：
   - 计算成功
   - 计算失败
   - 数据导入成功
4. 再接：
   - 角色/宠物选择
   - 打开控制面板
   - 打开仪表盘
5. 最后再处理统一打包。

### 当前专项结论
- 可以加入矩阵工作台与桌宠的互动逻辑。
- 可以加入上游已有的角色选择、宠物管理、控制面板等功能。
- 可以实现桌宠自己在屏幕里移动，这部分 DyberPet 上游已具备。
- 但最稳妥的接法不是“强行嵌入同一个窗体”，而是：
  - 矩阵工作台负责控制与状态桥接。
  - DyberPet 继续作为独立桌宠进程运行。
- 这样最符合“不要影响其它功能”的要求。

## 2026-04-17 桌宠发布策略补充（方案 B）

### 本轮结论
- 当前发布采用“矩阵核心内嵌 EXE + 桌宠目录外置”的方案。
- 优点：
  - 矩阵工作台主逻辑不再以 `.ps1` / `modules` 形式一起发出去。
  - 桌宠功能仍可保留。
  - 后续若只更新桌宠，主要替换 `DyberPet-main` 文件夹即可。
- 当前运行包里桌宠部分仍属于上游运行源码目录，不是单独编译好的桌宠 EXE。
- 因此：
  - “矩阵核心”暴露显著减少。
  - “桌宠上游代码”仍随运行目录存在，这是当前方案 B 的代价。

### 当前发布包
- `release\matrix_workspace_1.1.4_bundle.zip`

## 2026-04-17 桌宠透明框视觉隐藏（保守方案）

### 本轮目标
- 仅处理桌宠外围白色/透明矩形框的视觉问题。
- 不做像素级命中区裁切。
- 不改拖拽、扑鼠标、跟随鼠标等互动逻辑。

### 问题判断
- 结合上游代码与当前默认宠物 `Kitty` 素材检查，当前白框更像是：
  - 桌宠主窗体尺寸按 `pet_conf.json` 的最大配置尺寸保留。
  - 实际显示帧 `stand_0.png` 的可见内容比配置画布小。
  - 顶部 `status_frame` 默认也参与布局占位，即使平时没有显示番茄钟/专注条。
- 当前默认站立帧素材检查结果：
  - 图片大小：`72 x 64`
  - 非透明可见区域：`0,7 -> 67,63`
- 这说明外围“白框感”主要不是猫图本身大面积白底，而是窗体和布局留白过多。

### 本轮修改
- 修改文件：
  - `DyberPet-main/DyberPet/DyberPet.py`
- 处理内容：
  - 为桌宠主窗体补充更明确的透明属性：
    - `Qt.WA_NoSystemBackground`
    - `background: transparent; border: none;`
  - 为 `label`、`status_frame`、根布局补充透明样式。
  - `status_frame` 默认隐藏，避免未启用番茄钟/专注模式时仍占据上方空白。
  - `reset_size()` 改为优先按当前实际显示帧 `settings.current_img` 的尺寸计算窗体大小，而不是始终按 `pet_conf.width/height` 的配置最大画布。
  - 番茄钟/专注模式启用时再显示 `status_frame` 并重新计算大小；关闭时自动隐藏并回收占位。

### 设计取舍
- 当前方案只做“视觉隐藏”，不做命中区轮廓裁切。
- 好处：
  - 风险更低。
  - 不容易破坏拖拽和鼠标互动。
  - 不需要为每个角色单独做像素级遮罩。
- 仍可能保留的特征：
  - 点击区可能依然略大于猫图本身，但白色矩形感会明显减弱。

### 本轮验证说明
- 当前环境下 `python` / `py` 命令均不在 PATH 中，未能直接做本机 `py_compile`。
- 但本轮改动仅集中在 `DyberPet.py` 的窗体透明样式、状态条可见性和尺寸计算逻辑，没有改动上游动画/交互线程结构。

### 下一步建议
1. 先实际观察默认 `Kitty` 角色的白框是否已明显减弱。
2. 再切换其他角色，确认是否存在个别角色仍有明显留白。
3. 若后续还要继续收边，优先做“保留安全点击区的轻量裁边”，暂不建议直接上像素级 mask。

## 2026-04-18 记录补充

### 本轮说明
- 本轮继续完成的是：
  - `数据绘图`
  - `ScottPlot / LiveCharts2 桥接页`
  - `词云图`
- 当前没有修改桌宠相关文件。
- 本轮仍保持：
  - 不触碰 `DyberPet-main`
  - 不影响桌宠透明框、互动、控制面板与发布策略相关实现

### 当前结论
- 当前桌宠链路与本轮外部作图增强任务保持隔离。
- 若后续继续做桌宠功能，仍以上面各节记录为准。
