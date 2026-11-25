# 城市桥梁技术状况评定系统

基于 Python、PySide6 和 openpyxl 开发的桌面应用，用于录入桥梁构件损坏情况、计算技术状况指标、保存评定项目和导出 Excel 报告。

项目作者：Mingtao Zhang。

## 功能

- 按桥面系、上部结构、下部结构录入构件与损坏信息。
- 根据损坏程度计算扣分值、损坏权重、综合扣分值（MDP）、桥梁状况指数（BCI）和桥梁结构指数（BSI）。
- 支持多跨录入、构件存在状态选择和权重调整。
- 将项目保存为 JSON，加载历史项目，并导出 Excel 评定报告。
- 提供桥梁类型选择主界面和内置帮助文档。

程序内置帮助以《城市桥梁养护技术标准》CJJ 99—2017 为评定方法参考。具体计算规则和参数见各桥型模块及 `basic_data/` 配置。

## 支持的桥梁类型

| 桥梁类型 | 评定模块 |
| --- | --- |
| 混凝土梁桥 | `Concrete_beam_bridge.py` |
| 钢结构梁桥 | `Steel_structure_beam_bridge.py` |
| 混凝土刚构桥 | `Concrete_rigid_frame_bridge.py` |
| 实腹式拱桥 | `Solid_belly_arch_bridge.py` |
| 空腹式拱桥 | `Hollow_belly_arch_bridge.py` |
| 钢结构人行天桥 | `Steel_structure_pedestrian_overpass.py` |
| 混凝土人行天桥 | `Concrete_pedestrian_overpass.py` |

## 安装与启动

需要 Python 3 和可用的桌面图形环境。发布前已在 Python 3.12 下完成源码语法与 JSON 配置检查。

```powershell
git clone https://github.com/ZXLCQU/Urban_bridge_assessment.git
cd Urban_bridge_assessment
python -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
.\.venv\Scripts\python.exe bridge_assessment_main.py
```

上述命令适用于 Windows PowerShell，无需激活虚拟环境。在其他系统上，请使用虚拟环境对应的 Python 路径，例如 `.venv/bin/python`；这些平台的界面行为尚未验证。

在主界面选择桥梁类型后，程序使用同一个 Python 解释器启动对应的评定模块。也可以直接运行某个模块，例如：

```powershell
.\.venv\Scripts\python.exe Concrete_beam_bridge.py
```

## 使用流程

1. 启动主程序并选择桥梁类型。
2. 选择桥面系、上部结构或下部结构，录入损坏情况和程度。
3. 根据实际桥梁设置构件是否存在；多跨桥梁可逐跨录入。
4. 查看计算结果，并使用“保存项目”保存评定数据。
5. 使用“加载项目”继续编辑，或使用“导出 Excel”生成报告。

## 目录说明

```text
Urban_bridge_assessment/
├── bridge_assessment_main.py          # 主界面和内置帮助
├── *_bridge.py                       # 梁桥、刚构桥、拱桥评定模块
├── *_overpass.py                     # 人行天桥评定模块
├── basic_data/                       # 各桥型的评定基础配置
├── UrbanBridge.ico                   # 应用图标
├── requirements.txt                  # 第三方运行依赖
└── LICENSE                           # MIT 许可证
```

运行时，程序会按需创建以下目录：

- `project_data/<桥型>/`：保存的评定项目（JSON）。
- `export_data/<桥型>/`：导出的 Excel 报告。

历史工程项目、导出报告、Python 缓存和本地环境未纳入公开仓库；这些内容已加入 `.gitignore`。运行所需的七份基础配置随源码提供。

## 开发与验证

安装依赖后，可检查源码语法：

```powershell
.\.venv\Scripts\python.exe -m compileall -q -x "[\\/]\.venv[\\/]" .
```

当前版本未提供自动化计算结果测试或图形界面测试。修改评定参数或计算逻辑后，应结合已知算例核对结果。

## 许可证

本项目采用 [MIT License](LICENSE)。
