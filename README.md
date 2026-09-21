# IBGT 部门人员流动看板（IBGT Turnover Dashboard）

基于 **Excel 原生公式** 的部门人员流动看板生成器 —— 无宏、无插件依赖，粘贴数据即用，天然支持 SharePoint / Excel 网页版在线共享。

## 功能特性

- **双表合并统计**：`Data_Profile`（在职）+ `Data_Leavers`（离职）两张数据表自动合并，同一 `StaffID` 只计一次；两表状态冲突时以 Leavers 台账为准
- **Status 口径**：`Active` = 在职，`Inactive` = 离职（不看 `Contract end date`），离职月份取 `Last workday.`，离职当日仍计为在职
- **时间筛选**：全部时间 / 任一年份 / 任一月份，Headcount 取期间末、流入流出取期间内，明细表自动高亮选中期间
- **类型筛选**：按 `Perm / PC / MS` 列联动全部卡片与趋势图（不区分大小写、自动去空格）
- **可视化**：KPI 卡片 ×8、月末在职趋势线图、流出构成堆积柱状图、在职构成饼图，48 个月自动铺排
- **兼容性**：纯 `.xlsx` 无宏；公式仅用 SUMPRODUCT / COUNTIF / EDATE 等传统函数，Excel 2016+ 与网页版全部支持

## 快速开始

**方式 A · 有 Python 环境**

```bash
pip install openpyxl
python make_hr_dashboard.py        # 或双击 run_dashboard.bat
```

**方式 B · 无 Python 环境**

从 [Releases](../../releases) 下载 `IGBT看板生成器.exe`，双击即可（首次运行可能触发 SmartScreen，选择"仍要运行"）。

运行后生成两个文件：

| 文件 | 说明 |
|---|---|
| `IGBT_Turnover_Dashboard.xlsx` | 空白模板，粘贴真实数据使用 |
| `IGBT_Turnover_Dashboard_Demo.xlsx` | 内置演示数据，直接打开看效果 |

## 数据更新

把源工作簿（`'1. IBGT Profile'` 与 `'1.1 Leavers'` 两张表）的数据粘贴进看板的 `Data_Profile` / `Data_Leavers`（建议选择性粘贴为数值），看板自动重算。日期列必须为真日期格式；`Status` 只接受 `Active` / `Inactive`。

## 文件结构

```text
├── make_hr_dashboard.py        # 生成器主程序
├── base.py                     # 样式令牌（与主程序同目录放置）
├── run_dashboard.bat           # Windows 双击运行入口
├── IGBT人员流动指标计算依据.md   # 各指标的计算口径文档
└── IGBT_Turnover_Dashboard_Demo.xlsx   # 演示数据成品（示例输出）
```

## 指标口径

期末在职、当月流出/入职、期间筛选、两表去重、异常数据处理等完整规则见 [IGBT人员流动指标计算依据.md](IGBT人员流动指标计算依据.md)。

## 说明

- `base.py` 派生自 ZCode 内置 xlsx 技能的样式模板
- 演示数据均为虚构，不含任何真实人员信息
