# IBGT 部门人员流动看板 · 指标计算依据

> 适用范围：`IBGT_Turnover_Dashboard.xlsx`（含 Demo 版）中的全部人员流动指标。
> 文档口径与看板当前公式实现严格一致；字段名保留英文，说明用中文。

---

## 一、数据来源与基本约定

看板从两个数据表读取明细，两表结构相同（24 列），以 `StaffID` 作为人员的唯一标识：

| 数据表 | 对应源表 | 内容约定 |
|---|---|---|
| `Data_Profile` | `'1. IBGT Profile'` | 在职人员明细（Status = Active） |
| `Data_Leavers` | `'1.1 Leavers'` | 全部历史离职人员（Status = Inactive） |

**参与计算的四个关键字段：**

| 字段 | 用途 |
|---|---|
| `Staff start date` | 入职日期：决定入职月份、在职起始 |
| `Last workday.` | 最后工作日：决定离职月份、在职截止 |
| `Status` | 在职状态：`Active` / `Inactive`，是在职与离职的判定依据 |
| `Perm / PC / MS` | 招聘类型：分类筛选与分类统计的依据 |

其余字段（`Contract end date`、`Type`、`Tenure` 等）**不参与计算**。`Contract end date` 仅为预留字段——因为合同到期日可能不准确，离职判定一律以 `Status` 为准、以 `Last workday.` 为时点。

**合并与去重约定：** 两个表会合并扫描。同一个人如果同时出现在两表（`StaffID` 相同），在职与入职只计一次；流出的去重规则见第二节。

**日期要求：** `Staff start date` 与 `Last workday.` 必须是真日期格式（Excel 日期序列值）。文本格式的日期无法参与日期比较，会被「数据质量自检」计为异常。

---

## 二、三条核心判定规则

### 2.1 在职判定（某月末仍算在职的条件）

某员工在月末 M 仍计为在职，需同时满足：

1. `Staff start date` ≤ 当月月末（已入职）；
2. `Status` = `Active`，**或** 填有 `Last workday.` 且 `Last workday.` ≥ 当月月末。

即：在职人员一直算到今天；已离职人员算到 `Last workday.` 所在月为止（离职当天仍算在职，次月起不再计入）。

> 示例：某人 `Last workday.` = 2025-06-30，则 2025 年 6 月末仍计他在职，2025 年 7 月末起不再计入。

### 2.2 离职判定

`Status` = `Inactive` 即视为已离职，**不看 `Contract end date`**。离职月份取 `Last workday.` 所在月份。

- `Status` = `Inactive` 但 `Last workday.` 为空：算离职，但无法归属到具体月份——不计入任何月份的流出数，也不计入当前在职，由「数据概览」单独提示人数（提醒补数据）。
- `Status` 为空或其它值：不计入当前在职、不计入流出，由「数据质量自检」的 *Rows with blank / other Status* 提示。
  （边界情形：此类行若填有合法的 `Last workday.`，在月度历史中仍按 2.1 的第二条计入在职。）

### 2.3 类型归一化

类型取 `Perm / PC / MS` 列（注意不是 `Type` 列）。匹配时先去首尾空格、不区分大小写——因此 `perm`、`PC `（带尾空格）都能正确归入 `Perm` / `PC`。看板的类型下拉框选项固定为 **All / Perm / PC / MS**；选 All 表示不筛选。

---

## 三、指标定义

### 3.1 当前在职人数（Current Headcount，Period 选 All 时）

```text
当前在职 = Status = Active 且 Staff start date ≤ 今天 的人数（按类型筛选）
```

即"此刻正在职、且已入职报到"的人。未来才入职的人（start date 晚于今天）不计入。

### 3.2 期末在职人数（Headcount EOM，某月 / 某年）

按 2.1 的在职判定统计月末 M 时点仍在职的人数（年 = 该年 12 月末；All = 当前在职）。这是看板折线图「Month-End Headcount Trend」的数据来源。

### 3.3 当月流出人数（Leavers，某月）

```text
当月流出 = Status = Inactive 且 Last workday. 落在该月 的人数（按类型筛选）
```

`Last workday.` 当天离职的，计入当月。这是堆积柱状图「Monthly Leavers by Type」与月度明细表 Leavers 列的数据来源。

### 3.4 当月入职人数（Joiners，某月）

```text
当月入职 = Staff start date 落在该月 的人数（按类型筛选）
```

与 `Status` 无关——离职的人入职当月同样计入（用于还原历史规模）。

### 3.5 期间合计（Joiners / Leavers in Period）

Period 选某年或某月时：

- **Joiners in Period** = `Staff start date` 落在所选期间内的人数；
- **Leavers in Period** = `Status` = `Inactive` 且 `Last workday.` 落在所选期间内的人数。

Period 选 **All** 时：Joiners = 全部有 `Staff start date` 的人；Leavers = 全部 `Status` = `Inactive` 的人（含 `Last workday.` 缺失的，此时以"人"为单位，不按月归属）。

### 3.6 期间末在职（Period-End Headcount）

Period 选某年 → 该年 12 月末的期末在职；选某月 → 该月月末的期末在职（算法同 3.2）；选 **All** → 当前在职（算法同 3.1）。Perm / PC / MS 三张卡片与饼图「Current Headcount by Type」使用同一时点。

### 3.7 Latest Month Leavers

数据中最近一个发生流出的月份的流出人数（不受 Period 影响），用于快速看到最新动态。

---

## 四、时间筛选口径（Period）

| 选项 | 含义 | Headcount 取值 | Joiners / Leavers 统计范围 |
|---|---|---|---|
| All | 全部时间 | 当前在职（Status = Active） | 全部历史 |
| 某年份（如 2026） | 该自然年 | 该年 12 月末的期末在职 | `start date` / `Last workday.` 落在该年内 |
| 某月份（如 2025-06） | 该自然月 | 该月月末的期末在职 | 落在该月内 |

- 类型筛选（Type Filter）与时间筛选相互独立、可叠加。
- 选中的期间在「Monthly Detail」表中以浅蓝底色高亮。
- 月份范围自动从数据中最早的 `Staff start date` 起，共 48 个月；之后月份留空。

---

## 五、两表去重规则

同一个人（`StaffID` 相同）同时出现在 `Data_Profile` 和 `Data_Leavers` 时：

| 情形 | 处理 |
|---|---|
| 两表都是 `Inactive`（如 Profile 保留全员并更新状态，Leavers 另存一份） | 在职 / 入职 / 流出均只计一次 |
| Profile 是 `Active`、Leavers 是 `Inactive`（状态未同步回 Profile） | 在职只计一次（Profile 的 Active 生效）；**流出以 Leavers 表为准，计为已离职** |
| 只出现在其中一个表 | 正常计入该表 |

> 原则：在职状态以较新的在岗事实为准；离职事件以 Leavers 表（离职台账）为准。

---

## 六、数据质量自检项

「Guide」表的 Data Quality Checks 与看板「Data Overview」实时显示以下检查：

| 检查项 | 含义 | 建议处理 |
|---|---|---|
| 两个表的数据行数 | 确认数据已粘贴 | 为 0 时看板显示空 |
| Profile rows with Last workday | Profile 中已带离职日期的人数（即 Profile 内的离职人员） | 与 Leavers 行数对照 |
| Leavers also in Profile (deduped) | 两表重复的人员行数（已自动去重） | 一般应 ≤ Leavers 总行数 |
| Rows with blank / other Status | Status 缺失或非 Active/Inactive 的行数 | 补全 Status，否则不计入在职 |
| Date anomalies | 开始 / 离职日期中不是真日期的单元格数 | 改为日期格式（文本日期不参与统计） |
| Inactive without Last workday | 已离职但没填 `Last workday.` 的人数 | 补填后才能按月归属流出 |

---

## 七、在看板中的实现位置

| 位置 | 内容 |
|---|---|
| `Monthly Detail` 表 | 48 个月 × 各指标的全部计算公式（SUMPRODUCT 逐条件统计），合计行、饼图小表 |
| `Dashboard` 表 | KPI 卡公式（随 Type / Period 联动）、三张图表、数据概览 |
| `Calc`（隐藏表） | 辅助列：类型与 Status 规范化、两表去重标记、月份链、Period 选项与期间解析 |
| `Guide` 表 | 表头映射配置（若源表表头名变化，改 C 列即可）与数据质量自检 |

> 更新方式：把源表数据粘贴进 `Data_Profile` / `Data_Leavers`（建议选择性粘贴为数值），看板自动重算；如未刷新按 `Ctrl+Alt+F9` 强制全量重算。
