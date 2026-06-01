---
name: excel-beautification
description: "业务报表 Excel 美化（销售/财务/运营/库存等结构化二维表）。当用户要求制作/整理/导出表格、给 Excel 美化、生成报表、把数据做成漂亮表格，或提到「做 Excel」「生成报表」「美化表格」「表格样式」「business report」「spreadsheet」时，务必使用本 skill。本 skill 输出一份带品牌蓝表头、斑马纹、合计行、冻结表头与隐藏网格线的成品 xlsx，符合 references/styling-guide.md 的统一规范。即使用户只说『做 Excel』，只要内容是结构化数据报表，也应使用本 skill。"
license: Proprietary.
---

# 业务报表 Excel 美化

把用户提供的结构化数据，组织成符合统一视觉规范的成品 xlsx。**所有色号、字号、字段命名都来自 `references/styling-guide.md`，不要凭感觉自由发挥。** 实物参考 `assets/template.xlsx`。

## 前置依赖

- 依赖 `openpyxl`（>= 3.1）。当前工作目录有现成的 `venv/`，但 openpyxl 未装入；如执行工具链需要，可执行：
  ```bash
  D:/WorkAgent/venv/Scripts/python.exe -m pip install openpyxl
  ```
- 依赖 Python 标准库 `pathlib`、`re`，用于路径与文件名清洗。
- 配套 `multi_agent.py` 里现有的 `make_excel` 工具**只能写数据 + 加粗表头**（无样式能力），本 skill 的样式效果需要用 openpyxl 单独写一个增强版工具（如 `make_beautiful_excel`）来落地，**不要在 `make_excel` 上硬塞样式参数**，避免破坏现有调用方。

## 第一步：识别"是不是业务报表"

只有**结构化二维表**才走本 skill。其他场景不要触发：

| 场景                                   | 走本 skill？ | 原因                         |
| -------------------------------------- | ------------ | ---------------------------- |
| 销售/财务/库存/HR 名单 + 指标          | ✅           | 标准业务报表                 |
| 实验数据、问卷统计、对比表             | ✅           | 二维表，符合规范             |
| 自由排版的"卡片式"Excel / 嵌套合并小计 | ⚠️ 部分可用  | 按本规范打底，但允许局部定制 |
| 单字段列表（如待办、日程）             | ❌           | 退化为 markdown/纯文本更合适 |
| 用户只给了一句话没数据                 | ❌           | 先反问澄清，不要凭空生成     |
| 用户明确要 CSV                         | ❌           | 不动样式，按 CSV 导出        |

## 第二步：收集与确认输入

必需：

- **filename**：英文/拼音，不要带后缀（程序会清洗非法字符）
- **headers**：列名数组
- **rows**：与 headers 等列数的二维数组

确认这些，缺一个就问一次（用 `ask_user_input_v0` 一次性问完，不要拆成多次追问）。可选但常见：

- **大标题**（如 "2026 年 5 月销售月报"）—— 没有就退化为只展示表头
- **合计列**（如 "销售额(元)" 求和）—— 末行合计；用户不要求则不加
- **数字格式**（货币 / 整数 / 百分比）—— 默认按 styling-guide.md 第 7 节推断

## 第三步：按规范构建

把数据喂给一个能调用 openpyxl 的工具。**关键样式参数必须取自 `references/styling-guide.md`**：

1. **大标题行**：合并所有列 + 深蓝 `#293764` 底 + 白色 14pt 加粗 + 居中 + 行高 32
2. **表头行**：品牌蓝 `#4A70AE` 底 + 白色 11pt 加粗 + 居中 + 上下 medium 边框 + 行高 24
3. **数据行**：偶数行 `#F2F6FC` 斑马纹 + 微软雅黑 11pt + 四边 thin 边框 + 文字列左对齐、列右对齐 + `wrap_text=True`
4. **数字格式**：货币列 `'"¥"#,##0.00'`、整数列 `'#,##0'`、百分比 `'0.00%'`、日期 `'yyyy-mm-dd'`
5. **合计行**：合并前半段写"合计" + 淡蓝 `#D9E2F3` 底 + 加粗 + `SUM` 公式
6. **冻结表头**：`ws.freeze_panes = "A3"`（标题与表头都常驻）
7. **关闭网格线**：`ws.sheet_view.showGridLines = False`
8. **列宽**：按表头与数据长度估算，逐列 `ws.column_dimensions[col].width = N`
9. **打印**：横向 + `fitToWidth=1` + 水平居中

公式直接用 `=SUM(D3:D{last})` 写字符串，openpyxl 默认不计算；用户在 Excel/WPS 打开时自动求值。**不要**在生成端手动求和后写成数字。

## 第四步：质检

生成后用 openpyxl 重新打开做这几项校验：

```python
from openpyxl import load_workbook
wb = load_workbook(OUT, data_only=False)
ws = wb.active

# 1) 标题行存在且被合并
assert ws["A1"].value is not None
assert any("A1:F1" in str(r) for r in ws.merged_cells.ranges)

# 2) 表头列数 == 数据列数
hdr = [ws.cell(row=2, column=c).value for c in range(1, ws.max_column + 1)]
data_cols = {len(r) for r in rows}
assert len(data_cols) == 1 and data_cols.pop() == len(hdr), "数据列数不统一"

# 3) 任意一行单元格有 fill
assert ws.cell(row=3, column=1).fill.fgColor.rgb is not None

# 4) 合计行公式存在（若用户要求合计）
#    ws.cell(row=total_row, column=6).value 应以 "=SUM" 开头
```

视觉质检（不强制但推荐）：转图查看，确认无文字溢出、斑马纹连续、合计行末位对齐。

## 第五步：交付

- 输出路径：`./output/{filename}.xlsx`（与 `make_excel` 工具一致）
- 在聊天回复里给用户**文件路径 + 文件大小 + 行/列数**，不要只发"已生成"
- 如果用户没指定 sheet 名，默认用"销售月报"等业务语义名（不要用 `Sheet1`）

## 注意事项

- **统一性优先于个性**：本 skill 强调"风格统一"。如果用户明确说"换一个色"，改完同步更新 `references/styling-guide.md`，避免下次又回到默认。
- **数字格式推断**：参考 styling-guide.md 第 7 节；含 `元/¥/$` → 货币，含 `%` → 百分比，含 `日期/时间/年/月/日` → 日期，其余默认 `'#,##0'` 整数或 `'General'`。
- **避免合并陷阱**：合并单元格后 `value` 只能在左上角读取；不要在已被合并的次格写数据，会被 openpyxl 静默丢弃。
- **行数多时**（>1000 行）：斑马纹 PatternFill 仍 OK，但 freeze_panes 与列宽不变；不要为了省事关掉样式。
- **xlsx 与 CSV 的边界**：本 skill 只对 xlsx 生效；用户要 CSV 时不要套样式。
- **公式安全**：openpyxl 写入的公式以 `=` 开头，**字符串中不要有用户输入的未清洗内容**，否则打开时 Excel 会弹安全警告。合计行只写 `=SUM(...)` 之类，不接受用户给的任意表达式。
- **本 skill 不涉及的 Excel 能力**：图表（chart）、数据透视、宏（VBA）、条件格式、密码保护。如用户要这些，告诉他们本 skill 范围有限，建议改用 Python pandas + xlsxwriter 或导出 PDF。

## 排错

| 现象                            | 原因                                    | 解决                                            |
| ------------------------------- | --------------------------------------- | ----------------------------------------------- |
| `ModuleNotFoundError: openpyxl` | venv 里没装                             | `pip install openpyxl`                          |
| 数字列左对齐/文字列右对齐       | 推断列类型时反了                        | 按 styling-guide.md 第 3 节区分                 |
| 合计行不求和                    | 写成了字符串 `'=SUM(...)'` 后保存时丢失 | 确认 value 以 `=` 开头，单元格 `data_type` 留空 |
| 斑马纹只在奇数行                | `r_idx` 起始值算错                      | 从 3 起算，偶数行 `r_idx % 2 == 0`              |
| 用户看不到"合计"二字            | 合并范围把文字列也并了                  | 末行只合并"文字列段"，合计字写在最左格          |
| 文件双击提示"文件已损坏"        | 用 `data_only=True` 写回导致公式丢失    | 保存用 `data_only=False`（默认）                |
