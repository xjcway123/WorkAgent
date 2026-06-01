---
name: ppt-generation
description: "通用商务汇报 PPT 生成（产品介绍 / 公司介绍 / 项目汇报 / 季度复盘 / 销售提案 / 培训分享 等结构化演示文稿）。当用户要求制作一份商务/工作汇报/项目复盘/季度总结/产品发布/客户提案的 PPT，或提到「做 PPT」「生成汇报」「商务演示」「slides」「deck」「pitch deck」时，务必使用本 skill。本 skill 输出一份 16:9 宽屏、带品牌蓝主色与统一六种页型（封面 / 目录 / 章节分隔 / 内容 / 表格 / 结束）的成品 pptx，符合 references/style-guide.md 的统一规范。如果用户明确说『毕业答辩 / 学位论文 / 开题报告 / defense slides』，则改用 graduation-defense-pptx skill，本 skill 不覆盖学术答辩场景。"
license: Proprietary.
---

# 通用商务汇报 PPT 生成

把用户提供的素材（要点、指标、数据、章节标题、汇报人信息）组织成符合统一视觉规范的成品 pptx。**所有色号、字号、页型都来自 `references/style-guide.md`，不要凭感觉自由发挥。** 实物参考 `assets/template.pptx`（6 页型样例）。

## 与 `graduation-defense-pptx` 的边界

| 场景                                     | 用哪个 skill               | 原因                                  |
| ---------------------------------------- | -------------------------- | ------------------------------------- |
| 商务汇报 / 项目复盘 / 销售提案 / 培训    | **本 skill (ppt-generation)** | 通用、轻量、无模板包袱                |
| 毕业 / 学位论文 / 开题 / 中期答辩        | `graduation-defense-pptx`  | 学术风格 + 蓝色立方体模板 + 标准答辩结构 |
| 自由排版（海报 / 邀请函 / 朋友圈图文）   | ❌ 都不要                  | 退化为图片 / Word / Canva            |

**不要**试图在本 skill 里复刻答辩模板——答辩场景的复杂版式由独立 skill 负责。

## 前置依赖

- `python-pptx` >= 1.0（系统已装，1.0.2）
- 输出路径：`./output/{filename}.pptx`（与 `make_ppt` 工具一致）
- 标准库 `os` / `re` 用于文件名清洗

`multi_agent.py` 里现有的 `make_ppt` 工具**只能生成简单封面 + 标题/内容两栏**（不调样式），本 skill 的统一视觉需要用 python-pptl 写一个增强版工具（如 `make_business_ppt`）来落地，**不要在 `make_ppt` 上硬塞样式参数**。

## 第一步：识别"是不是商务汇报"

只有**结构化、有清晰章节**的内容才走本 skill：

| 场景                                              | 走本 skill？ | 备注                       |
| ------------------------------------------------- | ------------ | -------------------------- |
| 季度/年度业务汇报、产品发布、客户提案、销售 pitch | ✅           | 标准商务场景               |
| 培训课件、读书分享、工作坊                        | ✅           | 页数偏多（>20）也支持      |
| 自由排版的单页海报、活动邀请函                    | ❌           | 建议用 Canva / Figma       |
| 嵌入大量视频/动画的演示文稿                       | ⚠️ 部分支持  | 视频可插入，复杂动画不支持 |
| 答辩 / 学术汇报                                   | ❌           | 改用 `graduation-defense-pptx` |

## 第二步：收集与确认输入

必需（**缺一个就用 `ask_user_input_v0` 一次问完**）：

- **filename**：英文/拼音，不带后缀
- **title**：本次汇报的主标题
- **subtitle**（可选）：副标题或时间范围
- **presenter**：汇报人姓名
- **date**：汇报日期
- **sections**：章节列表，每项含 `name`（中文）、可选 `name_en`（英文副名）

内容（可让用户直接给大纲，也可在确认章节后逐章询问）：

- 每个章节下的 **要点列表**（每点 ≤ 30 字）
- 涉及的 **KPI**（数值 + 同比变化）
- 涉及的 **表格**（列名 + 行数据）
- 是否要 **Q & A 结束页**

> 不要问"你想要什么风格"——本 skill 自带品牌蓝风格，不接受改色（除非用户明确说"换一个色"，并接受同步更新 style-guide.md）。

## 第三步：按规范构建

按 `assets/template.pptx` 的结构生成 6 种页型：

### 页型 1：封面
- 左侧 5 英寸深蓝 `#293764` 大色块
- 主标题 44pt 白色加粗 + 副标题 36pt 白色加粗
- 副标题用 `——` 起头 + 关键信息
- 汇报人 + 日期 12pt 白色，距底 0.6"

### 页型 2：目录
- 顶部 0.18" 品牌蓝细条
- 4–6 个章节，每行：编号（32pt 品牌蓝）+ 章节名（18pt 深灰）+ 页码（12pt 灰）+ 0.01" 浅分割线

### 页型 3：章节分隔页
- 全屏深蓝底
- 大数字 140pt 品牌蓝 Arial（如 `01`）
- 章节名 36pt 白色加粗 + 0.04" 装饰条 + 英文副名 14pt 辅蓝
- **章节数 < 4 时不要插分隔页**——会让结构显得空洞

### 页型 4：内容页（KPI / 要点）
- 顶部色条 + 24pt 加粗深蓝标题
- 可选引言（14pt 灰，1–2 句）
- KPI 卡片：浅蓝底 + 左侧 0.04" 品牌蓝边条 + 12pt 标签 + 28pt 数值（Arial）+ 11pt 同比
- 要点列表：自定义蓝方块（`add_filled_rect` 0.1"×0.1"）替代默认 `•`

### 页型 5：内容页（表格）
- 顶部色条 + 标题
- 表头：深蓝底白字加粗 13pt
- 偶数行：浅蓝 `#F2F6FC` 底
- 行高 0.5"，单元格内边距 0.2"
- 不画边框线（用底色 + 对齐区分）

### 页型 6：结束页
- 全屏深蓝底 + 0.04" 品牌蓝横线
- "谢谢聆听" 54pt 白色居中
- "Q & A" 22pt 辅蓝居中
- 联系方式 12pt 辅蓝居中

所有色号、字号、字体见 `references/style-guide.md`，**禁止自由发挥**。

## 第四步：质检

生成后用 python-pptx 重新打开做这几项：

```python
from pptx import Presentation
prs = Presentation(OUT)
slides = prs.slides
total = len(slides)

# 1) 必须是 16:9
assert prs.slide_width / prs.slide_height == 16/9, "页面比例不是 16:9"

# 2) 至少 6 页
assert total >= 6, f"页数过少 ({total}), 至少 6 页"

# 3) 首页与末页必有特定标记
first_text = "\n".join(sh.text_frame.text for sh in slides[0].shapes if sh.has_text_frame)
last_text  = "\n".join(sh.text_frame.text for sh in slides[-1].shapes if sh.has_text_frame)
assert "汇报" in first_text or "Q" in first_text, "封面缺汇报主题"
assert "谢谢" in last_text or "Q & A" in last_text, "结束页缺致谢/Q&A"

# 4) 不出现 emoji / 艺术字残留
import re
for s in slides:
    for sh in s.shapes:
        if sh.has_text_frame:
            t = sh.text_frame.text
            assert not re.search(r"[\U0001F300-\U0001FAFF]", t), "检测到 emoji"
```

视觉质检（推荐）：用 LibreOffice 转 PNG 逐页看
```bash
libreoffice --headless --convert-to png --outdir /tmp/pptqa D:/WorkAgent/output/x.pptx
```

## 第五步：交付

- 输出路径：`./output/{filename}.pptx`
- 聊天回复给用户**文件路径 + 文件大小 + 页数 + 章节清单**，不要只发"已生成"
- 如果用户给了原始素材（数据/图表），明确指出哪几页基于这些素材

## 注意事项

- **风格统一优先于个性**：同部门同系列的汇报要看起来像一套；用户改色前先警告会破坏一致性。
- **页数控制**：商务汇报 10–20 页为宜；少于 8 页会显得"没什么可讲"，多于 25 页会让人走神。
- **章节数 ≥ 4 才加分隔页**，否则直接 封面→目录→内容×N→结束。
- **动画与转场**：本 skill **不**添加任何动画/转场；如有需求建议用 PowerPoint 客户端手动加，AI 加动画容易触发兼容性 bug。
- **嵌入字体**：用 PowerPoint 打开后"另存为"勾选"将字体嵌入文件"——但 .pptx 嵌入字体后文件会变大数倍，默认不嵌入。
- **图表**：本 skill 暂未提供图表页型，复杂图表建议用户用 Excel/数据可视化工具导出 PNG 后插入。如确需图表页型，扩展时新增"页型 7"并同步更新 style-guide.md。
- **本 skill 不涉及的 PPT 能力**：动画/转场、宏（VBA）、录制旁白、母版定制、协同编辑、密码保护。
- **素材边界**：本 skill 不生成原创图表、不写原创文案——所有内容必须由用户提供或经用户明确确认。

## 排错

| 现象                                          | 原因                                       | 解决                                                    |
| --------------------------------------------- | ------------------------------------------ | ------------------------------------------------------- |
| `ModuleNotFoundError: No module named 'pptx'` | venv 里没装                                | `pip install python-pptx`                               |
| 中文显示为方块                                | 系统缺中文字体                             | 安装微软雅黑，或改用 `Source Han Sans CN`               |
| 表格行高被内容撑开                            | 没设固定行高                               | 用 `add_filled_rect` 画行 + 文本框叠加，不要让 cell 自适应 |
| 封面副标题盖住主标题                          | top 偏移冲突                               | 主标题 y=2.6"、副标题 y=3.6"、副文 y=5.0"               |
| 页脚页码与内容重叠                            | 页脚 y=7.05" 但内容延伸到 7" 之下          | 内容页主体止于 6.5"，留 0.5" 给页脚                     |
| 16:9 播放时两侧被裁                           | 实际为 4:3                                 | 显式设 `prs.slide_width=Inches(13.333)`、`prs.slide_height=Inches(7.5)` |
| 章节分隔页大数字未显示                        | 数字框被推到 slide 外                       | 数字 y ≤ slide_height - 1.5"                            |
| 转 PDF 后字体丢失                             | 没嵌入字体                                 | 客户端另存时勾选嵌入，或先 `libreoffice --convert-to pdf` 中转 |
