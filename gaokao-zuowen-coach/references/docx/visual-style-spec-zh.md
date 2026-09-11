# 视觉篇：文档版式生成规范

> 面向 AI 生成文档的版式总览规范（LaTeX / Word 双实现）
> 版本 1.0 ｜ 用途：介绍 + 实操指南 ｜ 配套演示文档：`视觉篇-文档版式规范.docx`

---

## 0. 这份规范是什么

本规范回答一个问题：**AI 生成的文档应该"长什么样"**。它把文档的视觉设计拆成一套可复用的规则——从页面、字体、颜色，到文本框、表格、公式、页眉页脚——每一条都给出**规范值**与 **LaTeX / Word 两套实现写法**，并附**从 PDF 逆向版式的方法论**，供生成、评审与复刻三种场景使用。

三条总原则贯穿全文：

1. **层级优先**——视觉重量按信息重要性递减（题名 > 章 > 节 > 正文 > 注释），读者 3 秒内能抓住结构。
2. **一致性优先**——同一种元素全文只用一种样式；宁缺毋滥。
3. **克制优先**——全文主色不超过 3 种，强调色 1 种；颜色服务语义，不为装饰。

---

## 1. 设计总原则（细则）

### 1.1 信息层级
任何文档都应呈现 5 级视觉阶梯（由重到轻）：

| 层级 | 视觉手段 | 示例 |
|---|---|---|
| L1 文档级 | 封面、大字号题名 | 标题 14–16pt 加粗 |
| L2 章级 | 编号标题 + 页眉呼应 | 1 一级标题 |
| L3 节级 | 次级标题 | 1.1 二级标题 |
| L4 正文级 | 正文样式，最安静 | 五号/11pt 常规 |
| L5 注释级 | 小字号、浅色、楷体 | 脚注、图注、出处 |

### 1.2 留白与节奏
- 段落行距：1.5 倍（中文出版惯例）或 1.3–1.4 倍（紧凑文档）。
- 段后距：6pt（正文）／标题上 24pt 下 18pt（章）。
- 文本框内侧留白：4–6mm，四边一致；避免文字贴框。
- 页面不拥挤：一页正文 25–35 行汉字为宜。

### 1.3 行宽与可读性
- 中文正文每行 **30–40 字**；英文 60–90 字符。
- A4 + 2cm 边距 → 文字区约 165mm，天然满足。
- 超出此行宽（如宽屏排版）应分栏或收紧边距，严禁整行撑满。

### 1.4 对比度（无障碍底线）
- 正文文字与背景对比度 ≥ **4.5:1**（WCAG AA）。
- 深色框配浅色底；浅色文字只出现在深色底上。
- 不要用纯黑 #000 配纯白 #FFF 之外的高反差组合做正文（刺眼），正文用 #212529、#333。

---

## 2. 页面系统

### 2.1 纸张与边距

| 项目 | 规范值 |
|---|---|
| 纸张 | A4（210×297mm / 595.3×841.9pt） |
| 边距 | 上下左右 **2cm**（LaTeX `geometry`；Word 需手动改为 2cm） |
| 文字区宽 | ≈165mm / 468pt（2cm 边距下的 A4） |
| 装订侧 | 双面打印时左 2.5cm，右 2cm |

实现（LaTeX）：

```latex
\documentclass[10pt,A4]{ctexart}
\usepackage[margin=2cm]{geometry}
```

实现（Word / python-docx）：

```python
sec.page_width, sec.page_height = Inches(8.27), Inches(11.69)
sec.left_margin = sec.right_margin = Inches(0.79)
sec.top_margin = sec.bottom_margin = Inches(0.79)
```

### 2.2 页眉页脚
- 页眉：左侧文档标题、右侧当前章/栏目；页眉下方 **0.4pt 全宽横线**。
- 页脚：居中页码，格式"第 X 页，共 Y 页"（域代码，非手打数字）。
- 封面页单独成节，页眉页脚与正文**隔离**（Word 必须 `is_linked_to_previous = False`）。

### 2.3 目录
- 生成 2–3 级自动目录（Word 用 TOC 域，LaTeX 用 `\tableofcontents`）。
- Word 域目录初生成时为空白属正常，交付时提示用户 `Ctrl+A → F9` 刷新。

### 2.4 分节
- 封面 / 目录 / 正文 / 附录 各占一节；每节可独立控制页眉页脚、页码起始。

---

## 3. 字体体系

### 3.1 中文字体家族（按角色分配）

| 角色 | 推荐字体 | 可嵌入替代 | 说明 |
|---|---|---|---|
| 正文 | 宋体（中易宋体 / 思源宋体） | FandolSong（TeX Live 自带） | 最安静的正文；10.5–12pt |
| 标题 | 黑体（微软雅黑 / 思源黑体） | FandolSong-Bold、NotoSansCJK | 标题一律加粗黑体系 |
| 引文/注释/题解 | 楷体（楷体 / 思源楷体） | FandolKai | 引理陈述、引言、批注 |
| 公文式引文 | 仿宋 | 仿宋_GB2312 | 少用，仅规范性场景 |

### 3.2 西文与数学字体

| 场景 | 字体 |
|---|---|
| 西文正文（衬线） | Times New Roman / Latin Modern Roman |
| 西文标题（无衬线） | Arial / Latin Modern Sans |
| LaTeX 数学 | Computer Modern / Latin Modern Math（默认即美观） |
| Word 公式（OMML） | Cambria Math（Word 公式编辑器默认） |

### 3.3 字号阶梯（正文 10.5pt 为基准）

| 用途 | pt | Word 号数 |
|---|---|---|
| 题名 | 14–16 | 四号～三号 |
| 章标题 | 12–13 | 小四～四号 |
| 节标题 | 11–12 | 五号～小四 |
| 正文 | 10.5–11 | 五号 |
| 表内文字 | 9–10.5 | 小五～五号 |
| 脚注/图注 | 8–9 | 小五以下 |

### 3.4 中英混排的坑（AI 高频踩雷）
- Word：中文字体必须写在 `w:eastAsia` 属性上，否则中文回退为宋体、英文变 Times，两套字体打架。
- LaTeX：用 XeLaTeX + xeCJK/ctex 自动处理中文标点挤压与换行，**不要用 pdflatex 排中文**。
- 全角/半角：中文语境用全角标点（，。），公式与数字用半角。

---

## 4. 颜色体系

### 4.1 核心三色（"题目-答案-总结"语义色，源自实测逆向）

| 框 | 边框色（RGB） | 底色（RGB） | 语义 |
|---|---|---|---|
| 题目框 | 深蓝 (0, 0, 115) | 浅蓝 (247, 247, 255) | 问题陈述 |
| 答案框 | 深绿 (0, 115, 0) | 浅绿 (245, 255, 245) | 结论/答案 |
| 复盘框 | 橙 (166, 83, 0) | 米色 (255, 254, 243) | 总结/易错点 |

> 规律：**深色细边框 + 极浅色大底**。边框 1.5pt、圆角 2.5mm。这个"深框浅底"是中文 AI 文档最常见的识别特征。

### 4.2 中性色与强调色

| 角色 | 色值 |
|---|---|
| 正文 | #212529 |
| 次要文字 | #6C757D |
| 分割线 | #DEE2E6 |
| 浅底（代码/引用） | #F8F9FA |
| 语义成功/警告/危险 | #2E7D32 / #B26A00 / #C62828 |
| 标题强调蓝 | #1F3864（封面常用） |

### 4.3 配色纪律
- 全文主色 ≤ 3；文本颜色只允许中性色 + 1 个强调色。
- 色块不承载关键信息（同时用文字表达），照顾色盲与黑白打印。
- 浅底颜色选择 `fill` 色值必须用十六进制，Word 底纹 `w:val` 必须为 `clear`（`solid` 会渲染成黑块）。

---

## 5. 版式元素规范

### 5.1 封面
- 独占一页，不进目录；垂直方向约 1/3 处开始。
- 元素：大号题名（28pt 左右，深蓝 #1F3864 或黑色）→ 副标题（16pt 灰）→ 分隔线 → 作者/日期（12pt 浅灰）。
- 封面后加分节符，与正文页眉页脚隔离。

### 5.2 标题层级
- 编号：1 / 1.1 / 1.1.1；章节标题黑体加粗；节标题同色系但字号降一档。
- 可选装饰：章标题深色底白字整行反白，或"深蓝文字 + 下边框线"（推荐后者，克制）。
- 标题与正文之间留 16–24pt 空距，标题勿孤行（孤行控制开启）。

### 5.3 文本框（卡片）
统一参数：边框 1.5pt、圆角 2.5mm、内侧边距 4–6mm、底色极浅。

LaTeX（tcolorbox）：

```latex
\definecolor{probblue}{RGB}{0,0,115}
\definecolor{probback}{RGB}{247,247,255}
\newtcolorbox{problembox}{colframe=probblue, colback=probback,
  arc=2.5mm, boxrule=1.5pt, left=6mm, right=6mm, top=4mm, bottom=4mm}
```

Word（python-docx，oxml 注入：底纹 + 四边框）：

```python
def add_box(doc, text, fill_hex, border_hex):
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Inches(0.2)
    p.paragraph_format.right_indent = Inches(0.2)
    r = p.add_run(text)
    set_run_font(r, "微软雅黑", 11)
    add_shading(p, fill_hex)          # w:shd val=clear fill=fill_hex
    pPr = p._p.get_or_add_pPr()
    pBdr = OxmlElement('w:pBdr')
    for side in ['top','left','bottom','right']:
        b = OxmlElement(f'w:{side}'); b.set(qn('w:val'),'single')
        b.set(qn('w:sz'),'12')        # 12 半磅 = 1.5pt
        b.set(qn('w:color'), border_hex); pBdr.append(b)
    pPr.append(pBdr)
```

### 5.4 表格
- 推荐**三线表**（booktabs 风格）：顶线/底线 1.5pt，栏线仅表头下一条 0.5pt，无竖线。
- 复杂数据表用网格表：表头深蓝底（#2B579A 等）+ 白字加粗，正文行 9–10.5pt。
- **必须设置固定列宽布局**（`w:tblLayout type="fixed"`），否则 LibreOffice/网页打开列宽错乱。
- 数字列右对齐，文本列左对齐，表头居中；表宽不超过文字区。

### 5.5 公式
- 行内公式随文；独立公式居中、编号右对齐 `(1)`；多行对齐用 `align` + `&=`。
- 公式是**原生对象**不是图片：Word 必须是 OMML（`<m:oMath>`），LaTeX 用 amsmath。
- AI 生成 Word 含数学文档的标准路径：**写 LaTeX → pandoc 转 docx**（公式自动转 OMML 可编辑）。
- 编号交叉引用：LaTeX 用 `\eqref`（编译期动态），Pandoc 转 Word 后退化为静态文本，需人工核对。

### 5.6 列表与引用块
- 无序列表用圆点，二级用空心圆；有序列表连续编号。
- 引用块：左侧 4pt 竖线 + 浅灰底 + 楷体；代码块：等宽字体（Consolas / 等线）+ #F8F9FA 底。
- 长列表拆表：超过 6 条的并列信息优先转定宽表格，而非堆列表。

### 5.7 强调体系
- 只有三种强调手段：**加粗**（最常用）、**颜色**（1 种强调色）、**高亮底纹**（黄色系，Word 用 WD_COLOR_INDEX）。
- 一段文字只允许一种强调；全文高亮段落 ≤ 全文 5%。

### 5.8 图片与图注
- 图注置于图下居中，9pt 灰色，编号"图 1"；表注置于表上，同理。
- 图与正文间距上下各 6–8pt；图片尽量不跨页。

### 5.9 脚注与参考文献
- 脚注 8–9pt，与正文间用短横线分隔；参考文献统一编号 `[1]` 格式。
- Word 引用域：`{ REF }` / `{ CROSSREF }` 可更新；LaTeX 用 bibtex/biblatex 多趟编译。

### 5.10 目录与书签
- TOC 域 `TOC \o "1-3" \h \z \u`；多节文档每节可独立 TOC。
- 长文档启用书签与交叉引用，避免"见上文"式手写引用。

---

## 6. LaTeX 与 Word 双实现对照速查

| 视觉元素 | LaTeX（xelatex + ctex） | Word（python-docx / docx-js + oxml） |
|---|---|---|
| 页边距 | `geometry` | `section.left_margin=...` |
| 页眉线 | `fancyhdr`（headrule 默认 0.4pt） | 段落 `w:pBdr` bottom |
| 圆角色块 | `tcolorbox` | `w:shd` + 四边 `w:pBdr` |
| 三线表 | `booktabs` | 单元格边框逐侧设置 |
| 固定表宽 | `tabular` 自动 | `w:tblLayout type="fixed"` |
| 数学公式 | `amsmath`（Computer Modern） | OMML（pandoc texmath 转换） |
| 页码域 | 自动 | `fldChar` PAGE / NUMPAGES |
| 目录 | `\tableofcontents` | TOC 域，交付前 F9 |
| 中文字体 | xeCJK/ctex（Fandol 等） | `w:eastAsia` 属性 |
| 交叉引用 | `\ref/\eqref` | REF 域 |
| 高亮 | `\hl`/`soul` | `WD_COLOR_INDEX` |

---

## 7. 逆向方法论：从 PDF 反推别人的版式

拿到一份 PDF 想复刻，按顺序提取四类信息（工具：Python `pymupdf`/`fitz`）：

1. **引擎指纹（metadata）**：`creator` 含 "XeTeX output"/"LaTeX with hyperref" → LaTeX 系；`producer` 的 xdvipdfmx 版本号 → TeX Live 年代；PDF 1.5 vs 1.7 提示是否被后处理。
2. **字体指纹（get_fonts）**：字体名带 `XXXXXX+` 子集前缀 → 正常编译；**裸字体名（如 SimSun）→ 外部工具二次加工**；Fandol/Noto/CJK 变体（SC/JP/TC）提示字体配置来源。
3. **版式数值（get_drawings + spans）**：矢量框的宽高、填充色、圆角弧长 → 复刻 tcolorbox 参数；字号分布 → 字号阶梯；文字区宽度 → 边距。
4. **时间戳（creationDate/modDate）**：两者差数小时且 modDate 带本地时区 → 文档被编辑过；创建时间与考试/事件时刻对比可定位生成场景。

实战案例（本项目）：一份"IMO 讲义"PDF → metadata 锁定 XeLaTeX+TL2024；字体 FandolSong/Kai → ctex 模板；矢量框 468×124pt、边框 RGB(0,0,115)+底(247,247,255) → tcolorbox 三色框参数；页眉 468pt×0.4pt 横线 → fancyhdr。全部数值已收录进本文第 2、4、5 节。

---

## 8. 交付自查清单（AI 生成后逐项打勾）

- [ ] 中文字体已写入 `w:eastAsia`（Word）／使用 XeLaTeX + ctex（LaTeX）
- [ ] 所有底纹 `w:val="clear"`，无黑块
- [ ] 表格固定列宽布局，跨平台不变形
- [ ] 公式为原生 OMML / LaTeX 数学环境，非图片
- [ ] 页眉页脚分节隔离（封面独立）
- [ ] 页码、目录为域代码（Word 用户需 F9 刷新，已提示）
- [ ] 正文对比度 ≥ 4.5:1；色块不单独承载信息
- [ ] 字号阶梯齐全（题名/章/节/正文/注释五档）
- [ ] 主色 ≤ 3 种；文本框遵循"深框浅底"
- [ ] 行宽 30–40 字/行；行距 1.3–1.5 倍
- [ ] 用 `soffice --headless --convert-to pdf` 转图逐页目检（注意：LibreOffice 渲染 OMML 公式可能空白，需解压检查 `<m:oMath>` 节点数量）

---

## 附录 A：推荐色板

| 色名 | Hex | RGB | 用途 |
|---|---|---|---|
| 深蓝 | 1F3864 | (31,56,100) | 封面/主标题 |
| 题目蓝 | 000073 | (0,0,115) | 题目框边框 |
| 题目蓝底 | F7F7FF | (247,247,255) | 题目框底色 |
| 答案绿 | 007300 | (0,115,0) | 答案框边框 |
| 答案绿底 | F5FFF5 | (245,255,245) | 答案框底色 |
| 复盘橙 | A65300 | (166,83,0) | 复盘框边框 |
| 复盘米底 | FFFEF3 | (255,254,243) | 复盘框底色 |
| 表头蓝 | 2B579A | (43,89,154) | 表格表头 |
| 正文黑 | 212529 | (33,37,41) | 正文 |
| 次要灰 | 6C757D | (108,117,125) | 注释 |
| 分割线 | DEE2E6 | (222,226,230) | 分隔 |
| 浅灰底 | F8F9FA | (248,249,250) | 代码/引用底 |

## 附录 B：字号速查

| 号数 | pt | 典型用途 |
|---|---|---|
| 三号 | 16 | 大标题 |
| 四号 | 14 | 题名 |
| 小四 | 12 | 章标题/正文（宽松） |
| 五号 | 10.5 | 正文（基准） |
| 小五 | 9 | 表格、注释 |
| 六号 | 7.5 | 脚注（慎用） |

## 附录 C：LaTeX 宏包清单（中文数学文档）

```
ctex / xeCJK   中文字体与标点
geometry       页面设置
amsmath, amssymb, amsthm   数学与定理环境
tcolorbox      彩色圆角框（[most] 选项）
booktabs       三线表
fancyhdr       页眉页脚
hyperref       书签与链接
graphicx       图片
tabularx / longtable   表格
soul / xcolor  高亮与颜色
```
