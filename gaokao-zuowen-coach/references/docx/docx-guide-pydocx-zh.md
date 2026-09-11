# Claude 生成 Word 文档指南：基于 python-docx 的模块化排版详解

**副标题：** 文档部件、样式模块与代码实现全解析
**日期：** 2026 年 7 月 10 日

---

## 一、概述

Claude 在应对「帮我生成一份 Word 文档」这类需求时，不会把整篇文档当成一整块纯文本去处理，而是把 Word 文档拆解成一系列可复用、可组合的「文档部件」（building blocks），再根据文档的用途（报告、论文、说明书、简历等）挑选、拼装这些部件，最终形成一份结构清晰、样式统一的 `.docx` 文件。这种「模块化」思路的好处在于：

- 每个模块职责单一，方便复用和调试；
- 模块之间通过统一的样式表（Style）关联，保证全文风格一致；
- 出错时可以定位到具体模块，而不必推倒重来。

本文以 Python 生态中最常用的 `python-docx` 库为例，逐一介绍构建文档时可能用到的各类模块，涵盖「页面结构」「标题与目录」「列表」「表格」「字符格式」「图片与对象」「引用与注释」，以及一些需要绕过 `python-docx` 原生限制、直接操作底层 XML（业内称为 `oxml`）才能实现的高级效果。

> **说明：** 在 Claude 当前所处的沙盒环境中，实际生成 `.docx` 文件所使用的工具链是 Node.js 的 `docx`（docx-js）库，而非 `python-docx`；但二者最终都是在拼装同一套 OOXML（Office Open XML）结构，思路是相通的。本文选择 `python-docx` 作为讲解载体，是因为它 API 更直观、社区示例更丰富，便于读者在自己的 Python 环境中直接复现。

全文按以下九大类别组织：

1. 页面结构类模块（封面、分节符、页眉页脚、页码域、分栏、水印）
2. 标题与目录类模块（标题样式、正文样式、目录域）
3. 列表类模块（项目符号、编号、多级列表）
4. 表格类模块（基础表格、单元格底纹、合并单元格、表格题注）
5. 字符格式与强调类模块（加粗斜体下划线、颜色高亮、字体字号）
6. 图片与对象类模块（图片、图片题注、文本框）
7. 引用与注释类模块（脚注、尾注、超链接、书签与交叉引用）
8. 高级与自定义类模块（自定义样式、分隔线、首字下沉、批注与修订）
9. 模块速查表、常见报错汇总与完整骨架示例

---

## 二、页面结构类模块

### 2.1 封面（Cover Page）

**长什么样：** 封面通常独占一页，是全文中排版最「自由」的一页——大号加粗标题、居中副标题、可选的分隔线或色块、以及日期、作者、机构等信息的堆叠；不使用正文的标题层级（Heading），而是用手动设置字号、加粗、居中对齐的普通段落拼出来的「视觉标题」。

**实现代码：**

```python
from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH

doc = Document()

# 标题
title = doc.add_paragraph()
title.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = title.add_run("2025 年度技术报告")
run.font.size = Pt(32)
run.font.bold = True
run.font.color.rgb = RGBColor(0x1F, 0x4E, 0x79)

# 副标题
subtitle = doc.add_paragraph()
subtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER
sub_run = subtitle.add_run("—— 业务增长与技术演进")
sub_run.font.size = Pt(16)
sub_run.font.color.rgb = RGBColor(0x59, 0x59, 0x59)

# 用若干空行把落款信息推到页面下方
for _ in range(8):
    doc.add_paragraph()

info = doc.add_paragraph()
info.alignment = WD_ALIGN_PARAGRAPH.CENTER
info.add_run("撰写人：Claude\n日期：2026 年 7 月 10 日")

doc.add_page_break()  # 封面独占一页
```

**适用场景：** 正式报告、论文、标书、产品说明书等需要「第一眼」传达标题与机构信息的长文档；不适合简单的备忘录或聊天式短文。

**提醒：** `python-docx` 没有「封面模板」的概念，所有视觉效果都要靠手动堆砌段落和空行实现；段末别忘记 `add_page_break()`，否则正文会紧跟着挤到封面下方。

---

### 2.2 分节符（Section Break）

**长什么样：** 分节符本身不可见，但它把文档切成若干个「节」（Section），每一节可以有独立的页码格式、页眉页脚、纸张方向（纵向/横向）甚至纸张大小。最常见用法是让封面和目录不显示页码，正文从第 1 页重新开始编号。

**实现代码：**

```python
from docx.enum.section import WD_SECTION

# add_section 会在文档当前位置插入分节符，并返回新的 Section 对象
new_section = doc.add_section(WD_SECTION.NEW_PAGE)
new_section.orientation = 0  # 0=纵向 portrait，1=横向 landscape
```

**适用场景：** 封面/目录与正文页码格式不同、部分表格页需要横向排版（如宽表）、多章节报告需要分别设置页眉。

**提醒：** `python-docx` 对「节」内页码起始值等设置支持有限，若要实现「正文页码从 1 重新开始」，通常需要直接操作 `sectPr` 对应的底层 XML，而不是纯 API 调用。

---

### 2.3 页眉与页脚（Header & Footer）

**长什么样：** 页面顶部/底部的固定区域，常见内容是公司 Logo、文档标题、章节名、页码等，可以设置为「首页不同」或「奇偶页不同」。

**实现代码：**

```python
section = doc.sections[0]
section.header.is_linked_to_previous = False
header_para = section.header.paragraphs[0]
header_para.text = "机密文件 · 内部使用"
header_para.alignment = WD_ALIGN_PARAGRAPH.RIGHT

footer_para = section.footer.paragraphs[0]
footer_para.text = "XX 科技有限公司"
footer_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
```

**适用场景：** 企业报告、合同、说明书等需要在每一页重复展示标识信息的场景。

**提醒：** 若封面不希望出现页眉页脚，需要在封面所在的节里单独设置 `different_first_page_header_footer = True`。

---

### 2.4 页码域（Page Number Field）

**长什么样：** 页脚（或页眉）中央/靠边显示的当前页码，通常写作「第 X 页」。这是一个「域代码」（Field），会随文档翻页自动刷新，而不是写死的静态文字。

**实现代码：**

```python
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

def add_page_number(paragraph):
    run = paragraph.add_run()
    fld_char1 = OxmlElement('w:fldChar')
    fld_char1.set(qn('w:fldCharType'), 'begin')
    instr_text = OxmlElement('w:instrText')
    instr_text.text = "PAGE"
    fld_char2 = OxmlElement('w:fldChar')
    fld_char2.set(qn('w:fldCharType'), 'end')
    run._r.append(fld_char1)
    run._r.append(instr_text)
    run._r.append(fld_char2)

add_page_number(footer_para)
```

**适用场景：** 任何多页正式文档，几乎是「必备」模块。

**提醒：** `python-docx` 没有直接封装「插入页码」的 API，必须像上面这样手写 `w:fldChar` / `w:instrText` 域代码；Word 打开后往往需要用户按 `Ctrl+A`、`F9`「更新域」才能看到页码实际数字，这是 Word 本身的机制，不是代码 bug。

---

### 2.5 分栏（Columns）

**长什么样：** 把一页内容分成 2 栏或 3 栏，像报纸/杂志排版一样，文字自动从左栏排到底后跳到右栏。

**实现代码：**

```python
sectPr = doc.sections[0]._sectPr
cols = sectPr.xpath('./w:cols')[0]
cols.set(qn('w:num'), '2')       # 两栏
cols.set(qn('w:space'), '425')   # 栏间距（单位：twip，约 0.3 英寸）
```

**适用场景：** 词典类内容、新闻稿、学术期刊的双栏排版正文。

**提醒：** `python-docx` 未提供 `sections.columns` 这样的高层属性，只能通过底层 XML 直接改 `w:cols` 节点；分栏设置作用于「节」，如果只想让某一部分分栏，需要先用分节符把该部分独立成一节。

---

### 2.6 水印（Watermark）

**长什么样：** 一段以极浅灰色、大角度倾斜、贯穿整页背景的文字（如「机密」「草稿」），不影响正文阅读但能提示文档性质。

**实现代码：**

```python
# python-docx 没有原生水印 API，需要向页眉插入一个带旋转角度的文本路径图形，
# 这一步通常直接拼接底层 XML 字符串后注入 header 部分
watermark_xml = '''
<w:pict xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <v:shapetype id="_x0000_t136" coordsize="1600,21600" o:spt="136"/>
  <v:shape id="PowerPlusWaterMarkObject" type="#_x0000_t136"
     style="position:absolute;margin-left:0;margin-top:0;width:400pt;height:200pt;
            rotation:315;z-index:-1" fillcolor="silver" stroked="f">
    <v:textpath style="font-family:'微软雅黑';font-size:1pt" string="草稿"/>
  </v:shape>
</w:pict>
'''
```

**适用场景：** 合同/协议的「草稿」版本、内部保密文件、防止截图外传的标识场景。

**提醒：** 水印是 `python-docx` 支持度最低的模块之一，几乎必须手写 VML/DrawingML 片段；不同 Word 版本对该片段的兼容性也有差异，生成后强烈建议用 LibreOffice 或 Word 实际打开核对效果。

---

## 三、标题与目录类模块

### 3.1 标题样式（Heading Styles）

**长什么样：** Word 内置了 Heading 1~Heading 9 九级标题样式，字号逐级递减、通常带有主题色，并且天然具备「大纲级别」（outline level），是目录能够自动生成的基础。

**实现代码：**

```python
doc.add_heading("第一章 项目背景", level=1)
doc.add_heading("1.1 行业现状", level=2)
doc.add_heading("1.1.1 市场规模", level=3)
```

**适用场景：** 任何需要清晰层级结构的长文档：报告、论文、说明书、书籍章节。

**提醒：** 若要自定义标题颜色/字体，应该修改 `doc.styles['Heading 1']` 这个样式对象本身，而不是逐个 run 去改颜色，否则全文标题风格会不统一，后续也无法通过「一键改样式」批量调整。

---

### 3.2 正文样式（Normal / Body Text）

**长什么样：** 全文默认的段落样式，决定了正文的字体、字号、行距、段前段后间距，是「排版是否精美」的地基。

**实现代码：**

```python
normal = doc.styles['Normal']
normal.font.name = '微软雅黑'
normal.font.size = Pt(12)
# 中文字体需要额外设置东亚字体，否则中文可能仍显示为默认字体
normal.element.rPr.rFonts.set(qn('w:eastAsia'), '微软雅黑')
normal.paragraph_format.line_spacing = 1.5
normal.paragraph_format.space_after = Pt(6)
```

**适用场景：** 几乎每一份文档都要先定好 Normal 样式，再往上叠加其他模块。

**提醒：** `python-docx` 设置西文字体（`font.name`）和中文字体（`eastAsia`）是两套独立属性，只设置前者会出现「英文变了字体，中文没变」的现象，这是最常见的中文排版坑之一。

---

### 3.3 目录域（Table of Contents Field）

**长什么样：** 通常出现在正文最前面独立一页，列出各级标题及其对应页码，并带有「点状引导线」（dot leader），点击可跳转（需 Word 环境）。

**实现代码：**

```python
def add_toc(document):
    paragraph = document.add_paragraph()
    run = paragraph.add_run()
    fld_char_begin = OxmlElement('w:fldChar')
    fld_char_begin.set(qn('w:fldCharType'), 'begin')
    instr_text = OxmlElement('w:instrText')
    instr_text.set(qn('xml:space'), 'preserve')
    instr_text.text = 'TOC \\o "1-3" \\h \\z \\u'  # 抓取 1~3 级标题
    fld_char_separate = OxmlElement('w:fldChar')
    fld_char_separate.set(qn('w:fldCharType'), 'separate')
    fld_char_end = OxmlElement('w:fldChar')
    fld_char_end.set(qn('w:fldCharType'), 'end')
    for el in (fld_char_begin, instr_text, fld_char_separate, fld_char_end):
        run._r.append(el)

add_toc(doc)
doc.add_page_break()
```

**适用场景：** 超过 5~6 页的长文档，方便读者快速跳转章节。

**提醒：** 目录内容依赖 Heading 样式的大纲级别，若标题是用「加粗+大字号」伪造出来的普通段落，而不是真正的 Heading 样式，目录将无法抓取到该标题；生成后同样需要在 Word 中「更新域」才会显示真实页码。

---

## 四、列表类模块

### 4.1 项目符号列表（Bullet List）

**长什么样：** 每一项前面带一个圆点/方块等符号，项目之间通常没有先后顺序。

**实现代码：**

```python
doc.add_paragraph("支持 Markdown 语法", style="List Bullet")
doc.add_paragraph("支持实时协作编辑", style="List Bullet")
doc.add_paragraph("支持导出为 PDF", style="List Bullet")
```

**适用场景：** 罗列并列的特性、优点、注意事项等不强调顺序的内容。

**提醒：** 不要为了图省事直接在正文里手打「• 文字」，那只是一个普通字符，不具备列表的缩进/编号联动特性，在 Word 里也无法用「继续列表」功能自动延续。

---

### 4.2 编号列表（Numbered List）

**长什么样：** 每一项前带阿拉伯数字/罗马数字等，强调先后顺序或步骤。

**实现代码：**

```python
doc.add_paragraph("下载安装包", style="List Number")
doc.add_paragraph("运行安装向导", style="List Number")
doc.add_paragraph("重启计算机", style="List Number")
```

**适用场景：** 操作步骤、排名、法律条款等有明确顺序的内容。

**提醒：** `python-docx` 内置的 List Number 样式在不同模板下起始编号可能不是从 1 开始，如果连续两段之间被其他样式的段落打断，编号会「断开重排」，需要时可通过底层 XML 中的 `w:numId` 手动关联同一个编号序列。

---

### 4.3 多级列表（Multilevel List）

**长什么样：** 类似「1. / 1.1 / 1.1.1」这种带层级缩进的编号，常见于合同条款、标准规范文档。

**实现代码：**

```python
doc.add_paragraph("总则", style="List Number")
doc.add_paragraph("定义", style="List Number 2")     # 二级编号样式
doc.add_paragraph("术语解释", style="List Number 3")  # 三级
```

**适用场景：** 合同条款、国家标准/行业规范、复杂的技术手册目录结构。

**提醒：** 默认模板里 List Number 2/3 这类多级样式的编号格式未必是「1.1 / 1.1.1」这种常见形式，通常还需要在 `numbering.xml` 里自定义编号格式（`abstractNum`），纯 API 很难做到完全精确控制。

---

## 五、表格类模块

### 5.1 基础表格与内置样式（Table & Built-in Style）

**长什么样：** 由行（row）和列（column）组成的网格，Word 自带几十种表格样式（如「浅色列表 - 强调 1」），一键切换配色和边框风格。

**实现代码：**

```python
table = doc.add_table(rows=3, cols=3)
table.style = "Light Grid Accent 1"  # 使用内置样式
table.cell(0, 0).text = "季度"
table.cell(0, 1).text = "营收（万元）"
table.cell(0, 2).text = "同比增长"
```

**适用场景：** 财务数据、参数对比、时间安排表等结构化信息展示。

**提醒：** 内置样式名称必须与模板中实际存在的名称完全一致（区分大小写和空格），写错名字会直接抛出 `KeyError`，建议先用 `[s.name for s in doc.styles]` 打印确认可用样式列表。

---

### 5.2 单元格底纹（Cell Shading）

**长什么样：** 给指定单元格填充背景色，常用于表头行或需要高亮的关键数据格。

**实现代码：**

```python
def set_cell_background(cell, color_hex):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:fill'), color_hex)  # 例如 "1F4E79"
    tcPr.append(shd)

set_cell_background(table.cell(0, 0), "1F4E79")
```

**适用场景：** 表头行强调、超出预算的费用行标红、达标行标绿等「一眼看出重点」的场景。

**提醒：** 这里的 `w:val` 必须是 `"clear"`（表示无图案叠加），如果误用 `"solid"` 之类的值，在部分渲染器里反而会显示成纯黑色底纹，这是 OOXML 里一个经典的反直觉坑。

---

### 5.3 合并单元格（Merged Cells）

**长什么样：** 把相邻的多个单元格合并成一个更大的格子，常见于跨列表头或跨行分类标签。

**实现代码：**

```python
a = table.cell(0, 0)
b = table.cell(0, 1)
merged_cell = a.merge(b)  # 横向合并第一行的前两列
merged_cell.text = "合并后的表头"
```

**适用场景：** 多级表头（如「上半年」横跨「1月/2月/3月」三列）、分类标签跨行展示。

**提醒：** 合并后原本 `b` 单元格的引用仍然指向同一个合并区域，重复对 `b` 赋值会覆盖 `merged_cell` 的内容；合并单元格较多的复杂表格，建议先画草图确定行列结构再写代码，避免逻辑错乱。

---

### 5.4 表格题注（Table Caption）

**长什么样：** 表格上方或下方的一行小字，格式通常是「表 1-1 XX 数据统计」，用于配合正文交叉引用。

**实现代码：**

```python
caption = doc.add_paragraph()
caption.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = caption.add_run("表 1  2025 年各季度营收统计")
run.font.size = Pt(10)
run.font.italic = True
```

**适用场景：** 论文、技术报告中需要被正文引用（如「详见表 1」）的数据表格。

**提醒：** 严格意义上的「题注」在 Word 里是通过 SEQ 域自动编号的（表 1 / 表 2 会随插入顺序自动更新），`python-docx` 目前没有对 SEQ 域的高层封装，上面代码只是「静态文字」版的简化实现，数量较多的表格建议同样手写域代码以获得自动编号能力。

---

## 六、字符格式与强调类模块

### 6.1 字符强调：加粗 / 斜体 / 下划线

**长什么样：** 在同一段落内，通过对某一个 Run（而非整段）单独设置属性，实现局部加粗、倾斜或加下划线，而不影响同段其他文字。

**实现代码：**

```python
p = doc.add_paragraph("请特别注意：")
run1 = p.add_run("本条款具有法律约束力")
run1.bold = True
p.add_run("，逾期未确认视为默认同意。")
run2 = p.add_run("务必在 7 个工作日内回复。")
run2.underline = True
```

**适用场景：** 正文中需要突出局部关键词、警示语、法律条款等场景。

**提醒：** `python-docx` 的格式属性（`bold`/`italic`/`underline`/`color` 等）只能设置在 Run 级别，不能直接设置在 Paragraph 上；如果对整段文字一次性写入再统一设置格式，会导致「一句话里想强调的几个字」无法单独区分开。

---

### 6.2 文字颜色与高亮（Font Color & Highlight）

**长什么样：**「颜色」是改变文字本身的墨色（如把字变成红色）；「高亮」则更像用记号笔在文字背后涂一层底色（如荧光黄），两者可以叠加使用。

**实现代码：**

```python
from docx.enum.text import WD_COLOR_INDEX

run = p.add_run("这是需要重点关注的风险项")
run.font.color.rgb = RGBColor(0xC0, 0x00, 0x00)   # 深红色文字
run.font.highlight_color = WD_COLOR_INDEX.YELLOW  # 荧光黄高亮
```

**适用场景：** 风险提示、审校批注、需要读者第一时间注意到的结论性文字。

**提醒：** `font.highlight_color` 只能从 `WD_COLOR_INDEX` 枚举里的有限几种颜色中选择（类似 Word 荧光笔的预设色），不能像 `font.color.rgb` 那样自定义任意十六进制颜色。

---

### 6.3 自定义字体与字号（Font & Size）

**长什么样：** 针对某一段/某几个字单独调整字体家族（如宋体改成黑体）、字号（五号、小四等 Word 惯用字号或磅值）。

**实现代码：**

```python
run = p.add_run("重要提示")
run.font.name = "黑体"
run.element.rPr.rFonts.set(qn('w:eastAsia'), '黑体')
run.font.size = Pt(14)  # 相当于「四号」
```

**适用场景：** 需要区分中西文字体的中文文档、需要遵循特定格式规范（如论文要求宋体小四）的场景。

**提醒：** 同 3.2 节所述，中文字体一定要额外设置 `w:eastAsia`，否则在 Word 里看起来「设置了却没生效」，这是询问频率最高的 `python-docx` 问题之一。

---

## 七、图片与对象类模块

### 7.1 图片插入（Picture）

**长什么样：** 在段落中插入位图（png/jpg 等），可指定显示宽度或高度（等比缩放），支持居中/左右对齐。

**实现代码：**

```python
from docx.shared import Inches

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = p.add_run()
run.add_picture("chart.png", width=Inches(5))
```

**适用场景：** 报告中的数据图表截图、产品说明书中的实物照片、流程示意图。

**提醒：** 只指定 `width` 时会自动按原图比例计算 `height`，如果两者都指定但比例与原图不符，图片会被拉伸变形；插入前建议确认图片文件存在且路径正确，否则会抛出 `FileNotFoundError`。

---

### 7.2 图片题注（Figure Caption）

**长什么样：** 图片下方一行小字，格式通常是「图 1-1 XX 流程图」，与表格题注的排版逻辑基本一致。

**实现代码：**

```python
caption = doc.add_paragraph()
caption.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = caption.add_run("图 1  用户注册流程示意图")
run.font.size = Pt(10)
run.font.color.rgb = RGBColor(0x59, 0x59, 0x59)
```

**适用场景：** 论文插图、产品手册中需要被正文引用的示意图。

**提醒：** 与表格题注同理，严谨的自动编号需要 SEQ 域支持，这里给出的是无需自动编号时的简化写法。

---

### 7.3 文本框（Text Box）

**长什么样：** 一个可以脱离正文流、自由摆放在页面任意位置的矩形容器，常用来做旁批、引言框、侧边栏摘要等「悬浮」效果。

**实现代码：**

```python
# python-docx 没有原生 add_textbox() API，需要拼接 DrawingML/VML 片段后
# 插入到段落的 run 中（以下为概念性片段）
textbox_xml = (
    '<w:pict xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
    '<v:shape style="width:200pt;height:80pt" fillcolor="#F2F2F2">'
    '<v:textbox><w:txbxContent>'
    '<w:p><w:r><w:t>这是一段侧边批注文字</w:t></w:r></w:p>'
    '</w:txbxContent></v:textbox>'
    '</v:shape></w:pict>'
)
```

**适用场景：** 论文里的「术语解释」旁批框、杂志排版中的引言摘录、说明书里的「小贴士」提示框。

**提醒：** 文本框是 `python-docx` 支持最薄弱的模块之一，几乎完全依赖手写 XML，兼容性和渲染效果强烈建议在 Word 和 LibreOffice 中分别核对；如果只是想要「带边框的一段话」这种简单效果，优先考虑用 5.2 节的单元格底纹 + 无边框单元格表格来模拟，会比真正的文本框稳定得多。

---

## 八、引用与注释类模块

### 8.1 脚注（Footnote）

**长什么样：** 正文中一个小号上标数字，对应页面底部的注释说明，阅读时不打断正文视线。

**实现代码：**

```python
# python-docx 官方 API 不直接支持脚注，以下为需要底层 XML 支持的概念性示例
def add_footnote_ref(paragraph, footnote_id):
    run = paragraph.add_run()
    rPr = OxmlElement('w:rPr')
    rStyle = OxmlElement('w:rStyle')
    rStyle.set(qn('w:val'), 'FootnoteReference')
    rPr.append(rStyle)
    ref = OxmlElement('w:footnoteReference')
    ref.set(qn('w:id'), str(footnote_id))
    run._r.append(rPr)
    run._r.append(ref)
```

**适用场景：** 学术论文中的引用出处、术语首次出现时的补充说明、法律文书中的条款备注。

**提醒：** `python-docx` 原生完全不支持脚注/尾注，必须自己创建并维护 `word/footnotes.xml` 及其与正文的关系文件，操作复杂度较高，实践中更推荐借助社区维护的扩展库，或者退而求其次用「文末统一注释列表 + 手动上标数字」来模拟。

---

### 8.2 尾注（Endnote）

**长什么样：** 与脚注类似，但注释内容统一汇总在全文/章节末尾，而不是每页底部。

**实现代码：**

```python
# 原理与脚注几乎一致，区别在于对应的是 word/endnotes.xml，
# 域引用类型为 w:endnoteReference 而非 w:footnoteReference
ref = OxmlElement('w:endnoteReference')
ref.set(qn('w:id'), '1')
```

**适用场景：** 大部头著作、需要把大量参考文献集中放在书末的学术专著。

**提醒：** 与脚注同理，`python-docx` 无原生支持，工程量不小；如果文档篇幅不长，脚注的阅读体验通常优于尾注，可优先考虑脚注方案。

---

### 8.3 超链接（Hyperlink）

**长什么样：** 一段带下划线、通常显示为蓝色的可点击文字，点击后跳转到网址或文档内部书签。

**实现代码：**

```python
def add_hyperlink(paragraph, url, text):
    part = paragraph.part
    r_id = part.relate_to(
        url,
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
        is_external=True
    )
    hyperlink = OxmlElement('w:hyperlink')
    hyperlink.set(qn('r:id'), r_id)
    run = OxmlElement('w:r')
    rPr = OxmlElement('w:rPr')
    rStyle = OxmlElement('w:rStyle')
    rStyle.set(qn('w:val'), 'Hyperlink')
    rPr.append(rStyle)
    run.append(rPr)
    t = OxmlElement('w:t')
    t.text = text
    run.append(t)
    hyperlink.append(run)
    paragraph._p.append(hyperlink)

add_hyperlink(doc.add_paragraph(), "https://python-docx.readthedocs.io", "python-docx 官方文档")
```

**适用场景：** 参考文献列表、附录中的资料来源、文档内「点击跳转章节」的目录式导航。

**提醒：** `python-docx` 没有 `add_hyperlink()` 这样的高层方法，必须手动创建关系（relationship）并拼接 `w:hyperlink` 节点；文字颜色/下划线要靠引用内置的 Hyperlink 字符样式才会呈现「传统蓝色带下划线」的观感，否则默认是纯黑色纯文本外观。

---

### 8.4 书签与交叉引用（Bookmark & Cross-reference）

**长什么样：**「书签」是给文档中某个位置打的隐形标记（不显示任何符号）；「交叉引用」则是正文中引用该书签的文字，如「详见第 3.2 节」，并且会随实际章节变化自动更新。

**实现代码：**

```python
def add_bookmark(paragraph, bookmark_name, bookmark_id):
    start = OxmlElement('w:bookmarkStart')
    start.set(qn('w:id'), str(bookmark_id))
    start.set(qn('w:name'), bookmark_name)
    end = OxmlElement('w:bookmarkEnd')
    end.set(qn('w:id'), str(bookmark_id))
    paragraph._p.insert(0, start)
    paragraph._p.append(end)
```

**适用场景：** 长文档内部的「详见 XX 章节」式相互引用、需要跳转定位的电子合同。

**提醒：** 交叉引用本质上也是一种域代码（REF 域），同样需要在 Word 中「更新域」才会显示真实的章节号/页码；`python-docx` 对这一整套机制均无高层封装，复杂交叉引用建议评估投入产出比，或考虑生成后用 Word 宏批量处理。

---

## 九、高级与自定义类模块

### 9.1 自定义样式（Custom Style）

**长什么样：** 在内置样式（Normal、Heading 1 等）之外，自己定义一套新样式（如「警示框文字」），并赋予专属的字体、颜色、边框、底纹组合，可在全文任意位置复用。

**实现代码：**

```python
from docx.enum.style import WD_STYLE_TYPE

styles = doc.styles
warning_style = styles.add_style('警示文字', WD_STYLE_TYPE.PARAGRAPH)
warning_style.font.color.rgb = RGBColor(0xC0, 0x00, 0x00)
warning_style.font.bold = True
warning_style.paragraph_format.left_indent = Pt(20)

doc.add_paragraph("请勿在通电状态下拆卸设备", style='警示文字')
```

**适用场景：** 需要在文档中反复出现「同一种视觉效果」的场景（如多处警示语、多处引用块），定义一次样式比每次手动设置属性更高效、更易维护。

**提醒：** 自定义样式名称若与模板中已有样式重名，`add_style` 会抛出异常；建议先检查 `[s.name for s in doc.styles]` 确认没有冲突。

---

### 9.2 段落边框分隔线（Paragraph Border as Divider）

**长什么样：** 一条横贯页面宽度的细线，常用于章节之间的视觉分隔，效果类似 Markdown 里的 `---`。

**实现代码：**

```python
def add_horizontal_line(paragraph):
    p_pr = paragraph._p.get_or_add_pPr()
    p_borders = OxmlElement('w:pBdr')
    bottom = OxmlElement('w:bottom')
    bottom.set(qn('w:val'), 'single')
    bottom.set(qn('w:sz'), '6')
    bottom.set(qn('w:space'), '1')
    bottom.set(qn('w:color'), 'auto')
    p_borders.append(bottom)
    p_pr.append(p_borders)

add_horizontal_line(doc.add_paragraph())
```

**适用场景：** 封面与正文之间的视觉分隔、章节标题下方的装饰线、简历中模块之间的分割。

**提醒：** 不要用「单行单列的表格」去模拟分隔线——这在部分渲染器（尤其是移动端 Word 或在线预览）中可能出现意料之外的边距问题，官方也明确建议用段落边框（`pBdr`）实现，而不是拿表格「作弊」。

---

### 9.3 首字下沉（Drop Cap）

**长什么样：** 段落第一个字被放大数倍，并且后续几行文字围绕这个大字排列，常见于书籍章节的开篇段落，视觉上很有「翻开一本书」的仪式感。

**实现代码：**

```python
# python-docx 无原生支持，需要直接操作段落属性节点 w:framePr
p_pr = paragraph._p.get_or_add_pPr()
frame_pr = OxmlElement('w:framePr')
frame_pr.set(qn('w:dropCap'), 'drop')
frame_pr.set(qn('w:lines'), '3')  # 下沉三行高度
p_pr.append(frame_pr)
```

**适用场景：** 小说/散文类书籍排版、杂志专栏开篇、需要「精致感」的邀请函类文档。

**提醒：** 首字下沉对段落结构要求较严格（首字通常需要单独成一个 run），效果强烈依赖渲染引擎，建议生成后务必用真实 Word 打开核对，不要只信任代码「跑通了不报错」。

---

### 9.4 批注与修订模式（Comments & Track Changes）

**长什么样：**「批注」是附加在选中文字旁边的气泡式评论，不改变正文内容；「修订模式」则会把增删的文字用红色下划线/删除线直接体现在正文中，并记录修改者和时间。这两者是 Word「协作审阅」场景的核心功能。

**实现代码：**

```python
# python-docx 完全不支持批注和修订的生成/解析，
# 需要直接维护 word/comments.xml 及其配套的
# commentsExtended.xml / commentsIds.xml，
# 并在正文段落中插入 w:commentRangeStart / w:commentRangeEnd / w:commentReference
# 三个节点标记批注所覆盖的文字范围（以下为概念示意，非完整可运行代码）：
comment_range_start = OxmlElement('w:commentRangeStart')
comment_range_start.set(qn('w:id'), '0')
paragraph._p.insert(0, comment_range_start)
```

**适用场景：** 合同审阅稿、论文导师批注、多人协作的文档审校流程。

**提醒：** 这是 `python-docx` 支持度最低的模块之一，官方文档明确未覆盖；如果任务的核心诉求就是生成/操作批注和修订，更推荐直接对 OOXML 底层编辑（解压 docx、手改 XML、重新打包），而不是试图完全依赖 `python-docx` 的高层 API。

---

## 十、模块速查表

| 模块名称 | 所属类别 | python-docx 原生支持 | 典型适用场景 |
|---|---|---|---|
| 封面 | 页面结构 | 需手动堆砌段落 | 报告、论文、标书首页 |
| 分节符 | 页面结构 | 部分支持 | 分节独立页码/纸张方向 |
| 页眉页脚 | 页面结构 | 支持 | 企业报告、合同 |
| 页码域 | 页面结构 | 需手写域代码 | 多页正式文档 |
| 分栏 | 页面结构 | 需底层 XML | 词典、报纸式排版 |
| 水印 | 页面结构 | 需手写 VML | 草稿标识、保密文件 |
| 标题样式 | 标题目录 | 支持 | 章节层级 |
| 正文样式 | 标题目录 | 支持 | 全文默认排版 |
| 目录域 | 标题目录 | 需手写域代码 | 长文档导航 |
| 项目符号列表 | 列表 | 支持 | 并列特性罗列 |
| 编号列表 | 列表 | 支持 | 步骤、排名 |
| 多级列表 | 列表 | 部分支持 | 合同条款、规范文档 |
| 基础表格/内置样式 | 表格 | 支持 | 数据展示 |
| 单元格底纹 | 表格 | 需底层 XML | 表头/重点行强调 |
| 合并单元格 | 表格 | 支持 | 多级表头 |
| 表格题注 | 表格 | 需手写域代码（自动编号） | 论文、报告配图表 |
| 加粗/斜体/下划线 | 字符格式 | 支持 | 局部强调 |
| 颜色与高亮 | 字符格式 | 支持（高亮色有限） | 风险提示 |
| 自定义字体字号 | 字符格式 | 支持（需设东亚字体） | 格式规范类文档 |
| 图片插入 | 图片对象 | 支持 | 图表、照片 |
| 图片题注 | 图片对象 | 需手写域代码（自动编号） | 论文插图 |
| 文本框 | 图片对象 | 需手写 VML/DrawingML | 旁批、提示框 |
| 脚注 | 引用注释 | 不支持（需扩展/手写 XML） | 学术引用出处 |
| 尾注 | 引用注释 | 不支持（需扩展/手写 XML） | 书末参考文献 |
| 超链接 | 引用注释 | 需手动建关系 | 参考资料来源 |
| 书签与交叉引用 | 引用注释 | 需手写域代码 | 章节互相引用 |
| 自定义样式 | 高级自定义 | 支持 | 重复出现的视觉效果 |
| 段落边框分隔线 | 高级自定义 | 需底层 XML | 章节视觉分隔 |
| 首字下沉 | 高级自定义 | 需底层 XML | 书籍/杂志排版 |
| 批注与修订 | 高级自定义 | 不支持（需直接编辑 OOXML） | 协作审阅、合同批注 |

---

## 十一、常见报错与注意事项汇总

- **`KeyError: no style with name 'XXX'`** → 样式名称拼写错误或模板中不存在该样式，先打印 `doc.styles` 列表确认。
- **中文显示为默认字体而非指定字体** → 忘记设置 `w:eastAsia` 东亚字体节点。
- **表格底纹显示为纯黑色** → `w:shd` 的 `w:val` 误用了 `solid`，应使用 `clear`。
- **生成的页码/目录显示为空白或 0** → 域代码需要在 Word 中「更新域」（`Ctrl+A` → `F9`）才会计算出真实数值，这是 Word 客户端行为，不是代码问题。
- **`add_picture` 抛出 `FileNotFoundError`** → 图片路径错误或相对路径与运行目录不一致，建议统一用绝对路径。
- **合并单元格后内容错乱** → 合并操作会改变原有单元格对象的引用关系，合并后应统一对返回的新单元格对象赋值。
- **自定义样式抛出「样式已存在」异常** → 样式名称与内置或已定义样式重复，添加前先检查是否已存在。
- **插入的水印/文本框在 Word 中不显示** → 手写的 XML 片段命名空间或节点顺序有误，这类底层 XML 对结构非常敏感，建议参考真实 Word 生成的 `document.xml` 逐字段比对。

---

## 十二、完整文档骨架示例

将上述模块串联成一份可运行的最小骨架（省略了部分辅助函数的重复定义，实际使用时需将本文前述的 `add_page_number`、`set_cell_background` 等函数一并复制进脚本）：

```python
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

doc = Document()

# 1. 全局正文样式
normal = doc.styles['Normal']
normal.font.name = '微软雅黑'
normal.font.size = Pt(12)
normal.element.rPr.rFonts.set(qn('w:eastAsia'), '微软雅黑')

# 2. 封面
title = doc.add_paragraph()
title.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = title.add_run("示例报告标题")
run.font.size = Pt(32)
run.font.bold = True
doc.add_page_break()

# 3. 目录页（占位，实际域代码见 3.3 节）
doc.add_heading("目录", level=1)
doc.add_page_break()

# 4. 正文章节
doc.add_heading("第一章 概述", level=1)
doc.add_paragraph("这里是正文内容……")

doc.add_paragraph("要点一", style="List Bullet")
doc.add_paragraph("要点二", style="List Bullet")

table = doc.add_table(rows=2, cols=2)
table.style = "Light Grid Accent 1"
table.cell(0, 0).text = "指标"
table.cell(0, 1).text = "数值"

# 5. 页脚页码（函数定义见 2.4 节）
section = doc.sections[0]
footer_para = section.footer.paragraphs[0]
footer_para.alignment = WD_ALIGN_PARAGRAPH.CENTER

# 6. 保存
doc.save("report.docx")
```

**核心思路：** 先定好全局样式（Normal + Heading），再按「封面 → 目录 → 正文（标题/列表/表格/图片穿插）→ 页眉页脚」的顺序依次拼装，遇到 `python-docx` 没有高层封装的模块（页码、目录、脚注、水印、批注等）时，统一通过 `docx.oxml` 提供的 `OxmlElement` / `qn` 工具直接操作底层 XML 节点来补齐。
