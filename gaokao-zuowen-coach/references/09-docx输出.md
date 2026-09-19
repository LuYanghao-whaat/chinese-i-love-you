# Word 文档输出（.docx）

> 当用户在开场选“输出格式 = Word 文档”时，用本模块。
> 工具链：Python 3.12 + python-docx。完整排版规范已归档在 `references/docx/`：
> `docx-skill-pydocx-zh.md`、`docx-guide-pydocx-zh.md`、`visual-style-spec-zh.md`。

这个模块不是一套固定模板，而是一个组件库：按内容挑模块、拼装成文档。
先定“文档类型”，再定“视效档位”，然后从组件库里挑件。

---

## 零、先定文档类型（批改报告 / 讲评讲义）

同一个 `.docx`，可以是两种东西。生成前先问清楚（或按触发语判断）。

| | 批改报告 | 讲评讲义 |
|---|---|---|
| 面向 | 一篇作文的作者本人 | 一个班 / 多篇作文 / 一个题目 |
| 目的 | 诊断这一篇 + 给分 + 改法 | 教这一类题怎么写 + 横向对比 |
| 结构 | 总评 → 扣题 → 分段评语 → 语言 → 逻辑链 → 评分 → 建议 | 原题再现 → 审题 → 立意 → 陷阱 → 作文分析 → 总结 |
| 重点组件 | 评分卡、逐段色块、逻辑链表、建议清单 | 原题框、方法列表、范文/多篇对照、陷阱警示块 |
| 触发语 | “批改这篇”“打个分”“给份报告” | “讲评”“讲义”“给全班”“分析这几篇” |

两者不互斥：讲义里的“作文分析”一节，可以直接嵌一份批改报告。
拿不准时，默认按批改报告（单篇请求最常见）；用户提到“班/多篇/讲评”就转讲义。

---

## 一、再定视效档位（三档，都生成 .docx）

| 档位 | 保留什么 | 不用什么 |
|---|---|---|
| 普通（默认） | 封面 + 分节标题 + 正文 + 页码 | 目录、色块 |
| 美观 | 再加目录、点评色块、数据表、封面页 | — |
| 极简 | 只保留 Markdown 级格式：标题、列表、表格、加粗斜体 | 颜色、底纹、边框、自定义字体 |

---

## 二、环境准备（默认帮用户装好）

1. 查 Python 3.12：`py -3.12 --version`
2. 查库：`py -3.12 -m pip show python-docx`
3. 缺 Python 3.12 → `winget install Python.Python.3.12`
4. 缺库 → `py -3.12 -m pip install python-docx`

装完复验，再生成。不要把带报错的半成品丢给用户。

---

## 三、字体与配色（组件库的公共常量）

字体：正文标题用 `宋体`，引用/点评/色块内文字用 `楷体`。
配色：一套克制的主色 + 语义色，别每处随手换。
**标题一律用深色文字，不用大面积实心色块**（色块只用于批注、亮点、提示这类小面积元素）。

```python
FONT_SONG = "宋体"
FONT_KAI = "楷体"

C_PRIMARY = (44, 74, 102)      # 墨蓝：标题文字（柔和，不做实心背景块）
C_ACCENT  = (46, 117, 182)     # 蓝：二级标题
C_TEXT    = (51, 51, 51)       # 正文
C_GRAY    = (120, 120, 120)    # 辅助信息
C_GOOD    = (39, 111, 66)      # 绿：亮点
C_WARN    = (150, 84, 20)      # 棕：问题
C_SCORE   = (176, 58, 46)      # 红：分数

HEX_PRIMARY = "2C4A66"         # 标题文字色（深色只上文字，不做背景块）
HEX_HEADER  = "E3EBF3"         # 表头浅蓝底
HEX_ZEBRA   = "F5F8FB"         # 表格隔行浅底
HEX_RULE    = "C9D8E8"         # 分隔线
```

---

## 四、工作流（动态拼装）

1. 读原始内容，先判文档类型（批改报告 / 讲义）。
2. 按类型选骨架顺序（见第八节），但骨架只是顺序，具体每节用哪个组件看内容。
3. 逐节挑组件：这段是并列要点就用列表，是对照就用两栏/表格，是重点结论就用色块，是分数就用评分卡。
4. 写脚本、本地跑通。
5. 生成后处理（有 Word 就做）：打开文档、更新域、保存，把目录/页码烘焙进去，避免打开弹“是否更新域”；顺带导出 PDF 做审计。
6. 渲染审计（第七节）：逐页看图，有问题就改、重生成。

动态化原则：不要每篇都用同一串函数。内容决定组件——
- 有对比 → 两栏或表格；纯叙述 → 正文 + 色块。
- 有分数 → 评分卡；只有建议 → 编号列表。
- 讲义的章节标题用下划线（`h1_banner`）；批改报告的节标题用竖条（`h1_bar`），更克制。**标题都是纯文字，不铺背景色块。**

---

## 五、组件库

下面每个函数都可独立调用、自由组合。整体复制进脚本即可。

```python
import os
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_COLOR_INDEX
from docx.enum.section import WD_SECTION
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

# ============ A. 底层工具 ============

def set_run(run, font=FONT_SONG, size=12, color=None, bold=False, italic=False, highlight=None):
    run.font.name = font
    run._element.rPr.rFonts.set(qn('w:eastAsia'), font)
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.italic = italic
    if color:
        run.font.color.rgb = RGBColor(*color)
    if highlight:
        run.font.highlight_color = highlight

def shading(paragraph, hex_color):
    pPr = paragraph._p.get_or_add_pPr()
    for el in pPr.findall(qn('w:shd')):
        pPr.remove(el)
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')          # 必须 clear，否则黑块
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), hex_color)
    pPr.append(shd)

def cell_shading(cell, hex_color):
    """给整个单元格填色。表格填色一律用它，不要用 shading()，
    否则只填段落、四周留白边。"""
    tcPr = cell._tc.get_or_add_tcPr()
    for el in tcPr.findall(qn('w:shd')):
        tcPr.remove(el)
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), hex_color)
    tcPr.append(shd)

def border(paragraph, sides, color=HEX_RULE, sz="6", space="4"):
    pPr = paragraph._p.get_or_add_pPr()
    pBdr = OxmlElement('w:pBdr')
    for side in sides:
        b = OxmlElement(f'w:{side}')
        b.set(qn('w:val'), 'single'); b.set(qn('w:sz'), sz)
        b.set(qn('w:space'), space); b.set(qn('w:color'), color)
        pBdr.append(b)
    pPr.append(pBdr)

def page_number(paragraph):
    def fld(p, t):
        r = p.add_run()
        b = OxmlElement('w:fldChar'); b.set(qn('w:fldCharType'), 'begin')
        i = OxmlElement('w:instrText'); i.text = t
        e = OxmlElement('w:fldChar'); e.set(qn('w:fldCharType'), 'end')
        r._r.extend([b, i, e])
    paragraph.add_run("第 "); fld(paragraph, "PAGE")
    paragraph.add_run(" 页 / 共 "); fld(paragraph, "NUMPAGES"); paragraph.add_run(" 页")

# ============ B. 全局设置 ============

def setup_global_styles(doc, body_size=12, line=1.5):
    n = doc.styles['Normal']
    n.font.name = FONT_SONG; n.font.size = Pt(body_size)
    n.element.rPr.rFonts.set(qn('w:eastAsia'), FONT_SONG)
    n.font.color.rgb = RGBColor(*C_TEXT)
    n.paragraph_format.line_spacing = line
    n.paragraph_format.space_after = Pt(6)

def setup_page_section(doc, index=0, footer=True):
    s = doc.sections[index]
    s.page_width, s.page_height = Inches(8.27), Inches(11.69)
    s.left_margin = s.right_margin = s.top_margin = s.bottom_margin = Inches(1)
    if footer:
        s.footer.is_linked_to_previous = False
        fp = s.footer.paragraphs[0]; fp.clear()
        fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
        page_number(fp)

def set_doc_settings(doc):
    """兼容性模式 15。不要设 updateFields=true —— 那会让 Word 每次打开都弹“是否更新域”。"""
    st = doc.settings.element
    for el in st.findall(qn('w:compat')): st.remove(el)
    compat = OxmlElement('w:compat')
    cs = OxmlElement('w:compatSetting')
    cs.set(qn('w:name'), 'compatibilityMode')
    cs.set(qn('w:uri'), 'http://schemas.microsoft.com/office/word')
    cs.set(qn('w:val'), '15')
    compat.append(cs); st.append(compat)
    for el in st.findall(qn('w:updateFields')): st.remove(el)

def header_text(section, text, align=WD_ALIGN_PARAGRAPH.RIGHT):
    section.header.is_linked_to_previous = False
    p = section.header.paragraphs[0]; p.clear(); p.alignment = align
    set_run(p.add_run(text), FONT_SONG, 9, C_GRAY)

def set_columns(section, num=2):
    cols = section._sectPr.xpath('./w:cols')[0]
    cols.set(qn('w:num'), str(num)); cols.set(qn('w:space'), '425')

# ============ C. 封面（三选一） ============

def cover_centered(doc, title, subtitle, info):
    for _ in range(6): doc.add_paragraph()
    p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_run(p.add_run(title), FONT_SONG, 26, C_PRIMARY, bold=True)
    if subtitle:
        p2 = doc.add_paragraph(); p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
        set_run(p2.add_run(subtitle), FONT_KAI, 15, C_GRAY)
    rule = doc.add_paragraph(); rule.alignment = WD_ALIGN_PARAGRAPH.CENTER
    border(rule, ['bottom'], color=HEX_PRIMARY, sz="8")
    for _ in range(8): doc.add_paragraph()
    p3 = doc.add_paragraph(); p3.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_run(p3.add_run(info), FONT_SONG, 11, C_GRAY)
    doc.add_page_break()

def cover_banner(doc, title, subtitle, info):
    """封面：蓝色宋体大标题 + 一条细线（无背景色块）"""
    for _ in range(6): doc.add_paragraph()
    p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_run(p.add_run(title), FONT_SONG, 28, C_PRIMARY, bold=True)
    rule = doc.add_paragraph(); rule.alignment = WD_ALIGN_PARAGRAPH.CENTER
    border(rule, ['bottom'], color=HEX_PRIMARY, sz="8")
    if subtitle:
        p2 = doc.add_paragraph(); p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
        set_run(p2.add_run(subtitle), FONT_KAI, 15, C_GRAY)
    for _ in range(8): doc.add_paragraph()
    p3 = doc.add_paragraph(); p3.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_run(p3.add_run(info), FONT_SONG, 11, C_GRAY)
    doc.add_page_break()

def cover_minimal(doc, title, subtitle=""):
    p = doc.add_paragraph(); p.paragraph_format.space_before = Pt(160)
    set_run(p.add_run(title), FONT_SONG, 22, C_PRIMARY, bold=True)
    if subtitle:
        p2 = doc.add_paragraph()
        set_run(p2.add_run(subtitle), FONT_KAI, 13, C_GRAY)
    doc.add_page_break()

# ============ D. 标题与分隔 ============

def h1_bar(doc, text):
    """左竖条，批改报告的节标题（克制）"""
    p = doc.add_paragraph()
    p.paragraph_format.space_before, p.paragraph_format.space_after = Pt(18), Pt(8)
    p.paragraph_format.left_indent = Inches(0.12)
    set_run(p.add_run(text), FONT_SONG, 16, C_PRIMARY, bold=True)
    border(p, ['left'], color=HEX_PRIMARY, sz="18", space="6")

def h1_banner(doc, text):
    """章节标题：蓝色宋体纯文字 + 下划线（无背景色块）"""
    p = doc.add_paragraph()
    p.paragraph_format.space_before, p.paragraph_format.space_after = Pt(16), Pt(8)
    set_run(p.add_run(text), FONT_SONG, 17, C_PRIMARY, bold=True)
    border(p, ['bottom'], color=HEX_RULE, sz="6")

def h2(doc, text):
    p = doc.add_paragraph()
    p.paragraph_format.space_before, p.paragraph_format.space_after = Pt(12), Pt(4)
    set_run(p.add_run(text), FONT_SONG, 13, C_ACCENT, bold=True)
    border(p, ['bottom'], color=HEX_RULE)

def h3(doc, text):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(8)
    set_run(p.add_run(text), FONT_SONG, 12, C_PRIMARY, bold=True)

def divider(doc, text=""):
    p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_run(p.add_run(text or " "), FONT_SONG, 9, C_GRAY)
    border(p, ['bottom'], color=HEX_RULE)

# ============ E. 正文与文本 ============

def body(doc, text):
    doc.add_paragraph(text)

def body_indent(doc, text):
    """首行缩进两字，正文更中文"""
    p = doc.add_paragraph(text)
    p.paragraph_format.first_line_indent = Pt(24)

def lead(doc, text):
    """导语：楷体灰"""
    p = doc.add_paragraph()
    set_run(p.add_run(text), FONT_KAI, 12, C_GRAY)

def quote(doc, text):
    """引用：楷体 + 左竖条"""
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Inches(0.2)
    set_run(p.add_run(text), FONT_KAI, 12, (60, 60, 60))
    border(p, ['left'], color=HEX_RULE, sz="18", space="8")

def caption(doc, text):
    p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_run(p.add_run(text), FONT_SONG, 9, C_GRAY, italic=True)

# ============ F. 列表 ============

def bullet(doc, text, color=None):
    p = doc.add_paragraph(); p.paragraph_format.left_indent = Inches(0.3)
    set_run(p.add_run("· " + text), FONT_SONG, 12, color or C_TEXT)

def numbered(doc, idx, text):
    p = doc.add_paragraph(); p.paragraph_format.left_indent = Inches(0.3)
    set_run(p.add_run(f"{idx}. "), FONT_SONG, 12, C_ACCENT, bold=True)
    set_run(p.add_run(text), FONT_SONG, 12, C_TEXT)

def kv(doc, key, value):
    p = doc.add_paragraph()
    set_run(p.add_run(key + "："), FONT_SONG, 12, C_PRIMARY, bold=True)
    set_run(p.add_run(value), FONT_SONG, 12, C_TEXT)

# ============ G. 色块卡片（语义化） ============

_BLOCK = {
    "info":  ("EAF2FA", "B9D4EC", (40, 40, 40)),
    "good":  ("EAF5EC", "B7DCC0", (30, 80, 45)),
    "warn":  ("FBF3E7", "E6D2AE", (120, 70, 15)),
    "score": ("FDECEA", "F0C0B8", (150, 40, 30)),
    "plain": ("F7F8FA", "E2E6EA", (51, 51, 51)),
}

def block(doc, text, kind="info"):
    bg, bd, fg = _BLOCK.get(kind, _BLOCK["info"])
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Inches(0.15)
    p.paragraph_format.right_indent = Inches(0.15)
    p.paragraph_format.space_before, p.paragraph_format.space_after = Pt(6), Pt(6)
    set_run(p.add_run(text), FONT_KAI, 12, fg)
    shading(p, bg)
    border(p, ['top', 'left', 'bottom', 'right'], color=bd, sz="4")

def score_card(doc, total, level, dims):
    """评分卡：总分 + 等级 + 维度行。dims = [(名称, 得分, 说明), ...]"""
    p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_run(p.add_run(f"{total}"), FONT_SONG, 30, C_SCORE, bold=True)
    set_run(p.add_run(" / 60"), FONT_SONG, 14, C_GRAY)
    p2 = doc.add_paragraph(); p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_run(p2.add_run(level), FONT_SONG, 13, C_PRIMARY, bold=True)
    for name, sc, note in dims:
        kv(doc, name, f"{sc}　{note}")

# ============ H. 表格 ============

def table(doc, matrix, header_hex=HEX_HEADER, zebra=HEX_ZEBRA, widths=None):
    t = doc.add_table(rows=len(matrix), cols=len(matrix[0]))
    t.style = 'Table Grid'
    layout = OxmlElement('w:tblLayout'); layout.set(qn('w:type'), 'fixed')
    t._tbl.tblPr.append(layout)
    if widths:
        for r in t.rows:
            for i, w in enumerate(widths):
                r.cells[i].width = Inches(w)
    for ri, row in enumerate(matrix):
        for ci, val in enumerate(row):
            cell = t.cell(ri, ci); cell.text = str(val)
            if ri == 0:
                cell_shading(cell, header_hex)
            elif zebra and ri % 2 == 0:
                cell_shading(cell, zebra)
            for para in cell.paragraphs:
                for run in para.runs:
                    set_run(run, FONT_SONG, 10,
                            C_PRIMARY if ri == 0 else C_TEXT, bold=(ri == 0))
    return t

def logic_table(doc, rows):
    """逻辑链审查专用：环节 / 说法 / 判断 / 说明"""
    table(doc, [["环节", "文章的说法", "判断", "说明"]] + rows, widths=[0.9, 2.3, 0.9, 2.4])

def side_table(doc, rows, widths=(3.1, 2.4)):
    """旁批：两列，左列原文、右列批注。rows = [(原文段落, 批注), ...]"""
    t = doc.add_table(rows=len(rows), cols=2)
    t.style = 'Table Grid'
    layout = OxmlElement('w:tblLayout'); layout.set(qn('w:type'), 'fixed')
    t._tbl.tblPr.append(layout)
    for ri, (src, note) in enumerate(rows):
        c0, c1 = t.cell(ri, 0), t.cell(ri, 1)
        c0.width, c1.width = Inches(widths[0]), Inches(widths[1])
        set_run(c0.paragraphs[0].add_run(src), FONT_SONG, 10, C_TEXT)
        p1 = c1.paragraphs[0]
        set_run(p1.add_run(note), FONT_KAI, 10, C_WARN)
        cell_shading(c1, "FBF3E7")
    return t

# ============ I. 图片与高级（按需） ============

def picture(doc, path, width_in=5):
    p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.add_run().add_picture(path, width=Inches(width_in))

def hyperlink(paragraph, url, text):
    part = paragraph.part
    r_id = part.relate_to(url,
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
        is_external=True)
    link = OxmlElement('w:hyperlink'); link.set(qn('r:id'), r_id)
    r = OxmlElement('w:r'); t = OxmlElement('w:t'); t.text = text
    r.append(t); link.append(r); paragraph._p.append(link)

def dropcap(paragraph, lines=3):
    pPr = paragraph._p.get_or_add_pPr()
    fp = OxmlElement('w:framePr')
    fp.set(qn('w:dropCap'), 'drop'); fp.set(qn('w:lines'), str(lines))
    pPr.append(fp)
```

---

## 六、致命错误规避

1. 中英文字体失效：文档级改 `Normal` 的 `w:eastAsia`。
2. 底纹黑块：`w:shd` 的 `w:val` 必须是 `"clear"`。
3. 表格列宽错乱：加 `w:tblLayout type="fixed"`。
4. 页眉页脚串扰：多节文档先 `is_linked_to_previous = False`。
5. 换行崩溃：`add_run()` 里不能用 `\n`。
6. 兼容性模式：必须调 `set_doc_settings()`，把 `compatibilityMode` 设为 15。
7. 目录空白页 / 打开弹窗：TOC 和页码是域，python-docx 不会渲染，Word 打开时可能是空白。
   - 不要设 `updateFields=true`——那会让 Word 每次打开都弹“该文档包含的域可能引用了其他文件，是否更新”。
   - 正确做法：生成后用 Word 打开一次、更新域、保存（见第六节“生成后处理”），把目录烘焙进文档，之后打开不再弹窗。
   - 普通 / 极简档不放目录，只有美观档放。

---

## 七、生成后审计（强制，不许跳过）

生成完 `.docx` 不算完。有 Word 时，先更新域并保存（烘焙目录/页码，避免打开弹窗），再渲染成图审计：

```powershell
$w = New-Object -ComObject Word.Application; $w.Visible = $false; $w.DisplayAlerts = 0
$d = $w.Documents.Open("C:\path\报告.docx")
$d.Fields.Update() | Out-Null
$d.Save()                                   # 把更新后的域烘焙进 docx
$d.ExportAsFixedFormat("C:\path\报告.pdf", 17)
$d.Close($false); $w.Quit()
```

然后把文档渲染成逐页图片，自己逐页看一遍，确认排版没问题，再交付。
渲染方法自选——导出 PDF 再转图、用 Word 截图、或别的方式都行，关键是“真的看到每一页”。

逐页检查：

- 有没有整页空白（尤其目录页）。
- 标题层级、色块、表格有没有错位、压字、串行。
- 有没有孤行、断表、封面内容溢出到第二页。
- 中文有没有退化成默认字体、底纹有没有变成黑块。
- 旁批表格的左右两列是否对齐、有没有被挤扁。

发现问题就改、重新生成、再看，直到干净。没渲染看过就交付，等于没做。

---

## 八、两种文档类型的推荐拼装

### 批改报告（单篇）

```
cover_centered 或 cover_minimal
h1_bar 一、总评              → body + block(“info”)
h1_bar 二、审题与评价标准     → 材料在问什么 / 评价标准 / 本文扣题定位
                              （body + table 或 block；透明化 AI 的审题与尺度）
h1_bar 三、分段点评（批注式）  → side_table（左原文、右批注，不加斑马纹）+ 段内批【】 + 段末批
h1_bar 四、语言与综合         → bullet（优点/问题）
h1_bar 五、逻辑链审查         → logic_table
h1_bar 六、参考评分           → score_card
h1_bar 七、修改建议           → numbered
```
节标题用竖条（克制）；只有“亮点/问题/结论”用色块；分数用评分卡。
第二节“审题与评价标准”是给读者的“透明化”模块：让作者看到 AI 怎么读题、按什么尺子给分，可复核、也能学审题。

第三节“分段点评”一律用批注式，完整展示原文（一字不落），**三种批注都要有，缺一不可**：

| 形式 | 用在 | docx 实现 |
|---|---|---|
| 旁批（主力） | 逐段对照，每个主体段一条 | `side_table`（两列：左原文、右批注）。**旁批表不加斑马纹**：左列白底、右列米色即可 |
| 段内批 | 局部字句、逻辑跳步、错别字 | 左列原文里直接插 `【...】` |
| 段末批 | 整段的总结性判断 | 原文段之后另起一段 body/block |

**【】内是老师即时点评：口语、直接、带判断，不必端着。** 三种示例：
- 夸：`【这句话是金句。原因在于……】`
- 挑错：`【这里逻辑跳脱了。“XX”并不等同于“XX”。】`
- 提醒：`【“公立”疑为“功利”笔误】`

做法：把各段原文按顺序放进 `side_table` 左列（局部问题用 `【】` 标出），右列写旁批；
每段之后可再补一句段末批。**不要只给右列旁批、把【】丢了。**

### 讲评讲义（班级 / 多篇 / 一题）

```
cover_banner
h1_banner 一、原题再现   → quote（原题一字不改）
h1_banner 二、审题指导   → h2 + body + bullet
h1_banner 三、立意指导   → table（立意对照）
h1_banner 四、陷阱提示   → block(“warn”)
h1_banner 五、作文分析   → 可嵌多份“批改报告”的精简版
h1_banner 六、总结       → block(“good”)
```
章节标题用下划线（醒目）；原题用引用块；对立意/多篇用表格对照。

共同点：骨架只是顺序，具体每节挑哪个组件，由内容决定。
