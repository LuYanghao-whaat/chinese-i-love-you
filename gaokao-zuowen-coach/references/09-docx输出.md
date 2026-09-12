# Word 文档输出（.docx）

> 当用户在开场选"输出格式 = Word 文档"时，用本模块。
> 工具链：Python 3.12 + python-docx。完整排版规范已随本 Skill 归档在 `references/docx/`：
> `docx-skill-pydocx-zh.md`（SOP 6 步 + 全景模板）、`docx-guide-pydocx-zh.md`（模块详解）、`visual-style-spec-zh.md`（视觉规范）。

---

## 零、先确认文档视效（四档）

选"Word 文档"时，先问用户要哪一档：

| 档位 | 做什么 | 用哪些函数 |
|---|---|---|
| 普通（默认） | 标题 + 正文 + 页码，够看就行 | `setup_global_styles`、`setup_page_section`、`add_h1`、`add_h2`、`add_body` |
| 美观 | 加封面、目录、色块、分节，用于正式场合 | 再加 `add_cover`、`add_toc`、`add_custom_block`、`add_data_table` |
| 仅 Markdown | 不生成 `.docx`，只输出 Markdown 源码 | — |
| 纯文本 | 只输出无格式纯文本 | — |

用户没指定，就按"普通"来。选"仅 Markdown / 纯文本"时，不写 Python、不碰环境。

---

## 一、环境准备（默认帮用户装好）

生成前先确认环境，缺什么补什么：

1. 查 Python 3.12：`py -3.12 --version`
2. 查库：`py -3.12 -m pip show python-docx`
3. 缺 Python 3.12 → 默认帮装：`winget install Python.Python.3.12`（装不了就提示用户手动装）
4. 缺库 → `py -3.12 -m pip install python-docx`

装完复验一次，再生成。不要把带报错的半成品丢给用户。

---

## 二、字体：覆盖模板默认

本 Skill 的中文文档偏好是宋体 / 楷体，避免微软雅黑、等线。
标准模板里的 `微软雅黑` 一律替换：

- 正文、标题：`宋体`
- 引用、点评、色块内文字：`楷体`

（模板代码里所有出现 `"微软雅黑"` 的地方，按上述替换。）

---

## 三、工作流（SOP 6 步）

1. 确定内容：读原始文本，理清封面信息、章节、段落、表格、批注。
2. 模块化布局：提取标题做封面；规划章节层级；把并列句转成定宽表格或列表。
3. 套用模块样式：映射到标准模板的函数；强调处分离 Run、用全局高亮色。
4. 写代码生成：完整脚本（含标准模板骨架），本地 Shell 执行。
5. 验证与自愈：报错就回到第 4 步改，直到跑通。
6. 交付：输出最终 `.docx` 的绝对路径。

---

## 四、致命错误规避

1. 中英文字体失效：在文档级改 `Normal` 样式的 `w:eastAsia`。
2. 底纹黑块：`w:shd` 的 `w:val` 必须是 `"clear"`，不能是 `"solid"`。
3. 颜色：字体色用 `RGBColor`；高亮底色用 `WD_COLOR_INDEX`。
4. 表格自适应崩溃：必须设 `w:tblLayout type="fixed"`。
5. 页眉页脚串扰：多节文档改页眉页脚前，先 `is_linked_to_previous = False`。
6. 换行崩溃：`add_run()` 里不能用 `\n`，换行要新建段落或 `run.add_break()`。
7. 域代码：TOC / 页码生成后为空白正常，交付时提醒用户按 `Ctrl+A` 再 `F9` 刷新。

---

## 五、标准模板（整体使用，只在内容填充区写逻辑）

```python
import os
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING, WD_COLOR_INDEX
from docx.enum.section import WD_SECTION
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.opc.constants import RELATIONSHIP_TYPE as RT

# 本 Skill 字体偏好（覆盖模板默认的微软雅黑）
FONT_SONG = "宋体"
FONT_KAI = "楷体"

# ================= 1. 核心 OXML 工具库 =================

def set_run_font(run, font_name, size_pt, rgb_tuple=None):
    run.font.name = font_name
    run._element.rPr.rFonts.set(qn('w:eastAsia'), font_name)
    run.font.size = Pt(size_pt)
    if rgb_tuple:
        run.font.color.rgb = RGBColor(*rgb_tuple)

def add_shading(element, hex_color):
    pr = element._p.get_or_add_pPr() if hasattr(element, '_p') else element._tc.get_or_add_tcPr()
    for el in pr.findall(qn('w:shd')): pr.remove(el)
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), hex_color)
    pr.append(shd)

def add_bottom_border(paragraph, color="E2E8F0", sz="6"):
    pPr = paragraph._p.get_or_add_pPr()
    pBdr = OxmlElement('w:pBdr')
    bottom = OxmlElement('w:bottom')
    bottom.set(qn('w:val'), 'single')
    bottom.set(qn('w:sz'), sz)
    bottom.set(qn('w:space'), '4')
    bottom.set(qn('w:color'), color)
    pBdr.append(bottom)
    pPr.append(pBdr)

def add_horizontal_line(paragraph):
    pPr = paragraph._p.get_or_add_pPr()
    pBdr = OxmlElement('w:pBdr')
    bottom = OxmlElement('w:bottom')
    bottom.set(qn('w:val'), 'single')
    bottom.set(qn('w:sz'), '6')
    bottom.set(qn('w:space'), '1')
    bottom.set(qn('w:color'), 'auto')
    pBdr.append(bottom)
    pPr.append(pBdr)

def add_hyperlink(paragraph, url, display_text):
    r_id = paragraph.part.relate_to(url, RT.HYPERLINK, is_external=True)
    hyperlink = OxmlElement("w:hyperlink")
    hyperlink.set(qn("r:id"), r_id)
    wr = OxmlElement("w:r")
    rPr = OxmlElement("w:rPr")
    rSt = OxmlElement("w:rStyle")
    rSt.set(qn("w:val"), "Hyperlink")
    rPr.append(rSt)
    wr.append(rPr)
    t = OxmlElement("w:t")
    t.text = display_text
    wr.append(t)
    hyperlink.append(wr)
    paragraph._p.append(hyperlink)

def add_page_number(paragraph):
    def fld(para, instr_text):
        run = para.add_run()
        begin = OxmlElement('w:fldChar')
        begin.set(qn('w:fldCharType'), 'begin')
        instr = OxmlElement('w:instrText')
        instr.text = instr_text
        end = OxmlElement('w:fldChar')
        end.set(qn('w:fldCharType'), 'end')
        run._r.extend([begin, instr, end])
    paragraph.add_run("第 ")
    fld(paragraph, "PAGE")
    paragraph.add_run(" 页，共 ")
    fld(paragraph, "NUMPAGES")
    paragraph.add_run(" 页")

def add_toc(doc, title="目录"):
    doc.add_heading(title, level=1)
    para = doc.add_paragraph()
    run = para.add_run()
    begin = OxmlElement("w:fldChar")
    begin.set(qn("w:fldCharType"), "begin")
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = 'TOC \\o "1-3" \\h \\z \\u'
    end = OxmlElement("w:fldChar")
    end.set(qn("w:fldCharType"), "end")
    run._r.extend([begin, instr, end])
    doc.add_page_break()

# ================= 2. 全局样式与页面布局引擎 =================

def setup_global_styles(doc):
    """设定全局正文基准，防止中文字体退化"""
    normal = doc.styles['Normal']
    normal.font.name = FONT_SONG
    normal.font.size = Pt(12)
    normal.element.rPr.rFonts.set(qn('w:eastAsia'), FONT_SONG)
    normal.paragraph_format.line_spacing = 1.5
    normal.paragraph_format.space_after = Pt(6)

def set_table_width_fixed(table, width_twips=9360):
    tblPr = table._tbl.tblPr
    for el in tblPr.findall(qn("w:tblW")): tblPr.remove(el)
    tblW = OxmlElement("w:tblW")
    tblW.set(qn("w:w"), str(width_twips))
    tblW.set(qn("w:type"), "dxa")
    tblPr.insert(0, tblW)
    for el in tblPr.findall(qn("w:tblLayout")): tblPr.remove(el)
    tblLayout = OxmlElement("w:tblLayout")
    tblLayout.set(qn("w:type"), "fixed")
    tblPr.append(tblLayout)

def style_cell(cell, bg_hex=None):
    tcPr = cell._tc.get_or_add_tcPr()
    if bg_hex:
        add_shading(cell, bg_hex)
    for el in tcPr.findall(qn('w:vAlign')): tcPr.remove(el)
    vAlign = OxmlElement('w:vAlign')
    vAlign.set(qn('w:val'), 'center')
    tcPr.append(vAlign)

def setup_page_section(doc, index=0):
    sec = doc.sections[index]
    sec.page_width, sec.page_height = Inches(8.27), Inches(11.69)
    sec.left_margin = sec.right_margin = sec.top_margin = sec.bottom_margin = Inches(1)
    ftr = sec.footer
    ftr.is_linked_to_previous = False
    fp = ftr.paragraphs[0]
    fp.clear()
    fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    add_page_number(fp)

# ================= 3. 模块化内容生成器 =================

def add_cover(doc, title_text, subtitle_text, info_text):
    for _ in range(4): doc.add_paragraph()
    title = doc.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = title.add_run(title_text)
    set_run_font(run, FONT_SONG, 28, (31, 78, 121))
    run.bold = True
    if subtitle_text:
        sub = doc.add_paragraph()
        sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
        sub_run = sub.add_run(subtitle_text)
        set_run_font(sub_run, FONT_KAI, 16, (89, 89, 89))
    for _ in range(10): doc.add_paragraph()
    info = doc.add_paragraph()
    info.alignment = WD_ALIGN_PARAGRAPH.CENTER
    info_run = info.add_run(info_text)
    set_run_font(info_run, FONT_SONG, 12, (128, 128, 128))
    doc.add_section(WD_SECTION.NEW_PAGE)

def add_h1(doc, text):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before, p.paragraph_format.space_after = Pt(24), Pt(18)
    run = p.add_run(text)
    run.bold = True
    set_run_font(run, FONT_SONG, 18, (255, 255, 255))
    add_shading(p, '1A202C')

def add_h2(doc, text):
    p = doc.add_paragraph()
    p.paragraph_format.space_before, p.paragraph_format.space_after = Pt(16), Pt(8)
    run = p.add_run(text)
    run.bold = True
    set_run_font(run, FONT_SONG, 14, (46, 117, 182))
    add_bottom_border(p)

def add_body(doc, text):
    doc.add_paragraph(text)

def add_list_item(doc, text):
    p = doc.add_paragraph()
    p.paragraph_format.left_indent, p.paragraph_format.first_line_indent = Inches(0.4), Inches(-0.2)
    run = p.add_run("•\t" + text)
    set_run_font(run, FONT_SONG, 12, (64, 64, 64))

def add_data_table(doc, data_matrix):
    if not data_matrix: return
    table = doc.add_table(rows=len(data_matrix), cols=len(data_matrix[0]))
    set_table_width_fixed(table, 9360)
    table.style = 'Table Grid'
    for r_idx, row_data in enumerate(data_matrix):
        for c_idx, cell_text in enumerate(row_data):
            cell = table.cell(r_idx, c_idx)
            cell.text = str(cell_text)
            style_cell(cell, bg_hex='2B579A' if r_idx == 0 else None)
            for p in cell.paragraphs:
                for run in p.runs:
                    set_run_font(run, FONT_SONG, 10, (255,255,255) if r_idx==0 else (0,0,0))

def add_custom_block(doc, text, bg_hex="F8F9FA", font_rgb=(33, 37, 41), border_hex="DEE2E6"):
    """支持底纹、字体颜色及四面边框的色块，适合"亮点/问题"点评"""
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Inches(0.2)
    p.paragraph_format.right_indent = Inches(0.2)
    p.paragraph_format.space_before, p.paragraph_format.space_after = Pt(8), Pt(8)
    run = p.add_run(text)
    set_run_font(run, FONT_KAI, 12, font_rgb)
    add_shading(p, bg_hex)
    pPr = p._p.get_or_add_pPr()
    pBdr = OxmlElement('w:pBdr')
    for side in ['top', 'left', 'bottom', 'right']:
        bdr = OxmlElement(f'w:{side}')
        bdr.set(qn('w:val'), 'single')
        bdr.set(qn('w:sz'), '4')
        bdr.set(qn('w:space'), '4')
        bdr.set(qn('w:color'), border_hex)
        pBdr.append(bdr)
    pPr.append(pBdr)

def generate_document(output_path):
    doc = Document()
    setup_global_styles(doc)
    setup_page_section(doc, 0)

    # ================= 4. 内容填充区（按文本动态生成） =================
    # 建议顺序：封面(可选) → 初始化第二节页脚 setup_page_section(doc,1) → 目录(可选) → 正文
    # 正文：add_h1 / add_h2 / add_body / add_list_item / add_data_table / add_custom_block
    # 从这里开始构建文档...

    # ===================================================================
    doc.save(output_path)
    print(f"成功渲染文档: {output_path}")

if __name__ == "__main__":
    output_file = os.path.abspath("Rendered_Document.docx")
    generate_document(output_file)
```

---

## 六、与批改 / 讲义结合

- 用户要"讲评讲义"时，按 `03-批改与评分.md` 第九节的六段结构组织：
  原题再现 → 审题指导 → 立意指导 → 陷阱提示 → 作文分析 → 总结。
- 亮点、问题、金句点评用 `add_custom_block` 做成色块。
- 逻辑链审查（`08`）用 `add_data_table` 输出"环节 / 说法 / 判断 / 说明"表。
- 分数、维度得分用 `add_data_table` 或色块。
- 交付时给出 `.docx` 绝对路径，并提醒按 `Ctrl+A`、`F9` 刷新目录与页码域。
