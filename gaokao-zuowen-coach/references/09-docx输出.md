# Word 文档输出（.docx）

> 当用户在开场选"输出格式 = Word 文档"时，用本模块。
> 工具链：Python 3.12 + python-docx。完整排版规范已随本 Skill 归档在 `references/docx/`：
> `docx-skill-pydocx-zh.md`、`docx-guide-pydocx-zh.md`、`visual-style-spec-zh.md`。

---

## 零、先确认文档视效（三档）

选"Word 文档"时，先问用户要哪一档。三档都生成 `.docx`，区别只在保留多少视觉包装：

| 档位 | 保留什么 | 不用什么 |
|---|---|---|
| 普通（默认） | 封面 + 分节标题 + 正文 + 页码 | 目录、色块 |
| 美观 | 再加目录、点评色块、数据表 | — |
| 极简 | 只保留 Markdown 语法能表达的格式：标题、列表、表格、加粗斜体 | 颜色、底纹、边框、自定义字体 |

用户没指定，就按"普通"来。极简也生成 `.docx`，只是样式收敛到 Markdown 级别。

---

## 一、环境准备（默认帮用户装好）

1. 查 Python 3.12：`py -3.12 --version`
2. 查库：`py -3.12 -m pip show python-docx`
3. 缺 Python 3.12 → `winget install Python.Python.3.12`（装不了就提示用户手动装）
4. 缺库 → `py -3.12 -m pip install python-docx`

装完复验一次，再生成。不要把带报错的半成品丢给用户。

---

## 二、字体

本 Skill 偏好宋体 / 楷体，避免微软雅黑、等线。模板里的字体常量：
正文与标题用 `宋体`，引用、点评、色块内文字用 `楷体`。

---

## 三、工作流（SOP 6 步）

1. 确定内容：理清封面信息、章节、段落、表格、点评。
2. 模块化布局：标题做封面；规划章节层级；并列内容转表格或列表。
3. 套用模板函数，不要另起炉灶。
4. 写完整脚本，本地执行。
5. 验证与自愈：报错就改，直到跑通；跑通后再走第六节的"生成后自检"。
6. 交付：给出 `.docx` 绝对路径。

---

## 四、致命错误规避

1. 中英文字体失效：文档级改 `Normal` 样式的 `w:eastAsia`。
2. 底纹黑块：`w:shd` 的 `w:val` 必须是 `"clear"`。
3. 颜色：字体色用 `RGBColor`；高亮底色用 `WD_COLOR_INDEX`。
4. 表格列宽错乱：设 `w:tblLayout type="fixed"`。
5. 页眉页脚串扰：多节文档改页眉页脚前先 `is_linked_to_previous = False`。
6. 换行崩溃：`add_run()` 里不能用 `\n`，换行用新段落或 `run.add_break()`。
7. 兼容性模式：必须调 `set_doc_settings()`，把 `compatibilityMode` 设为 15，否则 Word 以兼容模式打开、排版漂移。
8. 目录空白页：TOC 是域，python-docx 生成后不会自动渲染，Word 首次打开是空白。`set_doc_settings()` 里已加 `updateFields=true`，让 Word 打开时自动更新域。普通/极简档不放目录，只有美观档放。

---

## 五、标准模板（整体使用，只在内容填充区写逻辑）

```python
import os
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_COLOR_INDEX
from docx.enum.section import WD_SECTION
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

FONT_SONG = "宋体"
FONT_KAI = "楷体"

# 统一配色
C_PRIMARY = (31, 78, 121)     # 深蓝：大标题
C_ACCENT  = (46, 117, 182)    # 蓝：小标题
C_TEXT    = (51, 51, 51)      # 正文
C_GRAY    = (120, 120, 120)   # 辅助信息
HEX_ACCENT = "1F4E79"
HEX_RULE   = "C9D8E8"

# ================= 1. OXML 工具 =================

def set_run_font(run, font_name, size_pt, rgb_tuple=None, bold=False):
    run.font.name = font_name
    run._element.rPr.rFonts.set(qn('w:eastAsia'), font_name)
    run.font.size = Pt(size_pt)
    run.font.bold = bold
    if rgb_tuple:
        run.font.color.rgb = RGBColor(*rgb_tuple)

def add_shading(paragraph, hex_color):
    pPr = paragraph._p.get_or_add_pPr()
    for el in pPr.findall(qn('w:shd')): pPr.remove(el)
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), hex_color)
    pPr.append(shd)

def add_bottom_border(paragraph, color=HEX_RULE, sz="6"):
    pPr = paragraph._p.get_or_add_pPr()
    pBdr = OxmlElement('w:pBdr')
    bottom = OxmlElement('w:bottom')
    bottom.set(qn('w:val'), 'single')
    bottom.set(qn('w:sz'), sz)
    bottom.set(qn('w:space'), '4')
    bottom.set(qn('w:color'), color)
    pBdr.append(bottom)
    pPr.append(pBdr)

def add_left_bar(paragraph, color=HEX_ACCENT, sz="18"):
    pPr = paragraph._p.get_or_add_pPr()
    pBdr = OxmlElement('w:pBdr')
    left = OxmlElement('w:left')
    left.set(qn('w:val'), 'single')
    left.set(qn('w:sz'), sz)
    left.set(qn('w:space'), '6')
    left.set(qn('w:color'), color)
    pBdr.append(left)
    pPr.append(pBdr)

def add_page_number(paragraph):
    def fld(para, instr_text):
        run = para.add_run()
        begin = OxmlElement('w:fldChar'); begin.set(qn('w:fldCharType'), 'begin')
        instr = OxmlElement('w:instrText'); instr.text = instr_text
        end = OxmlElement('w:fldChar'); end.set(qn('w:fldCharType'), 'end')
        run._r.extend([begin, instr, end])
    paragraph.add_run("第 ")
    fld(paragraph, "PAGE")
    paragraph.add_run(" 页 / 共 ")
    fld(paragraph, "NUMPAGES")
    paragraph.add_run(" 页")

def add_toc(doc, title="目录"):
    add_h1(doc, title)
    para = doc.add_paragraph()
    run = para.add_run()
    begin = OxmlElement("w:fldChar"); begin.set(qn("w:fldCharType"), "begin")
    instr = OxmlElement("w:instrText"); instr.set(qn("xml:space"), "preserve")
    instr.text = 'TOC \\o "1-2" \\h \\z \\u'
    end = OxmlElement("w:fldChar"); end.set(qn("w:fldCharType"), "end")
    run._r.extend([begin, instr, end])
    doc.add_page_break()

# ================= 2. 全局设置 =================

def setup_global_styles(doc):
    normal = doc.styles['Normal']
    normal.font.name = FONT_SONG
    normal.font.size = Pt(12)
    normal.element.rPr.rFonts.set(qn('w:eastAsia'), FONT_SONG)
    normal.font.color.rgb = RGBColor(*C_TEXT)
    normal.paragraph_format.line_spacing = 1.5
    normal.paragraph_format.space_after = Pt(6)

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

def set_doc_settings(doc):
    """兼容性模式 15 + 打开时更新域（修目录空白页）"""
    st = doc.settings.element
    for el in st.findall(qn('w:compat')): st.remove(el)
    compat = OxmlElement('w:compat')
    cs = OxmlElement('w:compatSetting')
    cs.set(qn('w:name'), 'compatibilityMode')
    cs.set(qn('w:uri'), 'http://schemas.microsoft.com/office/word')
    cs.set(qn('w:val'), '15')
    compat.append(cs)
    st.append(compat)
    for el in st.findall(qn('w:updateFields')): st.remove(el)
    uf = OxmlElement('w:updateFields'); uf.set(qn('w:val'), 'true')
    st.append(uf)

# ================= 3. 内容模块 =================

def add_cover(doc, title, subtitle, info):
    for _ in range(6): doc.add_paragraph()
    p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_run_font(p.add_run(title), FONT_SONG, 26, C_PRIMARY, bold=True)
    if subtitle:
        p2 = doc.add_paragraph(); p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
        set_run_font(p2.add_run(subtitle), FONT_KAI, 15, C_GRAY)
    rule = doc.add_paragraph(); rule.alignment = WD_ALIGN_PARAGRAPH.CENTER
    add_bottom_border(rule, color=HEX_ACCENT, sz="8")
    for _ in range(8): doc.add_paragraph()
    p3 = doc.add_paragraph(); p3.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_run_font(p3.add_run(info), FONT_SONG, 11, C_GRAY)
    doc.add_page_break()

def add_h1(doc, text):
    p = doc.add_paragraph()
    p.paragraph_format.space_before, p.paragraph_format.space_after = Pt(18), Pt(8)
    p.paragraph_format.left_indent = Inches(0.12)
    set_run_font(p.add_run(text), FONT_SONG, 16, C_PRIMARY, bold=True)
    add_left_bar(p)

def add_h2(doc, text):
    p = doc.add_paragraph()
    p.paragraph_format.space_before, p.paragraph_format.space_after = Pt(12), Pt(4)
    set_run_font(p.add_run(text), FONT_SONG, 13, C_ACCENT, bold=True)
    add_bottom_border(p)

def add_body(doc, text):
    doc.add_paragraph(text)

def add_list_item(doc, text):
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Inches(0.3)
    set_run_font(p.add_run("· " + text), FONT_SONG, 12, C_TEXT)

def add_block(doc, text, bg="EAF2FA", border="B9D4EC"):
    """点评色块：亮点用 EAF2FA，问题用 FBF3E7 / E6D2AE"""
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Inches(0.15)
    p.paragraph_format.right_indent = Inches(0.15)
    p.paragraph_format.space_before, p.paragraph_format.space_after = Pt(6), Pt(6)
    set_run_font(p.add_run(text), FONT_KAI, 12, (40, 40, 40))
    add_shading(p, bg)
    pPr = p._p.get_or_add_pPr()
    pBdr = OxmlElement('w:pBdr')
    for side in ['top', 'left', 'bottom', 'right']:
        b = OxmlElement(f'w:{side}')
        b.set(qn('w:val'), 'single'); b.set(qn('w:sz'), '4')
        b.set(qn('w:space'), '4'); b.set(qn('w:color'), border)
        pBdr.append(b)
    pPr.append(pBdr)

def add_data_table(doc, matrix):
    if not matrix: return
    table = doc.add_table(rows=len(matrix), cols=len(matrix[0]))
    table.style = 'Table Grid'
    tblPr = table._tbl.tblPr
    layout = OxmlElement('w:tblLayout'); layout.set(qn('w:type'), 'fixed')
    tblPr.append(layout)
    for r, row in enumerate(matrix):
        for c, val in enumerate(row):
            cell = table.cell(r, c)
            cell.text = str(val)
            if r == 0:
                add_shading(cell.paragraphs[0], HEX_ACCENT)
            for para in cell.paragraphs:
                for run in para.runs:
                    set_run_font(run, FONT_SONG, 10,
                                 (255, 255, 255) if r == 0 else C_TEXT,
                                 bold=(r == 0))

def generate(output_path):
    doc = Document()
    setup_global_styles(doc)
    setup_page_section(doc, 0)
    set_doc_settings(doc)

    # ============ 内容填充区（按文本动态生成） ============
    # 普通档：add_cover → add_h1/add_h2/add_body/add_list_item
    # 美观档：再加 add_toc / add_block / add_data_table
    # 极简档：只用 add_h1/add_h2/add_body/add_list_item/add_data_table，不传颜色、不加底纹边框

    # ======================================================
    doc.save(output_path)
    print(f"成功渲染文档: {output_path}")

if __name__ == "__main__":
    generate(os.path.abspath("批改报告.docx"))
```

---

## 六、生成后自检（重要）

跑通不算完，必须看一眼：

1. 有 Word 时，用 COM 打开、更新域、导出 PDF，翻一遍：
   - 目录页是否仍空白（若空，说明 `updateFields` 没生效，改用 COM 手动 `doc.Fields.Update()` 后保存）。
   - 标题层级、色块、表格是否错位。
   - 有没有整页空白、孤行。
2. 没有 Word 时，至少重新解压 docx，确认目录页不是"横幅 + 空白"。
3. 发现版面问题就改，不要直接交付。

PowerShell 更新域并导出 PDF 的参考：

```powershell
$w = New-Object -ComObject Word.Application; $w.Visible = $false; $w.DisplayAlerts = 0
$d = $w.Documents.Open("C:\path\批改报告.docx")
$d.Fields.Update() | Out-Null
$d.Save()
$d.ExportAsFixedFormat("C:\path\批改报告.pdf", 17)
$d.Close($false); $w.Quit()
```

---

## 七、与批改 / 讲义结合

- 讲义按 `03-批改与评分.md` 第九节的六段结构：原题再现 → 审题指导 → 立意指导 → 陷阱提示 → 作文分析 → 总结。
- 亮点/问题点评用 `add_block`（蓝底=亮点，暖底=问题）。
- 逻辑链审查（`08`）用 `add_data_table` 输出"环节 / 说法 / 判断 / 说明"。
- 分数、维度得分用 `add_data_table` 或色块。
- 交付时给出 `.docx` 绝对路径；用了目录就提醒按 `Ctrl+A`、`F9` 刷新。
