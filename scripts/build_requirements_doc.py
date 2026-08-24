from pathlib import Path

from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.table import WD_ALIGN_VERTICAL, WD_CELL_VERTICAL_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor


OUTPUT = Path("库存整理桌面软件需求规格说明书.docx")

PROJECT_ROOT = Path(__file__).resolve().parents[1]
OUTPUT = PROJECT_ROOT / "docs" / "库存整理桌面软件需求规格说明书.docx"

NAVY = "0B2545"
BLUE = "2E74B5"
BLUE_DARK = "1F4D78"
MUTED = "666666"
LIGHT_BLUE = "E8EEF5"
LIGHT_GRAY = "F2F4F7"
WHITE = "FFFFFF"
RED = "9B1C1C"


def set_run_font(run, size=11, color=None, bold=None, italic=None):
    run.font.name = "Microsoft YaHei"
    run._element.rPr.rFonts.set(qn("w:ascii"), "Calibri")
    run._element.rPr.rFonts.set(qn("w:hAnsi"), "Calibri")
    run._element.rPr.rFonts.set(qn("w:eastAsia"), "Microsoft YaHei")
    run.font.size = Pt(size)
    if color:
        run.font.color.rgb = RGBColor.from_string(color)
    if bold is not None:
        run.bold = bold
    if italic is not None:
        run.italic = italic


def set_cell_shading(cell, fill):
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = tc_pr.find(qn("w:shd"))
    if shd is None:
        shd = OxmlElement("w:shd")
        tc_pr.append(shd)
    shd.set(qn("w:fill"), fill)


def set_cell_border(cell, color="D5DCE5", size="6"):
    tc_pr = cell._tc.get_or_add_tcPr()
    borders = tc_pr.first_child_found_in("w:tcBorders")
    if borders is None:
        borders = OxmlElement("w:tcBorders")
        tc_pr.append(borders)
    for edge in ("top", "left", "bottom", "right"):
        tag = qn(f"w:{edge}")
        edge_el = borders.find(tag)
        if edge_el is None:
            edge_el = OxmlElement(f"w:{edge}")
            borders.append(edge_el)
        edge_el.set(qn("w:val"), "single")
        edge_el.set(qn("w:sz"), size)
        edge_el.set(qn("w:space"), "0")
        edge_el.set(qn("w:color"), color)


def set_cell_margins(cell, top=90, start=120, bottom=90, end=120):
    tc = cell._tc
    tc_pr = tc.get_or_add_tcPr()
    tc_mar = tc_pr.first_child_found_in("w:tcMar")
    if tc_mar is None:
        tc_mar = OxmlElement("w:tcMar")
        tc_pr.append(tc_mar)
    for margin_name, value in (("top", top), ("start", start), ("bottom", bottom), ("end", end)):
        margin = tc_mar.find(qn(f"w:{margin_name}"))
        if margin is None:
            margin = OxmlElement(f"w:{margin_name}")
            tc_mar.append(margin)
        margin.set(qn("w:w"), str(value))
        margin.set(qn("w:type"), "dxa")


def set_table_geometry(table, widths):
    table.autofit = False
    tbl_pr = table._tbl.tblPr
    tbl_w = tbl_pr.first_child_found_in("w:tblW")
    if tbl_w is None:
        tbl_w = OxmlElement("w:tblW")
        tbl_pr.append(tbl_w)
    tbl_w.set(qn("w:w"), str(sum(widths)))
    tbl_w.set(qn("w:type"), "dxa")
    tbl_layout = tbl_pr.first_child_found_in("w:tblLayout")
    if tbl_layout is None:
        tbl_layout = OxmlElement("w:tblLayout")
        tbl_pr.append(tbl_layout)
    tbl_layout.set(qn("w:type"), "fixed")
    tbl_ind = tbl_pr.first_child_found_in("w:tblInd")
    if tbl_ind is None:
        tbl_ind = OxmlElement("w:tblInd")
        tbl_pr.append(tbl_ind)
    tbl_ind.set(qn("w:w"), "120")
    tbl_ind.set(qn("w:type"), "dxa")
    grid = table._tbl.tblGrid
    for grid_col, width in zip(grid.gridCol_lst, widths):
        grid_col.set(qn("w:w"), str(width))
    for row in table.rows:
        for cell, width in zip(row.cells, widths):
            tc_pr = cell._tc.get_or_add_tcPr()
            tc_w = tc_pr.find(qn("w:tcW"))
            if tc_w is None:
                tc_w = OxmlElement("w:tcW")
                tc_pr.append(tc_w)
            tc_w.set(qn("w:w"), str(width))
            tc_w.set(qn("w:type"), "dxa")
            set_cell_margins(cell)
            set_cell_border(cell)


def format_table(table, widths, header_rows=1, font_size=9.5):
    set_table_geometry(table, widths)
    for row_idx, row in enumerate(table.rows):
        for cell in row.cells:
            cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            if row_idx < header_rows:
                set_cell_shading(cell, LIGHT_BLUE)
            for paragraph in cell.paragraphs:
                paragraph.paragraph_format.space_after = Pt(0)
                paragraph.paragraph_format.line_spacing = 1.08
                for run in paragraph.runs:
                    set_run_font(run, size=font_size, color=NAVY, bold=row_idx < header_rows)


def add_heading(doc, text, level=1):
    p = doc.add_paragraph(style=f"Heading {level}")
    p.paragraph_format.keep_with_next = True
    run = p.add_run(text)
    return p


def add_body(doc, text, bold_prefix=None):
    p = doc.add_paragraph()
    if bold_prefix and text.startswith(bold_prefix):
        run = p.add_run(bold_prefix)
        set_run_font(run, size=10.5, color="1F1F1F", bold=True)
        run = p.add_run(text[len(bold_prefix):])
        set_run_font(run, size=10.5, color="1F1F1F")
    else:
        run = p.add_run(text)
        set_run_font(run, size=10.5, color="1F1F1F")
    return p


def add_numbered(doc, items):
    for item in items:
        p = doc.add_paragraph(style="List Number")
        p.paragraph_format.space_after = Pt(4)
        p.paragraph_format.line_spacing = 1.25
        set_run_font(p.add_run(item), size=10.5, color="1F1F1F")


def add_bullets(doc, items):
    for item in items:
        p = doc.add_paragraph(style="List Bullet")
        p.paragraph_format.space_after = Pt(3)
        p.paragraph_format.line_spacing = 1.25
        set_run_font(p.add_run(item), size=10.5, color="1F1F1F")


def add_note(doc, label, text, color=BLUE_DARK):
    table = doc.add_table(rows=1, cols=1)
    set_table_geometry(table, [9360])
    cell = table.cell(0, 0)
    set_cell_shading(cell, "F4F6F9")
    p = cell.paragraphs[0]
    p.paragraph_format.space_after = Pt(0)
    run = p.add_run(f"{label}：")
    set_run_font(run, size=10, color=color, bold=True)
    run = p.add_run(text)
    set_run_font(run, size=10, color="1F1F1F")
    doc.add_paragraph().paragraph_format.space_after = Pt(1)


def add_matrix(doc, headers, rows, widths, font_size=9.2):
    table = doc.add_table(rows=1, cols=len(headers))
    table.style = "Table Grid"
    for cell, value in zip(table.rows[0].cells, headers):
        cell.text = value
    for row_values in rows:
        cells = table.add_row().cells
        for cell, value in zip(cells, row_values):
            cell.text = value
    format_table(table, widths, font_size=font_size)
    doc.add_paragraph().paragraph_format.space_after = Pt(1)
    return table


def add_page_number(paragraph):
    run = paragraph.add_run()
    fld_char1 = OxmlElement("w:fldChar")
    fld_char1.set(qn("w:fldCharType"), "begin")
    instr_text = OxmlElement("w:instrText")
    instr_text.set(qn("xml:space"), "preserve")
    instr_text.text = "PAGE"
    fld_char2 = OxmlElement("w:fldChar")
    fld_char2.set(qn("w:fldCharType"), "end")
    run._r.append(fld_char1)
    run._r.append(instr_text)
    run._r.append(fld_char2)


def configure_styles(doc):
    styles = doc.styles
    normal = styles["Normal"]
    normal.font.name = "Microsoft YaHei"
    normal._element.rPr.rFonts.set(qn("w:eastAsia"), "Microsoft YaHei")
    normal.font.size = Pt(10.5)
    normal.paragraph_format.space_after = Pt(6)
    normal.paragraph_format.line_spacing = 1.25

    for name, size, color, before, after in (
        ("Heading 1", 16, BLUE, 16, 8),
        ("Heading 2", 13, BLUE, 12, 6),
        ("Heading 3", 12, BLUE_DARK, 8, 4),
    ):
        style = styles[name]
        style.font.name = "Microsoft YaHei"
        style._element.rPr.rFonts.set(qn("w:eastAsia"), "Microsoft YaHei")
        style.font.size = Pt(size)
        style.font.bold = True
        style.font.color.rgb = RGBColor.from_string(color)
        style.paragraph_format.space_before = Pt(before)
        style.paragraph_format.space_after = Pt(after)
        style.paragraph_format.keep_with_next = True

    for name in ("List Bullet", "List Number"):
        style = styles[name]
        style.font.name = "Microsoft YaHei"
        style._element.rPr.rFonts.set(qn("w:eastAsia"), "Microsoft YaHei")
        style.font.size = Pt(10.5)
        style.paragraph_format.space_after = Pt(4)
        style.paragraph_format.line_spacing = 1.25


def setup_page(doc):
    section = doc.sections[0]
    section.top_margin = Inches(0.8)
    section.bottom_margin = Inches(0.72)
    section.left_margin = Inches(0.85)
    section.right_margin = Inches(0.85)
    section.header_distance = Inches(0.35)
    section.footer_distance = Inches(0.35)

    header = section.header.paragraphs[0]
    header.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    header.paragraph_format.space_after = Pt(0)
    run = header.add_run("库存整理桌面软件 | 需求规格说明书")
    set_run_font(run, size=8.5, color=MUTED)

    footer = section.footer.paragraphs[0]
    footer.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    footer.paragraph_format.space_before = Pt(0)
    run = footer.add_run("第 ")
    set_run_font(run, size=8.5, color=MUTED)
    add_page_number(footer)
    run = footer.add_run(" 页")
    set_run_font(run, size=8.5, color=MUTED)


def build_document():
    doc = Document()
    configure_styles(doc)
    setup_page(doc)

    title = doc.add_paragraph()
    title.paragraph_format.space_before = Pt(30)
    title.paragraph_format.space_after = Pt(6)
    title.alignment = WD_ALIGN_PARAGRAPH.LEFT
    set_run_font(title.add_run("库存整理桌面软件"), size=26, color=NAVY, bold=True)

    subtitle = doc.add_paragraph()
    subtitle.paragraph_format.space_after = Pt(18)
    set_run_font(subtitle.add_run("需求规格说明书"), size=15, color=BLUE_DARK)

    info = add_matrix(
        doc,
        ["项目", "内容"],
        [
            ["文档目的", "确认库存对比、跨品牌汇总和德清仓映射的业务规则，作为界面设计与开发验收依据。"],
            ["目标用户", "每周需整理多个品牌库存、制作库存对比表和汇总表的业务人员。"],
            ["版本", "V1.0（需求确认版）"],
            ["状态", "已确认；下一阶段为界面样式设计。"],
        ],
        [1850, 7510],
        font_size=10,
    )

    add_note(doc, "完成定义", "用户可手动导入符合规范的库存 Excel，完成同品牌周度对比或多品牌库存汇总；程序对关键数据异常进行阻断提示，并将结果导出为独立 Excel 文件。")

    add_heading(doc, "1. 背景与目标")
    add_body(doc, "业务人员每周一需要整理不同品牌商品在不同仓库的库存数量，并产出周度库存差异表和多品牌库存总表。当前软件目标是减少手工汇总与核对，保证数据规则一致、错误可见、结果可导出。")
    add_bullets(doc, [
        "支持品牌动态增删：不在程序中预置品牌，导入 Excel 的文件名即品牌名。",
        "支持同一品牌的本周、上周库存逐仓对比，完整导出新增、删除、增减和无变化记录。",
        "支持多个品牌库存汇总到一张宽表，德清仓固定优先展示，其余仓库按名称动态生成。",
        "对空数量、规格冲突、字段不匹配等关键问题明确提示并阻止错误导出。",
    ])

    add_heading(doc, "2. 范围与边界")
    add_matrix(
        doc,
        ["范围", "V1.0 包含", "V1.0 不包含"],
        [
            ["数据导入", "手动选择 Excel 文件；批量导入多个品牌文件。", "自动读取邮件、网盘或业务系统。"],
            ["历史管理", "每次手动上传本周和上周文件。", "自动归档、自动识别上周文件。"],
            ["仓库设置", "按品牌维护一个德清仓源名称映射。", "除德清仓外的仓库别名映射或自动语义识别。"],
            ["导出", "生成独立的对比表和汇总表 Excel。", "直接写回源文件或在线协作编辑。"],
        ],
        [1850, 3800, 3710],
        font_size=9.3,
    )

    add_heading(doc, "3. 源数据规范")
    add_body(doc, "第一版只接收与提供样例一致的 Excel 结构。列名、列数与列顺序必须完全一致；其中“库存状态”固定置于“库存量(主单位)”和“库存组织”之间。若缺少必填列、列名或列顺序不匹配、文件无法读取，则不允许进入计算。")
    add_matrix(
        doc,
        ["字段", "是否必填", "用途与规则"],
        [
            ["物料编码", "是", "与物料名称共同构成商品识别键；内部用于分组、对比和异常核验。"],
            ["物料名称", "是", "与物料编码共同构成商品识别键；用于对外展示和排序。"],
            ["规格型号", "是", "用于总表和对比表展示；同一商品键出现不同规格时为阻断性异常。"],
            ["仓库名称", "是", "用于逐仓对比和汇总列归属；德清仓映射前使用源名称。"],
            ["库存量(主单位)", "是", "数值字段；为空或非数字即为阻断性异常。"],
            ["数量（箱）", "是", "数值字段；为空或非数字即为阻断性异常。"],
            ["批号", "否", "不参与商品识别、汇总或对比。"],
            ["库存主单位、库存状态、库存组织", "否", "第一版不参与计算；库存状态为必备结构列，可保留为源数据辅助信息。"],
        ],
        [1800, 1100, 6460],
        font_size=9.15,
    )
    add_note(doc, "商品识别", "程序内部以“物料编码 + 物料名称”作为商品键；批号不参与区分。总表不展示物料编码，但仍按商品键逐行保留数据。")

    add_heading(doc, "4. 德清仓映射设置")
    add_body(doc, "各品牌的源仓库名称可能不同。设置模块需维护“品牌名 -> 一个源仓库名称”的映射；命中该映射的库存归入总表统一的“德清仓”列。")
    add_bullets(doc, [
        "品牌名默认取导入文件名（不含扩展名）；允许在设置中新增、修改、删除映射记录。",
        "每个品牌最多配置一个德清仓源名称；第一版不支持多个源仓库合并至德清仓。",
        "除德清仓外，不维护仓库别名：源文件中的仓库名称原样作为总表列名。",
        "生成汇总表前，程序核查所有已导入品牌是否具备德清仓映射。存在未配置品牌时，弹窗列出品牌并提供“继续生成”与“去配置映射”。",
        "选择继续生成时，该品牌在德清仓列留空，其他源仓库仍按原名称参与汇总。",
    ])

    add_heading(doc, "5. 同品牌库存对比")
    add_body(doc, "用户在库存对比页面手动填写品牌名，再分别上传或拖拽本周和上周库存 Excel。两个文件名不要求相同。程序以“物料编码 + 物料名称 + 仓库名称”为比较粒度。")
    add_matrix(
        doc,
        ["比较场景", "判定规则", "对比表变动类型"],
        [
            ["仅本周存在", "该商品键与仓库组合在上周不存在。", "新增商品"],
            ["仅上周存在", "该商品键与仓库组合在本周不存在。", "删除商品"],
            ["两周均存在，变化值大于 0", "本周数量减上周数量大于 0。", "库存增加"],
            ["两周均存在，变化值小于 0", "本周数量减上周数量小于 0。", "库存减少"],
            ["两周均存在，变化值为 0", "主单位数量与箱数均无变化。", "无变化"],
            ["跨仓移动", "同一商品从仓库 A 到仓库 B。", "仓库 A 按删除、仓库 B 按新增展示；不自动推断调拨。"],
        ],
        [2200, 3900, 3260],
        font_size=9.1,
    )
    add_body(doc, "对比表必须导出完整数据，包含无变化记录。除主单位变化值、箱数变化值外，另输出主单位同比变化和箱数同比变化；同比计算为“(本周 - 上周) / |上周|”，上周数量为 0 时显示“-”，不计算无限大百分比。")

    add_heading(doc, "6. 多品牌库存汇总")
    add_body(doc, "用户在汇总页面批量导入多个品牌库存文件。软件按品牌和商品整理库存，并生成一张仓库列动态扩展的宽表。总表参考现有模板的结构，但不沿用示例中的固定仓库数量和名称。")
    add_matrix(
        doc,
        ["项目", "汇总规则"],
        [
            ["总表固定商品列", "品牌、序号、名称、规格。总表不展示物料编码。"],
            ["仓库列生成", "第一组固定为德清仓；其余列取所有已导入品牌的非德清源仓库名称并集。每个仓库占两列：主单位数量、箱数。"],
            ["同名非德清仓", "不同品牌若仓库名称完全相同，共用同一组总表仓库列；数据不混算，因为每行保留品牌。"],
            ["仓库排序", "德清仓排第一；其余仓库按所有导入文件中的首次出现顺序排列。"],
            ["品牌排序", "按用户导入文件的顺序排列。"],
            ["商品排序", "同一品牌内按物料名称排序。"],
            ["同商品同仓多行", "将主单位数量与箱数分别求和。"],
            ["合计", "保留最后两列合计，分别汇总该商品在所有仓库中的主单位数量和箱数。"],
        ],
        [2350, 7010],
        font_size=9.35,
    )
    add_note(doc, "同名商品提示", "如果不同物料编码出现相同的物料名称和规格，程序先提醒用户核对；用户确认继续后，仍按两行展示，不合并。", color="7A5A00")

    add_heading(doc, "7. 数据校验与异常处理")
    add_body(doc, "所有阻断性异常都需要在界面中指出品牌、文件名、工作表、行号、字段名、当前值与修正建议。存在阻断性异常时，不允许导出对应的对比表或汇总表。")
    add_matrix(
        doc,
        ["级别", "触发条件", "系统处理"],
        [
            ["阻断", "源文件缺少必填列、列名不匹配、文件无法读取。", "停止导入或计算，列出缺失/不匹配字段。"],
            ["阻断", "库存量（主单位）或数量（箱）为空、非数字。", "停止导出，定位到对应数据行和字段。"],
            ["阻断", "同一商品键出现不同规格型号。", "停止导出，要求先修正源数据。"],
            ["确认", "不同物料编码出现相同物料名称和规格。", "提醒核对；用户确认后保留为两行并继续。"],
            ["确认", "导入品牌未配置德清仓。", "生成汇总前提示；用户选择继续时德清仓留空。"],
        ],
        [1000, 4300, 4060],
        font_size=9.15,
    )

    add_heading(doc, "8. 导出文件定义")
    add_heading(doc, "8.1 库存对比表", level=2)
    add_matrix(
        doc,
        ["字段类别", "建议导出字段"],
        [
            ["标识信息", "品牌、物料编码、物料名称、规格型号、仓库名称。"],
            ["上周库存", "上周主单位数量、上周箱数。"],
            ["本周库存", "本周主单位数量、本周箱数。"],
            ["变化", "主单位变化值、主单位同比变化、箱数变化值、箱数同比变化、变动类型。"],
            ["记录范围", "完整导出：新增、删除、库存增加、库存减少、无变化。"],
        ],
        [1800, 7560],
        font_size=9.4,
    )
    add_heading(doc, "8.2 库存汇总表", level=2)
    add_matrix(
        doc,
        ["区域", "字段或排版规则"],
        [
            ["表头", "第一行标题；第二、三行为动态仓库分组表头。德清仓为第一组，每仓库两列。"],
            ["商品信息", "品牌、序号、名称、规格。品牌列可按连续相同品牌合并展示，具体以界面设计阶段确认的模板样式为准。"],
            ["仓库数据", "每个仓库包含“数量（主单位）”和“数量（箱）”；无库存可留空。"],
            ["合计", "末尾固定为“合计”两列，分别为主单位数量合计和箱数合计。"],
        ],
        [1800, 7560],
        font_size=9.4,
    )

    add_heading(doc, "9. 软件页面与关键交互")
    add_matrix(
        doc,
        ["页面", "核心操作", "完成条件"],
        [
            ["首页", "进入库存对比、库存汇总或仓库映射设置。", "用户可清晰进入三个主功能。"],
            ["库存对比", "填写品牌；点击或拖拽分别导入本周、上周文件；查看结果；可清空当前文件与结果；导出。", "完成校验后可生成完整对比表 Excel。"],
            ["库存汇总", "点击或拖拽批量导入品牌文件；查看品牌、仓库和异常；确认未配置德清仓的处理；可清空当前文件与结果；导出。", "完成校验与用户确认后可生成动态宽表 Excel。"],
            ["德清仓映射设置", "新增、修改、删除“品牌 -> 源仓库名称”映射。", "保存后可在下一次汇总校验中生效。"],
        ],
        [1700, 4550, 3110],
        font_size=9.2,
    )

    add_heading(doc, "10. 验收标准")
    add_numbered(doc, [
        "导入满足源数据规范的单个或多个品牌文件时，软件能识别文件名品牌、商品、仓库和两种库存数量。",
        "对比两个同品牌文件时，软件按商品加仓库输出完整记录，准确标识新增、删除、增加、减少和无变化，并分别计算两种数量的变化值与同比变化；上周数量为 0 的同比显示“-”。",
        "对比文件名不同但用户指定同一品牌时，仍可完成对比。",
        "汇总多个品牌时，德清仓始终处于第一组仓库列；其他仓库按首次出现顺序动态生成。",
        "不同品牌相同名称的非德清仓共用同一仓库列，但每个品牌的库存仅展示在其自身商品行。",
        "同品牌同商品同仓库的多行库存会分别正确汇总主单位数量与箱数，合计列与仓库数据一致。",
        "未配置德清仓的品牌在生成前会被提示，用户可继续生成或进入设置；继续生成时该品牌德清仓为空。",
        "发生缺列、列名不匹配、数量为空/非数字、同一商品不同规格等阻断性异常时，软件不允许导出，并可定位问题数据。",
        "发生不同物料编码但名称和规格相同的情况时，软件要求用户确认；确认后仍以两行保留。",
        "用户可分别导出独立的库存对比表和库存汇总表 Excel 文件。",
        "对比与汇总页面均支持拖拽 .xlsx 文件导入，并可清空当前会话的文件与结果；清空不删除已保存的德清仓映射或导出记录。",
    ])

    add_heading(doc, "11. 后续阶段")
    add_body(doc, "本需求文档确认后，下一阶段进入桌面软件界面样式设计：确定技术实现方向、页面信息架构、关键状态、交互流程及视觉方案；样式获得确认后再进入开发。")

    doc.save(OUTPUT)
    print(OUTPUT.resolve())


if __name__ == "__main__":
    build_document()
