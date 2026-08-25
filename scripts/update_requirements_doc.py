from pathlib import Path

from docx import Document
from docx.enum.text import WD_BREAK
from docx.shared import Pt


DOCUMENT_PATH = Path("库存整理桌面软件需求规格说明书.docx")


PROJECT_ROOT = Path(__file__).resolve().parents[1]
DOCUMENT_PATH = PROJECT_ROOT / "docs" / "库存整理桌面软件需求规格说明书.docx"


def add_text(document, text, bold=False):
    paragraph = document.add_paragraph()
    paragraph.paragraph_format.space_after = Pt(6)
    run = paragraph.add_run(text)
    run.bold = bold
    return paragraph


def add_bullet(document, text):
    paragraph = document.add_paragraph(style="List Bullet")
    paragraph.paragraph_format.space_after = Pt(4)
    paragraph.add_run(text)
    return paragraph


def add_table(document, headers, rows):
    table = document.add_table(rows=1, cols=len(headers))
    table.style = "Table Grid"
    for cell, value in zip(table.rows[0].cells, headers):
        cell.text = value
        for run in cell.paragraphs[0].runs:
            run.bold = True
    for row_values in rows:
        cells = table.add_row().cells
        for cell, value in zip(cells, row_values):
            cell.text = value
    return table


def main():
    document = Document(DOCUMENT_PATH)
    document.add_paragraph().add_run().add_break(WD_BREAK.PAGE)

    document.add_heading("12. 导出记录", level=1)
    add_text(
        document,
        "软件应自动记录每次成功导出的单品牌库存对比表和库存汇总表，便于按文件回溯本次处理结果与保存位置。",
    )
    add_table(
        document,
        ["记录项", "规则"],
        [
            ["记录时机", "仅在 Excel 文件实际导出成功后新增一条记录；校验失败、用户取消或文件写入失败时不新增记录。"],
            ["记录内容", "文件名称、导出类型（单品牌库存对比/库存汇总）、生成时间、实际导出路径、关联品牌或品牌数量。"],
            ["列表能力", "支持按文件名称或品牌搜索，按导出类型筛选，并按生成时间倒序查看。"],
            ["打开位置", "用户可从记录中打开文件所在文件夹；文件被手动移动或删除时，软件应提示文件已不存在。"],
            ["重新导出", "用户可对记录执行重新导出。若关联的原始导入数据仍在当前会话中，则按原参数重新生成；数据已不可用时，提示重新导入源文件。"],
            ["保留策略", "V1 默认保留全部本地导出记录；不自动删除历史记录。"],
        ],
    )

    document.add_heading("13. 设置", level=1)
    add_text(document, "软件提供本地设置页面，用于管理 Excel 导出位置和导出文件命名规则。设置在当前电脑本地保存，下次启动后继续生效。")
    add_table(
        document,
        ["设置项", "规则"],
        [
            ["默认导出路径", "用户可选择本机文件夹作为单品牌库存对比表和库存汇总表的默认保存位置。首次使用时使用软件工作目录下的 exports 文件夹。"],
            ["路径校验", "保存设置及执行导出前均校验目标路径是否存在且可写；路径不存在时提示用户重新选择，不自动写入未知位置。"],
            ["单次调整", "每次导出前允许临时选择本次保存位置，不改变已保存的默认导出路径。"],
            ["对比表命名", "默认规则为 {品牌}_单品牌库存对比_{本周日期}.xlsx；支持修改为包含品牌、本周日期、生成日期的组合。"],
            ["汇总表命名", "默认规则为 库存汇总_{生成日期}.xlsx；支持修改为包含生成日期的组合。"],
            ["文件冲突", "目标目录已存在同名文件时，软件自动追加时间标识避免覆盖，并在导出记录中保存最终文件名。"],
            ["保存反馈", "用户点击保存设置后，软件显示保存成功反馈；设置未通过校验时，说明具体问题并不保存。"],
        ],
    )

    document.add_heading("14. 补充验收标准", level=1)
    for item in [
        "成功导出对比表或汇总表后，导出记录中应立即出现对应文件名称、类型、时间和实际保存路径。",
        "用户修改默认导出路径并保存后，下一次导出的文件应保存到新路径，重启软件后该设置仍有效。",
        "导出记录中打开文件夹操作应定位到文件实际所在目录；文件不存在时应给出明确提示。",
        "导出时出现同名文件，软件不得静默覆盖已有文件，应生成不重复的文件名。",
    ]:
        add_bullet(document, item)

    document.add_paragraph().add_run().add_break(WD_BREAK.PAGE)
    document.add_heading("15. 本轮交互与字段更新", level=1)
    add_text(document, "本节记录库存工作台当前版本已确认并纳入实现的模板、对比结果与导入交互更新。")
    add_table(
        document,
        ["项目", "规则"],
        [
            ["源表模板", "必填列固定为 10 列；“库存状态”位于“库存量(主单位)”和“库存组织”之间。库存状态不参与商品识别、数量计算、对比或汇总。"],
            ["对比环比", "对比表输出“环比变化”，位于“变动类型”前。计算方式为 (本周主单位 - 上周主单位) / |上周主单位|；上周主单位为 0 时显示“-”，箱数不计算环比变化。"],
            ["文件导入", "单品牌库存对比的上周、本周文件，以及库存汇总的多个品牌文件，均支持点击上传区选择或将 .xlsx 文件拖拽到上传区。"],
            ["会话清空", "单品牌库存对比和库存汇总均可清空当前已选文件和生成结果。清空同步清理服务端会话及应用保存的临时上传副本，不删除用户原始 Excel、已保存的德清仓映射、设置或导出记录。"],
            ["汇总重选", "重新选择或拖入汇总品牌文件时，替换当前汇总会话文件，并要求重新生成预览。"],
        ],
    )

    document.save(DOCUMENT_PATH)
    print(DOCUMENT_PATH.resolve())


if __name__ == "__main__":
    main()
