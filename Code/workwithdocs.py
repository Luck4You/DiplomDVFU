from docx.shared import Mm
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.enum.table import WD_ROW_HEIGHT_RULE

def MergeTable (table, from_index, to):
    table.cell(from_index, 0).merge(table.cell(from_index+to, 0))
    table.cell(from_index, 1).merge(table.cell(from_index+to, 1))
    table.cell(from_index, 2).merge(table.cell(from_index+to, 2))

def replace_text_in_paragraphs(paragraphs, data, index):
    for paragraph in paragraphs:
        for run in paragraph.runs:
            for column_name in data.columns:
                if f"{column_name}" in run.text:
                    run.text = run.text.replace(f"{column_name}", str(data.at[index, column_name]))

def replacetables(tables, data, index):
    for table in tables:
        for row in table.rows:
            for cell in row.cells:
                for column_name in data.columns:
                    if f"{column_name}" in cell.text:
                        if cell.paragraphs and cell.paragraphs[0].runs:
                            match column_name:
                                case 'FORM':
                                    cell.text = cell.text.replace(f"{column_name}", str(data.at[index, column_name])).lower()
                                    MakeFMT(cell,14,WD_ALIGN_PARAGRAPH.CENTER,WD_ALIGN_VERTICAL.TOP,0,False)
def MakeFMT(cell,
            font_size = 11,
            paragrapg_Align = WD_ALIGN_PARAGRAPH.CENTER,
            vertical_align = WD_ALIGN_VERTICAL.BOTTOM,
            font_space_after = 0,
            bond_type = False):
    cell.vertical_alignment = vertical_align
    rc = cell.paragraphs[0].runs[0]
    for paragraph in cell.paragraphs:
        fmt = paragraph.paragraph_format
        fmt.space_after = Mm(font_space_after)
        paragraph.alignment = paragrapg_Align
        for run in paragraph.runs:
            run.bold = bond_type
            run.font.size = Pt(font_size)

def Insert_Str_from (Table, index, disc, ze, mark):
    Table.cell(index, 0).text = disc
    cell = Table.cell(index, 0)
    MakeFMT(cell, 11, WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_VERTICAL.BOTTOM, 0, False)
        
    textForCell = ze
    textForCell = MakeZE(textForCell)
    Table.cell(index, 1).text = textForCell
    cell = Table.cell(index, 1)
    MakeFMT(cell, 11, WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_VERTICAL.BOTTOM, 0, False)

    textForCell = mark
    textForCell = textForCell.lower()
    Table.cell(index, 2).text = textForCell
    cell = Table.cell(index, 2)
    MakeFMT(cell, 11, WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_VERTICAL.BOTTOM, 0, False)

def MakeZE(old_str):
    startRec = False
    new_str = ""
    if old_str == "":
        return ""
    for word in old_str:
        if word == ")": break
        elif startRec: new_str = new_str + word
        elif word == "(": startRec = True
    new_str = new_str + " з.е."
    return new_str