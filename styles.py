from docx.shared import Cm, Pt
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.enum.section import WD_ORIENT
from docx.enum.text import WD_ALIGN_PARAGRAPH

DIMESSION_4A = [7772400, 10058400]

def set_column_width(cells, width):
    for cell in cells:
        cell.width = width
    
def set_header_style(table):
    set_column_width(table.columns[0].cells, Cm(5.))
    set_column_width(table.columns[1].cells, Cm(10.))
    set_column_width(table.columns[2].cells, Cm(5.))

def set_orientation(doc):
    for section in doc.sections:
        new_width, new_height = DIMESSION_4A[1], DIMESSION_4A[0]
        section.orientation = WD_ORIENT.LANDSCAPE
        section.page_width = new_width
        section.page_height = new_height

def set_table_style(table, columns_width):
    table.alignment = WD_TABLE_ALIGNMENT.CENTER

    table.allow_autofit = False
    columns = table.columns

    #libreoffice
    # columns[0].width = Cm(.8)
    # columns[1].width = Cm(9.)
    # columns[2].width = Cm(.9)
    # columns[-1].width = Cm(7.)

    # for i in range(len(columns)-2):
    #     columns[1+2].width = Cm(1.8)
        # set_column_width(columns[1+2].cells, Cm(1.8))
    
    # set_column_width(columns[0].cells, Cm(.5))
    # set_column_width(columns[1].cells, Cm(9.))
    # set_column_width(columns[2].cells, Cm(.9))
    # set_column_width(columns[-1].cells, Cm(7.))
    
    #Word
    for width, column in zip(columns_width, columns):
        set_column_width(column.cells, Cm(width))

    for c in table.rows[0].cells:
        p = c.paragraphs[0]
        r = p.runs[-1]
        font = r.font
        font.bold = True
        p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER

    table.style = 'Table Grid'

def set_doc_style(doc):
    style = doc.styles['Normal']
    font = style.font
    font.name = 'Arial'
    font.size = Pt(10)