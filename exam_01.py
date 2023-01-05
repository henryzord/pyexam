import os
from docx import Document
from docx.opc.coreprops import CoreProperties
from docx.shared import Cm
from docx.enum.text import WD_LINE_SPACING
from docx.shared import Pt
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
import json
import random


def configure_document():
    document = Document()
    style = document.styles['Normal']
    font = style.font
    font.name = 'Arial'
    font.size = Pt(12)
    return document


def main():
    filename = os.path.basename(__file__).split('.')[0]

    data = None
    with open(os.path.join('input', filename + '.json'), 'r', encoding='utf-8') as datafile:
        data = json.load(datafile)

    document = configure_document()  # type: Document

    p = document.add_paragraph(data['header'])

    random.seed(None)
    questions = data['questions']
    random.shuffle(questions)

    for i, question in enumerate(questions):
        p = document.add_paragraph('')

        p.add_run(question['name'].format(i + 1)).bold = True
        p.add_run(question['description'])

        if question['image'] is not None:
            width = eval(question['image']['width'])  # type: Cm

            document.add_picture(os.path.join('input', question['image']['path']), width=width)
            p = document.paragraphs[-1]
            p.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER

        for j, option in enumerate(question['options']):
            document.add_paragraph(option, style='List Number 2')

    # p = document.add_paragraph('A plain paragraph having some ')
    # p.add_run('bold').bold = True
    # p.add_run(' and some ')
    # p.add_run('italic.').italic = True
    #
    # document.add_heading('Heading, level 1', level=1)
    # document.add_paragraph('Intense quote', style='Intense Quote')
    #
    # document.add_paragraph(
    #     'first item in unordered list', style='List Bullet'
    # )
    # document.add_paragraph(
    #     'first item in ordered list', style='List Number'
    # )
    #
    # document.add_picture(os.path.join('input', 'monty-truth.jpg'), width=Inches(1.25))
    #
    # records = (
    #     (3, '101', 'Spam'),
    #     (7, '422', 'Eggs'),
    #     (4, '631', 'Spam, spam, eggs, and spam')
    # )
    #
    # table = document.add_table(rows=1, cols=3)
    # hdr_cells = table.rows[0].cells
    # hdr_cells[0].text = 'Qty'
    # hdr_cells[1].text = 'Id'
    # hdr_cells[2].text = 'Desc'
    # for qty, id, desc in records:
    #     row_cells = table.add_row().cells
    #     row_cells[0].text = str(qty)
    #     row_cells[1].text = id
    #     row_cells[2].text = desc
    #
    # document.add_page_break()

    document.save(os.path.join('output', filename + '.docx'))


if __name__ == '__main__':
    main()
