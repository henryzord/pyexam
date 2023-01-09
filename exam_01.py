import os
from docx import Document
from docx.opc.coreprops import CoreProperties
from docx.shared import Cm
from docx.enum.text import WD_LINE_SPACING
from docx.shared import Pt
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.enum.table import WD_TABLE_ALIGNMENT
import json
import random
from math import ceil


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

    n_cols = 5
    n_rows = 2 * ceil(len(questions) / n_cols)

    table = document.add_table(rows=n_rows, cols=n_cols)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER

    counter = 1
    for i in range(n_rows):
        if i % 2 == 0:
            for j in range(n_cols):
                if counter > len(questions):
                    break
                table.cell(i, j).text = 'Questão {0}'.format(counter)
                table.cell(i, j).vertical_alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
                # cell.text = 'Questão {0}'.format(counter)
                # cell.vertical_alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
                counter += 1
        else:
            for j in range(n_cols):
                table.cell(i, j).text = ' '

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

    # document.add_page_break()

    document.save(os.path.join('output', filename + '.docx'))


if __name__ == '__main__':
    main()
