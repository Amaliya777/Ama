
from flask import Flask, render_template, request
from docx import Document
from docx.shared import Pt
from docx.oxml.ns import qn

app = Flask(__name__)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/generate', methods=['POST'])
def generate():
    title = request.form['title']
    student = request.form['student']
    supervisor = request.form['supervisor']
    intro = request.form['intro']
    chapter1 = request.form['chapter1']
    chapter2 = request.form['chapter2']
    conclusion = request.form['conclusion']
    sources = request.form['sources'].strip().split('\n')

    doc = Document()

    style = doc.styles['Normal']
    font = style.font
    font.name = 'Times New Roman'
    font.size = Pt(14)
    style._element.rPr.rFonts.set(qn('w:eastAsia'), 'Times New Roman')

    doc.add_paragraph('Министерство науки и высшего образования РФ\n'
                      'Название ВУЗа\nФакультет / Кафедра\n\n'
                      'ДИПЛОМНАЯ РАБОТА\n\n'
                      f'на тему: "{title}"\n\n\n'
                      f'Студент: {student}\n'
                      f'Научный руководитель: {supervisor}\n\n'
                      'Город, 2025')

    doc.add_page_break()

    doc.add_heading('Введение', level=1)
    doc.add_paragraph(intro)

    doc.add_heading('Глава 1', level=1)
    doc.add_paragraph(chapter1)

    doc.add_heading('Глава 2', level=1)
    doc.add_paragraph(chapter2)

    doc.add_heading('Заключение', level=1)
    doc.add_paragraph(conclusion)

    doc.add_heading('Список использованных источников', level=1)
    for i, source in enumerate(sources, 1):
        doc.add_paragraph(f'{i}. {source.strip()}', style='List Number')

    filename = 'diploma.docx'
    doc.save(f'static/{filename}')
    return render_template('diploma.html', 
    filename=filename)

if __name__ == '__main__':
    app.run(debug=True)
