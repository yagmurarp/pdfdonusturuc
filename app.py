from flask import Flask, render_template, request, send_file
import os
from werkzeug.utils import secure_filename
from pdf2docx import Converter
from docx2pdf import convert
import tempfile

app = Flask(__name__)

UPLOAD_FOLDER = 'uploads'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

@app.route('/')
def index():
    return '''
    <h1>PDF Dönüştürücü</h1>
    <form action="/pdf_to_word" method="post" enctype="multipart/form-data">
        <p>PDF'yi Word'e Dönüştür:</p>
        <input type="file" name="file">
        <input type="submit" value="Dönüştür">
    </form>
    <br>
    <form action="/word_to_pdf" method="post" enctype="multipart/form-data">
        <p>Word'ü PDF'ye Dönüştür:</p>
        <input type="file" name="file">
        <input type="submit" value="Dönüştür">
    </form>
    '''

@app.route('/pdf_to_word', methods=['POST'])
def pdf_to_word():
    file = request.files['file']
    if file.filename == '':
        return "Dosya seçilmedi."
    
    filename = secure_filename(file.filename)
    filepath = os.path.join(UPLOAD_FOLDER, filename)
    file.save(filepath)

    # PDF -> Word
    word_path = os.path.join(UPLOAD_FOLDER, filename.replace('.pdf', '.docx'))
    cv = Converter(filepath)
    cv.convert(word_path)
    cv.close()

    return send_file(word_path, as_attachment=True)

@app.route('/word_to_pdf', methods=['POST'])
def word_to_pdf():
    file = request.files['file']
    if file.filename == '':
        return "Dosya seçilmedi."

    filename = secure_filename(file.filename)
    filepath = os.path.join(UPLOAD_FOLDER, filename)
    file.save(filepath)

    # Word -> PDF
    pdf_path = os.path.join(UPLOAD_FOLDER, filename.replace('.docx', '.pdf'))
    convert(filepath, pdf_path)

    return send_file(pdf_path, as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True)
