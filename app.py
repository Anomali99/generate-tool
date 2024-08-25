from flask import Flask, render_template, request, jsonify, send_file, abort
import os, uuid, shutil, pythoncom, logging
import win32com.client as win32
import pandas as pd

app = Flask(__name__)
app.static_folder = os.path.join(os.getcwd(), 'static')
os.makedirs(os.path.join(app.static_folder, 'output'), exist_ok=True)
logging.basicConfig(level=logging.DEBUG)


@app.route("/", methods=["GET"])
def index():
    return render_template('index.html')


@app.route('/download/<userUUID>', methods=["GET"])
def download(userUUID):
    path = os.path.join(app.static_folder, 'output', userUUID + '.zip')
    if os.path.exists(path):
        return send_file(path, as_attachment=True, download_name=userUUID + '.zip')
    else:
        abort(404)


@app.route("/generate", methods=["POST"])
def generate():
    try:
        pythoncom.CoInitialize()

        format = request.form.get('format')
        filename = request.form.get('filename')
        word_file = request.files['word']
        excel_file = request.files['excel']
        userUUID = str(uuid.uuid4())

        output = os.path.join(app.static_folder, 'output', userUUID)
        temp_folder = os.path.join(output, 'temp')
        os.makedirs(output, exist_ok=True)
        os.makedirs(temp_folder, exist_ok=True)
        word_folder = None
        if format != 'word':
            word_folder = os.path.join(output, 'word')
            os.makedirs(word_folder, exist_ok=True)

        word_temp_path = os.path.join(temp_folder, word_file.filename)
        word_file.save(word_temp_path)

        excel_temp_path = os.path.join(temp_folder, excel_file.filename)
        excel_file.save(excel_temp_path)

        data_frame = pd.read_excel(excel_temp_path)
        data_list = data_frame.to_dict(orient="records")

        for idx, data in enumerate(data_list):
            current_filename = str(idx + 1)
            if filename.strip() != '':
                try:
                    current_filename = data[filename]
                except KeyError:
                    current_filename = str(idx + 1)
            
            if format == 'word':
                _generate_handle(data, word_temp_path, output, current_filename)
            else:
                word_filename = _generate_handle(data, word_temp_path, word_folder, current_filename)
                _save_pdf(word_filename,output, current_filename)

        if word_folder != None:
            shutil.rmtree(temp_folder)
        shutil.make_archive(output, 'zip', output)
        shutil.rmtree(output)  

        return jsonify({
            "status": 200,
            "message": "success",
            "data": userUUID
        }), 200

    except Exception as e:
        logging.error(f"Error occurred: {str(e)}")
        return jsonify({
            "status": 500,
            "message": str(e)
        }), 500


def _generate_handle(data, word_temp_path, output, current_filename):
    wordApp = win32.Dispatch("Word.Application")
    wordApp.Visible = False
    wordApp.DisplayAlerts = 0

    doc = wordApp.Documents.Open(word_temp_path)
    for paragraph in doc.Paragraphs:
        for key, value in data.items():
            if f'{{{{{key}}}}}' in paragraph.Range.Text:
                paragraph.Range.Text = paragraph.Range.Text.replace(f'{{{{{key}}}}}', str(value))

    for shape in doc.Shapes:
        if shape.TextFrame.HasText:
            for key, value in data.items():
                if f'{{{{{key}}}}}' in shape.TextFrame.TextRange.Text:
                    shape.TextFrame.TextRange.Text = shape.TextFrame.TextRange.Text.replace(f'{{{{{key}}}}}', str(value))
    
    word_filename = os.path.join(output, f"{current_filename}.docx")
    doc.SaveAs(word_filename)

    doc.Close()
    wordApp.Quit()

    return word_filename


def _save_pdf(word_filename, output, current_filename):
    wordApp = win32.Dispatch('Word.Application')
    wordApp.Visible = False
    wordApp.DisplayAlerts = 0
    doc = wordApp.Documents.Open(word_filename)
    pdf_filename = os.path.join(output, f"{current_filename}.pdf")
    doc.SaveAs(pdf_filename, FileFormat=17)  
    doc.Close()
    wordApp.Quit()


if __name__ == "__main__":
    app.run(port=7125, debug=True)
