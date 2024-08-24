from flask import Flask, render_template, request, jsonify, send_file, abort
import os, uuid, shutil, pythoncom, tempfile, logging, time
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
    word_temp_path, excel_temp_path = None, None  
    wordApp = None
    try:
        pythoncom.CoInitialize()

        format = request.form.get('format')
        filename = request.form.get('filename')
        word_file = request.files['word']
        excel_file = request.files['excel']
        userUUID = str(uuid.uuid4())

        output = os.path.join(app.static_folder, 'output', userUUID)
        os.makedirs(output, exist_ok=True)

        word_temp_path = os.path.join(tempfile.gettempdir(), word_file.filename)
        word_file.save(word_temp_path)

        excel_temp_path = os.path.join(tempfile.gettempdir(), excel_file.filename)
        excel_file.save(excel_temp_path)

        logging.debug(f"Word file saved at: {word_temp_path}")
        logging.debug(f"Excel file saved at: {excel_temp_path}")

        wordApp = win32.Dispatch("Word.Application")
        wordApp.Visible = False  # Hide Word window
        wordApp.DisplayAlerts = 0  # Turn off Word alerts

        doc = wordApp.Documents.Open(word_temp_path)

        data_frame = pd.read_excel(excel_temp_path)
        data_list = data_frame.to_dict(orient="records")

        for idx, data in enumerate(data_list):
            for paragraph in doc.Paragraphs:
                for key, value in data.items():
                    if f'{{{{{key}}}}}' in paragraph.Range.Text:
                        paragraph.Range.Text = paragraph.Range.Text.replace(f'{{{{{key}}}}}', str(value))

            for shape in doc.Shapes:
                if shape.TextFrame.HasText:
                    for key, value in data.items():
                        if f'{{{{{key}}}}}' in shape.TextFrame.TextRange.Text:
                            shape.TextFrame.TextRange.Text = shape.TextFrame.TextRange.Text.replace(f'{{{{{key}}}}}', str(value))
            
            current_filename = str(idx + 1)
            if filename.strip() != '':
                try:
                    current_filename = data[filename]
                except Exception as e:
                    current_filename = str(idx + 1)
                
            if format == 'word':
                word_filename = os.path.join(output, f"{current_filename}.docx")
                doc.SaveAs(word_filename)
            else:
                pdf_filename = os.path.join(output, f"{current_filename}.pdf")
                doc.SaveAs(pdf_filename, FileFormat=17)

        doc.Close()
        
        wordApp.Quit()
        wordApp = None  

        shutil.make_archive(output, 'zip', output)
        shutil.rmtree(output)

        return jsonify({
            "status": 200,
            "message": "success",
            "data" : userUUID
        }), 200
    except Exception as e:
        logging.error(f"Error occurred: {str(e)}")
        if wordApp:
            try:
                wordApp.Quit()
            except Exception as quit_error:
                logging.error(f"Failed to quit Word application: {quit_error}")
        return jsonify({
            "status": 500,
            "message": str(e)
        }), 500
    finally:
        time.sleep(1) 
        try:
            if word_temp_path and os.path.exists(word_temp_path):
                os.remove(word_temp_path)
        except Exception as e:
            logging.error(f"Failed to remove word_temp_path: {str(e)}")
            
        try:
            if excel_temp_path and os.path.exists(excel_temp_path):
                os.remove(excel_temp_path)
        except Exception as e:
            logging.error(f"Failed to remove excel_temp_path: {str(e)}")


if __name__ == "__main__":
    app.run(port=7125, debug=True)
