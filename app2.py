from flask import Flask, request, redirect, url_for
import pandas as pd
from werkzeug.utils import secure_filename
import os

app = Flask(_name_)
UPLOAD_FOLDER = 'uploads'
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)

@app.route('/')
def index():
    return '''
        <html>
            <head>
                <title>How to upload files using HTML to website?</title>
            </head>
            <body style="text-align: center;">
                <h1 style="color: green;">Welcome to GeeksforGeeks</h1>
                <h2>How to upload files using HTML to website?</h2>
                <form action="/upload" method="post" enctype="multipart/form-data">
                    <input type="file" id="file1" name="upload">
                    <br><br>
                    <input type="submit" value="Upload File">
                </form>
            </body>
        </html>
    '''

@app.route('/upload', methods=['POST'])
def upload_file():
    if 'upload' not in request.files:
        return redirect(request.url)
    
    file = request.files['upload']
    if file.filename == '':
        return redirect(request.url)

    if file:
        filename = secure_filename(file.filename)
        file_path = os.path.join(app.config['C:\Users\T.Venkateshwar Reddy\OneDrive\Desktop\inter'], filename)
        file.save(file_path)
        
        # Reading the uploaded file content
        with open(file_path, 'r') as f:
            content = f.read()

        # Writing content to an Excel file
        df = pd.DataFrame([content.splitlines()])
        excel_path = os.path.join(app.config['C:\Users\T.Venkateshwar Reddy\OneDrive\Desktop\inter\output.xlsx'], 'output.xlsx')
        df.to_excel(excel_path, index=False, header=False)

        return f'File uploaded and content saved to {excel_path}'

if _name_ == "_main_":
    app.run(debug=True)