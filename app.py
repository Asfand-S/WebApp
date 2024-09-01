from flask import Flask, render_template, request, send_file
import pandas as pd
from utilities import *

app = Flask(__name__)

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        if 'file' not in request.files:
            return "No file uploaded", 400
        
        file = request.files['file']
        if file.filename == '':
            return "No selected file", 400
        
        if not (file.filename.endswith(".xlsx") or file.filename.endswith(".xls")):
            return "Uploaded file is not an Excel file", 400
        
        # Read Excel file
        df = pd.read_excel(file)
        
        # Generate PDF
        zip_output = generate_pdfs_zip(df)

        return send_file(zip_output, as_attachment=True, download_name='PDFs_compressed.zip', mimetype='application/zip')
    
    return render_template('index.html')
