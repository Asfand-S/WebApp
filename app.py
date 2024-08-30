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
        
        # Read Excel file
        df = pd.read_excel(file)
        
        # Generate PDF
        pdf_output = generate_pdf(df, "outputs")

        return send_file(pdf_output, as_attachment=True, download_name='output.pdf', mimetype='application/pdf')
    
    return render_template('index.html')

