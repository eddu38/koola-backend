import os
import datetime
import pdfplumber
import pandas as pd
from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
from werkzeug.utils import secure_filename
import tempfile

app = Flask(__name__)
# Enable CORS for all domains so your Vercel frontend can talk to this
CORS(app, resources={r"/*": {"origins": "*"}})

UPLOAD_FOLDER = tempfile.gettempdir()
ALLOWED_EXTENSIONS = {'pdf'}

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def clean_dataframe(df, page_num, table_num):
    df = df.replace(r'^\s*$', pd.NA, regex=True)
    df = df.dropna(how='all')
    df = df.dropna(axis=1, how='all')
    df = df.reset_index(drop=True)
    if len(df) > 0:
        if df.iloc[0].notna().any():
            first_row = df.iloc[0].astype(str).str.strip()
            if len(first_row.unique()) == len(first_row):
                df.columns = first_row
                df = df[1:].reset_index(drop=True)
    return df

@app.route('/', methods=['GET'])
def home():
    return "Koola Backend is Running!", 200

@app.route('/api/convert/excel', methods=['POST'])
def convert_pdf_to_excel():
    if 'files' not in request.files:
        return jsonify({'error': 'No file part'}), 400
    file = request.files['files']
    if file.filename == '':
        return jsonify({'error': 'No selected file'}), 400
    
    if file and allowed_file(file.filename):
        try:
            filename = secure_filename(file.filename)
            filepath = os.path.join(UPLOAD_FOLDER, filename)
            file.save(filepath)
            
            all_data_frames = []
            with pdfplumber.open(filepath) as pdf:
                for page_num, page in enumerate(pdf.pages, 1):
                    tables = page.extract_tables()
                    for table_num, table in enumerate(tables, 1):
                        if table and len(table) > 0:
                            df = pd.DataFrame(table)
                            df = clean_dataframe(df, page_num, table_num)
                            if not df.empty:
                                all_data_frames.append(df)
            
            if not all_data_frames:
                return jsonify({'error': 'No tables found in PDF'}), 404

            combined_df = pd.concat(all_data_frames, ignore_index=True)
            combined_df = combined_df.fillna('')
            timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
            output_filename = f"converted_{timestamp}.xlsx"
            output_path = os.path.join(UPLOAD_FOLDER, output_filename)
            
            with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
                combined_df.to_excel(writer, sheet_name="Combined_Data", index=False)
            
            return send_file(output_path, as_attachment=True, download_name=output_filename)
            
        except Exception as e:
            return jsonify({'error': str(e)}), 500
    return jsonify({'error': 'Invalid file type'}), 400

if __name__ == '__main__':
    port = int(os.environ.get("PORT", 5000))
    app.run(host='0.0.0.0', port=port)
