import os
from flask import Flask, request, render_template, redirect, url_for, send_file
from werkzeug.utils import secure_filename
from extract import collect_files, extract_macros, analyze_vba_code, generate_documentation, generate_flowchart
from fpdf import FPDF
import base64

app = Flask(__name__)

# Configuration
UPLOAD_FOLDER = 'uploads'
OUTPUT_FOLDER = 'outputs'
ALLOWED_EXTENSIONS = {'xls', 'xlsm'}

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['OUTPUT_FOLDER'] = OUTPUT_FOLDER

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def generate_pdf_report(analysis, flowchart_image, output_path):
    pdf = FPDF()
    pdf.set_auto_page_break(auto=True, margin=15)
    pdf.add_page()
    pdf.set_font("Arial", size=12)

    pdf.set_title("VBA Macro Analysis")

    pdf.set_font("Arial", size=16, style='B')
    pdf.cell(200, 10, txt="Functions and Subroutines", ln=True)
    pdf.set_font("Arial", size=12)
    for func in analysis["functions"]:
        pdf.cell(200, 10, txt=func, ln=True)

    pdf.set_font("Arial", size=16, style='B')
    pdf.cell(200, 10, txt="Variables", ln=True)
    pdf.set_font("Arial", size=12)
    for var in analysis["variables"]:
        pdf.cell(200, 10, txt=var, ln=True)

    pdf.set_font("Arial", size=16, style='B')
    pdf.cell(200, 10, txt="Logic Flow", ln=True)
    pdf.set_font("Arial", size=12)
    for logic in analysis["logic_flow"]:
        pdf.cell(200, 10, txt=logic, ln=True)

    pdf.set_font("Arial", size=16, style='B')
    pdf.cell(200, 10, txt="Data Flow", ln=True)
    pdf.set_font("Arial", size=12)
    for data in analysis["data_flow"]:
        pdf.cell(200, 10, txt=data, ln=True)

    # Check if flowchart_image exists and is a valid PNG file
    if os.path.isfile(flowchart_image) and flowchart_image.lower().endswith('.png'):
        pdf.add_page()
        pdf.image(flowchart_image, x=10, y=10, w=180)
    else:
        pdf.cell(200, 10, txt="Flowchart image not found or invalid", ln=True)

    pdf_output_path = f"{output_path}.pdf"
    pdf.output(pdf_output_path)
    return pdf_output_path

@app.route('/')
def index():
    output_files = os.listdir(app.config['OUTPUT_FOLDER'])
    return render_template('index.html', output_files=output_files)

@app.route('/analyze', methods=['POST'])
def analyze():
    # Clear existing files in output directory
    for file in os.listdir(app.config['OUTPUT_FOLDER']):
        file_path = os.path.join(app.config['OUTPUT_FOLDER'], file)
        try:
            if os.path.isfile(file_path):
                os.remove(file_path)
        except Exception as e:
            print(f"Failed to delete {file_path}. Reason: {e}")

    if 'files[]' not in request.files:
        return redirect(request.url)

    files = request.files.getlist('files[]')
    src_dir = app.config['UPLOAD_FOLDER']
    dest_dir = app.config['OUTPUT_FOLDER']
    min_file_id = 1

    for file in files:
        if file and allowed_file(file.filename):
            filename = secure_filename(file.filename)
            file_path = os.path.join(src_dir, filename)
            file.save(file_path)

            macros = extract_macros(file_path)
            for macro in macros:
                analysis = analyze_vba_code(macro)
                output_path = os.path.join(dest_dir, f'macro_{min_file_id}')
                generate_documentation(analysis, output_path + '.txt')
                generate_flowchart(analysis, output_path + '.png')
                min_file_id += 1

            os.remove(file_path)  # Remove uploaded file after processing

    return redirect(url_for('index'))

@app.route('/download/<filename>', methods=['GET'])
def download(filename):
    return send_file(os.path.join(app.config['OUTPUT_FOLDER'], filename), as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True)
