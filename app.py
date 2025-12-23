from flask import Flask, render_template, request, send_file, jsonify
import os
import io
import zipfile
import json
from werkzeug.utils import secure_filename
from PyPDF2 import PdfReader, PdfWriter, PdfMerger
from PIL import Image
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle
from reportlab.lib import colors
import img2pdf
import fitz  # PyMuPDF
import pdfplumber # For extracting tables
import pandas as pd # For Excel/CSV handling
from pdf2docx import Converter # For PDF to Word
import pytesseract # For OCR
from pdf2image import convert_from_bytes # For PDF to Images

# Note: For Word -> PDF, this works best on Windows/Mac with MS Word installed
try:
    from docx2pdf import convert as docx_convert
except ImportError:
    docx_convert = None

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024  # 50MB max file size

# Ensure upload folder exists
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)

ALLOWED_EXTENSIONS = {'pdf', 'png', 'jpg', 'jpeg', 'docx', 'xlsx', 'csv'}

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/about')
def about():
    return render_template('about.html')

@app.route('/privacy')
def privacy():
    return render_template('privacy.html') 
@app.route('/services')
def services():
    return render_template('services.html')

# ==========================================

@app.route('/merge', methods=['POST'])
def merge_pdfs():
    try:
        files = request.files.getlist('files')
        if len(files) < 2:
            return jsonify({'error': 'Please upload at least 2 PDF files'}), 400
        
        merger = PdfMerger()
        for file in files:
            if file and allowed_file(file.filename):
                merger.append(file)
        
        output_path = os.path.join(app.config['UPLOAD_FOLDER'], 'merged.pdf')
        merger.write(output_path)
        merger.close()
        
        return send_file(output_path, as_attachment=True, download_name='merged.pdf')
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/split', methods=['POST'])
def split_pdf():
    try:
        file = request.files['file']
        page_ranges = request.form.get('pages', '').strip()
        
        if not file or not allowed_file(file.filename):
            return jsonify({'error': 'Invalid file'}), 400
        
        reader = PdfReader(file)
        total_pages = len(reader.pages)
        
        pages_to_extract = []
        if page_ranges:
            for part in page_ranges.split(','):
                if '-' in part:
                    start, end = map(int, part.split('-'))
                    pages_to_extract.extend(range(start-1, end))
                else:
                    pages_to_extract.append(int(part)-1)
        else:
            pages_to_extract = list(range(total_pages))
        
        writer = PdfWriter()
        for page_num in pages_to_extract:
            if 0 <= page_num < total_pages:
                writer.add_page(reader.pages[page_num])
        
        output_path = os.path.join(app.config['UPLOAD_FOLDER'], 'split.pdf')
        with open(output_path, 'wb') as output_file:
            writer.write(output_file)
        
        return send_file(output_path, as_attachment=True, download_name='split.pdf')
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/compress', methods=['POST'])
def compress_pdf():
    try:
        file = request.files['file']
        if not file: return jsonify({'error': 'Invalid file'}), 400
        
        reader = PdfReader(file)
        writer = PdfWriter()
        
        for page in reader.pages:
            page.compress_content_streams()
            writer.add_page(page)
        
        output_path = os.path.join(app.config['UPLOAD_FOLDER'], 'compressed.pdf')
        with open(output_path, 'wb') as output_file:
            writer.write(output_file)
        
        return send_file(output_path, as_attachment=True, download_name='compressed.pdf')
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/watermark', methods=['POST'])
def add_watermark():
    try:
        file = request.files['file']
        watermark_text = request.form.get('text', 'WATERMARK')
        if not file: return jsonify({'error': 'Invalid file'}), 400
        
        reader = PdfReader(file)
        writer = PdfWriter()
        
        packet = io.BytesIO()
        can = canvas.Canvas(packet, pagesize=letter)
        can.setFont('Helvetica', 50)
        can.setFillColorRGB(0.5, 0.5, 0.5, alpha=0.3)
        can.rotate(45)
        can.drawString(200, 100, watermark_text)
        can.save()
        packet.seek(0)
        
        watermark = PdfReader(packet)
        watermark_page = watermark.pages[0]
        
        for page in reader.pages:
            page.merge_page(watermark_page)
            writer.add_page(page)
        
        output_path = os.path.join(app.config['UPLOAD_FOLDER'], 'watermarked.pdf')
        with open(output_path, 'wb') as output_file:
            writer.write(output_file)
        
        return send_file(output_path, as_attachment=True, download_name='watermarked.pdf')
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/protect', methods=['POST'])
def protect_pdf():
    try:
        file = request.files['file']
        password = request.form.get('password', 'password123')
        if not file: return jsonify({'error': 'Invalid file'}), 400
        
        reader = PdfReader(file)
        writer = PdfWriter()
        for page in reader.pages:
            writer.add_page(page)
        writer.encrypt(password)
        
        output_path = os.path.join(app.config['UPLOAD_FOLDER'], 'protected.pdf')
        with open(output_path, 'wb') as output_file:
            writer.write(output_file)
        return send_file(output_path, as_attachment=True, download_name='protected.pdf')
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/unlock', methods=['POST'])
def unlock_pdf():
    try:
        file = request.files['file']
        password = request.form.get('password', '')
        if not file: return jsonify({'error': 'Invalid file'}), 400
        
        reader = PdfReader(file)
        if reader.is_encrypted:
            reader.decrypt(password)
        
        writer = PdfWriter()
        for page in reader.pages:
            writer.add_page(page)
        
        output_path = os.path.join(app.config['UPLOAD_FOLDER'], 'unlocked.pdf')
        with open(output_path, 'wb') as output_file:
            writer.write(output_file)
        return send_file(output_path, as_attachment=True, download_name='unlocked.pdf')
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/remove-pages', methods=['POST'])
def remove_pages():
    try:
        file = request.files['file']
        pages_to_remove = request.form.get('pages', '').strip()
        
        if not file or not allowed_file(file.filename):
            return jsonify({'error': 'Invalid file'}), 400
        
        reader = PdfReader(file)
        writer = PdfWriter()
        
        remove_list = []
        for part in pages_to_remove.split(','):
            if '-' in part:
                start, end = map(int, part.split('-'))
                remove_list.extend(range(start-1, end))
            else:
                remove_list.append(int(part)-1)
        
        for i, page in enumerate(reader.pages):
            if i not in remove_list:
                writer.add_page(page)
        
        output_path = os.path.join(app.config['UPLOAD_FOLDER'], 'removed_pages.pdf')
        with open(output_path, 'wb') as output_file:
            writer.write(output_file)
        
        return send_file(output_path, as_attachment=True, download_name='removed_pages.pdf')
    except Exception as e:
        return jsonify({'error': str(e)}), 500


# ==========================================
# 🆕 16 NEW TOOLS (NO DUPLICATES)
# ==========================================

# 1) PDF → Word
@app.route('/pdf-to-word', methods=['POST'])
def pdf_to_word():
    try:
        file = request.files['file']
        if not file: return jsonify({'error': 'Invalid file'}), 400
        
        input_path = os.path.join(app.config['UPLOAD_FOLDER'], 'input.pdf')
        output_path = os.path.join(app.config['UPLOAD_FOLDER'], 'converted.docx')
        file.save(input_path)
        
        cv = Converter(input_path)
        cv.convert(output_path, start=0, end=None)
        cv.close()
        
        return send_file(output_path, as_attachment=True, download_name='converted.docx')
    except Exception as e:
        return jsonify({'error': str(e)}), 500

# 2) PDF → Excel
@app.route('/pdf-to-excel', methods=['POST'])
def pdf_to_excel():
    try:
        file = request.files['file']
        if not file: return jsonify({'error': 'Invalid file'}), 400
        
        input_path = os.path.join(app.config['UPLOAD_FOLDER'], 'input_tables.pdf')
        output_path = os.path.join(app.config['UPLOAD_FOLDER'], 'tables.xlsx')
        file.save(input_path)
        
        with pdfplumber.open(input_path) as pdf:
            all_tables = []
            for page in pdf.pages:
                tables = page.extract_tables()
                for table in tables:
                    df = pd.DataFrame(table[1:], columns=table[0])
                    all_tables.append(df)
            
            if not all_tables:
                return jsonify({'error': 'No tables found'}), 400
            
            with pd.ExcelWriter(output_path) as writer:
                for i, df in enumerate(all_tables):
                    df.to_excel(writer, sheet_name=f'Table_{i+1}', index=False)
                    
        return send_file(output_path, as_attachment=True, download_name='converted_tables.xlsx')
    except Exception as e:
        return jsonify({'error': str(e)}), 500

# 3) PDF → CSV
@app.route('/pdf-to-csv', methods=['POST'])
def pdf_to_csv():
    try:
        file = request.files['file']
        if not file: return jsonify({'error': 'Invalid file'}), 400
        
        input_path = os.path.join(app.config['UPLOAD_FOLDER'], 'input_csv.pdf')
        file.save(input_path)
        
        output_zip = os.path.join(app.config['UPLOAD_FOLDER'], 'tables_csv.zip')
        
        with pdfplumber.open(input_path) as pdf:
            with zipfile.ZipFile(output_zip, 'w') as zipf:
                count = 0
                for i, page in enumerate(pdf.pages):
                    tables = page.extract_tables()
                    for j, table in enumerate(tables):
                        df = pd.DataFrame(table)
                        csv_data = df.to_csv(index=False, header=False)
                        zipf.writestr(f'page_{i+1}_table_{j+1}.csv', csv_data)
                        count += 1
                        
                if count == 0:
                    return jsonify({'error': 'No tables found'}), 400

        return send_file(output_zip, as_attachment=True, download_name='extracted_csvs.zip')
    except Exception as e:
        return jsonify({'error': str(e)}), 500

# 4) Extract Tables (Same backend as Excel)
@app.route('/extract-tables', methods=['POST'])
def extract_tables():
    return pdf_to_excel()

# 5) OCR Text Extract
@app.route('/ocr-pdf', methods=['POST'])
def ocr_pdf():
    try:
        file = request.files['file']
        if not file: return jsonify({'error': 'Invalid file'}), 400
        
        images = convert_from_bytes(file.read())
        full_text = ""
        for img in images:
            text = pytesseract.image_to_string(img)
            full_text += text + "\n\n--- Page Break ---\n\n"
            
        output_path = os.path.join(app.config['UPLOAD_FOLDER'], 'ocr_text.txt')
        with open(output_path, 'w', encoding='utf-8') as f:
            f.write(full_text)
            
        return send_file(output_path, as_attachment=True, download_name='ocr_extracted.txt')
    except Exception as e:
        return jsonify({'error': 'Ensure Tesseract is installed. ' + str(e)}), 500

# 6) Images → PDF
@app.route('/images-to-pdf', methods=['POST'])
def images_to_pdf():
    try:
        files = request.files.getlist('files')
        if not files: return jsonify({'error': 'No files uploaded'}), 400
        
        image_paths = []
        for file in files:
            path = os.path.join(app.config['UPLOAD_FOLDER'], secure_filename(file.filename))
            file.save(path)
            image_paths.append(path)
            
        output_path = os.path.join(app.config['UPLOAD_FOLDER'], 'converted_images.pdf')
        with open(output_path, "wb") as f:
            f.write(img2pdf.convert(image_paths))
            
        return send_file(output_path, as_attachment=True, download_name='images_combined.pdf')
    except Exception as e:
        return jsonify({'error': str(e)}), 500

# 7) PDF → Images (Replaces old 'convert-to-images')
@app.route('/pdf-to-all-images', methods=['POST'])
def pdf_to_all_images():
    try:
        file = request.files['file']
        if not file: return jsonify({'error': 'Invalid file'}), 400
        
        images = convert_from_bytes(file.read())
        zip_path = os.path.join(app.config['UPLOAD_FOLDER'], 'all_images.zip')
        
        with zipfile.ZipFile(zip_path, 'w') as zipf:
            for i, img in enumerate(images):
                img_byte_arr = io.BytesIO()
                img.save(img_byte_arr, format='PNG')
                zipf.writestr(f'page_{i+1}.png', img_byte_arr.getvalue())
                
        return send_file(zip_path, as_attachment=True, download_name='all_pages_images.zip')
    except Exception as e:
        return jsonify({'error': str(e)}), 500

# 8) PDF → Text
@app.route('/pdf-to-text', methods=['POST'])
def pdf_to_text_simple():
    try:
        file = request.files['file']
        if not file: return jsonify({'error': 'Invalid file'}), 400
        
        reader = PdfReader(file)
        full_text = ""
        for page in reader.pages:
            full_text += page.extract_text() + "\n\n"
            
        output_path = os.path.join(app.config['UPLOAD_FOLDER'], 'extracted_text.txt')
        with open(output_path, 'w', encoding='utf-8') as f:
            f.write(full_text)
            
        return send_file(output_path, as_attachment=True, download_name='pdf_text.txt')
    except Exception as e:
        return jsonify({'error': str(e)}), 500

# 9) Word → PDF
@app.route('/word-to-pdf', methods=['POST'])
def word_to_pdf():
    try:
        file = request.files['file']
        if not file or not file.filename.endswith('.docx'): 
            return jsonify({'error': 'Invalid file. Upload .docx'}), 400
            
        input_path = os.path.join(app.config['UPLOAD_FOLDER'], 'input.docx')
        output_path = os.path.join(app.config['UPLOAD_FOLDER'], 'word_converted.pdf')
        file.save(input_path)
        
        if docx_convert:
            docx_convert(input_path, output_path)
            return send_file(output_path, as_attachment=True, download_name='converted.pdf')
        else:
            return jsonify({'error': 'docx2pdf library not installed'}), 500
    except Exception as e:
        return jsonify({'error': str(e)}), 500

# 10) Excel → PDF
@app.route('/excel-to-pdf', methods=['POST'])
def excel_to_pdf():
    try:
        file = request.files['file']
        if not file: return jsonify({'error': 'Invalid file'}), 400
        
        df = pd.read_excel(file)
        data = [df.columns.values.tolist()] + df.values.tolist()
        
        output_path = os.path.join(app.config['UPLOAD_FOLDER'], 'excel_converted.pdf')
        
        doc = SimpleDocTemplate(output_path, pagesize=letter)
        elements = []
        t = Table(data)
        t.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('GRID', (0, 0), (-1, -1), 1, colors.black)
        ]))
        elements.append(t)
        doc.build(elements)
        
        return send_file(output_path, as_attachment=True, download_name='excel_converted.pdf')
    except Exception as e:
        return jsonify({'error': str(e)}), 500

# 11) Add E-Signature
@app.route('/add-signature', methods=['POST'])
def add_signature():
    try:
        pdf_file = request.files['pdf_file']
        sig_file = request.files['signature_file']
        page_num = int(request.form.get('page', 1)) - 1
        x = float(request.form.get('x', 100))
        y = float(request.form.get('y', 100))
        
        if not pdf_file or not sig_file: return jsonify({'error': 'Missing files'}), 400
        
        pdf_path = os.path.join(app.config['UPLOAD_FOLDER'], 'input_sign.pdf')
        pdf_file.save(pdf_path)
        sig_path = os.path.join(app.config['UPLOAD_FOLDER'], 'signature.png')
        sig_file.save(sig_path)
        
        doc = fitz.open(pdf_path)
        if 0 <= page_num < len(doc):
            page = doc[page_num]
            rect = fitz.Rect(x, y, x + 100, y + 50) 
            page.insert_image(rect, filename=sig_path)
        
        output_path = os.path.join(app.config['UPLOAD_FOLDER'], 'signed.pdf')
        doc.save(output_path)
        doc.close()
        
        return send_file(output_path, as_attachment=True, download_name='signed.pdf')
    except Exception as e:
        return jsonify({'error': str(e)}), 500

# 12) Rotate PDF
@app.route('/rotate-pdf', methods=['POST'])
def rotate_pdf():
    try:
        file = request.files['file']
        rotation = int(request.form.get('rotation', 90))
        if not file: return jsonify({'error': 'Invalid file'}), 400
        
        reader = PdfReader(file)
        writer = PdfWriter()
        for page in reader.pages:
            page.rotate(rotation)
            writer.add_page(page)
            
        output_path = os.path.join(app.config['UPLOAD_FOLDER'], 'rotated.pdf')
        with open(output_path, 'wb') as f:
            writer.write(f)
        return send_file(output_path, as_attachment=True, download_name='rotated.pdf')
    except Exception as e:
        return jsonify({'error': str(e)}), 500

# 13) Extract All Text + Images (Replaces old 'extract-images')
@app.route('/extract-all-content', methods=['POST'])
def extract_all_content():
    try:
        file = request.files['file']
        if not file: return jsonify({'error': 'Invalid file'}), 400
        
        input_path = os.path.join(app.config['UPLOAD_FOLDER'], 'input_extract.pdf')
        file.save(input_path)
        
        doc = fitz.open(input_path)
        zip_path = os.path.join(app.config['UPLOAD_FOLDER'], 'content.zip')
        
        with zipfile.ZipFile(zip_path, 'w') as zipf:
            full_text = ""
            for i, page in enumerate(doc):
                full_text += f"--- Page {i+1} ---\n{page.get_text()}\n\n"
                img_list = page.get_images()
                for j, img in enumerate(img_list):
                    xref = img[0]
                    base_image = doc.extract_image(xref)
                    zipf.writestr(f"page_{i+1}_img_{j+1}.{base_image['ext']}", base_image["image"])
            
            zipf.writestr("full_text.txt", full_text)
            
        return send_file(zip_path, as_attachment=True, download_name='extracted_content.zip')
    except Exception as e:
        return jsonify({'error': str(e)}), 500

# 14) Reorder PDF Pages
@app.route('/reorder-pdf', methods=['POST'])
def reorder_pdf():
    try:
        file = request.files['file']
        order_str = request.form.get('order', '') 
        if not file: return jsonify({'error': 'Invalid file'}), 400
        
        reader = PdfReader(file)
        writer = PdfWriter()
        total_pages = len(reader.pages)
        
        if order_str:
            indices = [int(x.strip()) - 1 for x in order_str.split(',') if x.strip().isdigit()]
        else:
            return jsonify({'error': 'No order provided'}), 400
            
        for idx in indices:
            if 0 <= idx < total_pages:
                writer.add_page(reader.pages[idx])
                
        output_path = os.path.join(app.config['UPLOAD_FOLDER'], 'reordered.pdf')
        with open(output_path, 'wb') as f:
            writer.write(f)
        return send_file(output_path, as_attachment=True, download_name='reordered.pdf')
    except Exception as e:
        return jsonify({'error': str(e)}), 500

# 15) Crop PDF
@app.route('/crop-pdf', methods=['POST'])
def crop_pdf():
    try:
        file = request.files['file']
        top = float(request.form.get('top', 0))
        bottom = float(request.form.get('bottom', 0))
        left = float(request.form.get('left', 0))
        right = float(request.form.get('right', 0))
        
        if not file: return jsonify({'error': 'Invalid file'}), 400
        
        reader = PdfReader(file)
        writer = PdfWriter()
        
        for page in reader.pages:
            box = page.mediabox
            box.upper_right = (box.right - right, box.top - top)
            box.lower_left = (box.left + left, box.bottom + bottom)
            page.mediabox = box
            writer.add_page(page)
            
        output_path = os.path.join(app.config['UPLOAD_FOLDER'], 'cropped.pdf')
        with open(output_path, 'wb') as f:
            writer.write(f)
        return send_file(output_path, as_attachment=True, download_name='cropped.pdf')
    except Exception as e:
        return jsonify({'error': str(e)}), 500

# 16) Edit PDF Metadata
@app.route('/edit-metadata', methods=['POST'])
def edit_metadata():
    try:
        file = request.files['file']
        title = request.form.get('title', '')
        author = request.form.get('author', '')
        if not file: return jsonify({'error': 'Invalid file'}), 400
        
        reader = PdfReader(file)
        writer = PdfWriter()
        for page in reader.pages:
            writer.add_page(page)
            
        metadata = {'/Title': title, '/Author': author, '/Producer': 'My PDF App'}
        writer.add_metadata(metadata)
        
        output_path = os.path.join(app.config['UPLOAD_FOLDER'], 'metadata_edited.pdf')
        with open(output_path, 'wb') as f:
            writer.write(f)
        return send_file(output_path, as_attachment=True, download_name='metadata_edited.pdf')
    except Exception as e:
        return jsonify({'error': str(e)}), 500


# ==========================================
# VISUAL PDF EDITOR (COORDINATES/TEXT OVERLAY)
# ==========================================

@app.route('/get-pdf-text', methods=['POST'])
def get_pdf_text():
    try:
        file = request.files['file']
        if not file or not allowed_file(file.filename):
            return jsonify({'error': 'Invalid file'}), 400
        
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], 'temp_edit.pdf')
        file.save(filepath)
        
        doc = fitz.open(filepath)
        pages_data = []
        
        for page_num in range(len(doc)):
            page = doc[page_num]
            rect = page.rect
            blocks = page.get_text("dict")["blocks"]
            text_blocks = []
            
            for block in blocks:
                if "lines" in block:
                    for line in block["lines"]:
                        for span in line["spans"]:
                            text_blocks.append({
                                'text': span['text'],
                                'x': span['bbox'][0],
                                'y': span['bbox'][1],
                                'width': span['bbox'][2] - span['bbox'][0],
                                'height': span['bbox'][3] - span['bbox'][1],
                                'font': span['font'],
                                'size': span['size'],
                                'color': span['color']
                            })
            
            pages_data.append({
                'page_num': page_num,
                'width': rect.width,
                'height': rect.height,
                'text_blocks': text_blocks
            })
        
        doc.close()
        return jsonify({'success': True, 'pages': pages_data, 'total_pages': len(pages_data)})
        
    except Exception as e:
        return jsonify({'error': str(e)}), 500


@app.route('/edit-pdf', methods=['POST'])
def edit_pdf():
    try:
        data = request.get_json()
        edits = data.get('edits', [])
        
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], 'temp_edit.pdf')
        if not os.path.exists(filepath):
            return jsonify({'error': 'PDF not found. Please upload again.'}), 400
        
        doc = fitz.open(filepath)
        edits_by_page = {}
        for edit in edits:
            page_num = edit['page']
            if page_num not in edits_by_page:
                edits_by_page[page_num] = []
            edits_by_page[page_num].append(edit)
        
        for page_num, page_edits in edits_by_page.items():
            page = doc[page_num]
            for edit in page_edits:
                rect = fitz.Rect(edit['x'], edit['y'], edit['x'] + edit['width'], edit['y'] + edit['height'])
                page.draw_rect(rect, color=(1, 1, 1), fill=(1, 1, 1))
                
                text_color = edit.get('color', 0)
                if isinstance(text_color, int):
                    r = ((text_color >> 16) & 255) / 255.0
                    g = ((text_color >> 8) & 255) / 255.0
                    b = (text_color & 255) / 255.0
                    color = (r, g, b)
                else:
                    color = (0, 0, 0)
                
                page.insert_text(
                    (edit['x'], edit['y'] + edit['size']),
                    edit['new_text'],
                    fontsize=edit['size'],
                    color=color,
                    fontname=edit.get('font', 'helv')
                )
        
        output_path = os.path.join(app.config['UPLOAD_FOLDER'], 'edited.pdf')
        doc.save(output_path)
        doc.close()
        
        return send_file(output_path, as_attachment=True, download_name='edited.pdf')
        
    except Exception as e:
        return jsonify({'error': str(e)}), 500

if __name__ == '__main__':
    app.run(debug=True)