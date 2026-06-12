from concurrent.futures import ThreadPoolExecutor
import os
import re
import json
import pandas as pd
from pypdf import PdfReader, PdfWriter
from flask import Flask, render_template, request, send_file, jsonify
from io import BytesIO
import tempfile
import pypdfium2 as pdfium
import easyocr
import numpy as np
from PIL import Image

app = Flask(__name__)

# Initialize EasyOCR Reader once
# Using gpu=False explicitly to avoid overhead check if no GPU, though it's auto
reader = easyocr.Reader(['en'])

# Max workers for parallel processing - adjust based on CPU cores
MAX_WORKERS = os.cpu_count() or 4

def get_pdf_size(pdf_stream):
    try:
        pdf_stream.seek(0)
        r = PdfReader(pdf_stream)
        page = r.pages[0]
        w, h = float(page.mediabox.width), float(page.mediabox.height)
        
        # 1mm = 2.83465 points
        # Label (99x200): approx 280x567 pts
        # A4: approx 595x842 pts or 612x792 (Letter/A4 approx)
        if (250 < w < 380 and 500 < h < 700) or (250 < h < 380 and 500 < w < 700):
            return "label"
        elif (500 < w < 750 and 700 < h < 950) or (500 < h < 750 and 700 < w < 950):
            return "a4"
        return "other"
    except:
        return "error"

def expand_ref_numbers(ref_str):
    if not ref_str or ref_str == "Not Found":
        return ["Not Found"]
    
    # Split by semicolon or comma
    parts = [p.strip() for p in re.split(r'[;,]', ref_str)]
    expanded = []
    base_prefix = ""
    
    for p in parts:
        if not p: continue
        
        # Remove 'RMA' prefix if present
        clean_p = re.sub(r'^RMA-?', '', p, flags=re.IGNORECASE).strip()
        
        if clean_p.startswith('400'):
            expanded.append(clean_p)
            base_prefix = clean_p[:6]
        elif clean_p.startswith('-') and base_prefix:
            expanded.append(base_prefix + clean_p[1:])
        else:
            # Just digits
            digits = re.sub(r'\D', '', clean_p)
            if digits:
                expanded.append(digits)
            
    return expanded

def is_valid_label_page(page):
    """Checks if a page should be printed (excludes receipts/docs)."""
    text = page.extract_text()
    if not text or not text.strip():
        # If no text, we assume it's a label (likely the OCR case)
        return True
    
    text_lower = text.lower()
    exclude_keywords = ["waybill doc", "receipt", "archive doc", "copy for your records"]
    for kw in exclude_keywords:
        if kw in text_lower:
            return False
    return True

def extract_single_pdf(file_content, filename, manual_ref=None):
    """Worker function to process a single PDF in a thread."""
    stream = BytesIO(file_content)
    try:
        r = PdfReader(stream)
        full_text = ""
        
        # 1. Standard text extraction
        for page in r.pages:
            full_text += page.extract_text() + "\n"
        
        # 2. Rotation check
        if not full_text.strip():
            for page in r.pages:
                full_text += page.extract_text(orientations=(90,)) + "\n"

        # 3. OCR Fallback
        if not full_text.strip():
            stream.seek(0)
            doc = pdfium.PdfDocument(stream)
            for i in range(len(doc)):
                page = doc[i]
                for rot in [0, 90]:
                    bitmap = page.render(scale=1.5, rotation=rot)
                    ocr_result = reader.readtext(np.array(bitmap.to_pil()), detail=0)
                    page_text = " ".join(ocr_result)
                    if "Ref" in page_text or "WAYBILL" in page_text:
                        full_text += page_text + "\n"
                        break
            doc.close()

        normalized_text = " ".join(full_text.split())
        
        # Ref No Extraction (only if not manually provided)
        if manual_ref:
            ref_raw = manual_ref
        else:
            ref_match = re.search(r'Ref(?:\s*No)?[:\s]*([40RMA\d\s\-\;\,]+)', normalized_text, re.IGNORECASE)
            ref_raw = "Not Found"
            if ref_match:
                ref_raw = ref_match.group(1).strip()
                ref_raw = re.split(r'\s{2,}', ref_raw)[0]
        
        # Tracking Number Extraction (Strict 10 digits)
        waybill_match = re.search(r'(?:Waybill|Tracking\s*No)[:\s]*([\d\s]{10,15})', normalized_text, re.IGNORECASE)
        tracking_no = "Not Found"
        if waybill_match:
            digits_only = re.sub(r'\D', '', waybill_match.group(1))
            if len(digits_only) >= 10:
                tracking_no = digits_only[:10]
        
        if tracking_no == "Not Found":
            fallback_match = re.search(r'waybill\s*([\d\s]{10,15})', normalized_text, re.IGNORECASE)
            if fallback_match:
                digits_only = re.sub(r'\D', '', fallback_match.group(1))
                if len(digits_only) >= 10:
                    tracking_no = digits_only[:10]
        
        size = get_pdf_size(stream)
        return ref_raw, tracking_no, size
    except Exception as e:
        print(f"Error processing {filename}: {e}")
        return manual_ref if manual_ref else "Error", "Error", "Unknown"

@app.route('/')
def index():
    return render_template('index.html')

# Global storage for progress tracking
progress_data = {}

@app.route('/progress/<job_id>')
def get_progress(job_id):
    return jsonify(progress_data.get(job_id, {"current": 0, "total": 0}))

@app.route('/parse', methods=['POST'])
def parse():
    job_id = request.form.get('job_id')
    files = request.files.getlist('files')
    manual_refs_json = request.form.get('manual_refs', '{}')
    try:
        manual_refs = json.loads(manual_refs_json)
    except:
        manual_refs = {}
    
    if job_id:
        progress_data[job_id] = {"current": 0, "total": len(files)}

    file_data = []
    for f in files:
        content = f.read()
        m_ref = manual_refs.get(f.filename)
        file_data.append((content, f.filename, m_ref))
    
    results = []
    with ThreadPoolExecutor(max_workers=MAX_WORKERS) as executor:
        futures = [executor.submit(extract_single_pdf, d[0], d[1], d[2]) for d in file_data]
        for future in futures:
            ref_raw, track, size = future.result()
            expanded_refs = expand_ref_numbers(ref_raw)
            for ref in expanded_refs:
                results.append({
                    "Ref No": ref,
                    "Tracking Number": track,
                    "Size Type": size
                })
            if job_id and job_id in progress_data:
                progress_data[job_id]["current"] += 1
    
    if job_id and job_id in progress_data:
        del progress_data[job_id]

    return jsonify(results)

@app.route('/download_excel', methods=['POST'])
def download_excel():
    data = request.json
    if not data:
        return jsonify({"error": "데이터가 없습니다."}), 400
    
    # Create DataFrame with specific columns
    # A, B are empty, C is Tracking, D is SAP Order(s)
    df_data = []
    for item in data:
        df_data.append({
            "": "",               # Column A (Empty)
            " ": " ",              # Column B (Empty)
            "Tracking": item.get("Tracking Number", ""),
            "SAP Order(s)": item.get("Ref No", "")
        })
    
    df = pd.DataFrame(df_data)
    
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False)
    output.seek(0)
    
    return send_file(output, as_attachment=True, download_name="extraction_results.xlsx")



@app.route('/print_filter', methods=['POST'])
def print_filter():
    files = request.files.getlist('files')
    target_size = request.form.get('target_size')
    file_data = [(f.read(), f.filename) for f in files if f.filename.endswith('.pdf')]
    
    writer = PdfWriter()
    found_any = False
    
    def check_size_worker(data_pair):
        data, name = data_pair
        return get_pdf_size(BytesIO(data)) == target_size, data
    
    with ThreadPoolExecutor(max_workers=MAX_WORKERS) as executor:
        check_results = list(executor.map(check_size_worker, file_data))
        
    for is_target, data in check_results:
        if is_target:
            reader = PdfReader(BytesIO(data))
            for page in reader.pages:
                # Page filtering: exclude receipts and waybill docs
                if is_valid_label_page(page):
                    writer.add_page(page)
                    found_any = True
    
    if not found_any:
        return jsonify({"error": f"인쇄 가능한 {target_size} 규격의 페이지가 없습니다."}), 404
        
    output = BytesIO()
    writer.write(output)
    output.seek(0)
    
    return send_file(output, mimetype='application/pdf')

if __name__ == '__main__':
    app.run(debug=True, port=5000)
