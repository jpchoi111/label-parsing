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
import zipfile

app = Flask(__name__)

# Helper to extract images from template if they don't exist
def ensure_images_extracted():
    template_path = 'MEDIT_Delivery Note_template.xlsx'
    if not os.path.exists(template_path):
        return
    
    with zipfile.ZipFile(template_path, 'r') as z:
        # Map of zip path to local path
        media_map = {
            'xl/media/image1.png': 'image1.png',
            'xl/media/image2.png': 'image2.png'
        }
        for zip_p, local_p in media_map.items():
            if not os.path.exists(local_p):
                try:
                    with open(local_p, 'wb') as f:
                        f.write(z.read(zip_p))
                    print(f"Extracted {local_p} from template.")
                except:
                    pass

ensure_images_extracted()

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

from datetime import datetime

@app.route('/parse_picking_list', methods=['POST'])
def parse_picking_list_endpoint():
    file = request.files.get('picking_list')
    tracking_data_json = request.form.get('tracking_data', '[]')
    
    if not file:
        return jsonify({"error": "Picking List 파일을 업로드해주세요."}), 400
    
    try:
        tracking_map = {}
        try:
            tracking_list = json.loads(tracking_data_json)
            for item in tracking_list:
                ref = item.get('Ref No')
                track = item.get('Tracking Number')
                if ref and track and track != "Not Found":
                    tracking_map[ref] = track
        except:
            pass

        stream = BytesIO(file.read())
        reader = PdfReader(stream)
        orders = []
        
        # Deduplication tracker: (Packing No) - one row per order
        seen_packing_nos = set()

        for i in range(2, len(reader.pages)):
            text = reader.pages[i].extract_text()
            parts = re.split(r'(Packing No\.)', text)
            
            for j in range(1, len(parts), 2):
                if j+1 < len(parts):
                    segment = parts[j] + parts[j+1]
                    
                    p_match = re.search(r'Packing No\.\s*(\S+)', segment)
                    o_match = re.search(r'OrderNo\.\s*(\S+)', segment)
                    u_match = re.search(r'UserID\.\s*(\S+)', segment)
                    
                    if p_match and o_match and u_match:
                        packing_no = p_match.group(1)
                        order_no = o_match.group(1)
                        user_id = u_match.group(1)
                        
                        # Only one row per Packing No
                        if packing_no not in seen_packing_nos:
                            seen_packing_nos.add(packing_no)
                            
                            # Mapping UserID
                            company_name = ""
                            if user_id == "MEDITFRA":
                                company_name = "MEDIT EUROPE GMBH"
                            elif user_id == "MEDITRMA":
                                company_name = "MEDIT EUROPE"
                            
                            # Match Tracking
                            track_no = tracking_map.get(order_no, "")

                            orders.append({
                                "A": "", # Empty, A1 will be set separately
                                "B": packing_no,
                                "C": company_name,
                                "D": "",
                                "E": track_no
                            })
        
        if not orders:
            return jsonify({"error": "오더를 찾을 수 없습니다."}), 404

        # Create Excel
        df = pd.DataFrame(orders)
        
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            # Write data starting from B, C, D, E (A is empty in DF)
            df.to_excel(writer, index=False, header=False, sheet_name='Sheet1')
            
            # Set A1 date specifically
            workbook = writer.book
            worksheet = writer.sheets['Sheet1']
            worksheet['A1'] = datetime.now().strftime("%Y-%m-%d")
            
        output.seek(0)
        
        return send_file(output, as_attachment=True, download_name="picking_list_results.xlsx")
    except Exception as e:
        print(f"Error: {e}")
        return jsonify({"error": str(e)}), 500

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

@app.route('/generate_delivery_note', methods=['POST'])
def generate_delivery_note():
    source_file = request.files.get('source_file')
    if not source_file:
        return jsonify({"error": "원본 엑셀 파일(20260612.xls)을 업로드해주세요."}), 400
    
    template_path = 'MEDIT_Delivery Note_template.xlsx'
    if not os.path.exists(template_path):
        return jsonify({"error": "템플릿 파일이 서버에 존재하지 않습니다."}), 500

    try:
        from openpyxl import load_workbook
        from openpyxl.drawing.image import Image as XLImage
        from copy import copy
        
        # Load source data with resilience
        source_content = source_file.read()
        source_df = None
        
        # 1. Try default read (usually xlrd for .xls)
        try:
            source_df = pd.read_excel(BytesIO(source_content))
        except Exception as e1:
            print(f"Primary read failed: {e1}")
            # 2. Try xlrd with ignore_workbook_corruption=True (Definitive fix for seen[2]==4)
            try:
                import xlrd
                book = xlrd.open_workbook(file_contents=source_content, ignore_workbook_corruption=True)
                sheet = book.sheet_by_index(0)
                data = []
                for r in range(sheet.nrows):
                    data.append(sheet.row_values(r))
                if data:
                    source_df = pd.DataFrame(data[1:], columns=data[0])
                    print("Success with xlrd ignore_workbook_corruption fallback")
            except Exception as e_xlrd:
                print(f"xlrd ignore_corruption failed: {e_xlrd}")
                # 3. Try openpyxl (in case it's actually .xlsx renamed to .xls)
                try:
                    source_df = pd.read_excel(BytesIO(source_content), engine='openpyxl')
                except Exception as e2:
                    print(f"Openpyxl fallback failed: {e2}")
                    # 4. Try HTML (common for fake XLS exports)
                    try:
                        dfs = pd.read_html(BytesIO(source_content))
                        if dfs:
                            source_df = dfs[0]
                    except Exception as e3:
                        print(f"HTML fallback failed: {e3}")
        
        if source_df is None:
            return jsonify({"error": "엑셀 파일을 읽을 수 없습니다. 파일이 손상되었거나 지원되지 않는 형식입니다. 엑셀에서 '다른 이름으로 저장'을 통해 .xlsx 형식으로 저장 후 다시 시도해 주세요."}), 400

        data_to_fill = source_df[['TRKNO', 'ORDERNO', 'CUSITEMCODE', 'ITEMDETAIL', 'SRL_LOT']]
        
        # Calculate total carton count (Sum of Column D 'BOXCNT')
        total_cartons = 0
        if 'BOXCNT' in source_df.columns:
            try:
                total_cartons = pd.to_numeric(source_df['BOXCNT'], errors='coerce').sum()
            except:
                pass

        # Load template
        wb = load_workbook(template_path)
        ws = wb['Sheet1']
        
        # Set total carton count to J14
        ws['J14'] = total_cartons

        # Number of rows to add (we have row 14 already)
        num_new_rows = len(data_to_fill)
        
        # Capture sample row height
        sample_row_height = ws.row_dimensions[14].height

        if num_new_rows > 1:
            # Insert rows starting from row 15 to push the footer down
            ws.insert_rows(15, amount=num_new_rows - 1)
            
        def copy_style(src_cell, dst_cell):
            if src_cell.has_style:
                dst_cell.font = copy(src_cell.font)
                dst_cell.border = copy(src_cell.border)
                dst_cell.fill = copy(src_cell.fill)
                dst_cell.number_format = copy(src_cell.number_format)
                dst_cell.protection = copy(src_cell.protection)
                dst_cell.alignment = copy(src_cell.alignment)

        start_row = 14
        for idx, (_, row) in enumerate(data_to_fill.iterrows()):
            current_row = start_row + idx
            
            # Set row height to match template's row 14
            if sample_row_height is not None:
                ws.row_dimensions[current_row].height = sample_row_height

            # Fill columns A to E
            target_cols = ['TRKNO', 'ORDERNO', 'CUSITEMCODE', 'ITEMDETAIL', 'SRL_LOT']
            for col_idx, col_name in enumerate(target_cols, 1):
                cell = ws.cell(row=current_row, column=col_idx, value=row[col_name])
                # Copy style from row 14 of the same column
                copy_style(ws.cell(row=14, column=col_idx), cell)
            
            # Copy and adjust formulas for columns G, H, K, L
            formula_cols = [7, 8, 11, 12] # G, H, K, L
            for col_idx in formula_cols:
                source_cell = ws.cell(row=14, column=col_idx)
                if current_row > 14:
                    target_cell = ws.cell(row=current_row, column=col_idx)
                    if source_cell.data_type == 'f':
                        new_formula = source_cell.value.replace('14', str(current_row))
                        target_cell.value = new_formula
                    copy_style(source_cell, target_cell)

        # Re-insert images that were lost from Header/Footer
        logo_path = 'image1.png'
        footer_path = 'image2.png'
        
        try:
            from openpyxl.drawing.spreadsheet_drawing import OneCellAnchor, AnchorMarker, XDRPositiveSize2D
        except ImportError:
            # Fallback for different openpyxl versions
            from openpyxl.drawing.spreadsheet_drawing import OneCellAnchor, AnchorMarker
            from openpyxl.drawing.xdr import XDRPositiveSize2D
        from openpyxl.utils.units import pixels_to_EMU

        # Calculation:
        # Total width of A-E is approx 1000px.
        # Logo 400px wide -> Offset ~300px to center on page.
        # Footer 700px wide -> Center on Column C (center is at ~440px) -> Start at 440 - 350 = 90px.

        if os.path.exists(logo_path):
            img_logo = XLImage(logo_path)
            # User provided size and offset
            img_w, img_h = 310, 160
            h_offset = pixels_to_EMU(60)
            marker = AnchorMarker(col=0, colOff=h_offset, row=3, rowOff=100)
            size = XDRPositiveSize2D(cx=pixels_to_EMU(img_w), cy=pixels_to_EMU(img_h))
            img_logo.anchor = OneCellAnchor(_from=marker, ext=size)
            ws.add_image(img_logo)

            
        # Find the actual last row with content in columns A-E only
        last_data_row = 14 + num_new_rows
        for r in range(ws.max_row, 14, -1):
            if any(ws.cell(row=r, column=c).value for c in range(1, 6)): # A to E is columns 1-5
                last_data_row = r
                break

        footer_start_row = last_data_row + 2
        if os.path.exists(footer_path):
            img_footer = XLImage(footer_path)
            # Enlarged by 10% (700 * 1.1 = 770) and shifted left (offset 20px)
            f_w, f_h = 770, 77
            f_h_offset = pixels_to_EMU(20)
            f_marker = AnchorMarker(col=0, colOff=f_h_offset, row=footer_start_row-1, rowOff=0)
            f_size = XDRPositiveSize2D(cx=pixels_to_EMU(f_w), cy=pixels_to_EMU(f_h))
            img_footer.anchor = OneCellAnchor(_from=f_marker, ext=f_size)
            ws.add_image(img_footer)

        # Set A4 paper and Dynamic Print Area (Strictly A to E)
        ws.page_setup.paperSize = 9 # A4
        ws.page_setup.orientation = 'portrait'
        last_print_row = footer_start_row + 5
        ws.print_area = f'A1:E{last_print_row}'

        output = BytesIO()
        wb.save(output)
        output.seek(0)
        
        return send_file(output, as_attachment=True, download_name="MEDIT_Delivery Note.xlsx")
    except Exception as e:
        print(f"Error: {e}")
        return jsonify({"error": str(e)}), 500

if __name__ == '__main__':
    app.run(debug=True, port=5000)
