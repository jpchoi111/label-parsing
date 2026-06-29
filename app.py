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
import threading

app = Flask(__name__)

# Lock for thread-safe OCR access
ocr_lock = threading.Lock()

# Helper to extract images from template if they don't exist
def ensure_images_extracted():
    template_path = 'MEDIT_Delivery Note_template.xlsx'
    if not os.path.exists(template_path):
        return
    
    with zipfile.ZipFile(template_path, 'r') as z:
        media_map = {
            'xl/media/image1.png': 'image1.png',
            'xl/media/image2.png': 'image2.png'
        }
        for zip_p, local_p in media_map.items():
            if not os.path.exists(local_p):
                try:
                    with open(local_p, 'wb') as f:
                        f.write(z.read(zip_p))
                except:
                    pass

ensure_images_extracted()

# Initialize EasyOCR Reader once
reader = easyocr.Reader(['en'])

MAX_WORKERS = os.cpu_count() or 4

def get_pdf_size(pdf_stream):
    try:
        pdf_stream.seek(0)
        r = PdfReader(pdf_stream)
        page = r.pages[0]
        w, h = float(page.mediabox.width), float(page.mediabox.height)
        # Standard A4 is ~595x842 pts. Labels are ~280x560.
        # Increasing threshold to 600 to avoid misclassifying labels as A4.
        if w > 600 or h > 600:
            return "a4"
        return "label"
    except:
        return "error"

def expand_ref_numbers(ref_str):
    if not ref_str or ref_str == "Not Found":
        return ["Not Found"]
    parts = [p.strip() for p in re.split(r'[;,]', ref_str)]
    expanded = []
    base_prefix = ""
    for p in parts:
        if not p: continue
        clean_p = re.sub(r'^RMA-?', '', p, flags=re.IGNORECASE).strip()
        if clean_p.startswith('400'):
            expanded.append(clean_p)
            base_prefix = clean_p[:6]
        elif clean_p.startswith('-') and base_prefix:
            expanded.append(base_prefix + clean_p[1:])
        else:
            digits = re.sub(r'\D', '', clean_p)
            if digits:
                expanded.append(digits)
    return expanded

def is_valid_label_page(page):
    text = page.extract_text()
    if not text or not text.strip():
        return True
    text_lower = text.lower()
    exclude_keywords = ["waybill doc", "receipt", "archive doc", "copy for your records"]
    for kw in exclude_keywords:
        if kw in text_lower:
            return False
    return True

def extract_single_pdf(file_content, filename, manual_ref=None):
    stream = BytesIO(file_content)
    try:
        r = PdfReader(stream)
        first_page = r.pages[0]
        full_text = first_page.extract_text() or ""
        
        w, h = float(first_page.mediabox.width), float(first_page.mediabox.height)
        # Threshold changed to 600 to avoid misclassifying larger labels as A4
        size = "a4" if (w > 600 or h > 600) else "label"
        is_sideways_a4 = (size == "a4") and (w > h)
        
        if not full_text.strip():
            full_text = first_page.extract_text(orientations=(90,)) or ""


        def find_data(text):
            normalized = " ".join(text.split())

            # --- Ref No 파싱 ---
            ref_res = "Not Found"
            if manual_ref:
                ref_res = manual_ref
            else:
                # 1순위: 원본 텍스트에서 "Ref No:" 줄 파싱 (suffix -XXXX 포함 전체)
                ref_line = re.search(r'Ref\s*No[:\s]+([\d][0-9;\-,\s]{5,80})', text, re.IGNORECASE)
                if ref_line:
                    candidate = ref_line.group(1).strip()
                    # 숫자/구분자 이외 문자가 나오면 거기서 자름
                    candidate = re.split(r'[^\d;\-,\s]', candidate)[0].strip()
                    if len(re.sub(r'\D', '', candidate)) >= 7:
                        ref_res = candidate

                # 2순위: normalized에서 400xxxxxxx 형태 SAP 번호 직접 수집
                if ref_res == "Not Found" or len(re.sub(r'\D', '', ref_res)) < 7:
                    all_sap = re.findall(r'\b(400\d{7})\b', normalized)
                    if all_sap:
                        ref_res = "; ".join(sorted(list(set(all_sap))))

                # 3순위: 기타 Order/Reference 패턴 (fallback)
                if ref_res == "Not Found":
                    ref_match = re.search(
                        r'(?:Order|P/O|Shipment|Reference)[:\s]*([40RMA\d\s\-\;\,]{7,})',
                        normalized, re.IGNORECASE
                    )
                    if ref_match:
                        ref_res = re.split(r'\s{2,}', ref_match.group(1).strip())[0]

            # --- Tracking 파싱 (기존 유지) ---
            track_res = "Not Found"
            waybill_patterns = [
                r'WAYBILL\s*[:\s]*([\d\s]{10,25})',
                r'\b([71]\d[\d\s]{8,15}\d)\b',
                r'\b([71]\d{9})\b',
                r'\b(18\d{8})\b'
            ]
            for pattern in waybill_patterns:
                match = re.search(pattern, normalized, re.IGNORECASE)
                if match:
                    val = match.group(1 if "(" in pattern else 0)
                    digits = re.sub(r'\D', '', val)
                    if len(digits) >= 10:
                        temp_track = digits[:10]
                        if temp_track not in ref_res:
                            track_res = temp_track
                            break
            return ref_res, track_res

        ref_raw, tracking_no = find_data(full_text)

        # Fallback to OCR if tracking is missing or ref is missing (especially for A4 where extract_text might be partial)
        if tracking_no == "Not Found" or ref_raw == "Not Found":
            stream.seek(0)
            doc = pdfium.PdfDocument(stream)
            page = doc[0]
            rots = [90] if is_sideways_a4 else [0, 90]
            ocr_text = ""
            for rot in rots:
                bitmap = page.render(scale=2.0, rotation=rot)
                with ocr_lock:
                    ocr_result = reader.readtext(np.array(bitmap.to_pil()), detail=0)
                page_text = " ".join(ocr_result)
                ocr_text += page_text + "\n"
                
                temp_ref, temp_track = find_data(page_text)
                if temp_ref != "Not Found": ref_raw = temp_ref
                if temp_track != "Not Found": 
                    tracking_no = temp_track
                    break # Found it!
            doc.close()
        
        return ref_raw, tracking_no, size
    except Exception as e:
        print(f"Error processing {filename}: {e}")
        return manual_ref if manual_ref else "Error", "Error", "Unknown"

@app.route('/')
def index():
    return render_template('index.html')

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
        
        try:
            start_page = int(request.form.get('start_page', 1)) - 1
            if start_page < 0: start_page = 0
        except:
            start_page = 0

        orders = []
        seen_packing_nos = set()
        for i in range(start_page, len(reader.pages)):
            text = reader.pages[i].extract_text()
            if "Packing No." not in text or "OrderNo." not in text:
                continue
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
                        if packing_no not in seen_packing_nos:
                            seen_packing_nos.add(packing_no)
                            company_name = "MEDIT EUROPE GMBH" if user_id == "MEDITFRA" else ("MEDIT EUROPE" if user_id == "MEDITRMA" else "")
                            track_no = tracking_map.get(order_no, "")
                            orders.append({"A": "", "B": packing_no, "C": company_name, "D": "", "E": track_no})
        if not orders:
            return jsonify({"error": "오더를 찾을 수 없습니다."}), 404
        df = pd.DataFrame(orders)
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, header=False, sheet_name='Sheet1')
            worksheet = writer.sheets['Sheet1']
            worksheet['A1'] = datetime.now().strftime("%Y-%m-%d")
        output.seek(0)
        return send_file(output, as_attachment=True, download_name="picking_list_results.xlsx")
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route('/download_excel', methods=['POST'])
def download_excel():
    data = request.json
    if not data:
        return jsonify({"error": "데이터가 없습니다."}), 400
    df_data = [{"": "", " ": " ", "Tracking": item.get("Tracking Number", ""), "SAP Order(s)": item.get("Ref No", "")} for item in data]
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
                if is_valid_label_page(page):
                    writer.add_page(page)
                    found_any = True
    if not found_any:
        return jsonify({"error": f"인쇄 가능한 {target_size} 규격의 페이지가 없습니다."}), 404
    output = BytesIO()
    writer.write(output)
    output.seek(0)
    return send_file(output, mimetype='application/pdf')

SPECIAL_CODES = ['3A0113417C0', '3A0112485C0', '3A0113147C0', '3A0113418C0', '3A0113419C0', '3A0112340C0', '6A0111588C0', '5M0111852W0', '5M0112350E0', '3A0113412C0', '3A0113411C0', '3A0111853C0']

@app.route('/generate_delivery_note', methods=['POST'])
def generate_delivery_note():
    source_file = request.files.get('source_file')
    pallet_count = request.form.get('pallet_count', 1)
    box_count = request.form.get('box_count', 0)
    if not source_file:
        return jsonify({"error": "원본 엑셀 파일(20260612.xls)을 업로드해주세요."}), 400
    try:
        pallet_count = int(pallet_count)
        box_count = int(box_count)
    except:
        pallet_count, box_count = 1, 0
    template_path = 'MEDIT_Delivery Note_template.xlsx'
    if not os.path.exists(template_path):
        return jsonify({"error": "템플릿 파일이 서버에 존재하지 않습니다."}), 500
    try:
        from openpyxl import load_workbook
        from openpyxl.drawing.image import Image as XLImage
        from copy import copy
        source_content = source_file.read()
        source_df = None
        try:
            source_df = pd.read_excel(BytesIO(source_content))
        except:
            try:
                import xlrd
                book = xlrd.open_workbook(file_contents=source_content, ignore_workbook_corruption=True)
                sheet = book.sheet_by_index(0)
                data = [sheet.row_values(r) for r in range(sheet.nrows)]
                source_df = pd.DataFrame(data[1:], columns=data[0])
            except:
                try: source_df = pd.read_excel(BytesIO(source_content), engine='openpyxl')
                except:
                    dfs = pd.read_html(BytesIO(source_content))
                    if dfs: source_df = dfs[0]
        if source_df is None:
            return jsonify({"error": "엑셀 파일을 읽을 수 없습니다."}), 400
        data_to_fill = source_df[['TRKNO', 'ORDERNO', 'CUSITEMCODE', 'ITEMDETAIL', 'SRL_LOT']]
        special_match_count = sum(1 for code in source_df['CUSITEMCODE'].astype(str).str.strip() if code in SPECIAL_CODES) if 'CUSITEMCODE' in source_df.columns else 0
        total_cartons = int(pd.to_numeric(source_df['BOXCNT'], errors='coerce').sum()) if 'BOXCNT' in source_df.columns else 0
        wb = load_workbook(template_path)
        ws = wb['Sheet1']
        ws['A11'] = f"Carton : {total_cartons} stk"
        ws['B11'] = f"Pallet : {pallet_count} stk"
        num_new_rows = len(data_to_fill)
        sample_row_height = ws.row_dimensions[14].height
        if num_new_rows > 1:
            ws.insert_rows(15, amount=num_new_rows - 1)
        def copy_style(src, dst):
            if src.has_style:
                dst.font, dst.border, dst.fill, dst.number_format, dst.protection, dst.alignment = copy(src.font), copy(src.border), copy(src.fill), copy(src.number_format), copy(src.protection), copy(src.alignment)
        for idx, (_, row) in enumerate(data_to_fill.iterrows()):
            curr = 14 + idx
            if sample_row_height: ws.row_dimensions[curr].height = sample_row_height
            for c_idx, col in enumerate(['TRKNO', 'ORDERNO', 'CUSITEMCODE', 'ITEMDETAIL', 'SRL_LOT'], 1):
                cell = ws.cell(row=curr, column=c_idx, value=row[col])
                copy_style(ws.cell(row=14, column=c_idx), cell)
            for c_idx in [7, 8, 11, 12]:
                src_c = ws.cell(row=14, column=c_idx)
                if curr > 14:
                    dst_c = ws.cell(row=curr, column=c_idx)
                    if src_c.data_type == 'f': dst_c.value = src_c.value.replace('14', str(curr))
                    copy_style(src_c, dst_c)
        # Determine last valid data row and footer row
        last_d = 14 + num_new_rows
        for r in range(ws.max_row, 14, -1):
            if any(ws.cell(row=r, column=c).value for c in range(1, 6)): last_d = r; break
        foot_r = last_d + 2

        # Find the exact signature rows after insertion by scanning for "DATEN :"
        sig_start_row = None
        sig_end_row = None
        for r in range(14, ws.max_row + 1):
            val = ws.cell(row=r, column=1).value
            if val == "DATEN :":
                sig_start_row = r
            if val == "Unterschrift :":
                sig_end_row = r
                break

        if sig_start_row and sig_end_row:
            # Calculate A4 printable height dynamically based on margins (A4 height is 841.68 points)
            top_margin = ws.page_margins.top if ws.page_margins.top is not None else 0.75
            bottom_margin = ws.page_margins.bottom if ws.page_margins.bottom is not None else 1.0
            a4_page_height = 841.68 - (top_margin * 72.0) - (bottom_margin * 72.0) - 5.0
            default_h = ws.sheet_format.defaultRowHeight or 12.75

            # Map each row (up to foot_r) to its calculated page number
            page_map = {}
            current_page = 1
            current_height = 0.0
            for r in range(1, foot_r + 1):
                h = ws.row_dimensions[r].height or default_h
                if current_height + h > a4_page_height:
                    current_page += 1
                    current_height = h
                else:
                    current_height += h
                page_map[r] = current_page

            # If signature block spans pages, insert a manual break before it
            manual_break_row = None
            if page_map.get(sig_start_row, 1) != page_map.get(foot_r, 1):
                manual_break_row = sig_start_row - 1

            # Re-simulate with manual break
            actual_page_map = {}
            current_page = 1
            current_height = 0.0
            for r in range(1, foot_r + 1):
                h = ws.row_dimensions[r].height or default_h
                if manual_break_row and r == manual_break_row + 1:
                    current_page += 1
                    current_height = h
                elif current_height + h > a4_page_height:
                    current_page += 1
                    current_height = h
                else:
                    current_height += h
                actual_page_map[r] = current_page

            # Insert manual page breaks at page transitions
            break_points = [r for r in range(2, foot_r + 1)
                            if actual_page_map.get(r) != actual_page_map.get(r - 1)]
            from openpyxl.worksheet.pagebreak import Break
            for bp in reversed(break_points):
                ws.row_breaks.append(Break(id=bp - 1))

        # Set footer: &G inserts the VML-linked image in the center of the page footer
        ws.oddFooter.center.text = "&G"

        # Logo image (header area, sheet drawing)
        logo_p = 'image1.png'
        try: from openpyxl.drawing.spreadsheet_drawing import OneCellAnchor, AnchorMarker, XDRPositiveSize2D
        except: from openpyxl.drawing.spreadsheet_drawing import OneCellAnchor, AnchorMarker; from openpyxl.drawing.xdr import XDRPositiveSize2D
        from openpyxl.utils.units import pixels_to_EMU
        if os.path.exists(logo_p):
            img = XLImage(logo_p)
            img.anchor = OneCellAnchor(_from=AnchorMarker(col=0, colOff=pixels_to_EMU(60), row=3, rowOff=100), ext=XDRPositiveSize2D(cx=pixels_to_EMU(310), cy=pixels_to_EMU(160)))
            ws.add_image(img)

        ws.page_setup.paperSize, ws.page_setup.orientation, ws.print_area = 9, 'portrait', f'A1:E{foot_r + 5}'
        # Save workbook to buffer first
        output = BytesIO()
        wb.save(output)

        # Inject VML footer image into the saved xlsx (zip) directly from the template.
        # openpyxl does not support header/footer images natively, so we patch the zip.
        foot_p = 'image2.png'
        if os.path.exists(foot_p) and os.path.exists(template_path):
            # Read required VML content from template
            with zipfile.ZipFile(template_path, 'r') as tmpl_z:
                vml_content = tmpl_z.read('xl/drawings/vmlDrawing1.vml')
                vml_rels_content = tmpl_z.read('xl/drawings/_rels/vmlDrawing1.vml.rels')
                footer_img_content = tmpl_z.read('xl/media/image2.png')

            # Re-write the xlsx zip with VML entries added / patched
            VML_CT = '<Default Extension="vml" ContentType="application/vnd.openxmlformats-officedocument.vmlDrawing"/>'
            VML_REL_ENTRY = (
                '<Relationship Id="rIdVML" '
                'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing" '
                'Target="../drawings/vmlDrawing1.vml"/>'
            )
            output.seek(0)
            original_bytes = output.read()
            patched = BytesIO()
            with zipfile.ZipFile(BytesIO(original_bytes), 'r') as src_z, \
                 zipfile.ZipFile(patched, 'w', compression=zipfile.ZIP_DEFLATED) as dst_z:
                existing = {i.filename for i in src_z.infolist()}
                for item in src_z.infolist():
                    data = src_z.read(item.filename)
                    if item.filename == '[Content_Types].xml':
                        # Add VML content type if missing
                        txt = data.decode('utf-8', errors='replace')
                        if 'vmlDrawing' not in txt:
                            txt = txt.replace('</Types>', VML_CT + '</Types>')
                        data = txt.encode('utf-8')
                    elif item.filename == 'xl/worksheets/_rels/sheet1.xml.rels':
                        # Add legacyDrawingHF (VML) relationship
                        txt = data.decode('utf-8', errors='replace')
                        if 'vmlDrawing' not in txt:
                            txt = txt.replace('</Relationships>', VML_REL_ENTRY + '</Relationships>')
                        data = txt.encode('utf-8')
                    elif item.filename == 'xl/worksheets/sheet1.xml':
                        # Add <legacyDrawingHF> before </worksheet>.
                        # openpyxl's root <worksheet> element does NOT declare xmlns:r globally;
                        # r: only appears inline on the <drawing> tag. To avoid "undeclared prefix"
                        # errors, we include xmlns:r explicitly on legacyDrawingHF itself.
                        txt = data.decode('utf-8', errors='replace')
                        if 'legacyDrawingHF' not in txt:
                            R_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
                            legacy_tag = f'<legacyDrawingHF xmlns:r="{R_NS}" r:id="rIdVML"/>'
                            txt = txt.replace('</worksheet>', legacy_tag + '</worksheet>')
                        data = txt.encode('utf-8')
                    elif item.filename == 'xl/media/image2.png':
                        # Replace with template's footer image
                        data = footer_img_content
                    dst_z.writestr(item, data)
                # Add VML files if not already present
                if 'xl/drawings/vmlDrawing1.vml' not in existing:
                    dst_z.writestr('xl/drawings/vmlDrawing1.vml', vml_content)
                if 'xl/drawings/_rels/vmlDrawing1.vml.rels' not in existing:
                    dst_z.writestr('xl/drawings/_rels/vmlDrawing1.vml.rels', vml_rels_content)
                if 'xl/media/image2.png' not in existing:
                    dst_z.writestr('xl/media/image2.png', footer_img_content)
            patched.seek(0)
            output = patched
        else:
            output.seek(0)

        res = send_file(output, as_attachment=True, download_name="MEDIT_Delivery Note.xlsx")
        res.headers.update({'X-Special-Match-Count': str(special_match_count), 'X-Total-Cartons': str(total_cartons), 'X-Pallet-Count': str(pallet_count), 'X-Box-Count': str(box_count)})
        return res
    except Exception as e:
        return jsonify({"error": str(e)}), 500

if __name__ == '__main__':
    app.run(debug=True, port=5000)
