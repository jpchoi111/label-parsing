import os
import re
import pandas as pd
from pypdf import PdfReader

def extract_info_from_pdf(pdf_path):
    try:
        reader = PdfReader(pdf_path)
        full_text = ""
        for page in reader.pages:
            full_text += page.extract_text() + "\n"
        
        # Normalize text for easier matching (handle multiple spaces/newlines)
        normalized_text = " ".join(full_text.split())
        
        # Ref No Extraction
        # Patterns: "Ref No: 12345", "Ref: 12345", "Ref No 12345"
        ref_match = re.search(r'Ref(?:\s*No)?[:\s]*(\S+)', normalized_text, re.IGNORECASE)
        ref_no = ref_match.group(1) if ref_match else "Not Found"
        
        # Tracking Number (Waybill) Extraction
        # Patterns: "WAYBILL 47 3078 9593", "waybill 12345"
        # Often waybill contains spaces between digits.
        waybill_patterns = [
            r'WAYBILL\s*[:\s]*([\d\s]{10,25})',
            r'\b([71]\d[\d\s]{8,15}\d)\b', 
            r'\b([71]\d{9})\b',
            r'\b(18\d{8})\b'
        ]
        
        tracking_no = "Not Found"
        for pattern in waybill_patterns:
            match = re.search(pattern, normalized_text, re.IGNORECASE)
            if match:
                val = match.group(1 if "(" in pattern else 0)
                digits = re.sub(r'\D', '', val)
                if len(digits) >= 10:
                    temp_track = digits[:10]
                    if temp_track not in ref_no:
                        tracking_no = temp_track
                        break
        
        return ref_no, tracking_no
    except Exception as e:
        print(f"Error processing {pdf_path}: {e}")
        return "Error", "Error"

def main():
    pdf_dir = "pdf"
    results = []
    
    if not os.path.exists(pdf_dir):
        print(f"Directory '{pdf_dir}' not found.")
        return

    pdf_files = [f for f in os.listdir(pdf_dir) if f.lower().endswith('.pdf')]
    print(f"Found {len(pdf_files)} PDF files. Starting extraction...")

    for filename in pdf_files:
        pdf_path = os.path.join(pdf_dir, filename)
        ref_no, tracking_no = extract_info_from_pdf(pdf_path)
        results.append({
            "File Name": filename,
            "Ref No": ref_no,
            "Tracking Number": tracking_no
        })
        print(f"Processed: {filename} -> Ref: {ref_no}, Tracking: {tracking_no}")

    df = pd.DataFrame(results)
    output_file = "extraction_results.xlsx"
    df.to_excel(output_file, index=False)
    print(f"\nExtraction complete! Results saved to '{output_file}'.")

if __name__ == "__main__":
    main()
