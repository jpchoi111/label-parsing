import zipfile, re
from io import BytesIO
from openpyxl import load_workbook

template_path = 'MEDIT_Delivery Note_template.xlsx'

wb = load_workbook('MEDIT_Delivery Note3.xlsx')
ws = wb.active

# Remove floating footer images, keep only logo at row=3
ws._images = [img for img in ws._images if img.anchor._from.row == 3]

# Set real Excel footer with &G (image placeholder)
ws.oddFooter.center.text = "&G"
print('Footer set to:', ws.oddFooter.center.text)

# Save to buffer
output = BytesIO()
wb.save(output)

# Read VML from template
with zipfile.ZipFile(template_path, 'r') as tmpl_z:
    vml_content = tmpl_z.read('xl/drawings/vmlDrawing1.vml')
    vml_rels_content = tmpl_z.read('xl/drawings/_rels/vmlDrawing1.vml.rels')
    footer_img_content = tmpl_z.read('xl/media/image2.png')

VML_REL = (
    '<Relationship Id="rIdVML" '
    'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing" '
    'Target="../drawings/vmlDrawing1.vml"/>'
)
LEGACY_HF = '<legacyDrawingHF r:id="rIdVML"/>'
VML_CONTENT_TYPE = '<Default Extension="vml" ContentType="application/vnd.openxmlformats-officedocument.vmlDrawing"/>'

output.seek(0)
original_bytes = output.read()
patched = BytesIO()

with zipfile.ZipFile(BytesIO(original_bytes), 'r') as src_z, \
     zipfile.ZipFile(patched, 'w', compression=zipfile.ZIP_DEFLATED) as dst_z:

    existing = {i.filename for i in src_z.infolist()}

    for item in src_z.infolist():
        data = src_z.read(item.filename)

        if item.filename == '[Content_Types].xml':
            txt = data.decode('utf-8', errors='replace')
            # Add vml content type if missing
            if 'vmlDrawing' not in txt:
                txt = txt.replace('<Types ', VML_CONTENT_TYPE + '<Types ', 1)
                # Better: insert before </Types>
                txt = txt.replace('</Types>', VML_CONTENT_TYPE + '</Types>')
            data = txt.encode('utf-8')
            print('Content_Types patched, vml present:', 'vmlDrawing' in txt)

        elif item.filename == 'xl/worksheets/_rels/sheet1.xml.rels':
            txt = data.decode('utf-8', errors='replace')
            if 'vmlDrawing' not in txt:
                txt = txt.replace('</Relationships>', VML_REL + '</Relationships>')
                data = txt.encode('utf-8')
            print('sheet1.xml.rels vml rel added')

        elif item.filename == 'xl/worksheets/sheet1.xml':
            txt = data.decode('utf-8', errors='replace')
            if 'legacyDrawingHF' not in txt:
                R_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
                legacy_tag = f'<legacyDrawingHF xmlns:r="{R_NS}" r:id="rIdVML"/>'
                txt = txt.replace('</worksheet>', legacy_tag + '</worksheet>')
            print('legacyDrawingHF present:', 'legacyDrawingHF' in txt)
            data = txt.encode('utf-8')


        elif item.filename == 'xl/media/image2.png':
            data = footer_img_content

        dst_z.writestr(item, data)

    # Add VML files if missing
    if 'xl/drawings/vmlDrawing1.vml' not in existing:
        dst_z.writestr('xl/drawings/vmlDrawing1.vml', vml_content)
        print('vmlDrawing1.vml added')
    if 'xl/drawings/_rels/vmlDrawing1.vml.rels' not in existing:
        dst_z.writestr('xl/drawings/_rels/vmlDrawing1.vml.rels', vml_rels_content)
        print('vmlDrawing1.vml.rels added')
    if 'xl/media/image2.png' not in existing:
        dst_z.writestr('xl/media/image2.png', footer_img_content)
        print('image2.png added')

patched.seek(0)
out_path = 'MEDIT_Delivery Note3_footer_test2.xlsx'
with open(out_path, 'wb') as f:
    f.write(patched.read())
print('Saved:', out_path)

# Verify
with zipfile.ZipFile(out_path, 'r') as z:
    ct = z.read('[Content_Types].xml').decode()
    print()
    print('Content_Types vml entry:', 'vmlDrawing' in ct)
    print('VML file present:', 'xl/drawings/vmlDrawing1.vml' in z.namelist())
    print('VML rels present:', 'xl/drawings/_rels/vmlDrawing1.vml.rels' in z.namelist())
    print('image2.png present:', 'xl/media/image2.png' in z.namelist())
    rels = z.read('xl/worksheets/_rels/sheet1.xml.rels').decode()
    print('VML rel in sheet1.rels:', 'vmlDrawing' in rels)
