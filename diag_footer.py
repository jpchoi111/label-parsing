import zipfile, re

with zipfile.ZipFile('MEDIT_Delivery Note3_footer_test.xlsx', 'r') as z:
    print('=== [Content_Types].xml ===')
    ct = z.read('[Content_Types].xml').decode('utf-8', errors='replace')
    print(ct)
    print()

    print('=== sheet1.xml.rels ===')
    print(z.read('xl/worksheets/_rels/sheet1.xml.rels').decode('utf-8', errors='replace'))
    print()

    sheet1 = z.read('xl/worksheets/sheet1.xml').decode('utf-8', errors='replace')

    # Check r: namespace
    ns_match = re.search(r'xmlns:r="[^"]+"', sheet1)
    print('r: namespace:', ns_match.group(0) if ns_match else 'NOT FOUND')

    # Find legacyDrawingHF
    legacy = re.search(r'legacyDrawingHF[^>]*>', sheet1)
    print('legacyDrawingHF:', legacy.group(0) if legacy else 'NOT FOUND')

    # Find oddFooter
    footer = re.search(r'oddFooter[^<]*</oddFooter>', sheet1)
    print('oddFooter:', footer.group(0) if footer else 'NOT FOUND')
    print()

    print('=== vmlDrawing1.vml.rels ===')
    print(z.read('xl/drawings/_rels/vmlDrawing1.vml.rels').decode('utf-8', errors='replace'))

    # Check Content_Types for vml
    print()
    print('vml in Content_Types:', 'vmlDrawing' in ct)
    print('Files in zip:')
    for name in sorted(z.namelist()):
        print(' ', name)
