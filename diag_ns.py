import zipfile

with zipfile.ZipFile('MEDIT_Delivery Note3_footer_test2.xlsx', 'r') as z:
    sheet1 = z.read('xl/worksheets/sheet1.xml').decode('utf-8', errors='replace')

# Find the root element (first tag) to see namespace declarations
root_end = sheet1.index('>') + 1
print('Root element:')
print(sheet1[:root_end])
print()

# Find exact position around column 32741 in the single-line XML
col = 32741
start = max(0, col - 200)
end = min(len(sheet1), col + 200)
print(f'Context around col {col}:')
print(repr(sheet1[start:end]))
print()

# Find legacyDrawingHF
idx = sheet1.find('legacyDrawingHF')
if idx >= 0:
    print(f'legacyDrawingHF at col {idx+1}:')
    print(repr(sheet1[max(0,idx-50):idx+100]))
else:
    print('legacyDrawingHF NOT FOUND in sheet1.xml')

# Check what 'r' prefix is in the root
import re
ns_decls = re.findall(r'xmlns(?::(\w+))?="([^"]+)"', sheet1[:root_end])
print()
print('Namespace declarations in root:')
for prefix, uri in ns_decls:
    print(f'  {prefix or "(default)"}: {uri}')
