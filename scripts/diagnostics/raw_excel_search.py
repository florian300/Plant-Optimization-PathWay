import zipfile
import re

file_path = 'PathFinder input.xlsx'
try:
    with zipfile.ZipFile(file_path, 'r') as z:
        # Check all xml files inside the zip
        for name in z.namelist():
            if name.endswith('.xml'):
                with z.open(name) as f:
                    content = f.read().decode('utf-8', errors='ignore')
                    if 'BROWNIEN' in content.upper():
                        print(f"FOUND 'BROWNIEN' in internal file: {name}")
                        # Find the context
                        matches = re.findall(r'[^<]{1,50}BROWNIEN[^<]{1,50}', content, re.IGNORECASE)
                        for m in matches:
                            print(f"  Context: ...{m}...")
                    if '&' in content and 'LINEAR' in content.upper():
                        # Too many false positives for & alone, but let's check for LINEAR&
                        if 'LINEAR&' in content.upper():
                            print(f"FOUND 'LINEAR&' in internal file: {name}")
except Exception as e:
    print(f"Error: {e}")

print("\n--- End of raw search ---")
