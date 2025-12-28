
import zipfile
import re
import os

path = "d:/YAPAY ZEKALILAR/rapor_sistemi/rapor_sistemi_cpp/sablon/rapor_sablonu.docx"

if not os.path.exists(path):
    print(f"File not found: {path}")
    exit(1)

try:
    with zipfile.ZipFile(path, 'r') as docx:
        xml_content = docx.read('word/document.xml').decode('utf-8')

        # Search for some known placeholders or parts of them
        # Just printing the first 2000 chars might not be enough, lets look for {{ or firma terms

        print("--- XML Snippet around 'firma' ---")
        indices = [m.start() for m in re.finditer('firma', xml_content, re.IGNORECASE)]
        for idx in indices[:5]: # First 5 occurrences
            start = max(0, idx - 100)
            end = min(len(xml_content), idx + 100)
            print(f"...{xml_content[start:end]}...")
            print("-" * 50)

        print("\n--- XML Snippet around '{{' ---")
        indices = [m.start() for m in re.finditer('{{', xml_content)]
        if not indices:
             print("No '{{' found as a contiguous string!")
        for idx in indices[:5]:
            start = max(0, idx - 100)
            end = min(len(xml_content), idx + 100)
            print(f"...{xml_content[start:end]}...")
            print("-" * 50)

except Exception as e:
    print(f"Error: {e}")
