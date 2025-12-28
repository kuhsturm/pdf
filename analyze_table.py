import zipfile
from lxml import etree

path = 'sablon/rapor_sablonu.docx'
NS_W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
nsmap = {'w': NS_W}

with zipfile.ZipFile(path, 'r') as z:
    xml = z.read('word/document.xml')
    root = etree.fromstring(xml)

    tables = root.findall('.//w:tbl', nsmap)
    print(f'Toplam {len(tables)} tablo bulundu\n')

    # Tablo 3'ü detaylı incele (6.2 ve 6.3)
    if len(tables) > 3:
        tbl = tables[3]
        rows = tbl.findall('./w:tr', nsmap)
        print(f'=== Tablo 3 (6.2/6.3): {len(rows)} satir ===')

        for row_idx, row in enumerate(rows):
            cells = row.findall('./w:tc', nsmap)
            print(f'\nSatir {row_idx} ({len(cells)} hucre):')
            for i, cell in enumerate(cells):
                cell_texts = cell.findall('.//w:t', nsmap)
                text = ''.join([t.text or '' for t in cell_texts]).strip()
                text = text if text else "(bos)"
                print(f'  [{i}]: {text[:50]}')
