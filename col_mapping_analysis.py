"""
Python ve C++ DocxWriter karşılaştırma analizi.

Bu script, Python docx_writer.py ve C++ DocxWriter.cpp arasındaki
fonksiyon ve algoritma farklarını tespit eder.
"""

# PYTHON col_mapping (python-docx merged hücreleri sayıyor):
python_col_mapping = {
    'Linye Adi': 1,       # Merged hücre 1-2
    'Acma Egrisi': 3,
    'Kutup Sayisi': 4,
    'In (A)': 5,          # Merged hücre 5-6
    'Icu': 7,
    'Faz Kesiti': 8,      # Merged hücre 8-9
    'Notr Kesiti': 10,
    'Toprak Kesiti': 11,  # Merged hücre 11-12
    'Ib': 13,             # Merged hücre 13-14
    'Iz': 15,
    'RCD mA': 16,         # Merged hücre 16-17
    'RCD ms': 18,
    'Sonuc': 19,
}

# C++ col_mapping (XML'deki gerçek hücre indeksleri - 14 hücre):
# analyze_table.py çıktısı:
# [0]=No, [1]=Linye, [2]=Açma Eğrisi, [3]=Kutup, [4]=In(A)
# [5]=Icu, [6]=Faz kesiti, [7]=N/PEN, [8]=PE kesiti
# [9]=Ib, [10]=Iz, [11]=IΔ(mA), [12]=TΔ(ms), [13]=Sonuç
cpp_col_mapping = {
    'No': 0,
    'Linye': 1,
    'Acma_Egrisi': 2,
    'Kutup': 3,
    'In_A': 4,
    'Icu': 5,
    'Faz_Kesiti': 6,
    'Notr_Kesiti': 7,
    'PE_Kesiti': 8,
    'Ib': 9,
    'Iz': 10,
    'RCD_mA': 11,
    'RCD_ms': 12,
    'Sonuc': 13,
}

# python-docx, merged hücreleri tekrar sayar:
# Merged [1-2] -> python'da 1, XML'de 1
# Merged [5-6] -> python'da 5, XML'de 4
# vs.

# Dönüşüm tablosu:
# Python index -> XML index
conversion = {
    0: 0,   # No
    1: 1,   # Linye (merged 1-2)
    3: 2,   # Açma eğrisi
    4: 3,   # Kutup
    5: 4,   # In(A) (merged 5-6)
    7: 5,   # Icu
    8: 6,   # Faz kesiti (merged 8-9)
    10: 7,  # Notr kesiti
    11: 8,  # PE kesiti (merged 11-12)
    13: 9,  # Ib (merged 13-14)
    15: 10, # Iz
    16: 11, # RCD mA (merged 16-17)
    18: 12, # RCD ms
    19: 13, # Sonuç
}

print("Python -> XML Dönüşüm:")
for py_idx, xml_idx in conversion.items():
    print(f"  Python {py_idx} -> XML {xml_idx}")
