#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
sozlesme_parser.py

Sözleşme PDF dosyasını okuyup bilgileri çıkarır.
C++ PdfParser tarafından çağrılır.

Kullanım:
    python sozlesme_parser.py <pdf_dosyasi>

Çıktı:
    key: value formatında satırlar
"""

import sys
import re

try:
    import pdfplumber
except ImportError:
    print("HATA: pdfplumber modülü yüklü değil!", file=sys.stderr)
    print("Çözüm: pip install pdfplumber", file=sys.stderr)
    sys.exit(1)


def extract_text_from_pdf(pdf_path: str) -> str:
    """PDF dosyasından tüm metni çıkarır."""
    try:
        with pdfplumber.open(pdf_path) as pdf:
            text = ""
            for page in pdf.pages:
                page_text = page.extract_text()
                if page_text:
                    text += page_text + "\n"
            return text
    except Exception as e:
        print(f"PDF okuma hatası: {e}", file=sys.stderr)
        return ""


def parse_sozlesme_text(text: str) -> dict:
    """Sözleşme metninden bilgileri çıkarır."""
    data = {
        "sozlesme_id": "",
        "sozlesme_baslangic": "",
        "sozlesme_bitis": "",
        "firma_unvan": "",
        "firma_adres": "",
        "firma_il": "",
        "firma_sgk_no": "",
        "kontrol_eden_adsoyad": "",
        "kontrol_eden_tc": "",
        "pk_no": "",
        "tesisat_sayisi": "1",
    }

    if not text:
        return data

    # Sözleşme ID
    sozlesme_match = re.search(r'(\d{6,}-\d{6,}-\d{6,})', text)
    if sozlesme_match:
        data["sozlesme_id"] = sozlesme_match.group(1)

    # Tarih formatları
    tarih_pattern = r'(\d{2}[./]\d{2}[./]\d{4})'
    tarihler = re.findall(tarih_pattern, text)
    if len(tarihler) >= 2:
        data["sozlesme_baslangic"] = tarihler[0].replace('/', '.')
        data["sozlesme_bitis"] = tarihler[1].replace('/', '.')
    elif len(tarihler) == 1:
        data["sozlesme_baslangic"] = tarihler[0].replace('/', '.')

    # Firma unvanı
    firma_patterns = [
        r'(?:İş\s*Yeri|Firma)\s*Unvan[ıi]?\s*[:\-]?\s*(.+)',
        r'Hizmet\s*Alan\s*[:\-]?\s*(.+)',
        r'Firma\s*Adı\s*[:\-]?\s*(.+)',
        r'([A-ZÇĞİÖŞÜ][A-ZÇĞİÖŞÜa-zçğıöşü\s\.]+(?:A\.Ş\.|AŞ|LTD\. ŞTİ\.|LTD\.ŞTİ\.|LİMİTED ŞİRKETİ))',
        r'Firma\s*[:\-]?\s*(.+)',
    ]
    for pattern in firma_patterns:
        match = re.search(pattern, text, re.IGNORECASE)
        if match:
            val = match.group(1).strip()
            if val:
                data["firma_unvan"] = val
                break

    # Adres
    adres_match = re.search(r'Adres[i]?\s*[:\-]?\s*(.+?)(?:\n|İl\s*:|Şehir)', text, re.IGNORECASE | re.DOTALL)
    if adres_match:
        data["firma_adres"] = adres_match.group(1).strip().replace('\n', ' ')

    # İl
    il_match = re.search(r'(?:^|[\s])(?:İl|Şehir)\s*[:\-]\s*([A-ZÇĞİÖŞÜa-zçğıöşü]+)', text, re.IGNORECASE)
    if il_match:
        data["firma_il"] = il_match.group(1).strip()

    # SGK No
    sgk_patterns = [
        r'SGK\s*(?:Sicil)?\s*(?:No\.?|Numarası)?\s*[:\-]?\s*(\d[\d\.\-\s]+)',
        r'Sicil\s*No\.?\s*[:\-]?\s*(\d[\d\.\-\s]+)',
    ]
    for pattern in sgk_patterns:
        match = re.search(pattern, text, re.IGNORECASE)
        if match:
            data["firma_sgk_no"] = re.sub(r'[\.\-\s]', '', match.group(1))
            break

    # Kontrol eden kişi
    found_kontrol = False

    # 1. Direct match with Colon (Strong signal)
    kontrol_patterns_strong = [
        r'(?:Kontrol|Muayene)\s*(?:Eden|Yapan|Personeli|Elemanı|Uzmanı)\s*[:\-]\s*([A-ZÇĞİÖŞÜa-zçğıöşü\s]+)',
        r'Kontrol\s*Personeli\s*[:\-]\s*([A-ZÇĞİÖŞÜa-zçğıöşü\s]+)',
    ]
    for pattern in kontrol_patterns_strong:
        match = re.search(pattern, text, re.IGNORECASE)
        if match:
            val = match.group(1).strip().split('\n')[0].strip()
            # Filter out generic headers
            if len(val) > 2 and "Bilgi" not in val and "Personel" not in val:
                data["kontrol_eden_adsoyad"] = val
                found_kontrol = True
                break

    # 2. Section match (if not found yet)
    if not found_kontrol:
        section_pat = r'(?:Kontrol|Muayene)\s*(?:Eden|Yapan|Personeli|Uzmanı)(?:.+?)(?:Ad[ıi]\s*Soyad[ıi])\s*[:\-]?\s*([A-ZÇĞİÖŞÜa-zçğıöşü\s]+)'
        match = re.search(section_pat, text, re.IGNORECASE | re.DOTALL)
        if match:
            val = match.group(1).strip().split('\n')[0].strip()
            data["kontrol_eden_adsoyad"] = val

    # TC No
    tc_match = re.search(r'(?:T\.?C\.?\s*(?:No\.?|Kimlik No\.?)?\s*[:\-]?\s*)?(\d{11})', text)
    if tc_match:
        data["kontrol_eden_tc"] = tc_match.group(1)

    # PK No / Yetki No
    pk_patterns = [
        r'(?:PK|P\.K\.)\s*(?:No\.?|Numaras[ıi])?\s*[:\-]?\s*([A-Z0-9\-]+)',
        r'(?:Yetki|Kayıt)\s*(?:Belge)?\s*(?:No\.?|Numaras[ıi])?\s*[:\-]?\s*([A-Z0-9\-]+)',
        r'Ekipnet\s*(?:No\.?|Numaras[ıi])?\s*[:\-]?\s*([A-Z0-9\-]+)',
    ]
    for pattern in pk_patterns:
        pk_match = re.search(pattern, text, re.IGNORECASE)
        if pk_match:
            data["pk_no"] = pk_match.group(1)
            break

    # Tesisat sayısı
    tesisat_match = re.search(r'(?:Tesisat\s*Sayısı|Adet)\s*[:\-]?\s*(\d+)', text, re.IGNORECASE)
    if tesisat_match:
        data["tesisat_sayisi"] = tesisat_match.group(1)

    return data


def main():
    if len(sys.argv) < 2:
        print("Kullanım: python sozlesme_parser.py <pdf_dosyasi>", file=sys.stderr)
        sys.exit(1)

    pdf_path = sys.argv[1]

    # PDF'den metin çıkar
    text = extract_text_from_pdf(pdf_path)

    if not text:
        print("PDF okunamadı veya boş!", file=sys.stderr)
        sys.exit(1)

    # Metinden bilgileri parse et
    data = parse_sozlesme_text(text)

    # Çıktı (key: value formatı)
    for key, value in data.items():
        print(f"{key}: {value}")


if __name__ == "__main__":
    main()
