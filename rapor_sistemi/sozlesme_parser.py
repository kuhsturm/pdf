"""
İSG-KATİP Hizmet Sözleşmesi PDF Parser
PDF dosyasından firma bilgileri, SGK sicil, sözleşme ID ve kontrol eden kişi bilgilerini çıkarır.
"""

from pypdf import PdfReader
import re
from typing import Dict, Any, Optional


def parse_sozlesme_pdf(pdf_path: str) -> Dict[str, Any]:
    """
    İSG-KATİP hizmet sözleşmesi PDF'inden bilgileri çıkarır.

    Returns:
        Dict containing:
        - sozlesme_id: Sözleşme ID
        - sozlesme_baslangic: Sözleşme başlangıç tarihi
        - sozlesme_bitis: Sözleşme bitiş tarihi
        - firma_unvan: Hizmet alan işyeri unvanı
        - firma_adres: Hizmet alan işyeri adresi
        - firma_il: İl
        - firma_sgk_no: SGK/DETSİS numarası
        - kontrol_eden_adsoyad: Periyodik kontrol yapan kişi
        - kontrol_eden_tc: TC Kimlik No
        - pk_no: Periyodik kontrol numarası
        - tesisat_sayisi: Tesisat sayısı
    """
    result = {
        'sozlesme_id': '',
        'sozlesme_baslangic': '',
        'sozlesme_bitis': '',
        'firma_unvan': '',
        'firma_adres': '',
        'firma_il': '',
        'firma_sgk_no': '',
        'kontrol_eden_adsoyad': '',
        'kontrol_eden_tc': '',
        'pk_no': '',
        'tesisat_sayisi': ''
    }

    try:
        reader = PdfReader(pdf_path)

        if len(reader.pages) == 0:
            return result

        # Tüm sayfaların metnini birleştir
        text = ""
        for page in reader.pages:
            text += page.extract_text() + "\n"

        # Satır satır işle
        lines = text.split('\n')

        # Debug için
        full_text = text

        # === SÖZLEŞME ID ===
        # Önce "Sözleşme ID" satırından sonraki 8 haneli sayıyı bul
        for i, line in enumerate(lines):
            if 'Sözleşme ID' in line:
                # Sonraki birkaç satırda 8 haneli sayı ara
                for j in range(1, 6):
                    if i + j < len(lines):
                        match = re.search(r'\b(\d{8})\b', lines[i + j])
                        if match:
                            result['sozlesme_id'] = match.group(1)
                            break
                break

        # Alternatif: Regex ile ara
        if not result['sozlesme_id']:
            match = re.search(r'Sözleşme\s*ID.*?(\d{8})', text, re.IGNORECASE | re.DOTALL)
            if match:
                result['sozlesme_id'] = match.group(1)

        # === SÖZLEŞME TARİHLERİ ===
        # Başlangıç tarihi - "Sözleşme Başlangıç Tarihi" veya ":Sözleşme Başlangıç Tarihi"
        for i, line in enumerate(lines):
            if 'Başlangıç Tarihi' in line:
                # Aynı satırda tarih var mı?
                match = re.search(r'(\d{2}\.\d{2}\.\d{4})', line)
                if match:
                    result['sozlesme_baslangic'] = match.group(1)
                else:
                    # Önceki satırlarda ara
                    for j in range(1, 4):
                        if i - j >= 0:
                            match = re.search(r'(\d{2}\.\d{2}\.\d{4})', lines[i - j])
                            if match:
                                result['sozlesme_baslangic'] = match.group(1)
                                break
                break

        # Bitiş tarihi
        for i, line in enumerate(lines):
            if 'Bitiş Tarihi' in line:
                match = re.search(r'(\d{2}\.\d{2}\.\d{4})', line)
                if match:
                    result['sozlesme_bitis'] = match.group(1)
                else:
                    for j in range(1, 4):
                        if i - j >= 0:
                            match = re.search(r'(\d{2}\.\d{2}\.\d{4})', lines[i - j])
                            if match:
                                result['sozlesme_bitis'] = match.group(1)
                                break
                break

        # === GÖREVLENDİRİLEN KİŞİ BİLGİLERİ ===
        # Ad Soyad ve PK No
        in_gorevlendirilen = False
        for i, line in enumerate(lines):
            if 'GÖREVLENDİRİLEN KİŞİ BİLGİLERİ' in line:
                in_gorevlendirilen = True
                continue

            if in_gorevlendirilen:
                # "HİZMET ALAN" görünce bölüm bitti
                if 'HİZMET ALAN' in line:
                    break

                # Ad Soyad - büyük harfli isim ara (en az 2 kelime, rakam yok)
                if not result['kontrol_eden_adsoyad']:
                    # "Ad Soyad" satırından sonraki satırlarda isim ara
                    if 'Ad Soyad' in line or 'Ad-Soyad' in line:
                        # Sonraki satırlarda isim ara
                        for j in range(1, 5):
                            if i + j < len(lines):
                                candidate = lines[i + j].strip()
                                # İsim kriterleri: büyük harfli, en az 2 kelime, rakam yok, ":" ile başlamıyor
                                if (candidate and
                                    len(candidate) > 3 and
                                    ' ' in candidate and
                                    not any(c.isdigit() for c in candidate) and
                                    not candidate.startswith(':') and
                                    not candidate.startswith('Periyodik') and
                                    not candidate.startswith('TC')):
                                    result['kontrol_eden_adsoyad'] = candidate
                                    break
                    # Alternatif: Doğrudan büyük harfli isim satırı
                    elif (line.strip() and
                          ' ' in line.strip() and
                          not any(c.isdigit() for c in line.strip()) and
                          line.strip().isupper() and
                          len(line.strip()) > 5 and
                          len(line.strip()) < 50):
                        result['kontrol_eden_adsoyad'] = line.strip()

                # PK No - K ile başlayan numara
                if not result['pk_no']:
                    match = re.search(r':?\s*(K\d{8,})', line)
                    if match:
                        result['pk_no'] = match.group(1)

        # === HİZMET ALAN İŞYERİ BİLGİLERİ ===
        in_hizmet_alan = False
        for i, line in enumerate(lines):
            if 'HİZMET ALAN İŞYERİ BİLGİLERİ' in line:
                in_hizmet_alan = True
                continue

            if in_hizmet_alan:
                # "İMZA BİLGİLERİ" veya "HİZMET VEREN" görünce bölüm bitti
                if 'İMZA BİLGİLERİ' in line or 'HİZMET VEREN' in line:
                    break

                # Unvan - ": " ile başlayan ve şirket adı içeren satırlar
                if not result['firma_unvan']:
                    if line.strip().startswith(':') and ('MÜHENDİSLİK' in line or 'TİCARET' in line or 'SANAYİ' in line or 'LİMİTED' in line or 'ANONİM' in line or 'A.Ş' in line):
                        result['firma_unvan'] = line.strip().lstrip(':').strip()
                    elif 'ŞİRKETİ' in line and result.get('firma_unvan'):
                        # Unvan devam satırı
                        result['firma_unvan'] += ' ' + line.strip()
                    elif 'ŞİRKETİ' in line and not result.get('firma_unvan'):
                        # Önceki satırı kontrol et
                        if i > 0 and lines[i-1].strip().startswith(':'):
                            result['firma_unvan'] = lines[i-1].strip().lstrip(':').strip() + ' ' + line.strip()

                # Adres - ":Adres" ile başlayan satır
                if not result['firma_adres']:
                    if ':Adres' in line or 'Adres' in line:
                        # Adres satırını temizle
                        addr = line.replace(':Adres', '').replace('Adres', '').strip()
                        if addr:
                            result['firma_adres'] = addr
                    elif line.strip().startswith(':') and ('MAH' in line or 'CAD' in line or 'SOK' in line or 'NO:' in line):
                        result['firma_adres'] = line.strip().lstrip(':').strip()

                # SGK No - 24+ haneli sayı
                if not result['firma_sgk_no']:
                    match = re.search(r'(\d{24,26})', line)
                    if match:
                        result['firma_sgk_no'] = match.group(1)

                # İl
                if not result['firma_il']:
                    if ':İl' in line:
                        # Önceki satırda il adı olabilir
                        if i > 0:
                            prev = lines[i-1].strip()
                            if prev and prev.isupper() and len(prev) < 20 and not any(c.isdigit() for c in prev):
                                result['firma_il'] = prev
                    elif line.strip().isupper() and len(line.strip()) < 15 and not any(c.isdigit() for c in line.strip()):
                        # Tek kelimelik büyük harfli satır (il adı olabilir)
                        if line.strip() in ['İZMİR', 'ANKARA', 'İSTANBUL', 'BURSA', 'ANTALYA', 'KONYA', 'ADANA', 'MERSİN', 'GAZİANTEP', 'KOCAELİ', 'DENİZLİ', 'AYDIN', 'MUĞLA', 'ESKİŞEHİR', 'SAMSUN', 'TRABZON', 'KAYSERİ', 'BALIKESİR', 'SAKARYA', 'MANİSA']:
                            result['firma_il'] = line.strip()

        # === TESISAT SAYISI ===
        match = re.search(r'Tesisat Sayısı\s*:\s*(\d+)', text)
        if match:
            result['tesisat_sayisi'] = match.group(1)

    except Exception as e:
        print(f"PDF okuma hatası: {e}")

    return result


def format_sgk_no(sgk_no: str) -> str:
    """SGK numarasını okunabilir formata dönüştürür."""
    if not sgk_no:
        return ""
    # 5-8-5-6 formatında grupla
    if len(sgk_no) >= 24:
        return f"{sgk_no[:5]}-{sgk_no[5:13]}-{sgk_no[13:18]}-{sgk_no[18:]}"
    return sgk_no


if __name__ == "__main__":
    import sys
    import json

    if len(sys.argv) > 1:
        pdf_path = sys.argv[1]
    else:
        pdf_path = r"c:\Users\cmshe\OneDrive\Masaüstü\BUILDV\rapor_sistemi\hizmet_sozlesme_surecleri_detay_muayenekurulusu_periyodik_kontrol_202511241127401065.pdf"

    result = parse_sozlesme_pdf(pdf_path)

    print("=== ÇIKARILAN BİLGİLER ===")
    for key, value in result.items():
        if value:
            print(f"{key}: {value}")
