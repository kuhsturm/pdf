"""Rapor Üretici Modülü v2.1
Excel verilerini okuyarak DOCX raporu oluşturur.
Güncellemeler:
- Ana Dağıtım Pano desteği
- Otomatik kusur açıklaması
- Iz otomatik hesaplama
- Son 2 termal görsel
- PyInstaller gömülü şablon desteği
"""

from excel_reader import ExcelReader
from docx_writer import DocxWriter
from fluke_extractor import FlukeExtractor
from typing import Dict, Any, Optional, List
import os
import sys
import json
import tempfile
import shutil

# Şablon dosya adı (sabit)
TEMPLATE_FILENAME = "Elektrik Tesisatı Gözle Kontrol ve Fonksiyon Testleri Periyodik Kontrol Raporu.docx"


def get_base_path() -> str:
    """PyInstaller ile derlenmişse _MEIPASS, değilse script dizini döndürür."""
    if getattr(sys, 'frozen', False):
        # PyInstaller ile derlenmiş EXE
        return sys._MEIPASS
    else:
        # Normal Python çalıştırma
        return os.path.dirname(os.path.abspath(__file__))


def get_exe_directory() -> str:
    """EXE dosyasının bulunduğu dizini döndürür."""
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    else:
        return os.path.dirname(os.path.abspath(__file__))


def resolve_template_path(preferred: Optional[str] = None) -> str:
    """Find a usable DOCX template path with fallbacks.

    Priority:
    1) Explicit path passed in (preferred)
    2) PyInstaller gömülü şablon (_MEIPASS içinde)
    3) EXE ile aynı dizindeki şablon
    4) RAPOR_TEMPLATE_PATH environment variable
    5) config/system_config.json -> template_dosya
    6) Üst dizindeki varsayılan şablon (geliştirme modu)
    """

    def _exists(path: Optional[str]) -> Optional[str]:
        return path if path and os.path.exists(path) else None

    # 1) Explicit path
    if preferred and os.path.exists(preferred):
        return preferred

    # 2) rapor_sistemi/sablon/ klasöründeki şablon (öncelikli)
    script_dir = os.path.dirname(os.path.abspath(__file__))
    sablon_template = os.path.join(script_dir, "sablon", TEMPLATE_FILENAME)
    if os.path.exists(sablon_template):
        return sablon_template

    # 3) PyInstaller gömülü şablon
    base_path = get_base_path()
    embedded_template = os.path.join(base_path, TEMPLATE_FILENAME)
    if os.path.exists(embedded_template):
        return embedded_template

    # 4) EXE ile aynı dizindeki şablon
    exe_dir = get_exe_directory()
    exe_template = os.path.join(exe_dir, TEMPLATE_FILENAME)
    if os.path.exists(exe_template):
        return exe_template

    # 4) Ortam değişkeni
    env_path = os.getenv("RAPOR_TEMPLATE_PATH")
    if env_path and os.path.exists(env_path):
        return env_path

    # 5) Config dosyası
    repo_root = os.path.abspath(os.path.join(os.path.dirname(__file__), os.pardir))
    config_path = os.path.join(repo_root, "config", "system_config.json")
    if os.path.exists(config_path):
        try:
            with open(config_path, "r", encoding="utf-8") as f:
                cfg = json.load(f)
            tmpl_name = cfg.get("template_dosya")
            if tmpl_name:
                config_template = os.path.join(repo_root, tmpl_name)
                if os.path.exists(config_template):
                    return config_template
        except Exception:
            pass

    # 6) Üst dizindeki varsayılan şablon (geliştirme modu)
    common_default = os.path.join(repo_root, TEMPLATE_FILENAME)
    if os.path.exists(common_default):
        return common_default

    raise FileNotFoundError(
        f"DOCX şablonu bulunamadı!\n"
        f"Aranan konumlar:\n"
        f"  - EXE içi: {embedded_template}\n"
        f"  - EXE yanı: {exe_template}\n"
        f"  - Ortam değişkeni: RAPOR_TEMPLATE_PATH\n"
        f"  - Varsayılan: {common_default}\n\n"
        f"Lütfen '{TEMPLATE_FILENAME}' dosyasını EXE ile aynı dizine kopyalayın."
    )


class ReportGenerator:
    """Excel verilerinden veya sözlükten DOCX raporu oluşturur."""

    def __init__(self, template_path: Optional[str] = None):
        """
        Args:
            template_path: DOCX şablon dosyasının yolu (opsiyonel, otomatik bulunur)
        """
        self.template_path = resolve_template_path(template_path)
        self.temp_dir = None

    def generate(self, data_source: Any, output_path: str) -> str:
        """
        Rapor oluşturur.

        Args:
            data_source: Veri içeren Excel dosyasının yolu (str) VEYA veri sözlüğü (dict)
            output_path: Çıktı DOCX dosyasının yolu

        Returns:
            Oluşturulan raporun yolu
        """
        # Geçici klasör oluştur
        self.temp_dir = tempfile.mkdtemp(prefix="rapor_")

        try:
            # Veri kaynağını işle
            data = {}
            if isinstance(data_source, str):
                print("Excel verileri okunuyor...")
                reader = ExcelReader(data_source)
                data = reader.read_all()
            elif isinstance(data_source, dict):
                 print("Veri sozlukten aliniyor...")
                 data = data_source
            else:
                raise ValueError("data_source must be a file path (str) or a dictionary")

            # DOCX yazıcıyı başlat
            print("Sablon yukleniyor...")
            writer = DocxWriter(self.template_path)
            writer.load_template()

            # Firma bilgilerini doldur
            if data.get('firma_bilgileri'):
                print("Firma bilgileri dolduruluyor...")
                writer.fill_firma_bilgileri(data['firma_bilgileri'])

            # Ana Dağıtım Pano bilgilerini doldur
            if data.get('ana_dagitim_pano'):
                print("Ana Dagitim Pano bilgileri dolduruluyor...")
                # Pano adını gözle kontrol verilerinden al
                pano_adi = None
                if data.get('gozle_kontrol'):
                    pano_adi = data['gozle_kontrol'].get('pano_adi')
                writer.fill_ana_dagitim_pano(data['ana_dagitim_pano'], pano_adi)

            # Cihaz bilgilerini doldur
            if data.get('cihaz_bilgileri'):
                print("Cihaz bilgileri dolduruluyor...")
                writer.fill_cihaz_bilgileri(data['cihaz_bilgileri'])

            # Gözle kontrol bölümünü doldur
            if data.get('gozle_kontrol'):
                print("Gozle kontrol verileri dolduruluyor...")
                writer.fill_gozle_kontrol(data['gozle_kontrol'])

            # Fonksiyon testlerini doldur (Iz otomatik hesaplanmış)
            if data.get('fonksiyon_testleri'):
                print(f"Fonksiyon testleri dolduruluyor ({len(data['fonksiyon_testleri'])} satir)...")
                # ana_dagitim_pano verisini placeholder değiştirme için geçir
                writer.fill_fonksiyon_testleri(data['fonksiyon_testleri'], data.get('ana_dagitim_pano'))

            # Termal görüntüleri işle (SON 2 GÖRSEL)
            if data.get('termal_goruntuler'):
                print("Termal goruntuler isleniyor...")
                print(f"  DEBUG: termal_goruntuler listesi = {data['termal_goruntuler']}")
                all_images = []

                photo_date = None
                photo_no = None

                for termal in data['termal_goruntuler']:
                    fluke_path = termal.get('fluke_dosya')
                    print(f"  DEBUG: fluke_path = {fluke_path}, exists = {os.path.exists(fluke_path) if fluke_path else 'N/A'}")

                    if fluke_path and os.path.exists(fluke_path):
                        try:
                            extractor = FlukeExtractor(fluke_path)
                            # only_last_two=True parametresi ile son 2 görseli al
                            result = extractor.extract_all(self.temp_dir)
                            all_images.extend(result.get('images', []))

                            # Fotoğraf tarih/no'yu ilk bulunan dosyadan al
                            if not photo_date:
                                photo_date = result.get('photo_date')
                            if not photo_no:
                                photo_no = result.get('photo_no')
                            print(f"  - {len(result.get('images', []))} gorsel cikarildi: {os.path.basename(fluke_path)}")
                        except Exception as e:
                            print(f"  ! Fluke dosyasi islenemedi: {fluke_path}, Hata: {e}")

                if all_images:
                    print(f"Toplam {len(all_images)} termal goruntu rapora ekleniyor...")
                    writer.add_thermal_images(all_images)
                else:
                    print("  ! UYARI: Hic termal goruntu cikarilmadi!")

                # GK_24 / GK_25 alanlarını fluke bilgisinden doldur
                print(f"  DEBUG: photo_date={photo_date}, photo_no={photo_no}")
                if data.get('gozle_kontrol') and 'kontroller' in data['gozle_kontrol']:
                    kontroller = data['gozle_kontrol']['kontroller']
                    if photo_date:
                        kontroller['Fotograf Tarihi'] = photo_date
                        print(f"  DEBUG: kontroller['Fotograf Tarihi'] = {photo_date}")
                    if photo_no:
                        kontroller['Fotograf No'] = photo_no
                        print(f"  DEBUG: kontroller['Fotograf No'] = {photo_no}")
                    # GK_29 (Acil Durum Aydinlatma) GUI'den geliyor, zorla değiştirme

                    # Gözle kontrol tablosu termal veriden sonra yeniden yazılsın
                    print("  DEBUG: fill_gozle_kontrol ikinci cagri yapiliyor...")
                    writer.fill_gozle_kontrol(data['gozle_kontrol'])
                    print("  DEBUG: fill_gozle_kontrol ikinci cagri tamamlandi")

            # Kusur açıklamasını otomatik oluştur
            kusurlar = []
            if data.get('gozle_kontrol') and data['gozle_kontrol'].get('kusurlar'):
                kusurlar = data['gozle_kontrol']['kusurlar']
                print(f"Kusur aciklamasi olusturuluyor ({len(kusurlar)} kusur)...")

            # Sonuç bölümünü doldur
            sonuc_data = data.get('sonuc', {})
            writer.fill_sonuc(sonuc_data, kusurlar)

            # Uygunluk durumunu doldur (10. SONUÇ VE KANAAT)
            uygunluk = data.get('ana_dagitim_pano', {}).get('Uygunluk', 'Uygun')
            print(f"Uygunluk durumu dolduruluyor: {uygunluk}")
            writer.fill_uygunluk(uygunluk)

            # 6.2 Potansiyel Dengeleme ve 6.3 Zemin İzolasyonu bölümlerini doldur
            print("6.2 ve 6.3 bölümleri dolduruluyor...")
            # Pano adını bul (gozle_kontrol'dan veya ana_dagitim_pano'dan)
            pano_adi = data.get('gozle_kontrol', {}).get('pano_adi', '') or \
                       data.get('gozle_kontrol', {}).get('Pano Adi', '') or \
                       data.get('ana_dagitim_pano', {}).get('Pano Adi (PANO_adi1)', '')
            print(f"  DEBUG: pano_adi = '{pano_adi}'")
            pot_data = {
                'pano_adi': pano_adi,
                'gozle_kontrol': data.get('gozle_kontrol', {})
            }
            writer.fill_potansiyel_dengeleme_ve_zemin(pot_data, data.get('fonksiyon_testleri', []))

            # 6.2 Potansiyel Dengeleme tablosundan önce sayfa sonu ekle
            print("Sayfa sonu ekleniyor (Tablo 3 oncesi)...")
            writer.add_page_break_before_table(3)

            # Raporu kaydet
            print(f"Rapor kaydediliyor: {output_path}")
            writer.save(output_path)

            print("Rapor basariyla olusturuldu!")
            return output_path

        finally:
            # Geçici dosyaları temizle
            self._cleanup_temp()

    def _cleanup_temp(self):
        """Geçici dosyaları temizler."""
        if self.temp_dir and os.path.exists(self.temp_dir):
            import shutil
            try:
                shutil.rmtree(self.temp_dir)
            except Exception:
                pass


def main():
    """Ana fonksiyon - komut satırı kullanımı için."""
    import sys

    if len(sys.argv) < 3:
        print("Kullanim: python report_generator.py <excel_dosyasi> <cikti_dosyasi>")
        print("Ornek: python report_generator.py veri.xlsx rapor.docx")
        sys.exit(1)

    excel_path = sys.argv[1]
    output_path = sys.argv[2]

    try:
        template_path = resolve_template_path()
    except FileNotFoundError as e:
        print(f"Sablon bulunamadı: {e}")
        sys.exit(1)

    generator = ReportGenerator(template_path)

    try:
        result = generator.generate(excel_path, output_path)
        print(f"\n=== BASARILI ===")
        print(f"Rapor olusturuldu: {result}")
    except Exception as e:
        print(f"\n=== HATA ===")
        print(f"Rapor olusturulurken hata: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)


if __name__ == "__main__":
    main()
