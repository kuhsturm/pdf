/**
 * KisiBilgileriReader.cpp
 */

#include "KisiBilgileriReader.h"
#include "xlsxdocument.h"
#include <QFile>
#include <QDebug>

namespace RaporSistemi {

// Excel yapısı sabitleri
static const int BLOCK_SIZE = 15;     // Her kişi bloğu 15 satır
static const int NAME_COL = 4;        // D sütunu - alan adları
static const int VALUE_COL = 5;       // E sütunu - değerler
static const int PERSON_NAME_COL = 4; // D sütunu - kişi adı (Eskiden 8 idi, şimdi D sütununda)
static const int PERSON_TC_COL = 9;   // I sütunu - TC no

KisiBilgileriReader::KisiBilgileriReader() = default;
KisiBilgileriReader::~KisiBilgileriReader() = default;

QString KisiBilgileriReader::normalizeName(const QString& name) const {
    QString normalized = name.toUpper().trimmed();

    // Türkçe karakter normalizasyonu
    normalized.replace("İ", "I");
    normalized.replace("Ğ", "G");
    normalized.replace("Ü", "U");
    normalized.replace("Ş", "S");
    normalized.replace("Ö", "O");
    normalized.replace("Ç", "C");

    // Çoklu boşlukları tek boşluğa indir
    normalized = normalized.simplified();

    return normalized;
}

bool KisiBilgileriReader::load(const QString& excelPath) {
    m_persons.clear();
    m_loaded = false;

    if (!QFile::exists(excelPath)) {
        m_errorString = QString("Dosya bulunamadı: %1").arg(excelPath);
        return false;
    }

    QXlsx::Document doc(excelPath);
    if (!doc.load()) {
        m_errorString = QString("Excel dosyası açılamadı: %1").arg(excelPath);
        return false;
    }

    // İlk sayfa
    if (!doc.selectSheet(doc.sheetNames().first())) {
        m_errorString = "Sayfa seçilemedi";
        return false;
    }

    // Kişileri oku - her BLOCK_SIZE satırda bir kişi
    int row = 1;
    int maxRow = 1000;  // Güvenlik limiti

    while (row < maxRow) {
        // Kişi adını kontrol et
        QString personName = doc.read(row, PERSON_NAME_COL).toString().trimmed();

        // --- NAME FILTERING ---
        // Column D contains: Headers ("3. TERMAL..."), Keys ("termal_cihaz_adi"), and Names ("AHMET ISIK")
        bool isInvalidName = false;
        if (personName.isEmpty()) isInvalidName = true;
        else if (personName.contains("_")) isInvalidName = true; // snake_case keys
        else if (personName.at(0).isDigit()) isInvalidName = true; // "3. TERMAL..."
        else if (personName.contains("BILGI")) isInvalidName = true; // Generic header check

        if (isInvalidName) {
            row++;
            continue;
        }

        // TC numarası (TC is usually in Column I/9, check if valid row)
        QString tcNo = doc.read(row, PERSON_TC_COL).toString().trimmed();

        KisiBilgisi kisi;
        kisi.adSoyad = personName;
        kisi.tcNo = tcNo;

        // Cihaz bilgilerini oku (blok içinde)
        int blockStart = row;
        int deviceIndex = 0; // Generic counter

        for (int r = blockStart; r < blockStart + BLOCK_SIZE && r < maxRow; ++r) {
            QString fieldName = doc.read(r, NAME_COL).toString().trimmed();
            QString fieldValue = doc.read(r, VALUE_COL).toString().trimmed();

            // Normalize field name for check
            QString normalizedField = fieldName.toUpper();

            // Satırda cihaz adı varsa, hangi tip olduğunu anla
            // Genellikle Cihaz Adi satırında "Cihaz Adı" yazar, değeri önemlidir
            if (normalizedField.contains("CİHAZ AD") || normalizedField.contains("CIHAZ AD")) {
                deviceIndex++;
                QString deviceName = fieldValue.toUpper();

                // Termal mi?
                bool isTermal = deviceName.contains("TERMAL") || deviceName.contains("TC01") || deviceName.contains("Tİ4");
                // Ölçüm mü?
                bool isOlcum = deviceName.contains("ÇOK FONKSİYONLU") || deviceName.contains("ÖLÇÜM") || deviceName.contains("1663") || deviceName.contains("1664");

                // İlgili satırları bulmak için loop içinde ileriye bakmak zor,
                // bu yüzden basit bir mapping yapalım:
                // deviceIndex 1 -> genellikle Termal veya Ölçüm'dür.
                // Excel yapısı: Cihaz Adı, Seri No, Kalibrasyon sırayla gelir.

                // Basit logic: O anki field Value cihaz adı ise kaydet
                // Sonraki satırlarda seri no vb. gelecek. Bu yapıyı yönetmek için state kuralım.
            }
        }

        // --- DAHA İYİ YÖNTEM: Bloğu komple tara ve propertyleri topla ---
        // Her cihaz 5 satırdır genellikle: Adı, Seri No, Tarih, Geçerlilik, No
        // Ancak Python kodu fieldName üzerinden gidiyor.
        // Hangi cihazı okuduğumuzu anlamak için flag kullanabiliriz.

        QString currentDeviceType = ""; // "TERMAL", "OLCUM", "OTHER"

        for (int r = blockStart; r < blockStart + BLOCK_SIZE && r < maxRow; ++r) {
             QString rawName = doc.read(r, NAME_COL).toString();
             QString val = doc.read(r, VALUE_COL).toString().trimmed();

             if (rawName.isEmpty()) continue;
             QString key = rawName.toUpper();

             // --- EXACT MATCH FOR USER'S EXCEL FORMAT (snake_case) ---
             if (key.contains("TERMAL_CIHAZ_ADI")) { kisi.termalCihazAdi = val; continue; }
             if (key.contains("TERMAL_KALIBRASYON_TARIHI")) { kisi.termalKalibrasyonTarihi = val; continue; }
             if (key.contains("TERMAL_KALIBRASYON_GECERLILIK")) { kisi.termalKalibrasyonGecerlilik = val; continue; }
             if (key.contains("TERMAL_SERI_NUMARASI")) { kisi.termalSeriNo = val; continue; }
             if (key.contains("TERMAL_KALIBRASYON_NO")) { kisi.termalKalibrasyonNo = val; continue; }

             if (key.contains("OLCUM_CIHAZ_ADI")) { kisi.olcumCihazAdi = val; continue; }
             if (key.contains("OLCUM_KALIBRASYON_TARIHI")) { kisi.olcumKalibrasyonTarihi = val; continue; }
             if (key.contains("OLCUM_KALIBRASYON_GECERLILIK")) { kisi.olcumKalibrasyonGecerlilik = val; continue; }
             if (key.contains("OLCUM_SERI_NUMARASI")) { kisi.olcumSeriNo = val; continue; }
             if (key.contains("OLCUM_KALIBRASYON_NO")) { kisi.olcumKalibrasyonNo = val; continue; }

             // --- FALLBACK GENERIC MATCHING ---
             // CİHAZ ADI satırı tip belirler
             // RELAXED MATCHING: Sadece "C" "H" "Z" ve "AD" arayalım veya direkt değeri kontrol edelim.
             bool isDeviceNameRow = key.contains("CIHAZ") || key.contains("CİHAZ") || (key.contains("AD") && val.length() > 3);

             if (isDeviceNameRow && val.length() > 2) {
                 QString devNameUpper = val.toUpper();
                 if (devNameUpper.contains("TERMAL") || devNameUpper.contains("TC0") || devNameUpper.contains("TI4") || devNameUpper.contains("FLUKE")) {
                     if (kisi.termalCihazAdi.isEmpty()) { // Only if not set by exact match
                        currentDeviceType = "TERMAL";
                        kisi.termalCihazAdi = val;
                     }
                 }
                 else if (devNameUpper.contains("1663") || devNameUpper.contains("1664") || devNameUpper.contains("FONK") || devNameUpper.contains("OLCUM") || devNameUpper.contains("ÖLÇÜM")) {
                     if (kisi.olcumCihazAdi.isEmpty()) {
                        currentDeviceType = "OLCUM";
                        kisi.olcumCihazAdi = val;
                     }
                 } else {
                     currentDeviceType = "OTHER";
                     if (kisi.cihaz1Adi.isEmpty()) kisi.cihaz1Adi = val;
                 }
             }
             else if (key.contains("SERİ") || key.contains("SERI") || key.contains("NUMARA")) {
                 if (currentDeviceType == "TERMAL" && kisi.termalSeriNo.isEmpty()) kisi.termalSeriNo = val;
                 else if (currentDeviceType == "OLCUM" && kisi.olcumSeriNo.isEmpty()) kisi.olcumSeriNo = val;
                 else if (kisi.cihaz1SeriNo.isEmpty()) kisi.cihaz1SeriNo = val;
             }
             else if (key.contains("KALİBRASYON") || key.contains("KALIBRASYON")) {
                 if (key.contains("TARİH") || key.contains("TARIH")) {
                     if (currentDeviceType == "TERMAL" && kisi.termalKalibrasyonTarihi.isEmpty()) kisi.termalKalibrasyonTarihi = val;
                     else if (currentDeviceType == "OLCUM" && kisi.olcumKalibrasyonTarihi.isEmpty()) kisi.olcumKalibrasyonTarihi = val;
                     else if (kisi.cihaz1Kalibrasyon.isEmpty()) kisi.cihaz1Kalibrasyon = val;
                 }
                 else if (key.contains("GEÇER") || key.contains("GECER")) {
                     if (currentDeviceType == "TERMAL" && kisi.termalKalibrasyonGecerlilik.isEmpty()) kisi.termalKalibrasyonGecerlilik = val;
                     else if (currentDeviceType == "OLCUM" && kisi.olcumKalibrasyonGecerlilik.isEmpty()) kisi.olcumKalibrasyonGecerlilik = val;
                 }
                 else if (key.contains("NO")) {
                     if (currentDeviceType == "TERMAL" && kisi.termalKalibrasyonNo.isEmpty()) kisi.termalKalibrasyonNo = val;
                     else if (currentDeviceType == "OLCUM" && kisi.olcumKalibrasyonNo.isEmpty()) kisi.olcumKalibrasyonNo = val;
                 }
             }
        }

    // Normalize edilmiş isimle kaydet
        QString normalizedName = normalizeName(personName);
        m_persons[normalizedName] = kisi;

        qDebug() << "Loaded Person:" << normalizedName << "Termal:" << kisi.termalCihazAdi << "Olcum:" << kisi.olcumCihazAdi;

        row += BLOCK_SIZE;
    }

    m_loaded = !m_persons.isEmpty();
    qDebug() << "Total persons loaded:" << m_persons.size();
    return m_loaded;
}

QStringList KisiBilgileriReader::getPersonList() const {
    QStringList names;
    for (auto it = m_persons.constBegin(); it != m_persons.constEnd(); ++it) {
        names.append(it.value().adSoyad);
    }
    return names;
}

KisiBilgisi KisiBilgileriReader::getPersonByName(const QString& name) const {
    QString normalized = normalizeName(name);
    qDebug() << "Searching for person:" << name << "Normalized:" << normalized;

    // Tam eşleşme
    if (m_persons.contains(normalized)) {
        qDebug() << "Exact match found.";
        return m_persons[normalized];
    }

    // Kısmi eşleşme
    for (auto it = m_persons.constBegin(); it != m_persons.constEnd(); ++it) {
        if (it.key().contains(normalized) || normalized.contains(it.key())) {
            qDebug() << "Partial match found with:" << it.key();
            return it.value();
        }
    }

    qDebug() << "No match found for:" << name;
    return {};
}

void KisiBilgileriReader::fillCihazBilgileri(const QString& name, FirmaBilgileri& firma) const {
    KisiBilgisi kisi = getPersonByName(name);

    if (!kisi.isValid()) {
         qDebug() << "Person invalid, skipping device fill.";
         return;
    }

    firma.termalCihazAdi = kisi.termalCihazAdi;
    firma.termalKalibrasyonTarihi = kisi.termalKalibrasyonTarihi;
    firma.termalKalibrasyonGecerlilik = kisi.termalKalibrasyonGecerlilik;
    firma.termalSeriNo = kisi.termalSeriNo;
    firma.termalKalibrasyonNo = kisi.termalKalibrasyonNo;

    firma.olcumCihazAdi = kisi.olcumCihazAdi;
    firma.olcumKalibrasyonTarihi = kisi.olcumKalibrasyonTarihi;
    firma.olcumKalibrasyonGecerlilik = kisi.olcumKalibrasyonGecerlilik;
    firma.olcumSeriNo = kisi.olcumSeriNo;
    firma.olcumKalibrasyonNo = kisi.olcumKalibrasyonNo;

    // Legacy support
    if (firma.cihaz1Adi.isEmpty()) firma.cihaz1Adi = kisi.cihaz1Adi;
    if (firma.cihaz1SeriNo.isEmpty()) firma.cihaz1SeriNo = kisi.cihaz1SeriNo;
    if (firma.cihaz1KalibrasyonTarihi.isEmpty()) firma.cihaz1KalibrasyonTarihi = kisi.cihaz1Kalibrasyon;

    // Fallback if structured data empty but legacy has data
    if (firma.termalCihazAdi.isEmpty() && !kisi.cihaz1Adi.isEmpty()) {
       // Maybe assign legacy to termal or olcum based on guess?
       // For now keep as is.
    }

    if (firma.kontrolEdenTc.isEmpty()) {
        firma.kontrolEdenTc = kisi.tcNo;
    }
}

QString KisiBilgileriReader::getTcNo(const QString& name) const {
    KisiBilgisi kisi = getPersonByName(name);
    return kisi.tcNo;
}

void getCihazFromSozlesme(const QString& kisiExcelPath,
                          const QString& kontrolEdenAdSoyad,
                          FirmaBilgileri& firma) {
    KisiBilgileriReader reader;

    if (reader.load(kisiExcelPath)) {
        reader.fillCihazBilgileri(kontrolEdenAdSoyad, firma);
    }
}

} // namespace RaporSistemi
