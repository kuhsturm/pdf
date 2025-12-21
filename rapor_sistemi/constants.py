"""
Uygulama genelinde kullanılan sabitler ve konfigürasyon değerleri.
"""

# In değerine göre önerilen minimum kablo kesiti (bakır, Grup 2)
# Akım taşıma (Grup 2): 1.5→20A, 2.5→27A, 4→36A, 6→47A, 10→65A, 16→87A, 25→115A, 35→143A, 50→178A, 70→220A, 95→265A, 120→310A, 150→355A, 185→400A, 240→480A
IN_TO_KESIT = {
    1: "1.5", 2: "1.5", 3: "1.5", 4: "1.5", 6: "1.5",
    10: "1.5", 13: "1.5", 16: "1.5", 20: "1.5",
    25: "2.5", 32: "4", 40: "6", 50: "10",
    63: "10", 80: "16", 100: "25", 125: "35",
    160: "50", 200: "70", 250: "95", 315: "150", 400: "185"
}

# Kablo kesitine göre Iz tablosu (Grup 2, bakır)
KESIT_TO_IZ = {
    "1.5": 20, "2.5": 27, "4": 36, "6": 47, "10": 65,
    "16": 87, "25": 115, "35": 143, "50": 178, "70": 220,
    "95": 265, "120": 310, "150": 355, "185": 405, "240": 480
}

# Grup 2 akım taşıma tabloları (A) - sayısal kesit değerleri için
IZ_TABLE = {
    0.75: 13,
    1.0: 16,
    1.5: 20,
    2.5: 27,
    4.0: 36,
    6.0: 47,
    10.0: 65,
    16.0: 87,
    25.0: 115,
    35.0: 143,
    50.0: 178,
    70.0: 220,
    95.0: 265,
    120.0: 310,
    150.0: 355,
    185.0: 400,
    240.0: 480,
    300.0: 555,
    400.0: 770,
    500.0: 880,
}

# In (A) seçenekleri
IN_VALUES = ["1", "2", "3", "4", "6", "10", "13", "16", "20", "25", "32", "40", "50", "63", "80", "100", "125", "160", "200", "250", "315", "400"]

# Icu (kA) seçenekleri
ICU_VALUES = ["3kA", "4.5kA", "10kA", "25kA", "35kA", "55kA", "70kA"]

# Kablo kesitleri
BASE_SECTIONS = ["1.5", "2.5", "4", "6", "10", "16", "25", "35", "50", "70", "95", "120", "150", "185", "240"]
MULTIPLIERS = ["2", "3", "4", "5", "6", "7", "8"]

def get_all_sections():
    """Tüm kesit kombinasyonlarını döndürür."""
    return BASE_SECTIONS + [f"{m}x{s}" for m in MULTIPLIERS for s in ["16", "25", "35", "50", "70", "95", "120", "150", "185", "240"]]

ALL_SECTIONS = get_all_sections()

# Gözle kontrol maddeleri
GK_FIELDS = [
    "Kablo Sebeke Tarafi", "Kablo Donanim Tarafi",
    "Pano Sabitlenmesi", "Dis Darbelere Karsi Koruma Onlemi",
    "Elektrik Panosu Etrafinda Yabanci Malzemeler", "Zemin Izolasyonu",
    "Topraklama Iletkeni", "Ana Potansiyel Dengeleme Iletkeni",
    "Ek Potansiyel Dengeleme Iletkeni", "Pano Kapak Baglantisi Kontrolu 6 mm2",
    "Elektriksel Olmayan Tesislere Yaklasma", "Bant Ayrilmasi",
    "Guvenlik Devre Ayrilmasi", "Pano Ic Kapak",
    "Semalar Talimatlar", "Koruma Cihaz ve Terminal Etiket",
    "Tehlike Isaretleri", "Kablo Yollari",
    "Kablo Renk Kodlari", "Tesisat Yontemi",
    "Yangin Engeli", "Kontak Gevsekligi Isinmasi",
    "Asiri Yuk Isinmasi", "Yangin Sondurme",
    "Ekipman Temizlik", "Korozyon Kontrolu",
    "Acil Durum Aydinlatma",
]

def parse_kesit_value(section_val: str) -> float:
    """Parses a section string (e.g., '2x16', '2,5') into a float value."""
    if section_val is None:
        return 0.0
    val = str(section_val).lower().replace(',', '.').strip()
    if not val:
        return 0.0
    factor = 1
    base = val
    if 'x' in val:
        try:
            f_str, base = val.split('x', 1)
            factor = float(f_str)
        except ValueError:
            factor = 1
    try:
        size = float(base)
    except ValueError:
        return 0.0
    return size * factor

def calculate_iz(section_val: str) -> str:
    """Calculates Iz (Current Carrying Capacity) based on section value."""
    size = parse_kesit_value(section_val)
    if not size:
        return ""

    # Exact match
    if size in IZ_TABLE:
        return f"{IZ_TABLE[size]:.0f}"

    # Closest match (lower bound)
    for k in sorted(IZ_TABLE.keys()):
        if size <= k:
            return f"{IZ_TABLE[k]:.0f}"

    # If larger than max, return max
    return f"{list(IZ_TABLE.values())[-1]:.0f}"
