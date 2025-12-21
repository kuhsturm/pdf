# -*- mode: python ; coding: utf-8 -*-
"""
Elektrik Tesisatı Rapor Sistemi - PyInstaller Spec Dosyası
Çoklu Pano GUI uygulamasını derler.
DOCX şablonu EXE içine gömülüdür.
"""

import os
import sys
from PyInstaller.utils.hooks import collect_data_files, collect_submodules

block_cipher = None

# Çalışma dizini (spec dosyasının bulunduğu yer)
SPEC_DIR = os.getcwd()
print(f"[INFO] Çalışma dizini: {SPEC_DIR}")

# CustomTkinter veri dosyalarını topla
ctk_datas = collect_data_files('customtkinter')

# Sablon klasörü - EXE içine gömülecek
SABLON_DIR = os.path.join(SPEC_DIR, 'sablon')

# Şablon klasörünü veri olarak ekle (hedef: sablon/)
template_datas = []
if os.path.exists(SABLON_DIR):
    template_datas = [(SABLON_DIR, 'sablon')]
    print(f"[INFO] Sablon klasörü bulundu: {SABLON_DIR}")
else:
    print(f"[WARNING] Sablon klasörü bulunamadı: {SABLON_DIR}")

# Kisi Bilgileri dosyası (varsayılan veri)
KISI_BILGILERI_FILE = "kisi_bilgileri.xlsx"
KISI_BILGILERI_SOURCE = os.path.join(SPEC_DIR, KISI_BILGILERI_FILE)
if os.path.exists(KISI_BILGILERI_SOURCE):
    template_datas.append((KISI_BILGILERI_SOURCE, '.'))
    print(f"[INFO] Kisi bilgileri eklendi: {KISI_BILGILERI_SOURCE}")
else:
    print(f"[WARNING] kisi_bilgileri.xlsx bulunamadı!")

# Ana modül
main_module = 'multi_pano_gui.py'

# Yerel Python modülleri - PyInstaller data olarak ekle
local_py_files = [
    'report_generator.py',
    'excel_reader.py',
    'docx_writer.py',
    'fluke_extractor.py',
    'sozlesme_parser.py',
    'kisi_bilgileri_reader.py',
]

# Yerel modülleri data olarak ekle (çalışma zamanında import edilebilsin)
local_datas = []
for py_file in local_py_files:
    if os.path.exists(py_file):
        local_datas.append((py_file, '.'))
        print(f"[INFO] Yerel modül eklendi: {py_file}")
    else:
        print(f"[WARNING] Yerel modül bulunamadı: {py_file}")

# Gizli importlar
hidden_imports = [
    'customtkinter',
    'tkinter',
    'tkinter.messagebox',
    'tkinter.filedialog',
    'openpyxl',
    'openpyxl.workbook',
    'openpyxl.utils',
    'docx',
    'docx.document',
    'docx.shared',
    'docx.enum.text',
    'pypdf',
    'PIL',
    'PIL.Image',
    'lxml',
    'lxml.etree',
    'json',
    'datetime',
    'tempfile',
    'shutil',
    'zipfile',
    're',
    'random',
]

# Analiz - tüm yerel modülleri dahil et
a = Analysis(
    [main_module],
    pathex=[SPEC_DIR, '.'],  # Yerel modüllerin bulunduğu dizin
    binaries=[],
    datas=ctk_datas + template_datas + local_datas,  # Şablon + yerel modüller
    hiddenimports=hidden_imports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        'matplotlib',
        'numpy',
        'pandas',
        'scipy',
        'google',
        'google-generativeai',
    ],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

# PYZ arşivi
pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

# EXE oluştur
exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='ElektrikRaporSistemi',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,  # GUI uygulama olduğu için konsolu gizle
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=None,  # İkon dosyası eklenebilir: icon='icon.ico'
)
