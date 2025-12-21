"""
Elektrik Rapor Sistemi - Çoklu Pano GUI
Aynı firma için birden fazla pano raporu oluşturma.
Ortak bilgiler 1 kez girilir, her pano için ayrı rapor üretilir.
"""

import customtkinter as ctk
from tkinter import messagebox, filedialog
import tkinter as tk
from typing import Dict, Any, List, Optional
import os
import json
import sys
import random
import threading
from datetime import datetime

# Import yapılandırması - Modül olarak veya script olarak çalışmaya uyum
if __name__ == "__main__":
    # Eğer bu script doğrudan çalıştırılıyorsa, sys.path'e parent ekle
    # Böylece 'import constants' çalışabilir
    current_dir = os.path.dirname(os.path.abspath(__file__))
    parent_dir = os.path.dirname(current_dir)
    # Ayrıca current_dir'i ekleyelim ki gui paketi import edilebilsin
    if current_dir not in sys.path:
        sys.path.append(current_dir)

try:
    from report_generator import resolve_template_path
    from gui.utils import ScrollableFrame
    from gui.pano_frame import PanoDataFrame
    import constants as const
except ImportError:
    # Modül olarak çalıştırılıyorsa
    from .report_generator import resolve_template_path
    from .gui.utils import ScrollableFrame
    from .gui.pano_frame import PanoDataFrame
    from . import constants as const

# Tema ayarları
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")


class MultiPanoApp(ctk.CTk):
    """Çoklu pano rapor uygulaması."""

    def __init__(self):
        super().__init__()

        self.title("Elektrik Tesisatı Denetim Raporu v2.0")
        self.geometry("1500x900")
        # Grid layout 1x2

        self.pano_frames = []

        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)

        self.create_sidebar()
        self.create_main_area()

        # İlk pano
        self.add_pano()

    def create_sidebar(self):
        """Sol panel - ortak bilgiler."""
        sidebar = ctk.CTkFrame(self, width=350, fg_color="#1f1f1f")
        sidebar.grid(row=0, column=0, sticky="nsew", padx=5, pady=5)
        sidebar.grid_propagate(False)

        # Başlık
        ctk.CTkLabel(sidebar, text="🏢 Ortak Bilgiler", font=ctk.CTkFont(size=16, weight="bold")).pack(pady=10)

        # Sekmeler
        self.common_tabs = ctk.CTkTabview(sidebar, height=500)
        self.common_tabs.pack(fill="both", expand=True, padx=5, pady=5)

        self.tab_firma = self.common_tabs.add("Firma")
        self.tab_pano = self.common_tabs.add("Ana Pano")

        self.create_firma_section()
        self.create_pano_section()

        # Alt butonlar
        btn_frame = ctk.CTkFrame(sidebar, fg_color="transparent")
        btn_frame.pack(fill="x", padx=10, pady=10)

        ctk.CTkButton(btn_frame, text="💾 Kaydet", command=self.save_project, width=80).pack(side="left", padx=2)
        ctk.CTkButton(btn_frame, text="📂 Aç", command=self.load_project, width=80).pack(side="left", padx=2)

        # Sözleşme PDF yükleme butonu
        ctk.CTkButton(
            sidebar,
            text="📄 Sözleşme PDF Yükle",
            command=self.load_sozlesme_pdf,
            fg_color="#5c6bc0",
            hover_color="#3949ab"
        ).pack(fill="x", padx=10, pady=5)

        # Toplu rapor butonu
        self.generate_btn = ctk.CTkButton(
            sidebar,
            text="📋 TOPLU RAPOR OLUŞTUR",
            command=self.start_report_generation,
            height=50,
            font=ctk.CTkFont(size=14, weight="bold"),
            fg_color="#2e7d32",
            hover_color="#1b5e20"
        )
        self.generate_btn.pack(fill="x", padx=10, pady=10)

        # Progress Bar
        self.progress_bar = ctk.CTkProgressBar(sidebar)
        self.progress_bar.set(0)
        self.progress_bar.pack(fill="x", padx=10, pady=5)
        self.progress_bar.pack_forget() # Başlangıçta gizle

    def create_firma_section(self):
        """Firma bilgileri."""
        scroll = ScrollableFrame(self.tab_firma)
        scroll.pack(fill="both", expand=True)

        self.firma_entries = {}

        fields = [
            ("Firma Adi", ""),
            ("Tesis Adresi", ""),
            ("SGK Sicil No", ""),
            ("Sozlesme ID", ""),
            ("Teklif Numarasi", ""),
            ("Rapor Tarihi", datetime.now().strftime("%d.%m.%Y")),
            ("Kontrol Baslangic", datetime.now().strftime("%d.%m.%Y 09:00")),
            ("Kontrol Bitis", datetime.now().strftime("%d.%m.%Y 18:00")),
            ("Kontrol Eden", ""),
            ("Belge No", ""),
        ]

        for field, default in fields:
            row = ctk.CTkFrame(scroll, fg_color="transparent")
            row.pack(fill="x", pady=3, padx=5)

            ctk.CTkLabel(row, text=field, width=120, anchor="w").pack(side="left")

            entry = ctk.CTkEntry(row, width=200)
            entry.pack(side="left", fill="x", expand=True)
            if default:
                entry.insert(0, default)

            self.firma_entries[field] = entry

        # Rapor No Prefix (TPK2025-x formatı)
        row = ctk.CTkFrame(scroll, fg_color="transparent")
        row.pack(fill="x", pady=10, padx=5)

        ctk.CTkLabel(row, text="Rapor No Prefix", width=100, anchor="w",
                    font=ctk.CTkFont(weight="bold")).pack(side="left")

        self.rapor_prefix_entry = ctk.CTkEntry(row, width=150, placeholder_text="TPK2025-4001")
        self.rapor_prefix_entry.pack(side="left")

        ctk.CTkLabel(row, text="-Y", text_color="#888888").pack(side="left", padx=5)

        # Açıklama
        info = ctk.CTkLabel(scroll, text="💡 Her pano için rapor numarası otomatik artar:\n   Örnek: TPK2025-4001-1, TPK2025-4001-2, ...",
                           font=ctk.CTkFont(size=10), text_color="#888888", justify="left")
        info.pack(fill="x", padx=5, pady=5)

        # === PROJE VE TEK HAT ŞEMASI SEÇİMLERİ ===
        ctk.CTkLabel(scroll, text="📋 Tesis Belgeleri", font=ctk.CTkFont(size=12, weight="bold"),
                    text_color="#03a9f4").pack(fill="x", padx=5, pady=(15, 5))

        def on_proje_tekhat_change(choice=None):
            """Proje ve Tek Hat ikisi de Yok ise Ib alanlarını devre dışı bırak."""
            proje = self.proje_var_combo.get()
            tekhat = self.tekhat_var_combo.get()

            # İkisi de Yok ise Ib devre dışı
            ib_disabled = (proje == "Yok" and tekhat == "Yok")

            # Tüm pano frame'lerindeki Ib alanlarını güncelle
            for pano_frame in self.pano_frames:
                for row_data in pano_frame.ft_rows:
                    ib_entry = row_data['entries'].get('Ib')
                    if ib_entry:
                        if ib_disabled:
                            ib_entry.configure(state='normal')
                            ib_entry.delete(0, 'end')
                            ib_entry.insert(0, "-")
                            ib_entry.configure(state='readonly')
                        else:
                            # Ib'yi yeniden hesapla
                            in_val = row_data['entries'].get('In (A)')
                            if in_val:
                                try:
                                    in_float = float(in_val.get())
                                    ib_val = in_float * 0.7
                                    ib_entry.configure(state='normal')
                                    ib_entry.delete(0, 'end')
                                    ib_entry.insert(0, f"{ib_val:.1f}")
                                    ib_entry.configure(state='readonly')
                                except:
                                    pass

        # Proje Var mı?
        proje_row = ctk.CTkFrame(scroll, fg_color="transparent")
        proje_row.pack(fill="x", pady=3, padx=5)
        ctk.CTkLabel(proje_row, text="Tesise ait proje var mı?", width=180, anchor="w").pack(side="left")
        self.proje_var_combo = ctk.CTkComboBox(proje_row, values=["Var", "Yok"], width=100, command=on_proje_tekhat_change)
        self.proje_var_combo.set("Var")
        self.proje_var_combo.pack(side="left", padx=5)
        self.firma_entries["Proje Var"] = self.proje_var_combo

        # Tek Hat Şeması Var mı?
        tekhat_row = ctk.CTkFrame(scroll, fg_color="transparent")
        tekhat_row.pack(fill="x", pady=3, padx=5)
        ctk.CTkLabel(tekhat_row, text="Tek hat şeması var mı?", width=180, anchor="w").pack(side="left")
        self.tekhat_var_combo = ctk.CTkComboBox(tekhat_row, values=["Var", "Yok"], width=100, command=on_proje_tekhat_change)
        self.tekhat_var_combo.set("Var")
        self.tekhat_var_combo.pack(side="left", padx=5)
        self.firma_entries["Tek Hat Var"] = self.tekhat_var_combo

        # Yapı Cinsi
        yapi_row = ctk.CTkFrame(scroll, fg_color="transparent")
        yapi_row.pack(fill="x", pady=3, padx=5)
        ctk.CTkLabel(yapi_row, text="Yapı cinsi", width=180, anchor="w").pack(side="left")
        self.yapi_cinsi_combo = ctk.CTkComboBox(yapi_row, values=["Ev", "Ticari", "Endüstri", "Diğer"], width=120)
        self.yapi_cinsi_combo.set("Ticari")
        self.yapi_cinsi_combo.pack(side="left", padx=5)
        self.firma_entries["Yapi Cinsi"] = self.yapi_cinsi_combo

        # === CİHAZ BİLGİLERİ BÖLÜMÜ ===
        ctk.CTkLabel(scroll, text="📷 Termal Kamera", font=ctk.CTkFont(size=12, weight="bold"),
                    text_color="#ff9800").pack(fill="x", padx=5, pady=(15, 5))

        termal_fields = [
            ("Termal Cihaz Adi", ""),
            ("Termal Kalibrasyon Tarihi", ""),
            ("Termal Kalibrasyon Gecerlilik", ""),
            ("Termal Seri No", ""),
            ("Termal Kalibrasyon No", ""),
        ]

        for field, default in termal_fields:
            row = ctk.CTkFrame(scroll, fg_color="transparent")
            row.pack(fill="x", pady=2, padx=5)
            label_text = field.replace("Termal ", "")
            ctk.CTkLabel(row, text=label_text, width=130, anchor="w", font=ctk.CTkFont(size=10)).pack(side="left")
            entry = ctk.CTkEntry(row, width=180, font=ctk.CTkFont(size=10))
            entry.pack(side="left", fill="x", expand=True)
            if default:
                entry.insert(0, default)
            self.firma_entries[field] = entry

        ctk.CTkLabel(scroll, text="🔌 Ölçüm Cihazı", font=ctk.CTkFont(size=12, weight="bold"),
                    text_color="#2196f3").pack(fill="x", padx=5, pady=(15, 5))

        olcum_fields = [
            ("Olcum Cihaz Adi", ""),
            ("Olcum Kalibrasyon Tarihi", ""),
            ("Olcum Kalibrasyon Gecerlilik", ""),
            ("Olcum Seri No", ""),
            ("Olcum Kalibrasyon No", ""),
        ]

        for field, default in olcum_fields:
            row = ctk.CTkFrame(scroll, fg_color="transparent")
            row.pack(fill="x", pady=2, padx=5)
            label_text = field.replace("Olcum ", "")
            ctk.CTkLabel(row, text=label_text, width=130, anchor="w", font=ctk.CTkFont(size=10)).pack(side="left")
            entry = ctk.CTkEntry(row, width=180, font=ctk.CTkFont(size=10))
            entry.pack(side="left", fill="x", expand=True)
            if default:
                entry.insert(0, default)
            self.firma_entries[field] = entry

    def create_pano_section(self):
        """Ana pano bilgileri."""
        scroll = ScrollableFrame(self.tab_pano)
        scroll.pack(fill="both", expand=True)

        self.pano_entries = {}

        fields = [
            ("Enerji Saglayan Kurulus", "TEDAŞ", "text"),
            ("Sebeke Tipi", "TN-S", "dropdown", ["TT", "IT", "TN", "TN-CS", "TN-C", "TN-S"]),
            ("Temel Topraklama Direnci (Ohm)", "", "text"),
            ("Dis Cevrim Empedansi Z_E (Ohm)", "", "text"),
            ("Ana Kesici Tipi", "C", "dropdown", ["B", "C", "D"]),
            ("Ana Kesici Nominal Akimi", "", "text"),
            ("Ana RCD Tipi", "TOROİD", "dropdown", ["KAKR", "TOROİD"]),
            ("Ana RCD Anma Akimi (A)", "", "text"),  # mA yerine A, dropdown yerine text
            ("Ana RCD Test Akimi (mA)", "", "text"),
            ("Ana RCD Acma Suresi (ms)", "", "text"),
            ("Sistem Topraklama Kesiti (mm2)", "16", "dropdown", ["6", "10", "16", "25", "35", "50", "70", "95", "120"]),
            ("Ana Espotansiyel Kesiti (mm2)", "6", "dropdown", ["4", "6", "10", "16", "25", "35", "50"]),
        ]

        for field_data in fields:
            field = field_data[0]
            default = field_data[1]
            field_type = field_data[2] if len(field_data) > 2 else "text"

            row = ctk.CTkFrame(scroll, fg_color="transparent")
            row.pack(fill="x", pady=3, padx=5)

            ctk.CTkLabel(row, text=field[:25], width=150, anchor="w", font=ctk.CTkFont(size=11)).pack(side="left")

            if field_type == "dropdown":
                values = field_data[3]
                entry = ctk.CTkComboBox(row, values=values, width=120)
                entry.set(default)
            else:
                entry = ctk.CTkEntry(row, width=120)
                if default:
                    entry.insert(0, default)

            entry.pack(side="left")
            self.pano_entries[field] = entry

        # RCD Tipi değiştiğinde Anma Akımı alanını kontrol et
        def on_rcd_tipi_change(choice):
            anma_entry = self.pano_entries.get("Ana RCD Anma Akimi (A)")
            if anma_entry:
                if choice == "TOROİD":
                    # TOROİD seçildiğinde anma akımı alanını devre dışı bırak ve temizle
                    anma_entry.configure(state="disabled", fg_color="#2a2a2a")
                    anma_entry.delete(0, 'end')
                else:
                    # KAKR seçildiğinde aktif et
                    anma_entry.configure(state="normal", fg_color="#343638")

        # RCD Tipi combobox'una callback ekle
        rcd_tipi_combo = self.pano_entries.get("Ana RCD Tipi")
        if rcd_tipi_combo:
            rcd_tipi_combo.configure(command=on_rcd_tipi_change)
            # Başlangıçta kontrol et
            on_rcd_tipi_change(rcd_tipi_combo.get())

    def create_main_area(self):
        """Sağ panel - panolar."""
        main = ctk.CTkFrame(self, fg_color="#242424")
        main.grid(row=0, column=1, sticky="nsew", padx=5, pady=5)

        # Üst başlık
        header = ctk.CTkFrame(main, fg_color="transparent")
        header.pack(fill="x", padx=10, pady=10)

        ctk.CTkLabel(header, text="📊 Panolar", font=ctk.CTkFont(size=18, weight="bold")).pack(side="left")

        ctk.CTkButton(header, text="+ Pano Ekle", command=self.add_pano, width=120,
                     fg_color="#1976d2", hover_color="#1565c0").pack(side="right", padx=5)

        ctk.CTkButton(header, text="📋 Pano Kopyala", command=self.copy_pano_dialog, width=130,
                     fg_color="#7b1fa2", hover_color="#6a1b9a").pack(side="right", padx=5)

        # Pano listesi
        self.pano_scroll = ScrollableFrame(main)
        self.pano_scroll.pack(fill="both", expand=True, padx=10, pady=5)

    def add_pano(self):
        """Yeni pano ekle."""
        pano_frame = PanoDataFrame(
            self.pano_scroll,
            len(self.pano_frames),
            on_delete=self.delete_pano
        )
        pano_frame.pack(fill="x", pady=5)
        self.pano_frames.append(pano_frame)

    def delete_pano(self, pano_frame):
        """Panoyu sil."""
        if len(self.pano_frames) > 1:
            pano_frame.destroy()
            self.pano_frames.remove(pano_frame)
        else:
            messagebox.showwarning("Uyarı", "En az bir pano olmalı!")

    def copy_pano_dialog(self):
        """Hangi panonun kopyalanacağını seçtiren dialog."""
        if not self.pano_frames:
            messagebox.showwarning("Uyarı", "Kopyalanacak pano yok!")
            return

        dialog = ctk.CTkToplevel(self)
        dialog.title("Pano Kopyala")
        dialog.geometry("400x200")
        dialog.transient(self)
        dialog.grab_set()

        ctk.CTkLabel(dialog, text="Kopyalanacak Panoyu Seçin:",
                    font=ctk.CTkFont(size=14, weight="bold")).pack(pady=15)

        # Pano listesi
        pano_names = [f"{i+1}. {pano.get_name()}" for i, pano in enumerate(self.pano_frames)]

        selected_pano = ctk.CTkComboBox(dialog, values=pano_names, width=300)
        selected_pano.set(pano_names[0])
        selected_pano.pack(pady=10)

        # Yeni pano adı
        ctk.CTkLabel(dialog, text="Yeni Pano Adı (opsiyonel):", font=ctk.CTkFont(size=12)).pack(pady=(10, 5))
        new_name_entry = ctk.CTkEntry(dialog, width=300, placeholder_text="Boş bırakılırsa otomatik numara verilir")
        new_name_entry.pack(pady=5)

        def do_copy():
            # Seçilen panonun indeksini bul
            selected_text = selected_pano.get()
            idx = int(selected_text.split(".")[0]) - 1

            if 0 <= idx < len(self.pano_frames):
                new_name = new_name_entry.get().strip()
                self.copy_pano(self.pano_frames[idx], new_name if new_name else None)
                dialog.destroy()

        btn_frame = ctk.CTkFrame(dialog, fg_color="transparent")
        btn_frame.pack(pady=15)

        ctk.CTkButton(btn_frame, text="Kopyala", command=do_copy, width=100,
                     fg_color="#7b1fa2", hover_color="#6a1b9a").pack(side="left", padx=10)
        ctk.CTkButton(btn_frame, text="İptal", command=dialog.destroy, width=100,
                     fg_color="#757575").pack(side="left", padx=10)

    def copy_pano(self, source_pano, new_name: str = None):
        """Mevcut bir panoyu kopyalayarak yeni pano oluştur."""
        # Kaynak pano verilerini al
        source_data = source_pano.get_data()

        # Yeni pano oluştur
        new_pano = PanoDataFrame(
            self.pano_scroll,
            len(self.pano_frames),
            on_delete=self.delete_pano
        )
        new_pano.pack(fill="x", pady=5)
        self.pano_frames.append(new_pano)

        # Yeni pano adını ayarla
        if new_name:
            new_pano.name_entry.insert(0, new_name)
        else:
            new_pano.name_entry.insert(0, f"{source_data['pano_adi']} (Kopya)")

        # Verileri yeni panoya aktar
        new_pano.set_data(source_data)

        messagebox.showinfo("Başarılı", f"Pano kopyalandı: {new_pano.get_name()}")

    def get_common_data(self) -> Dict[str, Any]:
        """Ortak verileri al."""
        firma = {}
        for field, entry in self.firma_entries.items():
            firma[field] = entry.get()

        pano = {}
        for field, entry in self.pano_entries.items():
            pano[field] = entry.get()

        return {
            'firma_bilgileri': firma,
            'ana_dagitim_pano': pano
        }

    def save_project(self):
        """Projeyi kaydet."""
        file_path = filedialog.asksaveasfilename(
            defaultextension=".json",
            filetypes=[("JSON dosyası", "*.json")],
            initialfile=f"Proje_{datetime.now().strftime('%Y%m%d')}.json"
        )
        if file_path:
            project = {
                'common': self.get_common_data(),
                'rapor_prefix': self.rapor_prefix_entry.get(),
                'panolar': [pano.get_data() for pano in self.pano_frames]
            }
            with open(file_path, 'w', encoding='utf-8') as f:
                json.dump(project, f, ensure_ascii=False, indent=2)
            messagebox.showinfo("Başarılı", f"Proje kaydedildi:\n{file_path}")

    def load_project(self):
        """Projeyi yükle."""
        file_path = filedialog.askopenfilename(filetypes=[("JSON dosyası", "*.json")])
        if file_path:
            with open(file_path, 'r', encoding='utf-8') as f:
                project = json.load(f)

            # Ortak verileri yükle
            if 'common' in project:
                common = project['common']
                if 'firma_bilgileri' in common:
                    for field, value in common['firma_bilgileri'].items():
                        if field in self.firma_entries:
                            self.firma_entries[field].delete(0, 'end')
                            self.firma_entries[field].insert(0, value or '')

            # Rapor prefix yükle
            if 'rapor_prefix' in project:
                self.rapor_prefix_entry.delete(0, 'end')
                self.rapor_prefix_entry.insert(0, project['rapor_prefix'])

            messagebox.showinfo("Başarılı", "Proje yüklendi!")

    def load_sozlesme_pdf(self):
        """Hizmet sözleşmesi PDF'inden firma bilgilerini yükle ve kişiye göre cihaz bilgilerini çek."""
        file_path = filedialog.askopenfilename(
            filetypes=[("PDF dosyası", "*.pdf")],
            title="Hizmet Sözleşmesi PDF Seçin"
        )
        if not file_path:
            return

        try:
            # Conditional import
            try:
                from sozlesme_parser import parse_sozlesme_pdf
            except ImportError:
                from .sozlesme_parser import parse_sozlesme_pdf

            data = parse_sozlesme_pdf(file_path)

            # Alanları doldur
            field_mapping = {
                'Firma Adi': data.get('firma_unvan', ''),
                'Tesis Adresi': data.get('firma_adres', ''),
                'SGK Sicil No': data.get('firma_sgk_no', ''),
                'Sozlesme ID': data.get('sozlesme_id', ''),
                'Kontrol Eden': data.get('kontrol_eden_adsoyad', ''),
                'Belge No': data.get('pk_no', ''),
            }

            for field, value in field_mapping.items():
                if field in self.firma_entries and value:
                    self.firma_entries[field].delete(0, 'end')
                    self.firma_entries[field].insert(0, value)

            loaded_fields = [f for f, v in field_mapping.items() if v]

            # === CİHAZ BİLGİLERİNİ KİŞİYE GÖRE OTOMATİK ÇEK ===
            cihaz_loaded = []
            kontrol_eden = data.get('kontrol_eden_adsoyad', '')
            if kontrol_eden:
                try:
                    # Conditional import
                    try:
                        from kisi_bilgileri_reader import KisiBilgileriReader
                    except ImportError:
                        from .kisi_bilgileri_reader import KisiBilgileriReader

                    # Dosya adı varsayılanı
                    target_filename = "kisi_bilgileri.xlsx"

                    # Çalışma dizinlerini belirle
                    if getattr(sys, 'frozen', False):
                        # PyInstaller ile derlenmiş
                        base_path = sys._MEIPASS
                        exe_dir = os.path.dirname(sys.executable)
                    else:
                        # Normal python script
                        base_path = os.path.dirname(os.path.abspath(__file__))
                        exe_dir = base_path

                    # Config'den dosya adını okumaya çalış
                    try:
                        config_path = os.path.join(exe_dir, "config", "system_config.json")
                        if os.path.exists(config_path):
                            with open(config_path, "r", encoding="utf-8") as f:
                                cfg = json.load(f)
                                target_filename = cfg.get("kisi_bilgileri_dosya", target_filename)
                    except Exception as e:
                        print(f"[UYARI] Config okunamadı: {e}")

                    # Dosyayı ara (Öncelik: EXE yanı > Gömülü/Script yanı)
                    search_paths = [
                        os.path.join(exe_dir, target_filename),        # 1. EXE ile aynı dizinde (Kullanıcı verisi)
                        os.path.join(base_path, target_filename),      # 2. Gömülü/Script dizininde (Varsayılan)
                        os.path.join(exe_dir, "config", target_filename), # 3. Config klasöründe
                    ]

                    kisi_excel = None
                    for p in search_paths:
                        if os.path.exists(p):
                            kisi_excel = p
                            print(f"[INFO] Kisi bilgileri dosyasi bulundu: {p}")
                            break

                    if kisi_excel:
                        reader = KisiBilgileriReader(kisi_excel)
                        if reader.load():
                            cihaz = reader.get_cihaz_bilgileri(kontrol_eden)

                            # Cihaz alanlarını GUI'ye eşle
                            cihaz_mapping = {
                                'Termal Cihaz Adi': cihaz.get('termal_cihaz_adi', ''),
                                'Termal Kalibrasyon Tarihi': cihaz.get('termal_kalibrasyon_tarihi', ''),
                                'Termal Kalibrasyon Gecerlilik': cihaz.get('termal_kalibrasyon_gecerlilik', ''),
                                'Termal Seri No': cihaz.get('termal_seri_numarasi', ''),
                                'Termal Kalibrasyon No': cihaz.get('termal_kalibrasyon_no', ''),
                                'Olcum Cihaz Adi': cihaz.get('olcum_cihaz_adi', ''),
                                'Olcum Kalibrasyon Tarihi': cihaz.get('olcum_kalibrasyon_tarihi', ''),
                                'Olcum Kalibrasyon Gecerlilik': cihaz.get('olcum_kalibrasyon_gecerlilik', ''),
                                'Olcum Seri No': cihaz.get('olcum_seri_numarasi', ''),
                                'Olcum Kalibrasyon No': cihaz.get('olcum_kalibrasyon_no', ''),
                            }

                            for field, value in cihaz_mapping.items():
                                if field in self.firma_entries and value:
                                    self.firma_entries[field].delete(0, 'end')
                                    self.firma_entries[field].insert(0, value)
                                    cihaz_loaded.append(field)

                            if cihaz_loaded:
                                print(f"[INFO] {kontrol_eden} için {len(cihaz_loaded)} cihaz bilgisi yüklendi")
                            else:
                                print(f"[UYARI] {kontrol_eden} için cihaz bilgisi bulunamadı veya doldurulmamış")
                    else:
                        print(f"[UYARI] kisi_bilgileri.xlsx bulunamadı: {kisi_excel}")

                except ImportError as e:
                    print(f"[UYARI] kisi_bilgileri_reader modülü yüklenemedi: {e}")
                except Exception as e:
                    print(f"[HATA] Cihaz bilgileri okunurken hata: {e}")

            # Özet mesaj
            if loaded_fields or cihaz_loaded:
                msg = "PDF'den yüklenen alanlar:\n" + "\n".join(f"✓ {f}" for f in loaded_fields)
                if cihaz_loaded:
                    msg += f"\n\n📷🔌 {kontrol_eden} için cihaz bilgileri:\n"
                    msg += "\n".join(f"✓ {f}" for f in cihaz_loaded)
                messagebox.showinfo("Başarılı", msg)
            else:
                messagebox.showwarning("Uyarı", "PDF'den veri çıkarılamadı!")

        except ImportError:
            messagebox.showerror("Hata", "pypdf modülü yüklü değil!\npip install pypdf")
        except Exception as e:
            messagebox.showerror("Hata", f"PDF okuma hatası:\n{str(e)}")

    def start_report_generation(self):
        """Rapor oluşturma işlemini başlatır - Veriyi hazırla sonra thread başlat."""
        output_dir = filedialog.askdirectory(title="Raporların kaydedileceği klasörü seçin")
        if not output_dir:
            return

        # Verileri ana thread'de topla
        try:
            common = self.get_common_data()
            rapor_prefix = self.rapor_prefix_entry.get() or "TPK2025-0001"
            panos_data = [pano.get_data() for pano in self.pano_frames]
        except Exception as e:
            messagebox.showerror("Hata", f"Veri okunurken hata oluştu: {e}")
            return

        # UI güncelleme
        self.generate_btn.configure(state="disabled", text="Raporlar Oluşturuluyor...")
        self.progress_bar.pack(fill="x", padx=10, pady=5)
        self.progress_bar.start()

        # Thread başlat
        thread = threading.Thread(target=self.generate_all_reports_thread,
                                  args=(output_dir, common, rapor_prefix, panos_data),
                                  daemon=True)
        thread.start()

    def generate_all_reports_thread(self, output_dir, common, rapor_prefix, panos_data):
        """Thread içinde rapor oluşturma mantığı."""
        try:
            generated = []

            # ReportGenerator import - already imported at top but safety check
            try:
                # Modül olarak çalıştırılıyorsa
                from report_generator import ReportGenerator, resolve_template_path
            except ImportError:
                from .report_generator import ReportGenerator, resolve_template_path

            try:
                template_path = resolve_template_path()
            except FileNotFoundError as e:
                self.after(0, lambda: messagebox.showerror("Şablon bulunamadı", str(e)))
                return

            generator = ReportGenerator(template_path)

            for i, pano_data in enumerate(panos_data, 1):
                pano_name = pano_data['pano_adi']
                rapor_no = f"{rapor_prefix}-{i}"

                rapor_no_for_file = rapor_no
                if rapor_no.upper().startswith("TPK"):
                    rapor_no_for_file = rapor_no[3:]

                safe_pano_name = "".join(c if c.isalnum() or c in "_ -" else "_" for c in pano_name)
                output_path = os.path.join(output_dir, f"{rapor_no_for_file} {safe_pano_name} Elektrik Tesisat PK.docx")

                firma_with_rapor = dict(common['firma_bilgileri'])
                firma_with_rapor['Rapor No'] = rapor_no

                pano_overrides = dict(pano_data.get('ana_pano_overrides', {}))
                pano_overrides.setdefault('Pano Adi (PANO_adi1)', pano_name)

                # Hesaplamalar
                ik3_raw = pano_overrides.get('Ik3 (kA)', '')
                zln_raw = pano_overrides.get('Faz-Notr Cevrim Empedansi Z_ln (Ohm)', '')
                if (not str(ik3_raw).strip()) and str(zln_raw).strip():
                    try:
                        zln_val = float(str(zln_raw).replace(',', '.'))
                        if zln_val != 0:
                            pano_overrides['Ik3 (kA)'] = f"{int(round(380.0 / zln_val))}"
                    except ValueError:
                        pass

                full_data = {
                    'firma_bilgileri': firma_with_rapor,
                    'ana_dagitim_pano': {
                        **common.get('ana_dagitim_pano', {}),
                        **pano_overrides,
                    },
                    'gozle_kontrol': pano_data['gozle_kontrol'],
                    'fonksiyon_testleri': pano_data['fonksiyon_testleri'],
                    'termal_goruntuler': pano_data['termal_goruntuler']
                }

                generator.generate(full_data, output_path)
                generated.append(pano_name)

            # Başarılı bitiş
            def on_success():
                messagebox.showinfo(
                    "Başarılı",
                    f"{len(generated)} rapor oluşturuldu:\n" +
                    "\n".join(f"✓ {name}" for name in generated) +
                    f"\n\nKonum: {output_dir}"
                )
                os.startfile(output_dir)

            self.after(0, on_success)

        except Exception as e:
            self.after(0, lambda: messagebox.showerror("Hata", f"Rapor oluşturulurken hata:\n{str(e)}"))
            import traceback
            traceback.print_exc()
        finally:
            self.after(0, self.finish_generation)

    def finish_generation(self):
        """İşlem bitince UI'ı eski haline getir."""
        self.progress_bar.stop()
        self.progress_bar.pack_forget()
        self.generate_btn.configure(state="normal", text="📋 TOPLU RAPOR OLUŞTUR")


def main():
    app = MultiPanoApp()
    app.mainloop()


if __name__ == "__main__":
    main()
