import customtkinter as ctk
import random
from typing import Dict, Any, List, Optional
from rapor_sistemi.gui.utils import ScrollableFrame
from rapor_sistemi.gui.drag_drop import DragDropFrame
import rapor_sistemi.constants as const

class PanoDataFrame(ctk.CTkFrame):
    """Tek bir pano için veri girişi alanı."""

    def __init__(self, master, pano_index: int, on_delete=None, **kwargs):
        super().__init__(master, **kwargs)
        self.pano_index = pano_index
        self.on_delete = on_delete

        self.configure(fg_color="#1a1a1a", corner_radius=10)

        # Başlık ve silme butonu
        header = ctk.CTkFrame(self, fg_color="transparent")
        header.pack(fill="x", padx=10, pady=5)

        self.name_entry = ctk.CTkEntry(header, width=200, placeholder_text=f"Pano {pano_index + 1} Adı")
        self.name_entry.pack(side="left")

        if on_delete:
            ctk.CTkButton(header, text="🗑️", width=30, fg_color="#c62828", hover_color="#b71c1c",
                         command=lambda: on_delete(self)).pack(side="right")

        # Sekmeler
        # Daha fazla dikey alan için yüksekliği artırıldı
        self.tabview = ctk.CTkTabview(self, height=650)
        self.tabview.pack(fill="both", expand=True, padx=5, pady=5)

        self.tab_gk = self.tabview.add("Gözle Kontrol")
        self.tab_ft = self.tabview.add("Fonksiyon Testleri")
        self.tab_termal = self.tabview.add("Termal")

        # Parafudr bilgileri (her pano için)
        parafudr_frame = ctk.CTkFrame(self, fg_color="#1f1f1f")
        parafudr_frame.pack(fill="x", padx=10, pady=(0,5))
        ctk.CTkLabel(parafudr_frame, text="Parafudr Tipi", width=120, anchor="w", font=ctk.CTkFont(size=11)).pack(side="left", padx=(5,2), pady=4)
        self.parafudr_tip_entry = ctk.CTkEntry(parafudr_frame, width=140, placeholder_text="Örn: T1+T2")
        self.parafudr_tip_entry.pack(side="left", padx=4)
        ctk.CTkLabel(parafudr_frame, text="Parafudr Imax (kA)", width=140, anchor="w", font=ctk.CTkFont(size=11)).pack(side="left", padx=(12,2), pady=4)
        self.parafudr_imax_entry = ctk.CTkEntry(parafudr_frame, width=100, placeholder_text="Örn: 40")
        self.parafudr_imax_entry.pack(side="left", padx=4)

        # Döngü empedansı ve gerilim ölçümleri (her pano için)
        loop_frame = ctk.CTkFrame(self, fg_color="#1f1f1f")
        loop_frame.pack(fill="x", padx=10, pady=(0,6))
        ctk.CTkLabel(loop_frame, text="Zx (Ω)", width=80, anchor="w", font=ctk.CTkFont(size=11)).pack(side="left", padx=(5,2), pady=4)
        self.zx_entry = ctk.CTkEntry(loop_frame, width=80, placeholder_text="" )
        self.zx_entry.pack(side="left", padx=3)
        ctk.CTkLabel(loop_frame, text="Zln (Ω)", width=80, anchor="w", font=ctk.CTkFont(size=11)).pack(side="left", padx=(8,2), pady=4)
        self.zln_entry = ctk.CTkEntry(loop_frame, width=80, placeholder_text="" )
        self.zln_entry.pack(side="left", padx=3)
        ctk.CTkLabel(loop_frame, text="F-F (V)", width=80, anchor="w", font=ctk.CTkFont(size=11)).pack(side="left", padx=(10,2), pady=4)
        self.ff_entry = ctk.CTkEntry(loop_frame, width=80, placeholder_text="" )
        self.ff_entry.pack(side="left", padx=3)
        ctk.CTkLabel(loop_frame, text="L-N (V)", width=80, anchor="w", font=ctk.CTkFont(size=11)).pack(side="left", padx=(8,2), pady=4)
        self.ln_entry = ctk.CTkEntry(loop_frame, width=80, placeholder_text="" )
        self.ln_entry.pack(side="left", padx=3)
        ctk.CTkLabel(loop_frame, text="N-PE (V)", width=90, anchor="w", font=ctk.CTkFont(size=11)).pack(side="left", padx=(8,2), pady=4)
        self.npe_entry = ctk.CTkEntry(loop_frame, width=80, placeholder_text="" )
        self.npe_entry.pack(side="left", padx=3)
        ctk.CTkLabel(loop_frame, text="Ik3 (kA)", width=90, anchor="w", font=ctk.CTkFont(size=11)).pack(side="left", padx=(10,2), pady=4)
        self.ik3_entry = ctk.CTkEntry(loop_frame, width=80, placeholder_text="380/Zln")
        self.ik3_entry.pack(side="left", padx=3)

        # Ik3 otomatik: 380 / Zln
        def update_ik3(event=None):
            val = self.zln_entry.get().strip().replace(',', '.')
            ik3_text = ""
            try:
                zln_val = float(val)
                if zln_val != 0:
                    ik3_text = f"{int(round(380.0 / zln_val))}"
            except ValueError:
                pass
            self.ik3_entry.configure(state='normal')
            self.ik3_entry.delete(0, 'end')
            if ik3_text:
                self.ik3_entry.insert(0, ik3_text)
            self.ik3_entry.configure(state='readonly')

        self.zln_entry.bind("<KeyRelease>", update_ik3)
        self.zln_entry.bind("<FocusOut>", update_ik3)
        update_ik3()

        # Uygunluk seçimi (her pano için ayrı)
        uygunluk_frame = ctk.CTkFrame(self, fg_color="#1f1f1f")
        uygunluk_frame.pack(fill="x", padx=10, pady=(0,6))
        ctk.CTkLabel(uygunluk_frame, text="📋 Sonuç:", width=80, anchor="w",
                    font=ctk.CTkFont(size=12, weight="bold"), text_color="#4caf50").pack(side="left", padx=(5,2), pady=4)
        self.uygunluk_combo = ctk.CTkComboBox(uygunluk_frame, values=["Uygun", "Uygun Değil"], width=150)
        self.uygunluk_combo.set("Uygun")
        self.uygunluk_combo.pack(side="left", padx=10)

        self.create_gk_tab()
        self.create_ft_tab()
        self.create_termal_tab()

    def create_gk_tab(self):
        """Gözle kontrol sekmesi - 2 sütunlu kompakt layout."""
        scroll = ScrollableFrame(self.tab_gk)
        scroll.pack(fill="both", expand=True)

        self.gk_entries = {}

        fields = const.GK_FIELDS

        # Sabit değer sınıfı
        class FixedValue:
            def __init__(self, val):
                self._val = val
            def get(self):
                return self._val

        # 2'li gruplar halinde işle
        for i in range(0, len(fields), 2):
            row = ctk.CTkFrame(scroll, fg_color="transparent")
            row.pack(fill="x", pady=1, padx=5)

            # Sol madde
            left_field = fields[i]
            ctk.CTkLabel(row, text=left_field[:28], width=170, anchor="w", font=ctk.CTkFont(size=10)).pack(side="left")

            if left_field == "Tesisat Yontemi":
                label = ctk.CTkLabel(row, text="A1", width=90, font=ctk.CTkFont(size=10), text_color="#4CAF50")
                label.pack(side="left", padx=2)
                self.gk_entries[left_field] = FixedValue("A1")
            else:
                combo = ctk.CTkComboBox(row, values=["Uygun", "Uygun Değil", "Uygulanamaz"], width=90, font=ctk.CTkFont(size=10))
                combo.set("Uygun")
                combo.pack(side="left", padx=2)
                self.gk_entries[left_field] = combo

            # Sağ madde (varsa)
            if i + 1 < len(fields):
                right_field = fields[i + 1]
                ctk.CTkLabel(row, text=right_field[:28], width=170, anchor="w", font=ctk.CTkFont(size=10)).pack(side="left", padx=(15, 0))

                if right_field == "Tesisat Yontemi":
                    label = ctk.CTkLabel(row, text="A1", width=90, font=ctk.CTkFont(size=10), text_color="#4CAF50")
                    label.pack(side="left", padx=2)
                    self.gk_entries[right_field] = FixedValue("A1")
                else:
                    combo = ctk.CTkComboBox(row, values=["Uygun", "Uygun Değil", "Uygulanamaz"], width=90, font=ctk.CTkFont(size=10))
                    combo.set("Uygun")
                    combo.pack(side="left", padx=2)
                    self.gk_entries[right_field] = combo

    def create_ft_tab(self):
        """Fonksiyon testleri sekmesi."""
        # Üst butonlar
        btn_frame = ctk.CTkFrame(self.tab_ft, fg_color="transparent")
        btn_frame.pack(fill="x", padx=5, pady=5)

        ctk.CTkButton(btn_frame, text="+ Satır", command=self.add_ft_row, width=90, height=36).pack(side="left", padx=4)
        ctk.CTkButton(btn_frame, text="- Sil", command=self.remove_ft_row, width=90, height=36).pack(side="left", padx=4)
        ctk.CTkButton(btn_frame, text="++ Linye Grubu", command=self.add_multiple_ft_rows, width=110, height=36, fg_color="#2E7D32", hover_color="#1B5E20").pack(side="left", padx=4)

        # Başlıklar
        header_frame = ctk.CTkFrame(self.tab_ft, fg_color="transparent", height=40)
        header_frame.pack(fill="x", padx=8, pady=(8,2))

        # Sütun genişlikleri
        headers = [
            ("", 36), ("Linye Adı", 130), ("Eğri", 60), ("Kutup", 60),
            ("In", 60), ("Icu", 70), ("Ib", 60), ("Faz", 60), ("Nötr", 60), ("Toprak", 60), ("Iz", 70), ("Sonuç", 90),
            ("KAKR", 50), ("IΔn", 70), ("mA", 60), ("mS", 60), ("KAKR Yok", 80)
        ]

        header_font = ctk.CTkFont(size=11, weight="bold")
        for text, width in headers:
            ctk.CTkLabel(header_frame, text=text, width=width, font=header_font).pack(side="left", padx=3)

        self.ft_scroll = ScrollableFrame(self.tab_ft)
        self.ft_scroll.pack(fill="both", expand=True, padx=8, pady=8)

        self.ft_rows = []

        # İlk 3 satır
        for _ in range(3):
            self.add_ft_row()

    def add_ft_row(self):
        row_num = len(self.ft_rows) + 1
        row_height = 44
        entry_height = 36
        combo_height = 34  # Ok ikonunu küçültmek için combobox yüksekliğini düşürdük
        row = ctk.CTkFrame(self.ft_scroll, fg_color="#2b2b2b", height=row_height)
        row.pack(fill="x", pady=3, padx=8)
        row.pack_propagate(False)

        entry_font = ctk.CTkFont(size=11)
        compact_check_height = combo_height - 4
        compact_check_width = 50

        entries = {}

        # --- Helper Fonksiyonlar ---
        def update_ib(choice=None):
            try:
                # In değerini al (A) - hem combobox'tan hem entry'den
                in_str = entries['In (A)'].get() if entries.get('In (A)') else choice
                if not in_str:
                    return
                in_val = float(in_str)
                # Ib = In * 0.7 formülü
                ib_val = in_val * 0.7
                # Ib alanını güncelle
                entries['Ib'].configure(state='normal')
                entries['Ib'].delete(0, 'end')
                entries['Ib'].insert(0, f"{ib_val:.1f}")
                entries['Ib'].configure(state='readonly')

                # Otomatik kablo kesiti seç
                in_int = int(in_val)
                # En yakın değeri bul
                kesit = None
                for k in sorted(const.IN_TO_KESIT.keys()):
                    if in_int <= k:
                        kesit = const.IN_TO_KESIT[k]
                        break
                if kesit is None:
                    kesit = "240"  # En büyük

                entries['Faz Kesiti'].set(kesit)
                entries['Notr Kesiti'].set(kesit)
                entries['Toprak Kesiti'].set(kesit)

                # Iz'yi de güncelle
                if kesit in const.KESIT_TO_IZ:
                    entries['Iz'].delete(0, 'end')
                    entries['Iz'].insert(0, str(const.KESIT_TO_IZ[kesit]))
            except ValueError:
                pass

        def on_kakr_check():
            if kakr_var.get():
                # KAKR var -> Alanları aç
                entries['RCD Acma'].configure(state='normal')
                entries['RCD mA'].configure(state='normal')
                entries['RCD ms'].configure(state='normal')
                # Varsayılan değer ata
                if not entries['RCD Acma'].get():
                    entries['RCD Acma'].set("30mA")
                    generate_rcd_values("30mA")
            else:
                # KAKR yok -> Alanları temizle ve kapat
                entries['RCD Acma'].set("")
                entries['RCD Acma'].configure(state='disabled')

                entries['RCD mA'].delete(0, 'end')
                entries['RCD mA'].configure(state='disabled')

                entries['RCD ms'].delete(0, 'end')
                entries['RCD ms'].configure(state='disabled')

        def generate_rcd_values(choice):
            if not kakr_var.get():
                return

            # Eğer 30mA seçildiyse otomatik değer üret
            if choice == "30mA":
                # mA: 20-30 arası
                ma_val = random.uniform(20.0, 30.0)
                # mS: 17-40 arası
                ms_val = random.uniform(17.0, 40.0)

                entries['RCD mA'].delete(0, 'end')
                entries['RCD mA'].insert(0, f"{ma_val:.0f}")

                entries['RCD ms'].delete(0, 'end')
                entries['RCD ms'].insert(0, f"{ms_val:.0f}")
            elif choice == "300mA":
                 # 300mA: mA 250-290 arası, birler basamağı 0 olacak şekilde; mS 20-40 arası
                 ma_val = random.choice(range(250, 291, 10))
                 ms_val = random.uniform(20.0, 40.0)
                 entries['RCD mA'].delete(0, 'end')
                 entries['RCD mA'].insert(0, f"{ma_val:.0f}")

                 entries['RCD ms'].delete(0, 'end')
                 entries['RCD ms'].insert(0, f"{ms_val:.0f}")

        # --- Widgetlar ---
        ctk.CTkLabel(row, text=str(row_num), width=36, height=entry_height, font=entry_font).pack(side="left", padx=3)

        # Linye Adı
        e = ctk.CTkEntry(row, width=130, height=entry_height, font=entry_font, placeholder_text="Linye")
        e.pack(side="left", padx=3)
        entries['Linye Adi'] = e

        # Eğri - AAA dahil
        e = ctk.CTkComboBox(row, values=["B", "C", "D", "K", "Z", "AAA"], width=60, height=combo_height, font=entry_font)
        e.set("C")
        e.pack(side="left", padx=3)
        entries['Acma Egrisi'] = e

        # Kutup
        e = ctk.CTkComboBox(row, values=["1", "2", "3", "4"], width=60, height=combo_height, font=entry_font)
        e.set("1")
        e.pack(side="left", padx=3)
        entries['Kutup Sayisi'] = e

        # In (A) - Daha geniş değer listesi + manuel giriş
        e = ctk.CTkComboBox(row, values=const.IN_VALUES, width=70, height=combo_height, font=entry_font, command=update_ib)
        e.set("16")  # Varsayılan değer
        e.pack(side="left", padx=3)
        entries['In (A)'] = e
        # Manuel giriş için Enter tuşuna basınca güncelle
        e.bind('<Return>', lambda event: update_ib())
        e.bind('<FocusOut>', lambda event: update_ib())

        # Icu (kA) - listeden seç veya yaz
        e = ctk.CTkComboBox(row, values=const.ICU_VALUES, width=70, height=combo_height, font=entry_font)
        e.set(const.ICU_VALUES[0])
        e.pack(side="left", padx=3)
        entries['Icu'] = e

        # Ib (A) - Readonly
        e = ctk.CTkEntry(row, width=60, height=entry_height, font=entry_font)
        e.configure(state='readonly')
        e.pack(side="left", padx=3)
        entries['Ib'] = e

        # Kesitler (Faz, Nötr, Toprak)

        def calc_iz_from_section(section_val: str) -> str:
            val = (section_val or "").lower().replace(",", ".").strip()
            if not val:
                return ""
            factor = 1
            base = val
            if "x" in val:
                try:
                    f_str, base = val.split("x", 1)
                    factor = int(f_str)
                except ValueError:
                    factor = 1
            try:
                size = float(base)
            except ValueError:
                return ""
            base_iz = const.IZ_TABLE.get(size)
            if base_iz is None:
                return ""
            return f"{base_iz * factor:.0f}"

        def update_iz_from_faz(choice=None):
            iz_val = calc_iz_from_section(entries['Faz Kesiti'].get())
            entries['Iz'].delete(0, 'end')
            if iz_val:
                entries['Iz'].insert(0, iz_val)

        e_faz = ctk.CTkComboBox(row, values=const.ALL_SECTIONS, width=60, height=combo_height, font=entry_font, command=update_iz_from_faz)
        e_faz.pack(side="left", padx=3)
        entries['Faz Kesiti'] = e_faz

        e_notr = ctk.CTkComboBox(row, values=const.ALL_SECTIONS, width=60, height=combo_height, font=entry_font)
        e_notr.pack(side="left", padx=3)
        entries['Notr Kesiti'] = e_notr

        e_toprak = ctk.CTkComboBox(row, values=const.ALL_SECTIONS, width=60, height=combo_height, font=entry_font)
        e_toprak.pack(side="left", padx=3)
        entries['Toprak Kesiti'] = e_toprak

        # Iz (A) - akım taşıma kapasitesi (otomatik dolar)
        e = ctk.CTkEntry(row, width=70, height=entry_height, font=entry_font, placeholder_text="Iz")
        e.pack(side="left", padx=3)
        entries['Iz'] = e

        # Sonuç
        e = ctk.CTkComboBox(row, values=["Uygun", "Uygun Değil"], width=90, height=combo_height, font=entry_font)
        e.set("Uygun")
        e.pack(side="left", padx=4)
        entries['Sonuc'] = e

        # KAKR (Checkbox)
        kakr_var = ctk.BooleanVar(value=False)
        kakr_cb = ctk.CTkCheckBox(row, text="", variable=kakr_var, width=compact_check_width, height=compact_check_height, command=on_kakr_check)
        kakr_cb.pack(side="left", padx=4)
        entries['KAKR Var'] = kakr_var # BooleanVar saklıyoruz
        entries['KAKR_Checkbox'] = kakr_cb  # Checkbox widget'ı

        # RCD Açma Akımı
        e = ctk.CTkComboBox(row, values=["30mA", "300mA"], width=70, height=combo_height, font=entry_font, command=generate_rcd_values)
        e.set("")
        e.configure(state='disabled')
        e.pack(side="left", padx=3)
        entries['RCD Acma'] = e

        # Ölçülen mA
        e = ctk.CTkEntry(row, width=60, height=entry_height, font=entry_font)
        e.configure(state='disabled')
        e.pack(side="left", padx=3)
        entries['RCD mA'] = e

        # Ölçülen mS
        e = ctk.CTkEntry(row, width=60, height=entry_height, font=entry_font)
        e.configure(state='disabled')
        e.pack(side="left", padx=3)
        entries['RCD ms'] = e

        # KAKR Yok (checkbox) -> devralmayı kırar, mA/mS boş, Sonuç=Uygun Değil
        kakr_none_var = ctk.BooleanVar(value=False)
        e = ctk.CTkCheckBox(row, text="", variable=kakr_none_var, width=70, height=compact_check_height)
        e.pack(side="left", padx=6)
        entries['KAKR Yok'] = kakr_none_var

        self.ft_rows.append({'frame': row, 'entries': entries})

        # Varsayılan In değerine göre Ib'yi hesapla
        update_ib()

    def remove_ft_row(self):
        if self.ft_rows:
            row = self.ft_rows.pop()
            row['frame'].destroy()

    def add_multiple_ft_rows(self):
        """Linye grubu ekleme dialog'u."""
        dialog = ctk.CTkToplevel(self)
        dialog.title("Linye Grubu Ekle")
        dialog.geometry("500x620")
        dialog.transient(self)
        dialog.grab_set()

        # Ana frame - scrollable
        main_frame = ctk.CTkScrollableFrame(dialog, fg_color="transparent")
        main_frame.pack(fill="both", expand=True, padx=15, pady=15)

        # Adet
        ctk.CTkLabel(main_frame, text="Kaç Adet Eklenecek:", font=ctk.CTkFont(size=12, weight="bold")).pack(anchor="w", pady=(0,5))
        adet_entry = ctk.CTkEntry(main_frame, width=100, placeholder_text="10")
        adet_entry.pack(anchor="w", pady=(0,10))
        adet_entry.insert(0, "10")

        # Linye Adı Şablonu
        ctk.CTkLabel(main_frame, text="Linye Adı Şablonu:", font=ctk.CTkFont(size=12, weight="bold")).pack(anchor="w", pady=(0,5))
        ctk.CTkLabel(main_frame, text="(Sondaki sayı otomatik artar)", font=ctk.CTkFont(size=10), text_color="#888").pack(anchor="w")
        linye_entry = ctk.CTkEntry(main_frame, width=250, placeholder_text="Aydınlatma 1")
        linye_entry.pack(anchor="w", pady=(0,10))
        linye_entry.insert(0, "Aydınlatma 1")

        # Eğri
        ctk.CTkLabel(main_frame, text="Açma Eğrisi:", font=ctk.CTkFont(size=12, weight="bold")).pack(anchor="w", pady=(0,5))
        egri_combo = ctk.CTkComboBox(main_frame, values=["B", "C", "D", "K", "Z"], width=100)
        egri_combo.set("C")
        egri_combo.pack(anchor="w", pady=(0,10))

        # Kutup
        ctk.CTkLabel(main_frame, text="Kutup Sayısı:", font=ctk.CTkFont(size=12, weight="bold")).pack(anchor="w", pady=(0,5))
        kutup_combo = ctk.CTkComboBox(main_frame, values=["1", "2", "3", "4"], width=100)
        kutup_combo.set("1")
        kutup_combo.pack(anchor="w", pady=(0,10))

        # In değerine göre önerilen minimum kablo kesiti (Grup 2)
        # Using string keys for compatibility with the dialog logic which expects string matching from constant if keys match
        # However, the constant IN_TO_KESIT uses integers. Let's create a local mapping or convert.
        # The original code redefined it here with string keys.

        in_to_kesit_dialog = {str(k): v for k, v in const.IN_TO_KESIT.items()}

        def on_in_change(choice=None):
            in_val = in_combo.get() if choice is None else choice
            if in_val in in_to_kesit_dialog:
                faz_combo.set(in_to_kesit_dialog[in_val])

        # In (A) - Genişletilmiş liste + manuel giriş
        ctk.CTkLabel(main_frame, text="In (A):", font=ctk.CTkFont(size=12, weight="bold")).pack(anchor="w", pady=(0,5))
        in_combo = ctk.CTkComboBox(main_frame, values=const.IN_VALUES, width=100, command=on_in_change)
        in_combo.set("16")
        in_combo.pack(anchor="w", pady=(0,10))
        in_combo.bind('<Return>', lambda e: on_in_change())
        in_combo.bind('<FocusOut>', lambda e: on_in_change())

        # Faz Kesiti - varsayılan 2.5 (16A için)
        ctk.CTkLabel(main_frame, text="Faz Kesiti (mm²):", font=ctk.CTkFont(size=12, weight="bold")).pack(anchor="w", pady=(0,5))
        faz_combo = ctk.CTkComboBox(main_frame, values=const.ALL_SECTIONS, width=100)
        faz_combo.set("2.5")
        faz_combo.pack(anchor="w", pady=(0,10))

        # KAKR Checkbox
        kakr_var = ctk.BooleanVar(value=True)
        kakr_check = ctk.CTkCheckBox(main_frame, text="KAKR Var (30mA)", variable=kakr_var)
        kakr_check.pack(anchor="w", pady=(0,10))

        # RCD mA değeri (sabit)
        ctk.CTkLabel(main_frame, text="RCD mA Değeri:", font=ctk.CTkFont(size=12, weight="bold")).pack(anchor="w", pady=(0,5))
        rcd_ma_entry = ctk.CTkEntry(main_frame, width=100, placeholder_text="25")
        rcd_ma_entry.pack(anchor="w", pady=(0,10))
        rcd_ma_entry.insert(0, "25")

        # RCD mS değeri (sabit)
        ctk.CTkLabel(main_frame, text="RCD mS Değeri:", font=ctk.CTkFont(size=12, weight="bold")).pack(anchor="w", pady=(0,10))
        rcd_ms_entry = ctk.CTkEntry(main_frame, width=100, placeholder_text="20")
        rcd_ms_entry.pack(anchor="w", pady=(0,10))
        rcd_ms_entry.insert(0, "20")

        def apply_multiple():
            try:
                adet = int(adet_entry.get())
            except ValueError:
                adet = 1

            if adet < 1 or adet > 100:
                adet = 10

            linye_template = linye_entry.get() or "Linye {n}"
            egri = egri_combo.get()
            kutup = kutup_combo.get()
            in_val = in_combo.get()
            faz_kesit = faz_combo.get()
            kakr = kakr_var.get()
            rcd_ma = rcd_ma_entry.get() or "25"
            rcd_ms = rcd_ms_entry.get() or "20"

            # KAKR seçiliyse önce KAKR linyesini ekle (grubun başında)
            if kakr:
                self.add_ft_row()
                kakr_row = self.ft_rows[-1]['entries']
                kakr_row['Linye Adi'].delete(0, 'end')
                kakr_row['Linye Adi'].insert(0, "KAKR")
                kakr_row['Acma Egrisi'].set("AAA")
                kakr_row['Kutup Sayisi'].set("4")
                kakr_row['In (A)'].set("40")
                # Ib hesapla
                try:
                    ib_val = 40 * 0.7
                    kakr_row['Ib'].configure(state='normal')
                    kakr_row['Ib'].delete(0, 'end')
                    kakr_row['Ib'].insert(0, f"{ib_val:.1f}")
                    kakr_row['Ib'].configure(state='readonly')
                except:
                    pass
                kakr_row['Faz Kesiti'].set("2.5")
                # KAKR linyesine de aynı mA ve mS değerlerini ata
                kakr_row['KAKR Var'].set(True)
                kakr_row['KAKR_Checkbox'].select()
                kakr_row['RCD Acma'].configure(state='normal')
                kakr_row['RCD Acma'].set("30mA")
                kakr_row['RCD mA'].configure(state='normal')
                kakr_row['RCD ms'].configure(state='normal')
                kakr_row['RCD mA'].delete(0, 'end')
                kakr_row['RCD mA'].insert(0, rcd_ma)
                kakr_row['RCD ms'].delete(0, 'end')
                kakr_row['RCD ms'].insert(0, rcd_ms)

            # Mevcut satır sayısını al (numaralandırma için)
            start_num = len(self.ft_rows) + 1

            for i in range(adet):
                # Yeni satır ekle
                self.add_ft_row()

                # Son eklenen satıra değerleri ata
                row_entries = self.ft_rows[-1]['entries']

                # Linye adı - sondaki sayıyı artır veya {n} kullan
                import re
                # Sondaki sayıyı bul (örn: "1.Sıra Sigorta-1" → base="1.Sıra Sigorta-", num=1)
                match = re.match(r'^(.*?)(\d+)$', linye_template)
                if match:
                    base_name = match.group(1)
                    start_num_from_name = int(match.group(2))
                    linye_name = f"{base_name}{start_num_from_name + i}"
                elif "{n}" in linye_template:
                    linye_name = linye_template.replace("{n}", str(start_num + i))
                else:
                    linye_name = f"{linye_template} {i + 1}"
                row_entries['Linye Adi'].delete(0, 'end')
                row_entries['Linye Adi'].insert(0, linye_name)

                # Eğri
                row_entries['Acma Egrisi'].set(egri)

                # Kutup
                row_entries['Kutup Sayisi'].set(kutup)

                # In - bu Ib'yi de otomatik hesaplayacak
                row_entries['In (A)'].set(in_val)
                # Ib hesapla
                try:
                    ib_val = float(in_val) * 0.7
                    row_entries['Ib'].configure(state='normal')
                    row_entries['Ib'].delete(0, 'end')
                    row_entries['Ib'].insert(0, f"{ib_val:.1f}")
                    row_entries['Ib'].configure(state='readonly')
                except:
                    pass

                # Faz Kesiti
                row_entries['Faz Kesiti'].set(faz_kesit)

                # KAKR
                if kakr:
                    row_entries['KAKR Var'].set(True)
                    row_entries['KAKR_Checkbox'].select()
                    row_entries['RCD Acma'].configure(state='normal')
                    row_entries['RCD Acma'].set("30mA")
                    row_entries['RCD mA'].configure(state='normal')
                    row_entries['RCD ms'].configure(state='normal')
                    # Sabit RCD değerleri
                    row_entries['RCD mA'].delete(0, 'end')
                    row_entries['RCD mA'].insert(0, rcd_ma)
                    row_entries['RCD ms'].delete(0, 'end')
                    row_entries['RCD ms'].insert(0, rcd_ms)

            dialog.destroy()

        # Butonlar
        btn_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
        btn_frame.pack(fill="x", pady=(15, 0))

        ctk.CTkButton(btn_frame, text="Ekle", command=apply_multiple, width=100, fg_color="#2E7D32", hover_color="#1B5E20").pack(side="left", padx=5)
        ctk.CTkButton(btn_frame, text="İptal", command=dialog.destroy, width=100, fg_color="#757575").pack(side="left", padx=5)

    def create_termal_tab(self):
        """Termal görüntüler sekmesi."""
        self.drop_zone = DragDropFrame(self.tab_termal, height=150)
        self.drop_zone.pack(fill="x", padx=10, pady=10)

        ctk.CTkButton(self.tab_termal, text="Temizle", command=self.drop_zone.clear, width=100).pack(pady=5)

    def get_data(self) -> Dict[str, Any]:
        """Pano verilerini al."""
        # Gözle kontrol
        gk_data = {}
        for field, entry in self.gk_entries.items():
            gk_data[field] = entry.get()

        # Zln varsa Ik3'ü (380/Zln) yeniden hesapla (GUI event kaçarsa yakala)
        try:
            val = self.zln_entry.get().strip().replace(',', '.')
            zln_val = float(val)
            if zln_val != 0:
                self.ik3_entry.configure(state='normal')
                self.ik3_entry.delete(0, 'end')
                self.ik3_entry.insert(0, f"{int(round(380.0 / zln_val))}")
                self.ik3_entry.configure(state='readonly')
        except Exception:
            pass

        # Fonksiyon testleri
        ft_data = []
        for row in self.ft_rows:
            row_data = {}
            has_data = False
            for field, entry in row['entries'].items():
                val = entry.get()
                row_data[field] = val
                if val:
                    has_data = True
            if has_data:
                ft_data.append(row_data)

        # DEBUG: Termal dosya yollarını kontrol et
        termal_files = self.drop_zone.get_files()
        print(f"DEBUG [PanoDataFrame.get_data]: drop_zone.get_files() = {termal_files}")

        return {
            'pano_adi': self.name_entry.get() or f"Pano {self.pano_index + 1}",
            'gozle_kontrol': {'kontroller': gk_data, 'pano_adi': self.name_entry.get()},
            'fonksiyon_testleri': ft_data,
            'termal_goruntuler': [{'fluke_dosya': f} for f in termal_files],
            'ana_pano_overrides': {
                'Parafudr Tipi': self.parafudr_tip_entry.get(),
                'Parafudr Imax (kA)': self.parafudr_imax_entry.get(),
                'Faz-Toprak Cevrim Empedansi Z_x (Ohm)': self.zx_entry.get(),
                'Faz-Notr Cevrim Empedansi Z_ln (Ohm)': self.zln_entry.get(),
                'Gerilim F-F (V)': self.ff_entry.get(),
                'Gerilim L-N (V)': self.ln_entry.get(),
                'Gerilim N-PE (V)': self.npe_entry.get(),
                'Ik3 (kA)': self.ik3_entry.get(),
                'Uygunluk': self.uygunluk_combo.get(),
            },
        }

    def set_data(self, data: Dict[str, Any]):
        """Pano verilerini yükle (kopyalama için)."""
        # Gözle kontrol verilerini yükle
        if 'gozle_kontrol' in data and 'kontroller' in data['gozle_kontrol']:
            kontroller = data['gozle_kontrol']['kontroller']
            for field, value in kontroller.items():
                if field in self.gk_entries:
                    entry = self.gk_entries[field]
                    if hasattr(entry, 'set'):  # ComboBox
                        entry.set(value or "Uygun")

        # Ana pano override değerlerini yükle
        if 'ana_pano_overrides' in data:
            overrides = data['ana_pano_overrides']

            if overrides.get('Parafudr Tipi'):
                self.parafudr_tip_entry.delete(0, 'end')
                self.parafudr_tip_entry.insert(0, overrides['Parafudr Tipi'])

            if overrides.get('Parafudr Imax (kA)'):
                self.parafudr_imax_entry.delete(0, 'end')
                self.parafudr_imax_entry.insert(0, overrides['Parafudr Imax (kA)'])

            if overrides.get('Faz-Toprak Cevrim Empedansi Z_x (Ohm)'):
                self.zx_entry.delete(0, 'end')
                self.zx_entry.insert(0, overrides['Faz-Toprak Cevrim Empedansi Z_x (Ohm)'])

            if overrides.get('Faz-Notr Cevrim Empedansi Z_ln (Ohm)'):
                self.zln_entry.delete(0, 'end')
                self.zln_entry.insert(0, overrides['Faz-Notr Cevrim Empedansi Z_ln (Ohm)'])
                # Ik3'ü güncelle (event tetikle)
                self.zln_entry.event_generate('<KeyRelease>')

            if overrides.get('Gerilim F-F (V)'):
                self.ff_entry.delete(0, 'end')
                self.ff_entry.insert(0, overrides['Gerilim F-F (V)'])

            if overrides.get('Gerilim L-N (V)'):
                self.ln_entry.delete(0, 'end')
                self.ln_entry.insert(0, overrides['Gerilim L-N (V)'])

            if overrides.get('Gerilim N-PE (V)'):
                self.npe_entry.delete(0, 'end')
                self.npe_entry.insert(0, overrides['Gerilim N-PE (V)'])

            if overrides.get('Uygunluk'):
                self.uygunluk_combo.set(overrides['Uygunluk'])

        # Fonksiyon testleri verilerini yükle
        if 'fonksiyon_testleri' in data:
            ft_data = data['fonksiyon_testleri']

            # Mevcut satırları temizle (ilk 3 satır hariç)
            while len(self.ft_rows) > 0:
                self.remove_ft_row()

            # Yeni satırları ekle ve verileri yükle
            for row_data in ft_data:
                self.add_ft_row()
                row_entries = self.ft_rows[-1]['entries']

                # Linye Adı
                if row_data.get('Linye Adi'):
                    row_entries['Linye Adi'].delete(0, 'end')
                    row_entries['Linye Adi'].insert(0, row_data['Linye Adi'])

                # Açma Eğrisi
                if row_data.get('Acma Egrisi'):
                    row_entries['Acma Egrisi'].set(row_data['Acma Egrisi'])

                # Kutup Sayısı
                if row_data.get('Kutup Sayisi'):
                    row_entries['Kutup Sayisi'].set(str(row_data['Kutup Sayisi']))

                # In (A)
                if row_data.get('In (A)'):
                    row_entries['In (A)'].set(str(row_data['In (A)']))

                # Icu
                if row_data.get('Icu'):
                    row_entries['Icu'].set(str(row_data['Icu']))

                # Ib (readonly olduğu için önce normal yap)
                if row_data.get('Ib'):
                    row_entries['Ib'].configure(state='normal')
                    row_entries['Ib'].delete(0, 'end')
                    row_entries['Ib'].insert(0, str(row_data['Ib']))
                    row_entries['Ib'].configure(state='readonly')

                # Kesitler
                if row_data.get('Faz Kesiti'):
                    row_entries['Faz Kesiti'].set(str(row_data['Faz Kesiti']))
                if row_data.get('Notr Kesiti'):
                    row_entries['Notr Kesiti'].set(str(row_data['Notr Kesiti']))
                if row_data.get('Toprak Kesiti'):
                    row_entries['Toprak Kesiti'].set(str(row_data['Toprak Kesiti']))

                # Iz
                if row_data.get('Iz'):
                    row_entries['Iz'].delete(0, 'end')
                    row_entries['Iz'].insert(0, str(row_data['Iz']))

                # Sonuç
                if row_data.get('Sonuc'):
                    row_entries['Sonuc'].set(row_data['Sonuc'])

                # KAKR
                if row_data.get('KAKR Var'):
                    row_entries['KAKR Var'].set(True)
                    row_entries['KAKR_Checkbox'].select()
                    row_entries['RCD Acma'].configure(state='normal')
                    row_entries['RCD mA'].configure(state='normal')
                    row_entries['RCD ms'].configure(state='normal')

                    if row_data.get('RCD Acma'):
                        row_entries['RCD Acma'].set(row_data['RCD Acma'])
                    if row_data.get('RCD mA'):
                        row_entries['RCD mA'].delete(0, 'end')
                        row_entries['RCD mA'].insert(0, str(row_data['RCD mA']))
                    if row_data.get('RCD ms'):
                        row_entries['RCD ms'].delete(0, 'end')
                        row_entries['RCD ms'].insert(0, str(row_data['RCD ms']))

                # KAKR Yok
                if row_data.get('KAKR Yok'):
                    row_entries['KAKR Yok'].set(True)

    def get_name(self) -> str:
        return self.name_entry.get() or f"Pano {self.pano_index + 1}"
