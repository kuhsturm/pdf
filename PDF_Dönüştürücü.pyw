
import tkinter as tk
from tkinter import ttk, filedialog, scrolledtext, messagebox
import os
import threading
import win32com.client as win32
import subprocess
import re
import time
import json
import tempfile
import shutil
import zipfile
from tkinterdnd2 import DND_FILES, TkinterDnD

class MultiToolConverter(TkinterDnD.Tk):
    def __init__(self):
        super().__init__()

        self.title("PDF Araç Kutusu v9.5")
        self.geometry("750x900")

        self.state = "IDLE"
        self.pause_event = threading.Event()
        self.cancel_event = threading.Event()
        self.pause_event.set()

        self.source_items = []
        self.zip_source_items = []
        self.output_folder_path = None
        self.libreoffice_exe_path = None
        self.debug_mode_var = tk.BooleanVar(value=False)

        self.main_frame = tk.Frame(self, padx=15, pady=15)
        self.main_frame.pack(fill="both", expand=True)

        self.notebook = ttk.Notebook(self.main_frame)
        self.notebook.pack(fill="both", expand=True, pady=(0, 10))

        self.tab1 = ttk.Frame(self.notebook)
        self.notebook.add(self.tab1, text="  Belgeden PDF'e Dönüştür  ")
        self.setup_tab1()

        self.tab2 = ttk.Frame(self.notebook)
        self.notebook.add(self.tab2, text="  ZIP'ten PDF Çıkar  ")
        self.setup_tab2()

        self.setup_help_tab()

        self.setup_common_bottom_ui()
        
        self.load_settings()
        self.update_button_states()

    def setup_tab1(self):
        source_list_frame = tk.LabelFrame(self.tab1, text="1. Adım: Kaynak Dosya ve Klasörleri Ekleyin", padx=10, pady=10)
        source_list_frame.pack(fill="x", pady=(5, 10))
        source_listbox_frame = tk.Frame(source_list_frame)
        source_listbox_frame.pack(fill="x")
        self.source_listbox = tk.Listbox(source_listbox_frame, selectmode=tk.EXTENDED, height=5)
        self.source_listbox.pack(side="left", fill="x", expand=True)
        source_scrollbar = tk.Scrollbar(source_listbox_frame, orient="vertical", command=self.source_listbox.yview)
        source_scrollbar.pack(side="right", fill="y")
        self.source_listbox.config(yscrollcommand=source_scrollbar.set)
        self.source_listbox.drop_target_register(DND_FILES)
        self.source_listbox.dnd_bind('<<Drop>>', self.handle_drop_tab1)
        source_button_frame = tk.Frame(source_list_frame, padx=10)
        source_button_frame.pack(side="left", fill="y", anchor="n")
        tk.Button(source_button_frame, text="Dosya Ekle...", command=self.add_files).pack(fill="x", pady=2)
        tk.Button(source_button_frame, text="Klasör Ekle...", command=self.add_folder).pack(fill="x", pady=2)
        tk.Button(source_button_frame, text="Seçileni Sil", command=self.delete_selected_items).pack(fill="x", pady=2)
        tk.Button(source_button_frame, text="Listeyi Temizle", command=self.clear_source_list).pack(fill="x", pady=2)
        
        output_frame_parent = tk.LabelFrame(self.tab1, text="2. Adım: Hedef Klasörü Seçin (İsteğe Bağlı)", padx=10, pady=10)
        output_frame_parent.pack(fill="x", pady=5)
        output_frame = tk.Frame(output_frame_parent)
        output_frame.pack(fill="x", pady=5)
        self.select_output_button = tk.Button(output_frame, text="Hedef Klasör Seç", command=self.select_output_folder)
        self.select_output_button.pack(side="left", padx=(0, 10))
        self.output_label = tk.Label(output_frame, text="Tüm PDF'leri tek bir klasöre kaydetmek için seçin...", fg="gray", anchor="w")
        self.output_label.pack(side="left", fill="x", expand=True)
        
        settings_frame = tk.LabelFrame(self.tab1, text="3. Adım: Ayarları Yapılandırın", padx=10, pady=10)
        settings_frame.pack(fill="x", pady=10)
        self.engine_var = tk.StringVar(value="word")
        tk.Radiobutton(settings_frame, text="Microsoft Office Motoru", variable=self.engine_var, value="word").pack(anchor="w")
        tk.Radiobutton(settings_frame, text="LibreOffice Motoru", variable=self.engine_var, value="libreoffice").pack(anchor="w")
        libre_path_frame = tk.Frame(settings_frame)
        libre_path_frame.pack(fill="x", padx=20, pady=(5,0))
        tk.Button(libre_path_frame, text="Gözat...", command=self.select_libreoffice_path).pack(side="left", padx=(0, 10))
        self.libre_path_label = tk.Label(libre_path_frame, text="LibreOffice yolu (gerekirse)...", fg="gray", anchor="w")
        self.libre_path_label.pack(side="left", fill="x", expand=True)
        self.recursive_search_var = tk.BooleanVar()
        tk.Checkbutton(settings_frame, text="Klasörler için alt klasörleri de tara", variable=self.recursive_search_var).pack(anchor="w", pady=5)

    def setup_tab2(self):
        zip_source_frame = tk.LabelFrame(self.tab2, text="İşlenecek .zip Dosyalarını Ekleyin", padx=10, pady=10)
        zip_source_frame.pack(fill="both", expand=True, pady=5)
        zip_listbox_frame = tk.Frame(zip_source_frame)
        zip_listbox_frame.pack(fill="both", expand=True, side="left")
        self.zip_listbox = tk.Listbox(zip_listbox_frame, selectmode=tk.EXTENDED)
        self.zip_listbox.pack(side="left", fill="both", expand=True)
        zip_scrollbar = tk.Scrollbar(zip_listbox_frame, orient="vertical", command=self.zip_listbox.yview)
        zip_scrollbar.pack(side="right", fill="y")
        self.zip_listbox.config(yscrollcommand=zip_scrollbar.set)
        self.zip_listbox.drop_target_register(DND_FILES)
        self.zip_listbox.dnd_bind('<<Drop>>', self.handle_drop_tab2)
        zip_button_frame = tk.Frame(zip_source_frame, padx=10)
        zip_button_frame.pack(side="left", fill="y", anchor="n")
        tk.Button(zip_button_frame, text="Dosya Ekle...", command=self.add_zip_files).pack(fill="x", pady=2)
        tk.Button(zip_button_frame, text="Seçileni Sil", command=self.delete_selected_zip_items).pack(fill="x", pady=2)
        tk.Button(zip_button_frame, text="Listeyi Temizle", command=self.clear_zip_source_list).pack(fill="x", pady=2)

    def setup_help_tab(self):
        help_frame = ttk.Frame(self.notebook)
        self.notebook.add(help_frame, text="  Yardım  ")
        help_text_widget = scrolledtext.ScrolledText(help_frame, wrap=tk.WORD, padx=10, pady=10)
        help_text_widget.pack(fill="both", expand=True)

        help_text = """
PDF Araç Kutusu Kullanım Kılavuzu

Bu program, çeşitli dosya türlerini PDF formatına dönüştürmenize ve ZIP arşivlerindeki PDF dosyalarını çıkarmanıza olanak tanır.

ÖZELLİKLER

1. Belgeden PDF'e Dönüştürme:
   - Desteklenen Formatlar:
     - Microsoft Office Motoru: .docx, .doc, .rtf, .xls, .xlsx
     - LibreOffice Motoru: .docx, .doc, .rtf, .odt, .xls, .xlsx, .ods
   - Nasıl Kullanılır:
     1. 'Belgeden PDF'e Dönüştür' sekmesini seçin.
     2. 'Dosya Ekle' veya 'Klasör Ekle' butonlarını kullanarak veya dosyaları/klasörleri doğrudan listeye sürükleyip bırakarak işlenecek öğeleri ekleyin.
     3. (İsteğe bağlı) 'Hedef Klasör Seç' butonu ile tüm PDF'lerin kaydedileceği ortak bir klasör belirleyin. Belirlenmezse, PDF'ler orijinal dosyaların yanına kaydedilir.
     4. Kullanmak istediğiniz dönüştürme motorunu seçin (Microsoft Office veya LibreOffice).
        - LibreOffice motorunu ilk kez kullanıyorsanız, sistemde kurulu olan 'soffice.exe' dosyasının yolunu 'Gözat...' butonu ile göstermeniz gerekebilir. Bu ayar otomatik olarak kaydedilecektir.
     5. 'İşlemi Başlat' butonuna tıklayın.

2. ZIP'ten PDF Çıkarma:
   - Bir .zip arşivi içindeki tüm .pdf uzantılı dosyaları arşivin bulunduğu dizine çıkarır.
   - Nasıl Kullanılır:
     1. 'ZIP'ten PDF Çıkar' sekmesini seçin.
     2. 'Dosya Ekle' butonunu kullanarak veya .zip dosyalarını listeye sürükleyip bırakarak işlenecek arşivleri ekleyin.
     3. 'İşlemi Başlat' butonuna tıklayın.

Genel İpuçları:
- 'İşlem Geçmişi' alanından tüm adımları ve olası hataları takip edebilirsiniz.
- Dönüştürme sırasında bir hata oluşursa, 'Başarısız Öğeler' sekmesinde ilgili dosyaların bir listesini bulabilirsiniz.
- 'Duraklat' ve 'İptal Et' butonları ile uzun süren işlemleri kontrol edebilirsiniz.
- 'Hata Ayıklama Günlüğünü Aktif Et' seçeneği, sürükle-bırak gibi özelliklerde sorun yaşanması durumunda daha detaylı bilgi almak için kullanılabilir.
"""
        help_text_widget.insert(tk.END, help_text)
        help_text_widget.config(state="disabled")

    def setup_common_bottom_ui(self):
        self.control_frame = tk.Frame(self.main_frame)
        self.control_frame.pack(fill="x", pady=5)
        self.convert_button = tk.Button(self.control_frame, text="İşlemi Başlat", command=self.start_process, font=("Segoe UI", 10, "bold"))
        self.convert_button.pack(side="left", fill="x", expand=True, ipady=5, padx=(0,5))
        self.pause_resume_button = tk.Button(self.control_frame, text="Duraklat", command=self.toggle_pause, state="disabled")
        self.pause_resume_button.pack(side="left", ipady=2, padx=5)
        self.cancel_button = tk.Button(self.control_frame, text="İptal Et", command=self.cancel_conversion, state="disabled")
        self.cancel_button.pack(side="left", ipady=2, padx=(5,0))
        self.summary_frame = tk.LabelFrame(self.main_frame, text="İşlem Özeti ve İlerleme", padx=10, pady=5)
        self.summary_frame.pack(fill="x", pady=(10, 5))
        self.summary_label = tk.Label(self.summary_frame, text="İşlem başlatılmadı.", font=("Segoe UI", 10))
        self.summary_label.pack(pady=5)
        self.progress_bar = ttk.Progressbar(self.summary_frame, orient="horizontal", length=100, mode="determinate")
        self.progress_bar.pack(fill="x", expand=True, pady=5)
        
        # --- YENİ: Sekmeli Çıktı Alanı ---
        self.output_notebook = ttk.Notebook(self.main_frame)
        self.output_notebook.pack(fill="both", expand=True, pady=(10,0))

        log_tab = ttk.Frame(self.output_notebook)
        self.output_notebook.add(log_tab, text="İşlem Geçmişi")
        self.log_area = scrolledtext.ScrolledText(log_tab, wrap=tk.WORD, height=6)
        self.log_area.pack(fill="both", expand=True, padx=2, pady=2)
        self.log_area.config(state="disabled")

        self.failed_tab = ttk.Frame(self.output_notebook)
        self.output_notebook.add(self.failed_tab, text="Başarısız Öğeler")
        self.failed_files_listbox = tk.Listbox(self.failed_tab)
        self.failed_files_listbox.pack(side="left", fill="both", expand=True, padx=2, pady=2)
        scrollbar = tk.Scrollbar(self.failed_tab, orient="vertical", command=self.failed_files_listbox.yview)
        scrollbar.pack(side="right", fill="y")
        self.failed_files_listbox.config(yscrollcommand=scrollbar.set)
        
        debug_frame = tk.Frame(self.main_frame)
        debug_frame.pack(fill="x", pady=5)
        self.debug_check = tk.Checkbutton(debug_frame, text="Hata Ayıklama Günlüğünü Aktif Et", variable=self.debug_mode_var)
        self.debug_check.pack(anchor="w")

    def log(self, message):
        self.log_area.config(state="normal")
        self.log_area.insert(tk.END, message + "\n")
        self.log_area.config(state="disabled")
        self.log_area.see(tk.END)

    def start_process(self):
        current_tab = self.notebook.index(self.notebook.select())
        if current_tab == 0: self.start_conversion_thread()
        elif current_tab == 1: self.start_zip_extraction_thread()

    def handle_drop_tab1(self, event): self.add_items_to_source_list(self.parse_drop_event(event))
    def handle_drop_tab2(self, event): self.add_items_to_zip_list(self.parse_drop_event(event))

    def parse_drop_event(self, event):
        if self.debug_mode_var.get(): self.log(f"DEBUG: Drop event data: '{event.data}'")
        paths = re.findall(r'\{.*?\}', event.data)
        if paths: cleaned_paths = [p[1:-1] for p in paths]
        else: cleaned_paths = [event.data]
        if self.debug_mode_var.get(): self.log(f"DEBUG: Parsed paths: {cleaned_paths}")
        return cleaned_paths

    def add_items_to_source_list(self, items):
        for item in items:
            clean_item = os.path.normpath(item)
            if os.path.exists(clean_item) and clean_item not in self.source_items:
                self.source_items.append(clean_item)
                self.source_listbox.insert(tk.END, clean_item)
        self.update_button_states()

    def add_items_to_zip_list(self, items):
        for item in items:
            clean_item = os.path.normpath(item)
            if os.path.isfile(clean_item) and clean_item.lower().endswith('.zip') and clean_item not in self.zip_source_items:
                self.zip_source_items.append(clean_item)
                self.zip_listbox.insert(tk.END, clean_item)
        self.update_button_states()

    def add_files(self): 
        files = filedialog.askopenfilenames(title="Dönüştürülecek dosyaları seçin");
        if files: self.add_items_to_source_list(files)

    def add_folder(self):
        folder = filedialog.askdirectory(title="Dönüştürülecek dosyaların olduğu klasörü seçin")
        if folder: self.add_items_to_source_list([folder])

    def delete_selected_items(self):
        selected_indices = self.source_listbox.curselection()
        for index in sorted(selected_indices, reverse=True): self.source_listbox.delete(index); del self.source_items[index]
        self.update_button_states()

    def clear_source_list(self): self.source_items.clear(); self.source_listbox.delete(0, tk.END); self.update_button_states()
    def add_zip_files(self): 
        files = filedialog.askopenfilenames(title="İşlenecek .zip dosyalarını seçin", filetypes=[("ZIP files", "*.zip")]);
        if files: self.add_items_to_zip_list(files)

    def delete_selected_zip_items(self):
        selected_indices = self.zip_listbox.curselection()
        for index in sorted(selected_indices, reverse=True): self.zip_listbox.delete(index); del self.zip_source_items[index]
        self.update_button_states()

    def clear_zip_source_list(self): self.zip_source_items.clear(); self.zip_listbox.delete(0, tk.END); self.update_button_states()

    def select_output_folder(self):
        folder_path = filedialog.askdirectory()
        if folder_path: self.output_folder_path = folder_path; self.output_label.config(text=f"{folder_path}", fg="blue"); self.log(f"Hedef klasör ayarlandı: {folder_path}")

    def select_libreoffice_path(self):
        exe_path = filedialog.askopenfilename(title="soffice.exe dosyasını seçin", filetypes=[("Executable", "*.exe")])
        if exe_path and exe_path.lower().endswith("soffice.exe"): self.libreoffice_exe_path = exe_path; self.libre_path_label.config(text=f"{exe_path}", fg="blue"); self.log(f"LibreOffice yolu ayarlandı: {exe_path}"); self.save_settings()

    def set_state(self, new_state): self.state = new_state; self.after(0, self.update_button_states)

    def update_button_states(self):
        try: current_tab = self.notebook.index(self.notebook.select())
        except: current_tab = 0
        is_list_empty = (current_tab == 0 and not self.source_items) or (current_tab == 1 and not self.zip_source_items)
        if self.state == "IDLE":
            self.convert_button.config(state="disabled" if is_list_empty else "normal")
            self.pause_resume_button.config(state="disabled", text="Duraklat")
            self.cancel_button.config(state="disabled")
        else:
            self.convert_button.config(state="disabled")
            if self.state == "RUNNING": self.pause_resume_button.config(state="normal", text="Duraklat"); self.cancel_button.config(state="normal")
            elif self.state == "PAUSED": self.pause_resume_button.config(state="normal", text="Devam Et"); self.cancel_button.config(state="normal")
            elif self.state == "CANCELLING": self.pause_resume_button.config(state="disabled", text="İptal Ediliyor..."); self.cancel_button.config(state="disabled")

    def toggle_pause(self):
        if self.state == "RUNNING": self.pause_event.clear(); self.set_state("PAUSED"); self.log(">> İşlem duraklatıldı.")
        elif self.state == "PAUSED": self.pause_event.set(); self.set_state("RUNNING"); self.log(">> İşlem devam ediyor...")

    def cancel_conversion(self):
        if self.state in ["RUNNING", "PAUSED"]: self.set_state("CANCELLING"); self.log(">> İşlem iptal ediliyor..."); self.cancel_event.set()
        if self.state == "PAUSED": self.pause_event.set()

    def start_conversion_thread(self):
        files_to_process = self.find_files_to_convert()
        if not files_to_process: messagebox.showinfo("Bilgi", "Listede dönüştürülecek geçerli bir dosya bulunamadı."); return
        self.run_process(self.run_msoffice_conversion if self.engine_var.get() == "word" else self.run_libreoffice_conversion, files_to_process)

    def start_zip_extraction_thread(self):
        if not self.zip_source_items: messagebox.showinfo("Bilgi", "Listede işlenecek .zip dosyası bulunamadı."); return
        self.run_process(self.run_zip_extraction, self.zip_source_items)

    def run_process(self, target_function, files_to_process):
        self.pause_event.set(); self.cancel_event.clear(); self.set_state("RUNNING")
        self.failed_files_listbox.delete(0, tk.END)
        self.summary_label.config(text="İşlem devam ediyor...")
        self.progress_bar["value"] = 0
        self.progress_bar["maximum"] = len(files_to_process)
        self.log(f"\n===== {len(files_to_process)} öğe için işlem başlatılıyor... =====")
        thread = threading.Thread(target=target_function, args=(files_to_process,))
        thread.daemon = True
        thread.start()

    def update_ui_after_conversion(self, success_count, failed_files, was_cancelled=False):
        self.set_state("IDLE")
        self.log("===== İşlem tamamlandı. =====")
        if was_cancelled: summary_text = f"İşlem İptal Edildi | Tamamlanan: {success_count} | Başarısız: {len(failed_files)}"
        else: summary_text = f"Başarıyla İşlenen: {success_count}  |  Başarısız: {len(failed_files)}"
        self.summary_label.config(text=summary_text)
        if failed_files:
            self.log(f"ÖZET: {success_count} öğe başarıyla işlendi, {len(failed_files)} öğe işlenemedi.")
            self.output_notebook.select(self.failed_tab) # Başarısız sekmesini otomatik seç
        else: self.log(f"ÖZET: Toplam {success_count} öğenin tümü başarıyla işlendi.")
        for item in failed_files: self.failed_files_listbox.insert(tk.END, os.path.basename(item))

    def find_files_to_convert(self):
        supported_formats = (".docx", ".doc", ".rtf", ".xls", ".xlsx")
        if self.engine_var.get() == "libreoffice": supported_formats += (".odt", ".ods")
        files_to_process = set()
        for item_path in self.source_items:
            if os.path.isfile(item_path):
                if item_path.lower().endswith(supported_formats) and not os.path.basename(item_path).startswith("~"):
                    files_to_process.add(item_path)
            elif os.path.isdir(item_path):
                if self.recursive_search_var.get():
                    for root, _, files in os.walk(item_path):
                        for file in files:
                            if file.lower().endswith(supported_formats) and not file.startswith("~"): files_to_process.add(os.path.join(root, file))
                else:
                    for file in os.listdir(item_path):
                        if file.lower().endswith(supported_formats) and not file.startswith("~"): files_to_process.add(os.path.join(item_path, file))
        return list(files_to_process)

    def run_msoffice_conversion(self, files_to_process):
        word_app = None
        excel_app = None
        failed_files, success_count = [], 0
        was_cancelled = False
        try:
            for doc_path_raw in files_to_process:
                self.pause_event.wait()
                if self.cancel_event.is_set():
                    was_cancelled = True
                    break

                doc_path = os.path.normpath(os.path.abspath(doc_path_raw))
                filename = os.path.basename(doc_path)
                output_dir = self.output_folder_path if self.output_folder_path else os.path.dirname(doc_path)
                pdf_path = os.path.normpath(os.path.join(output_dir, os.path.splitext(filename)[0] + ".pdf"))

                try:
                    self.log(f"-> İşleniyor: {filename}")
                    if doc_path.lower().endswith((".xls", ".xlsx")):
                        if not excel_app:
                            excel_app = win32.Dispatch("Excel.Application")
                            excel_app.Visible = False
                        workbook = excel_app.Workbooks.Open(doc_path)
                        if not workbook:
                            raise RuntimeError("Excel belgesi açılamadı.")
                        
                        workbook.ActiveSheet.ExportAsFixedFormat(0, pdf_path)

                        workbook.Close(False) # Close without saving changes
                    else: # Assume Word document
                        if not word_app:
                            word_app = win32.Dispatch("Word.Application")
                            word_app.Visible = False
                        doc = word_app.Documents.Open(doc_path)
                        if not doc:
                            raise RuntimeError("Word belgesi açılamadı.")
                        doc.SaveAs(pdf_path, FileFormat=17)
                        doc.Close(0)

                    if not os.path.exists(pdf_path):
                        raise RuntimeError("PDF dosyası oluşturulamadı (sessiz hata).")

                    self.log(f"   BAŞARILI -> {pdf_path}")
                    success_count += 1
                except Exception as e:
                    self.log(f"   HATA: {filename} dönüştürülemedi. Sebep: {e}")
                    failed_files.append(doc_path)
                finally:
                    self.after(0, self.update_progress)
        except Exception as e:
            self.log(f"!! KRİTİK HATA: Microsoft Office başlatılamadı. Detay: {e}")
        finally:
            if word_app:
                word_app.Quit()
            if excel_app:
                excel_app.Quit()
            self.after(0, self.update_ui_after_conversion, success_count, failed_files, was_cancelled)

    def run_libreoffice_conversion(self, files_to_process):
        failed_files, success_count = [], 0; was_cancelled = False
        soffice_cmd = f'"{self.libreoffice_exe_path}"' if self.libreoffice_exe_path else "soffice"
        try:
            for docx_path in files_to_process:
                self.pause_event.wait()
                if self.cancel_event.is_set(): was_cancelled = True; break
                filename = os.path.basename(docx_path); process = None
                try:
                    output_dir = self.output_folder_path if self.output_folder_path else os.path.dirname(docx_path)
                    pdf_path = os.path.join(output_dir, os.path.splitext(filename)[0] + ".pdf")
                    self.log(f"-> İşleniyor: {filename}")
                    command = f'{soffice_cmd} --headless --convert-to pdf --outdir "{output_dir}" "{docx_path}"'
                    process = subprocess.Popen(command, shell=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
                    while process.poll() is None:
                        if self.cancel_event.is_set(): process.kill(); raise InterruptedError("Kullanıcı tarafından iptal edildi.")
                        time.sleep(0.2)
                    stdout, stderr = process.communicate()
                    if process.returncode == 0 and os.path.exists(pdf_path): self.log(f"   BAŞARILI -> {pdf_path}"); success_count += 1
                    else: raise RuntimeError(f"LibreOffice hatası. Kod: {process.returncode}. Mesaj: {stderr.strip()}")
                except InterruptedError as e: self.log(f"   İPTAL EDİLDİ: {filename} işlemi sonlandırıldı."); failed_files.append(docx_path); was_cancelled = True; break
                except Exception as e: self.log(f"   HATA: {filename} dönüştürülemedi."); failed_files.append(docx_path)
                finally: self.after(0, self.update_progress)
        finally: self.after(0, self.update_ui_after_conversion, success_count, failed_files, was_cancelled)

    def run_zip_extraction(self, files_to_process):
        failed_items, success_count = [], 0; was_cancelled = False
        try:
            for zip_path in files_to_process:
                self.pause_event.wait()
                if self.cancel_event.is_set(): was_cancelled = True; break
                filename = os.path.basename(zip_path)
                self.log(f"-> İşleniyor: {filename}")
                temp_dir = tempfile.mkdtemp(prefix="pdf_ext_")
                try:
                    with zipfile.ZipFile(zip_path, 'r') as zip_ref:
                        pdf_files_in_zip = [name for name in zip_ref.namelist() if name.lower().endswith('.pdf')]
                        if not pdf_files_in_zip: self.log(f"   UYARI: {filename} içinde PDF bulunamadı."); success_count += 1; continue
                        zip_ref.extractall(temp_dir, members=pdf_files_in_zip)
                    extracted_pdfs_paths = []
                    for root, _, files in os.walk(temp_dir):
                        for file in files: 
                            if file.lower().endswith('.pdf'): extracted_pdfs_paths.append(os.path.join(root, file))
                    output_dir = os.path.dirname(zip_path)
                    base_name = os.path.splitext(filename)[0]
                    if len(extracted_pdfs_paths) == 1: shutil.move(extracted_pdfs_paths[0], os.path.join(output_dir, f"{base_name}.pdf"))
                    else: 
                        for i, pdf_path in enumerate(extracted_pdfs_paths, 1): shutil.move(pdf_path, os.path.join(output_dir, f"{base_name}-{i}.pdf"))
                    self.log(f"   BAŞARILI: {filename} içinden {len(extracted_pdfs_paths)} PDF çıkarıldı."); success_count += 1
                except InterruptedError as e: self.log(f"   İPTAL EDİLDİ: {filename} işlemi sonlandırıldı."); failed_items.append(zip_path); was_cancelled = True; break
                except Exception as e: self.log(f"   HATA: {filename} işlenemedi. Sebep: {e}"); failed_items.append(zip_path)
                finally: shutil.rmtree(temp_dir, ignore_errors=True); self.after(0, self.update_progress)
        finally: self.after(0, self.update_ui_after_conversion, success_count, failed_items, was_cancelled)
    
    def update_progress(self):
        self.progress_bar["value"] += 1

    def save_settings(self):
        settings = {"libreoffice_path": self.libreoffice_exe_path}
        try:
            with open("config.json", "w") as f: json.dump(settings, f)
        except Exception as e: self.log(f"Ayarlar kaydedilemedi: {e}")

    def load_settings(self):
        try:
            if os.path.exists("config.json") and os.path.getsize("config.json") > 0:
                with open("config.json", "r") as f:
                    settings = json.load(f)
                    if settings.get("libreoffice_path"):
                        self.libreoffice_exe_path = settings["libreoffice_path"]
                        self.libre_path_label.config(text=f"{self.libreoffice_exe_path}", fg="blue")
        except Exception as e: self.log(f"Ayarlar yüklenemedi: {e}")

if __name__ == "__main__":
    app = MultiToolConverter()
    app.mainloop()
