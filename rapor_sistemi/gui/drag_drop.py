import customtkinter as ctk
from tkinter import filedialog
import os
from typing import List

class DragDropFrame(ctk.CTkFrame):
    """Sürükle-bırak dosya yükleme alanı."""

    def __init__(self, master, file_types: List[str] = None, on_drop=None, **kwargs):
        super().__init__(master, **kwargs)
        self.file_types = file_types or ['.docx']
        self.on_drop = on_drop
        self.dropped_files = []

        self.configure(fg_color="#2b2b2b", border_width=2, border_color="#565656")

        self.label = ctk.CTkLabel(
            self,
            text="📁 Fluke dosyalarını sürükleyin\nveya tıklayarak seçin",
            font=ctk.CTkFont(size=12),
            text_color="#888888"
        )
        self.label.pack(expand=True, fill="both", padx=10, pady=20)

        self.files_label = ctk.CTkLabel(self, text="", font=ctk.CTkFont(size=10))
        self.files_label.pack(pady=(0, 5))

        self.bind("<Button-1>", self.browse_files)
        self.label.bind("<Button-1>", self.browse_files)

    def browse_files(self, event=None):
        filetypes = [("Fluke dosyaları", "*.docx")]
        files = filedialog.askopenfilenames(filetypes=filetypes)
        if files:
            self.dropped_files = list(files)
            self.update_display()

    def update_display(self):
        if self.dropped_files:
            names = [os.path.basename(f) for f in self.dropped_files[:2]]
            text = "\n".join(f"✓ {n}" for n in names)
            if len(self.dropped_files) > 2:
                text += f"\n+{len(self.dropped_files) - 2} dosya"
            self.files_label.configure(text=text, text_color="#4CAF50")

    def clear(self):
        self.dropped_files = []
        self.files_label.configure(text="")

    def get_files(self):
        return self.dropped_files
