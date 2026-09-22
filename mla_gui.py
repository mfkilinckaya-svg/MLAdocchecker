"""
MLA DOCX Checker — pencere (GUI) arayüzü.
Çift tıklayınca açılır; .docx seçilir, rapor renkli olarak gösterilir.
.docx dosyası exe'nin üzerine sürüklenirse doğrudan analiz edilir.
"""

import os
import sys
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from mla_docx_checker import MLADocxChecker, __version__

COLORS = {"HATA": "#c62828", "UYARI": "#b26a00", "BAŞARILI": "#2e7d32"}
ICONS = {"HATA": "✖", "UYARI": "⚠", "BAŞARILI": "✔"}


class App(tk.Tk):
    def __init__(self, initial_file=None):
        super().__init__()
        self.title(f"MLA DOCX Checker v{__version__}")
        self.geometry("900x620")
        self.minsize(640, 420)
        self.report = []
        self.current_file = None

        top = ttk.Frame(self, padding=10)
        top.pack(fill="x")
        ttk.Button(top, text="Word dosyası seç (.docx)", command=self.choose).pack(side="left")
        self.save_btn = ttk.Button(top, text="Raporu kaydet (.txt)", command=self.save, state="disabled")
        self.save_btn.pack(side="left", padx=8)
        self.file_lbl = ttk.Label(top, text="Henüz dosya seçilmedi.", foreground="#555")
        self.file_lbl.pack(side="left", padx=8)

        self.summary = ttk.Label(self, text="", font=("Segoe UI", 11, "bold"), padding=(10, 0))
        self.summary.pack(fill="x")

        body = ttk.Frame(self, padding=10)
        body.pack(fill="both", expand=True)
        self.text = tk.Text(body, wrap="word", font=("Segoe UI", 10), padx=10, pady=8,
                            relief="solid", borderwidth=1)
        scroll = ttk.Scrollbar(body, command=self.text.yview)
        self.text.configure(yscrollcommand=scroll.set)
        scroll.pack(side="right", fill="y")
        self.text.pack(side="left", fill="both", expand=True)
        for status, color in COLORS.items():
            self.text.tag_configure(status, foreground=color, font=("Segoe UI", 10, "bold"))
        self.text.tag_configure("cat", font=("Segoe UI", 10, "bold"))
        self._write_intro()

        if initial_file:
            self.after(100, lambda: self.analyze(initial_file))

    def _write_intro(self):
        self.text.configure(state="normal")
        self.text.delete("1.0", "end")
        self.text.insert("end",
                         "MLA Handbook 9. baskı kurallarına göre Word (.docx) denetimi.\n\n"
                         "Başlamak için yukarıdaki 'Word dosyası seç' düğmesine tıklayın.\n"
                         "İpucu (Windows): .docx dosyasını programın simgesinin üzerine sürükleyip bırakabilirsiniz.")
        self.text.configure(state="disabled")

    def choose(self):
        path = filedialog.askopenfilename(title="Word dosyası seçin",
                                          filetypes=[("Word belgeleri", "*.docx"), ("Tüm dosyalar", "*.*")])
        if path:
            self.analyze(path)

    def analyze(self, path):
        if not path.lower().endswith(".docx"):
            messagebox.showerror("Desteklenmeyen dosya",
                                 "Yalnızca .docx dosyaları desteklenir.\n"
                                 "Eski .doc dosyalarını Word'de 'Farklı Kaydet → .docx' ile dönüştürün.")
            return
        try:
            self.report = MLADocxChecker(path).run_all_checks()
        except Exception as e:
            messagebox.showerror("Hata", f"Dosya okunamadı:\n{e}\n\nDosya Word'de açıksa kapatıp tekrar deneyin.")
            return

        self.current_file = path
        self.file_lbl.configure(text=os.path.basename(path))
        self.save_btn.configure(state="normal")

        order = {"HATA": 0, "UYARI": 1, "BAŞARILI": 2}
        stats = {k: 0 for k in order}
        self.text.configure(state="normal")
        self.text.delete("1.0", "end")
        for item in sorted(self.report, key=lambda r: order[r["durum"]]):
            stats[item["durum"]] += 1
            self.text.insert("end", f"{ICONS[item['durum']]} {item['durum']}  ", item["durum"])
            self.text.insert("end", f"{item['kategori']}: ", "cat")
            self.text.insert("end", f"{item['mesaj']}\n\n")
        self.text.configure(state="disabled")

        self.summary.configure(
            text=f"{stats['HATA']} hata   ·   {stats['UYARI']} uyarı   ·   {stats['BAŞARILI']} başarılı",
            foreground=COLORS["HATA"] if stats["HATA"] else (COLORS["UYARI"] if stats["UYARI"] else COLORS["BAŞARILI"]))

    def save(self):
        if not self.report:
            return
        default = os.path.splitext(os.path.basename(self.current_file))[0] + "_MLA_rapor.txt"
        path = filedialog.asksaveasfilename(defaultextension=".txt", initialfile=default,
                                            filetypes=[("Metin dosyası", "*.txt")])
        if not path:
            return
        with open(path, "w", encoding="utf-8") as f:
            f.write(f"MLA DOCX Checker v{__version__}\nBelge: {self.current_file}\n\n")
            for item in self.report:
                f.write(f"[{item['durum']}] {item['kategori']}: {item['mesaj']}\n")
        messagebox.showinfo("Kaydedildi", f"Rapor kaydedildi:\n{path}")


def main():
    initial = sys.argv[1] if len(sys.argv) > 1 else None
    App(initial).mainloop()


if __name__ == "__main__":
    main()
