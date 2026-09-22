# MLA DOCX Checker

Word (.docx) belgelerini **MLA 9th Edition** biçim kurallarına göre denetleyen ve hataları/uyarıları raporlayan basit bir Python aracı.

*A simple Python tool that checks Word (.docx) documents against MLA 9th Edition formatting rules. English summary below.*

## Neleri denetler?

Her mesajın sonunda ilgili MLA Handbook (9. baskı) bölüm numarası yer alır, ör. `[MLA 1.2]`.

| Kategori | Kontrol | Bölüm |
|---|---|---|
| Kenar boşlukları | Tüm kenarlar 1 inç (2,54 cm) | 1.1 |
| Metin biçimi | 11–13 punto, tek yazı tipi ve boyut, çift satır aralığı, iki yana yaslama yok, otomatik heceleme kapalı, 0,5 inç ilk satır girintisi, paragraflar arasında boş satır yok, cümle sonrası tek boşluk | 1.2 |
| İlk sayfa | Kimlik bloğu (ad, öğretim üyesi, ders, tarih) sola yaslı; başlık ortalı, noktasız, tırnaksız, tamamı büyük harf / kalın / altı çizili / italik değil | 1.3 |
| Üst bilgi | Sağ üstte "Soyadı SayfaNo", üstten 0,5 inç, ilk sayfa dahil; "p." veya işaret yok | 1.4 |
| Ara başlıklar | Sola yaslı, noktasız, tamamı büyük harf değil | 1.5 |
| Metin içi alıntılar | `(Baron 194)`: yazar ile sayfa arasında virgül yok; `p.`/`pp.` yok; nokta parantez dışında; `et al.` yazımı | 6.16, 6.30, 6.43 |
| Works Cited | Yeni sayfada, ortalı başlık; tek kaynakta "Work Cited"; harf harf alfabetik sıra (aksanlar ve baştaki A/An/The yok sayılır, `---.` desteklenir); 0,5 inç asılı girinti; çift aralık; her girdi nokta ile biter | 1.6, 5.120–5.126 |

Çıkış kodu: hata yoksa `0`, en az bir hata varsa `1` (CI/otomasyon için kullanılabilir).

## Kurulum

Python 3.8 veya üzeri gerekir.

```bash
git clone https://github.com/mfkilinckaya-svg/MLAdocchecker.git
cd MLAdocchecker
pip install -r requirements.txt
```

## Kullanım

```bash
python mla_docx_checker.py makalem.docx
```

Örnek çıktı:

```
✅ [BAŞARILI] Kenar Boşlukları: Bölüm 1: Tüm kenar boşlukları standart 1 inç (2.54 cm).
❌ [HATA] Metin İçi Alıntı Hata: Paragraf 6: Yazar ve sayfa numarası arasında virgül olmamalıdır! ...
⚠️ [UYARI] Paragraf Boşlukları: Paragraf 8 öncesinde veya sonrasında ekstra boşluk var (0 pt olmalıdır).
```

Kendi kodunuzda da kullanabilirsiniz:

```python
from mla_docx_checker import MLADocxChecker

report = MLADocxChecker("makalem.docx").run_all_checks()
for item in report:
    print(item["durum"], item["kategori"], item["mesaj"])
```

## Sınırlamalar

- Yalnızca `.docx` dosyalarını okur (`.doc`, `.pdf` desteklenmez).
- Biçim ayarları paragraf, stil zinciri ve belge varsayılanlarından okunur; tema yazı tipleri (ör. "+Body") adıyla çözümlenemeyebilir.
- Kimlik bloğunun ilk 4 dolu paragraf, başlığın 5. dolu paragraf olduğu varsayılır (başlık sayfası kullanan grup projelerinde sonuçlar yanıltıcı olabilir).
- Metin içi alıntı kontrolleri regex tabanlıdır; olağan dışı alıntı biçimlerini kaçırabilir.
- Sonuçlar yol gösterici niteliktedir; son kontrol için güncel MLA Handbook'a başvurun.

## English

Checks margins, running head, first-page heading block and title, double spacing, indentation, fonts, hyphenation, headings, in-text citation punctuation (`(Author Page)`, no `p.`, `et al.`), and the Works Cited page (new page, centered heading, letter-by-letter alphabetical order, hanging indent). Each message cites the relevant MLA Handbook (9th ed.) section. Messages are in Turkish.

```bash
pip install -r requirements.txt
python mla_docx_checker.py paper.docx
```

## Sürüm geçmişi

- **2.0.0**: Üst bilgi, heceleme, ara başlık, `p./pp.` ve Works Cited (yeni sayfa, Work/Works, çift aralık, nokta) kontrolleri eklendi; stil zincirinden biçim okuma; MLA 5.124'e uygun alfabetik sıralama; 11–13 pt aralığı; bölüm numaralı mesajlar.
- **1.0.0**: İlk sürüm.

## Lisans

[MIT](LICENSE) © 2026 Muhammed Fevzi Kılınçkaya
