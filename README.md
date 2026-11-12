# elektrik-panel

Azərişıq və Aztelekom üçün hesabat paneli. Birinci bölmə: elektrik sərfiyyatı.

## 2025–2026 müqayisəsi (iyun–oktyabr)

`index.html` açıldıqda hazır data ilə müqayisə görünür (kVt·saat, AZN, abunə kodları, EATS üzrə, tarif təhlili, Excel yoxlaması, çap/PDF). Data `assets/data.js` faylındadır və heç bir serverə göndərilmir.

Mənbələr: 12 aylıq hesab-faktura (04/200-068 və 04/100-107) və 12 aylıq EATS Excel cədvəli.

### Datanı yenidən yaratmaq (yeni ay əlavə olunanda)

```
python3 tools/build_data.py <açılmış_qovluq> <rtf_pdf_qovluğu> data.json
python3 tools/make_site_data.py data.json assets/data.js
```

Tələb: `openpyxl`, `pdftotext` (poppler). 2025-in RTF fakturaları əvvəlcə PDF-ə çevrilir (`soffice --headless --convert-to pdf`).

## Aylıq Excel əlavə et

Ayrıca tab: istənilən aylıq Excel faylını yükləyib aylar üzrə sərfiyyat fərqini izləmək üçün. Məlumat yalnız brauzerdə saxlanılır.
