# İSTREGISTER — Tədris Reyestri İdarəetmə Sistemi

Industrial Support and Training MMC təlim mərkəzinin tələbə qeydiyyat reyestrinin (REGİSTR-2026.xlsx) produksiyaya hazır veb tətbiqi.

## Xüsusiyyətlər

- **🔍 Qlobal axtarış** — bütün sütunlar üzrə real-time axtarış
- **📊 Sütun üzrə filtr** — Excel-vari filtr paneli, bir neçə sütun eyni anda (AND məntiqi)
- **⚡ Virtualised cədvəl** — AG-Grid vasitəsilə 3000+ sətir axıcı scroll
- **📋 Training Plan** — filtrlənmiş məlumatdan Excel sənəd yaradılması (Training_plan.xlsx şablonu əsasında)
- **📄 Kontekst menyu** — sağ klik / mobil uzun basış
- **📱 Responsive** — masaüstü, planşet və mobil tam dəstəklənir

## Quraşdırma

```bash
npm install
npm run dev      # development server
npm run build    # production build
npm run preview  # preview production build
```

## Firebase İnteqrasiyası

Tətbiq hazırda statik məlumat rejimində işləyir (REGİSTR-2026.xlsx-dan idxal edilən JSON).

Firebase Realtime Database-i aktivləşdirmək üçün `src/lib/firebase.js` faylında
konfiqurasiya dəyərlərini daxil edin, sonra `src/services/registryService.js` 
məlumatları Firebase-dən oxumağa başlayacaq.

## Data Strukturu

- `src/data/registrData.json` — REGİSTR-2026.xlsx-dan idxal edilən 3265 qeyd
- `src/data/adminPanel.json` — Kurs kodu → Gün sayı lüğəti
- `src/data/courses.js` — Kurs kodları, adları və saatları
- `src/data/teachers.js` — Müəllim siyahısı
- `src/data/Training_plan_template.xlsx` — Training Plan şablonu
