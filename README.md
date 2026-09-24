# Grāmatvedības darba pārvaldības aplikācija

Latviešu valodas mobilajai lietošanai paredzēts Vite + React prototips.

## Kas jau darbojas

- Klientu un darbu izvēle.
- Jauna klienta/jauna darba pievienošana.
- Darba laika taimeris.
- Piezīmes par paveikto darbu.
- Mēneša darba tabeles skatījums: klienti rindās, datumi kolonnās.
- Excel eksports ar tabeli un atsevišķu piezīmju lapu.
- Mēneša kopējais laiks tiek noapaļots līdz pilnām stundām.
- Latvijas kalendāra skats ar nedēļas nogalēm un svētku dienām.
- Gada pārskata kontrolsaraksts ar standarta un individuāliem darbiem.
- PVN klientu atšķiršana: mēnesis / ceturksnis / nav PVN.
- Atskaites un nodokļu informācijas atzīmēšana.
- Dati tiek glabāti pārlūkprogrammas localStorage.

## Palaišana lokāli

```bash
npm install
npm run dev
```

## GitHub Pages

Projektu var izvietot GitHub Pages. Vēlāk var pievienot GitHub Actions automātiskai `dist` publicēšanai.

## Nākamais izstrādes posms

1. Pilns Latvijas grāmatveža kalendārs ar pārceltajām darba dienām.
2. Sistēmas paziņojumi telefonā, apvienojot vienāda tipa atskaites vienā paziņojumā.
3. Klienta profilā visu atskaišu konfigurācija.
4. Gada pārskata darbu saglabāšana un individuālie kontrolsaraksti.
5. Datu rezerves kopija / eksports.
6. PWA instalēšana Android telefonā.
