# 05 — NORMATİF REFERANSLAR (PDF içeriğinin dayanağı)

## ГОСТ IEC 61439-1-2024 (≡ IEC 61439-1:2020) / ГОСТ IEC 61439-2
- KZ'de ГОСТ IEC 61439-1-2024, 2013 sürümünün yerine (EASC). Proje başka sürüm istiyorsa `teknik.standart` alanı değiştirilir.
- **§6.1 Маркировка (табличка НКУ)**: üreticinin adı / ticari markası; tip tanımı veya seri no; üretim tarihini belirleme imkânı; ilgili IEC 61439-X. Diğer veriler (Ue, InA, Icw, IP, topraklama sistemi…) tablada veya teknik dokümanda.
- **§6.2 Документация**: kurulum, işletme, bakım talimatı; teknik veriler.
- **§11 Rutin doğrulama** (her pano, üretici sorumluluğu): IP; hava aralığı / yüzeysel kaçak; elektrik çarpmasına karşı koruma ve PE sürekliliği; komponent entegrasyonu; iç devreler ve bağlantılar; dış iletken terminalleri; mekanik çalışma; dielektrik dayanım; kablolama ve fonksiyon. → "Общие технические требования" sayfasındaki ОТК tablosu bu listeden.
- Kullanıcı tarafından tanımlanması gereken arayüz verileri (ГОСТ IEC/TR 61439-0): Ue, topraklama sistemi, Ik, InA, IP, montaj yeri/şekli, ortam koşulları, form → titul tablosu alanları.

## Kazakistan / EAEU
- ТР ТС 004/2011 (düşük gerilim ekipman güvenliği), ТР ТС 020/2011 (EMC) → tablada EAC işareti.
- ПУЭ п.1.1.29–1.1.30 (RK ПУЭ aynı numaralandırma): **bara** renkleri 3 faz AC: L1 (A) sarı, L2 (B) yeşil, L3 (C) kırmızı; N mavi; PE sarı-yeşil şerit. → `genel.bara_renk = "ПУЭ"`.
- ГОСТ 33542-2015 (≡ IEC 60445:2010) iletken renkleri: N mavi, PE sarı-yeşil; faz iletkenleri kahverengi / siyah / gri (L1/L2/L3 eşlemesi standartta zorunlu değil, yaygın tercih BN/BK/GY). → `genel.bara_renk = "IEC"` seçilirse baralar da BN/BK/GY.
- ГОСТ 15150 iklim sınıfı (УХЛ4 iç mekân, УХЛ1 dış mekân).
- ГОСТ 2.710-81 harf kodları (QF, QS, QD, KM, KK, FV, FU, PI, TA, HL, SB, SA, UZ, XT).
- ГОСТ 12.4.026 uyarı işareti "Осторожно! Электрическое напряжение".

## Kumanda elemanları renkleri
- IEC 60073 (lamba): kırmızı = acil/arıza, sarı = anormal/uyarı, yeşil = normal, mavi = zorunlu eylem, beyaz = nötr (şebeke var).
- IEC 60204-1 §10.2 (buton): START = beyaz (tercih), gri, siyah veya yeşil — kırmızı olamaz; STOP = siyah, gri, beyaz veya kırmızı (tercih) — yeşil olamaz; ACİL STOP = kırmızı, sarı zemin.
- TEKLED pratiği: Пуск yeşil, Стоп kırmızı, Сеть beyaz, Работа yeşil, Авария kırmızı.

## XPAN pratiğinden alınan (TEKLED onayı bekleyen) tablolar
`genel.kesit_tablosu_onayli` / `tork_tablosu_onayli` true olmadan PDF'e basılmaz.
- İç tesisat kesiti (Cu, H07V-K / ПВ3): ≤25A 2,5 · 32A 4 · 40A 6 · 50A 10 · 63A 16 · 80A 16 · 100A 25 · 125A 35 · 160A 50 · 200A 70 · 250A 95 mm².
- Bara: 250A 20×5 · 400A 30×5 · 630A 30×10 · 800A 50×10 · 1000A 2×50×5 · 1250–1600A 2×50×10 · 2000A 3×60×10 · 2500A 2×100×10 · 3200A 3×100×10 · 4000A 3×120×10.
- Cıvata torku (8.8): M4 1,8/1,5 · M5 5/4,3 · M6 8/6,8 · M8 20/17 · M10 40/35 · M12 70/60 · M16 155/130 N·m (üretim/kontrol). XPAN kaynağı: ELAN RA 5/20.
- Kablo renkleri (XPAN sayfa 5, IEC 60757 kodları): güç fazları 400 V L1/L2/L3 **hepsi siyah (BK)**; N mavi (BU); PE yeşil-sarı; 230 V L siyah; 24 V AC L kahverengi (BN), N yeşil (GN); 24 V DC + kırmızı (RD), − beyaz (WH); harici gerilim turuncu (OG); akım trafosu L1/L2/L3 siyah/kahverengi/gri. Kumanda kablosu 0,75 mm², halojensiz, H07V-K.
- TEKLED seçimi gereken alanlar: `genel.ic_faz_renk` = "IEC" (BN/BK/GY) veya "BK" (XPAN gibi hepsi siyah); `genel.kumanda_renk` = serbest metin (örn. "~230 В: L чёрный, N синий; =24 В: + красный, − белый"). Boşsa PDF'te KONTROL (G03/G04).
