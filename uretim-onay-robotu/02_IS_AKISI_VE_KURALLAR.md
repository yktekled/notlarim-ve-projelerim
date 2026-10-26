# 02 — İŞ AKIŞI, JSON ŞEMASI VE KONTROL KURALLARI

## 1. BOM okuma (`bom_oku.py`)
Şablon "fiyat burda" A–O. Okunan kolonlar: C (tip), D (Артикул), E (Номенклатура), I (Кол.), J (Ед.), K (завод/marka). Fiyat kolonları okunmaz.
- **Pano satırı**: J içinde "компл", D boş, C bilinen malzeme tipi değil, metin montaj/işçilik değil → E = pano adı, I = adet.
- "Итого / toplam" satırları atlanır.
- **Kategori**: önce C kısaltması (korpus, TMŞ→MCCB, OTO→MCB, DİF→УЗО, İmpuls/Röle→реле, Sinyal Lambası→лампа, Klemens/LOOBAR→sarf, rubilnik, tt, sayac, opn, montaj, işçilik), yoksa E metninden regex (NXM/NM1/ВА88→MCCB, NXB/ВА47→MCB, NXBLE/АВДТ→RCBO, NXC/КМИ→kontaktör, NXR→termik, NP8…B→buton, ND16→lamba…). Bulunamazsa `belirsiz` + KONTROL.
- Cihazlar adet kadar devre satırına açılır, otomatik poz: QF, QFD, QD, QS, KM, KK, FV, FU, PI, TA, HL, SB, SA, UZ, KL. **Poz'lar şemada farklıysa şemadaki poz kullanılır.**
- Nomenklatürden çekilen: In, kutup (1P/3P/1P+N), eğri (B/C/D), Icu kA, IΔn mA, УЗО tipi. Çekilemeyen alan null.
- Korpus satırından H×W×D ve IP çekilir.

## 2. JSON şeması (özet)
```json
{
 "proje": {"no":"261008","sehir":"Актау","musteri":"…","obje":"…","rev":"R0","tarih":"08.10.2026",
           "firma":"TEKLED","kaynak_bom":"….xlsx","kaynak_sema":"….pdf","hazirlayan":"","kontrol_eden":"",
           "rev_gecmisi":[{"rev":"R0","tarih":"08.10.2026","aciklama":"Первый выпуск"}]},
 "genel": {"bara_renk":"ПУЭ|IEC|null","ic_faz_renk":"IEC|BK|null","kumanda_renk":"metin|null","kesit_tablosu_onayli":false,"tork_tablosu_onayli":false,"sarf_goster":false},
 "panolar": [{
   "kod":"P-01","ad":"ЩР1","aciklama":"…","adet":1,
   "teknik":{"Ue":"400/230 В","Ui":"690 В","f":"50 Гц","sistem":"TN-S","fazlar":"3P+N+PE","In_pano":160,
             "Ik_proje_kA":20,"Icw_kA":null,"kumanda":"~230 В","form":"1","IK":"IK08","klimatik":"УХЛ4",
             "yer":"внутри помещения","bara":"Cu","standart":"ГОСТ IEC 61439-1-2024, ГОСТ IEC 61439-2-2015"},
   "korpus":{"ad":"…","artikul":"…","marka":"…","H":1320,"W":750,"D":300,"IP":"IP54","montaj":"навесной",
             "RAL":"7035","kablo_giris":"снизу","din_sira":2,"sira_modul":24,"ek_modul":0,"kapi":"…","kilit":"…"},
   "devreler":[{"poz":"QF1","grup":"giris|cikis","kategori":"mccb","cihaz":"…","artikul":"…","marka":"CHINT",
                "In":160,"kutup":"3P","egri":null,"Icu_kA":36,"rcd_mA":null,"rcd_tip":null,"modul":null,
                "yuk":"Ввод","kablo":"ВВГнг-LS 5x95","ic_kablo":null,"P_kW":null,"sira":1,"not":""}],
   "kapi_elemanlari":[{"poz":"SB1","kategori":"buton","cihaz":"…","artikul":"…","renk":"зелёный","yazi":"Пуск"}],
   "malzemeler":[ … bom_oku çıktısı … ],
   "notlar":["pano özel montaj notu"],
   "kontrol":[{"seviye":"STOP|KONTROL","mesaj":"…","poz":""}]
 }]
}
```
Kurallar: bilinmeyen = `null`. `grup:"giris"` en az bir cihazda. `modul` sadece katalog değeri farklıysa (örn. RCBO 1P+N tek modül tipleri) girilir. `ek_modul` = DIN'de yer kaplayan ama devre olmayan (zaman rölesi, OPN sigortası vb.).

## 3. Otomatik kontrol kuralları (`uretim_onay.py`)
| Kod | Seviye | Kural |
|---|---|---|
| G01 | KONTROL | Bara renk kodu seçilmemiş (ПУЭ ↔ IEC) |
| G02 | KONTROL | Tek hat şeması yok — hat adları/kablolar yalnız BOM'dan |
| G03 | KONTROL | Pano içi faz iletken rengi seçilmemiş (IEC BN/BK/GY ↔ hepsi BK) |
| G04 | KONTROL | Kumanda devresi iletken renkleri tanımsız |
| R01 | STOP | Zorunlu veri yok: H, W, D, IP, Ue, topraklama sistemi, InA |
| R02 | KONTROL | Ik, montaj, RAL, kablo girişi, form, klimatik, Ui eksik |
| R03 | STOP | Artikül bazında BOM adedi ≠ devre+kapı tablosu adedi |
| R04 | STOP | Tabloda olup BOM'da olmayan artikül |
| R05 | STOP / KONTROL | Modül > kapasite (STOP); yedek < %20 (KONTROL); DIN sıra/modül tanımsız |
| R06 | STOP | Giriş cihazı Icu < proje Ik |
| R07 | KONTROL | Çıkış In > giriş In; giriş cihazı işaretlenmemiş |
| R08 | KONTROL | Pano içi iletken < kesit tablosu (yalnız tablo onaylıysa, `ic_kablo`) |
| R09 | STOP / KONTROL | СТОП butonu yeşil / ПУСК kırmızı / acil stop kırmızı değil (STOP); "Авария" lambası kırmızı değil (KONTROL); renk boş |
| R10 | STOP | Poz tekrarı |
| R11 | KONTROL | УЗО/АВДТ için IΔn veya tip (AC/A/B) yok |
| R12 | KONTROL | Çıkış hattının adı yok |
| M | elle | `kontrol` listesine girilen maddeler (korpus çelişkisi vb.) |

## 4. Elle STOP girilmesi gereken durumlar
- Korpus: proje tanımı ≠ BOM korpus satırı ≠ BOM notu (boyut / IP / marka).
- Şemadaki cihaz adedi ≠ BOM adedi (bom tarafında düzeltmeden önce).
- Kumanda gerilimi değişikliği (örn. 230VAC → 24VDC bobin).
- Marka/seri dönüşümü (Schneider→CHINT) sonrası Icu veya In karşılığı yoksa.
- Korpus büyütüldü/küçültüldü (Aktau ЯУО9601: 400×300×200 → 500×400×220): **KONTROL** + gerekçe notu.

## 5. Yasaklar
- JSON'a uydurma değer yazmak (IP, Ik, kablo, yük adı).
- PDF'i script dışında elle üretmek veya script'in sayfa düzenini her projede değiştirmek (değişiklik gerekiyorsa script'i güncelle, sürüm notu yaz, projeye yeniden yükle).
- Fiyat bilgisini PDF'e taşımak.
