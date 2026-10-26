# 06 — TEKLED SABİT STANDARTLARI (Yusuf onayı, 08.10.2026)

Bu seçimler `bom_oku.py` tarafından her yeni projenin `genel` alanına otomatik yazılır. Bir projede farklı olması gerekiyorsa yalnız o projenin JSON'unda değiştirilir ve ARA LİSTE 2'de belirtilir.

| Konu | Seçim | JSON |
|---|---|---|
| Bara renkleri | ПУЭ 1.1.30: L1 жёлтый, L2 зелёный, L3 красный; N синий; PE жёлто-зелёный | `bara_renk: "ПУЭ"` |
| Pano içi faz kabloları | IEC 60445: L1 коричневый, L2 чёрный, L3 серый | `ic_faz_renk: "IEC"` |
| Kumanda kabloları | ~230 В: L чёрный, N синий; =24 В: + красный, − белый; ~24 В: L коричневый, N зелёный (XPAN pratiği) | `kumanda_renk` |
| İç kablo kesit tablosu (In→mm²) | PDF'e basılır (XPAN değerleri, docs/05) | `kesit_tablosu_onayli: true` |
| Cıvata tork tablosu | PDF'e basılır (XPAN değerleri, docs/05) | `tork_tablosu_onayli: true` |
| Malzeme listesinde sarf (bara, kablo, klemens, kanal) | Gizli — yalnız cihazlar ve korpus | `sarf_goster: false` |
| Logo / firma | TEKLED | `proje.firma: "TEKLED"` |
| Разработал / Проверил | Boş — elle imzalanır | `hazirlayan: ""`, `kontrol_eden: ""` |

Değişiklik gerekirse: bu dosyayı ve `tools/bom_oku.py` içindeki `TEKLED_STANDART` sözlüğünü birlikte güncelle.
