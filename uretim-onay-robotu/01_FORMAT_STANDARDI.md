# 01 — FORMAT STANDARDI (ПРОИЗВОДСТВЕННОЕ СОГЛАСОВАНИЕ)

Kaynak: TEKLED örnekleri (260925 Aktau R1, 261006 Atyrau R4) + XPAN Pano Sistemleri EPLAN dosyası (004-ЩОЛ-1). Üretici: `tools/uretim_onay.py`.

## Genel
- A4 yatay, dil RU, font DejaVu (Kiril). Logo TEKLED (gömülü).
- Her sayfa: üst başlık (logo, belge adı, "Проект № | şehir | müşteri | P-xx | pano"), bölüm adı, sağda `R | tarih | DURUM`.
- Her sayfa altında antet: Разработал / Проверил / Производство / ОТК / Согласовано заказчиком (Ф.И.О./подпись/дата) + pano kodu, sayfa adı, kaynak BOM, Проект / Лист n/N / Рев. / Дата.
- Durum otomatik: STOP varsa **НЕ ЗАПУСКАТЬ В ПРОИЗВОДСТВО ДО ПОДТВЕРЖДЕНИЯ** (kırmızı); sadece KONTROL varsa **НА СОГЛАСОВАНИЕ (ЕСТЬ ВОПРОСЫ)**; hiç yoksa **НА СОГЛАСОВАНИЕ**.

## Proje sayfaları (1 kez)
| Sayfa | İçerik |
|---|---|
| Kapak | Müşteri, proje no/şehir/obje, durum bandı (STOP/KONTROL sayısı), pano tablosu (№, kod, ad, gabarit, IP, giriş, adet, sayfa aralığı, STOP/KONTROL), revizyon geçmişi, kaynaklar ve notlar |
| Реестр вопросов | Tüm STOP/KONTROL maddeleri K-01… numaralı, "Решение заказчика" boş kolon |
| Общие технические требования | Renk kodu (iletken + bara), ГОСТ 2.710 harf kodları, etiketleme (61439-1 §6.1), rutin testler (§11) + ОТК kutusu, iç kablo kesit tablosu ve tork tablosu (yalnız onaylıysa) |

## Pano sayfaları (her pano 5 sayfa; uzun tablolar devam sayfası açar)
| # | Sayfa | İçerik |
|---|---|---|
| 1 | Титульный лист щита | Pano adı/açıklama, oranlı ön görünüş (kapı elemanları renkli), IEC 61439 teknik parametre tablosu (19 satır), "ДО ЗАПУСКА" kutusu, müşteri imza satırı |
| 2 | Габаритные виды | Ön/yan/üst/alt görünüş ölçülü, kablo girişi, kapı-kilit-PE bilgisi, kapı elemanları tablosu (поз/eleman/renk/yazı) |
| 3 | Внутренняя компоновка | Montaj plakası: üstte DIN-dışı bloklar, DIN sıraları (modül hesaplı), N/PE bara; doluluk tablosu (≥%20 yedek), montaj gereklilikleri, STOP kutusu |
| 4 | Таблица линий и маркировки | Поз, cihaz, artikul, marka, In, kutup, eğri, Icu, УЗО IΔn/tip, hat/yük, P, kablo, etiket metni. Sorunlu poz kırmızı |
| 5 | Комплектация + ОТК | Fiyatsız malzeme listesi (montaj/işçilik ve varsayılan olarak sarf hariç) + 10 maddelik çıkış kontrolü |

## Teknik parametre tablosu (pano titul) — alanlar
Наименование, Количество, Корпус (тип/арт./изготовитель), Габариты В×Ш×Г, Установка, IP / IK, Цвет RAL, Климатическое исполнение / место, Ue / Ui, Частота, Система заземления / фазы, InA, Вводной аппарат, Ik / Icw, Напряжение управления, Форма секционирования, Ввод/вывод кабелей, Материал шин, Нормативный документ.
Zorunlu (boşsa STOP): Габариты, IP, Ue, Система заземления, InA.

## Dosya adı
`<YYMMDD>_<Şehir>_<Müşteri>_Uretim_Onay_RU_R<n>.pdf` (örn. `261008_Aktau_BATYS_TEHNO_Uretim_Onay_RU_R0.pdf`)

## Örneklerden alınan dersler (Atyrau R1→R4 düzeltmeleri)
- Çok bölmeli ВРУ: bölme sırası (sol–giriş–sağ) ve cihazların bölmelere dağılımı şemadan alınır; şemada yoksa KONTROL.
- Tek BOM satırında birden fazla pano (ЩОА-1/ЩОА-2, Я1/Я2) → ayrı pano kaydı, adet 2.
- Buton/lamba renkleri BOM artikülünden gelir (NP8-10BN/3 = yeşil, NP8-01BN/4 = kırmızı); şema metni ile çelişirse KONTROL.
- Gabariti olmayan modüler kutular (NEXUS vb.): "габарит по каталогу" + KONTROL, delik açma onaydan sonra.
- Korpus çelişkisi (proje adı ≠ BOM satırı ≠ not): STOP — "до резки не запускать".
