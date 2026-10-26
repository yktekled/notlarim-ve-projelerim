---
name: bom-robotu
description: YSF HICH pano projesiyle BOM veya ARA istendiğinde güncel proje talimatına göre okuma, sayım, mutabakat ve Excel üretim akışını yürütür. YENİLE, DÖNÜŞTÜR ve regresyon TEST işlemleri ayrı skill kapsamındadır.
---

# BOM Robotu

Bu skill yalnız akıştır. Kural kaynağı kullanıcının güncel YSF HICH — BOM ROBOTU proje talimatıdır. Atıflar v26.11 düzenine göredir. Önce güncel talimatı oku; kuralları bu dosyaya kopyalama. Çelişki varsa talimatı uygula ve farkı bildir. Kullanıcının istediği kapsamı izle; yalnız envanter veya sayım istenmişse sonraki üretim adımlarına kendiliğinden geçme.

## Adım 0 · Kaynakları hazırla

Talimat §3 sırasıyla proje dosyalarını bul. `bom_kur_v*.zip` desenindeki betik arşivinin sürümünü güncel talimatla karşılaştır; sürümü skill içinde sabitleme. Betik ve veri dosyaları proje dosyaları olarak kalır, skill paketine alınmaz. Çalışma ve veri klasörlerini, `BOM_VERI` değişkenini ve sürüm kontrolünü §3’e göre hazırla. Eksik dosya ve sürüm uyuşmazlıklarını §3’e göre işle; bulunmayan bir araç veya dosyayı erişilmiş gibi gösterme.

## Adım 1 · Oku ve envanter çıkar

Talimat §4 ve §9 Adım 1’i uygula. Kaynak dosyaları, sayfa/layout bölgelerini ve panoları envantere al. Okuma yöntemini dosya türüne göre §4’ten seç. Sembolleri bitişik teknik etiketleriyle birlikte incele; okuma sorunlarını talimattaki ilgili kontrol akışına taşı.

## Adım 2 · Pano pano say ve mutabakat kur

Talimat §5, §5.1 ve §9 Adım 2 / Adım 2a / Adım 2b’yi izle.

Tur 1 · Bara bölümü envanterini çıkar, besleyen cihaz ve giriş korumasını belirle; bölüm bölüm, hat hat COUNT_A sayımını yap.

Tur 2 · Cetvel dışı cihaz taramasını tamamla. BOM’a ve Tur 1’e bakmadan kaynak şemayı cihaz gruplarına göre yeniden okuyarak COUNT_B sayımını yap. Aynı çıkarılmış listenin yeniden gruplanması bağımsız ikinci okuma yerine geçmez.

Fark varsa ilgili bölgeyi yeniden oku ve §5.1’deki çözüm akışını uygula. Eksik ikinci okumayı yapılmış, kısmi sayımı tüm pano için tamamlanmış gösterme.

Her cihaz grubunun ŞEMA · CETVEL · BOM · COUNT_A · COUNT_B · lokatör mutabakatını §5.1 ve §7’ye göre kur. Kaynak izlerini grubun bütün cihazlarını kapsayacak şekilde denetle. ŞEMA ile BOM uyuşmuyorsa §9 Adım 2b üzerinden sayıma dön.

## Adım 3 · Seçim, yerleşim ve ARA çıktısı

Artikül seçimi, korpus, sıra planı, bara ve kumanda işlemlerini talimat §6 ve §7 üzerinden yürüt. Projede okunan değerlerle seçim sırasında verilen kararları ayırt et; bunların açıklama ve kontrol kayıtlarını ilgili talimat bölümüne göre oluştur.

Komut ARA ise §1 ve §9 Adım 3’teki pano listesini, mutabakatı, bara kararını ve kontrolleri sun. Excel üretimine geçmeden bu adımın gerektirdiği kullanıcı onayını bekle. BOM komutuna ayrıca ARA onayı ekleme.

## Adım 4 · Çalışma verisini hazırla

Talimat §7 ve §9 Adım 4’e göre çalışma JSON’unu yalnız çalışma klasöründe oluştur. Pano adları, proje eki, kaynak izleri, mutabakat ve kontrol kayıtlarını ilgili alan tanımlarından denetle. Kullanıcının daha önce verdiği teklif bilgilerini taşı. Dosyanın paylaşımı ve saklanmasında §9 Adım 10 ile §11’i uygula.

## Adım 5 · Excel’i betikle üret

Talimat §9 Adım 5 / Adım 6’daki üretim ve hata akışını yürüt. Doğrulama HATA’sını kaynak ve çalışma verisini düzelterek ele al. Python hatası ile doğrulama HATA’sını ayır; betik üzerinde izin verilmemiş değişiklik yapma. Düzeltme sınırlarını §9’dan al.

## Adım 6 · Yeniden hesapla ve doğrula

Talimat §9 Adım 7 / Adım 8 / Adım 9’u uygula. Kullanılan yeniden hesaplama aracının mevcut ve çalışmış olduğunu doğrula; erişilemeyen aracı çalıştırılmış gibi bildirme.

Adım 8 listesini kaynağından kontrol et: sayım mutabakatı, cihaz/EBAT özellikleri, pano ve tel/uç kapsamı, korpus yedeği, bara, tarak, sıra planı, otomatik kontroller, özet ve teklif. Sorun varsa talimatın düzeltme döngüsüne dön. Yalnız dosyanın oluşmasını kontrol başarısı sayma.

## Adım 7 · Teslim

Talimat §9 Adım 10 ve §10’a göre teslim et. Tam BOM üretimi ile kullanıcının yalnız bir ara adımı istediği çalışmayı ayırt et; tamamlanmayan üretimi tamamlanmış gibi sunma.
