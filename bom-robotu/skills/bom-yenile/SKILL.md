---
name: bom-yenile
description: YSF HICH BOM Excel dosyası için YENİLE veya DÖNÜŞTÜR istendiğinde güncel proje talimatına göre geri okuma, çalışma verisi kontrolü ve yeni revizyon üretimini yürütür.
---

# BOM Yenile

Bu skill yalnız akıştır. Kural kaynağı kullanıcının güncel YSF HICH — BOM ROBOTU proje talimatıdır. Atıflar v26.11 düzenine göredir. İşleme başlamadan güncel talimatın §1, §2, §3, §7, §8.4, §8.5 ve §9 bölümlerini oku. Çelişkide talimatı uygula ve farkı bildir; teknik kuralları bu skill’e taşıma.

## Adım 0 · Kaynakları hazırla

Talimat §3’e göre proje betiğini ve veri dosyalarını bul. `bom_kur_v*.zip` sürümünü güncel talimattan kontrol et. Çalışma klasörünü ve `BOM_VERI` değişkenini hazırla. Kullanıcının Excel dosyasını kaynak olarak koru, işlem için çalışma kopyası kullan. Betik ve veri dosyalarını skill paketine alma.

## Adım 1 · Geri oku

Komutu talimat §1’e göre ayır: YENİLE için `--yenile`, DÖNÜŞTÜR için `--donustur` akışını kullan. Çalışma JSON’unu talimatın belirttiği çalışma alanında oluştur; çıktı/paylaşım hükümleri §9 Adım 10 ve §11’dedir.

## Adım 2 · Geri okunan veriyi kontrol et

Talimat §8.5 kontrol listesini uygula: değiştirilmiş artikülün eski özellikleri, elle fiyatlar, artikül son ekleri, eski kontrol kayıtları, teklif bilgileri, pano/bara alanları ve sürüm göçleri. Kullanıcının değişikliklerini §2’ye göre koru; standart eski varsayılan ile kullanıcının bilerek değiştirdiği değeri ayırt et.

Mutabakat ve ELLE kayıtlarını §7 ve §8.5’e göre denetle. Yeni şema verilmişse ilgili pano için §5.1 sayım akışını yeniden yürüt; şema değişikliğini yalnız elle adet değişikliği gibi işleme. Eski dosyaların muafiyet durumunu güncel talimattan kontrol et. Eksik pano adlandırma alanlarını §7 üzerinden tamamla.

## Adım 3 · Revizyonu üret

Talimat §9’daki üretim ve yeniden hesaplama akışını uygula; çıktıyı yeni revizyon adıyla kaydet. HATA durumunda talimatın çalışma verisini düzeltme akışını izle; betik veya oluşturulmuş Excel üzerinde başka bir betikle düzeltme yapma. Kullanılabilir olmayan bir yeniden hesaplama aracını çalıştırılmış gibi gösterme.

## Adım 4 · Kontrol et ve teslim et

Talimat §9 Adım 8 listesini ve §8.4’teki değişiklik zincirini denetle. Özellikle sarf kapsamı, otomatik satır kapasitesi, canlı mutabakat ve kullanıcının değiştirdiği alanların korunmasını kontrol et. Düzeltme döngüsünü ve teslim biçimini §9 / §10’dan al. Tamamlanmayan bir kontrolü geçmiş gösterme.
