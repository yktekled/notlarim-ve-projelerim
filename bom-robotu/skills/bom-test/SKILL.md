---
name: bom-test
description: YSF HICH için TEST istendiğinde veya kullanıcı bom_kur.py sürümünü değiştirme çalışması başlattığında proje test paketiyle regresyon akışını yürütür ve beklenen sonuçların güncellenmesini yönetir.
---

# BOM Test

Bu skill yalnız akıştır. Kural kaynağı kullanıcının güncel YSF HICH — BOM ROBOTU proje talimatıdır. Atıflar v26.11 düzenine göredir. Başlamadan §3 ve §9.1’i oku; betik sürümü ve test kapsamını güncel talimattan al. Test talebini kendiliğinden betik değişikliği izni olarak yorumlama.

## Adım 1 · Test ortamını hazırla

Talimat §3 ve §9.1 sırasıyla proje betiğini, veri dosyalarını ve `bom_test.zip` paketini bul. `bom_kur_v*.zip` sürümünü talimatla karşılaştır; sürümü skill’e sabitleme. Arşiv içeriğini ve beklenen sonuç dosyasını kontrol ederek §9.1’deki çalışma/test yerleşimini hazırla. Test verisini skill paketine gömme; eksik paketi yeniden yazma. Eksiklik halinde §9.1’in durma akışını uygula.

## Adım 2 · Testi çalıştır

Talimat §9.1’deki test komutunu çalıştır. Test betiğinin üretim, yeniden hesaplama ve beklenen sonuç karşılaştırmasını gerçekten tamamladığını doğrula. Karşılaştırma alanlarını ve örnekleri §9.1’den al; komutun başlamasını veya dosya oluşmasını başarılı regresyon sayma. Çalışmayan aracı, eksik örneği ve çalıştırma hatasını sonuçta açıkça ayır.

## Adım 3 · Sonucu işle

Farkları ilgili test ve ölçütleriyle raporla. Fark yoksa §9.1’e göre kısa sonucu ver.

Fark varsa §9.1’in kullanıcıya farkları gösterme ve onay akışını izle. Beklenen sonuçları gerçek çıktıya eşitleyerek testi kendiliğinden geçirme. Yeni sürümü projeye yazma ve `--kaydet` ile beklenen sonuçları değiştirme işlemlerini yalnız talimatın gerektirdiği kullanıcı onayından sonra yap.

Onay verilmişse §9.1’e göre beklenen sonuç dosyasını güncelle, test paketini yeniden oluştur ve teslim et. Projedeki paket ile talimattaki betik arşivi adının güncellenme durumunu bildir. Yalnız test verisi için geçerli teslim istisnasını üretim JSON’una uygulama.
