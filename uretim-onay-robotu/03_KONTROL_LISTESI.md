# 03 — TESLİM ÖNCESİ KONTROL LİSTESİ

Her PDF teslim edilmeden önce sırayla uygulanır. Bir madde geçmezse düzelt → yeniden üret.

## A. Veri
- [ ] Pano sayısı = BOM'daki pano başlığı sayısı (bom_oku çıktısı ile karşılaştır)
- [ ] Her panonun adedi (компл.) BOM ile aynı
- [ ] Rapor: R03/R04 (BOM↔tablo) maddesi yok veya bilinçli olarak STOP bırakıldı
- [ ] `belirsiz` kategorili satır kalmadı (ya sınıflandı ya KONTROL olarak raporda)
- [ ] Doldurulan her değerin kaynağı ARA LİSTE 2'de gösterildi; kaynaksız değer yok
- [ ] Giriş cihazı her panoda `grup: giris`
- [ ] Kapı elemanlarının renkleri artikülden doğrulandı

## B. Belge
- [ ] Kapaktaki durum bandı raporla aynı (STOP / KONTROL sayıları)
- [ ] Sayfa numaraları "n / N" doğru, kapaktaki sayfa aralıkları doğru
- [ ] Revizyon no: dosya adı = PDF antedi = rev_gecmisi son satırı
- [ ] Fiyat / iskonto / işçilik hiçbir sayfada yok
- [ ] Kazakça yok; Türkçe yalnızca özel ad olarak (firma adı)
- [ ] Tablolar sayfa altına taşmıyor, antetle çakışma yok (en az 3 sayfa PNG ile görsel kontrol: kapak, bir titul, en kalabalık devre tablosu)

## C. Teknik mantık
- [ ] Çıkış In ≤ giriş In; MCCB/MCB Icu ≥ Ik (Ik biliniyorsa)
- [ ] УЗО/АВДТ: IΔn ve tip dolu
- [ ] DIN doluluk yedeği ≥ %20 veya KONTROL olarak işaretli
- [ ] Çok bölmeli panolarda bölme sırası ve cihaz dağılımı şemaya göre
- [ ] Korpus çelişkisi varsa STOP olarak reestrde

## D. Teslim mesajı (Türkçe, kısa)
```
<dosya adı> — <N> sayfa, <P> pano
Durum: <durum> | STOP: <s> | KONTROL: <k>
Müşteriye sorulacak kritik maddeler: K-xx, K-yy
```
