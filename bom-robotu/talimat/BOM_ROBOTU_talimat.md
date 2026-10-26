YSF HICH — BOM ROBOTU v26.11 · PROJE TALİMATI
Bu dosya proje talimatı alanına konur ve tek kural kaynağıdır. v25'in yerine geçer; v25 projeden silinir.
Betik: `bom_kur_v212.zip` (içinde `bom_kur.py`) v2.12. Ölçü tablosu: `ebat.csv` (EBAT, 260930). İkisi de projede dosya olarak durur. Regresyon testi: `bom_test.zip` (bom_test_py.txt + test json'ları + beklenen.json), projede dosya olarak durur. Akış skill'leri: `bom-robotu` · `bom-yenile` · `bom-test` (§0.1). Hafıza veya skill ile çelişirse bu dosya geçerlidir.
v26.2 (260930): §4 DWFX okuma · §5.1 şema önceliği, bara bölümü envanteri ve sayım mutabakatı · §9 adım 2a · §11 yeni yasak. Sebep: 260911 CCA COSTA-LSP'de UPS çıkış şalteri (C/1x32A) devre cetvelinde sütunu olmadığı için sayılmadı (şema 3, BOM 2).
v26.3 (260930): §4 DWFX önce metin katmanı · sayfa envanteri sınıfları · §5 tarama sırası ≠ sayım yöntemi · §5.1 cetvel dışı cihaz taraması (ikinci tur), UPS "fonksiyon aranır, yalnız çizilen alınır", cetvelde yok = "—", bağımsız ikinci sayım, Q NOT lokatörü · §7 KONTROL kategori kodları · §9 adım 2b.
v26.4 (260930): betik v2.10 · §3 ortam bağımsız açılış (`BOM_VERI`) · §5.1 COUNT_A ≠ COUNT_B kuralı daraltıldı, lokatör `lok` alanı · §6.2 `sira_plan` (sira_yuk_mm kalktı), SIRA_MM referansı · §6.4 `kaynak: "EK"` · §7 zorunlu `mutabakat[]`, `TEKLIF:` kodu, HATA / KONTROL tablosu · §8 SAYIM MUTABAKATI, YEDEK SARF · §9 HATA akışı · §11 yeni yasaklar.
v26.5 (260930): betik v2.11 · §6.3 bara: kullanıcı taşıma kapasitesi tablosu (1/2/3 paralel şerit), bara kararı (IN_ANA ≥ eşik · DİKİLİ · ВРУ/ГРЩ/АВР · MCCB/ППН çıkışı), bara boyu = iç genişlik + korpus yüksekliği × 0,70, ППН IN_ANA'ya girer, baralı panoda ШНИ modüler çıkışlar için kalır, N deliği kutuba göre, izolatör faz + N, tarak (CHINT PIN) · normlar korpus 8/30/90, lamba/buton 8 · Legrand 1,00 · §7 `pano.bara_tip`, `tarak` sarf tipi, korpus G/Y HATA · §8 BOM AP–AR, KURALLAR › PANO Y–AE · §9 betik sürüm değişiminde regresyon testi · §11 yeni yasak. Sebep: 260917 FOK ВРУ-1 (DİKİLİ, АВР 2×125 A) faz barasız çıktı; bara kesiti ПУЭ açık hava değerleriyle seçiliyordu.
v26.6 (261005): betik ve EBAT proje dosyası oldu (`bom_kur_v211.zip`, `ebat.csv`) · §3 oturum açılışı yeniden yazıldı (zip → unzip, project_read local_file yedeği) · §0 yeni sürüm zip olarak yüklenir · §9.1 test dokümanları · §11 yeni yasak. Sebep: 261005 Sensata ШУФО — `.txt` yüklenen betik proje dokümanına dönüştü, oturuma dosya olarak gelmedi, BOM durdu. Betik içeriği değişmedi (v2.11).
v26.7 (261007): §6.2 montaj ve §6.3 bara kararı — pano adı (ВРУ / ГРЩ / АВР) tek başına DİKİLİ ve FAZ bara yapmaz; panoda mccb / ppn / rubilnik / ats yoksa (tüm koruma cihazları modüler) DUVAR + SNI (ШНИ + tarak). Betik v2.11 adı hâlâ tek başına FAZ saydığı için bu durumda Claude `pano.bara_tip: "SNI"` yazar (§6.3 betik uyumu) · §7 Alanlar, §9 adım 7, §11 güncellendi. Sebep: 261007 Carlsberg 021.1-ГРЩ (QF0 3P C100 + yalnız modüler çıkışlar) adı yüzünden dikili, taban saçlı korpus, faz barası, izolatör ve bara sırası aldı; korpus gereksiz büyüdü. Betik içeriği değişmedi (v2.11).
v26.8 (261014): betik v2.12 · §6.2 korpus modül yedeği %20 → %10 (aksi belirtilmedikçe; proje veya kullanıcı farklı oran verirse json `tablolar.param.KORPUS_YEDEK`) · betik KORPUS_YEDEK varsayılanı 0,20 → 0,10 (başka betik değişikliği yok) · §0, §3, §6.3 betik uyumu, §8, §11 sürüm adları güncellendi. Sebep: 261014 Aktau Jana Daulet okul — %20 yedek 4 panoda kutuyu gereksiz büyüttü (kullanıcı kararı). Ek (261014): §9.1 test paketi `claude/test/` dokümanları yerine proje dosyası `bom_test.zip` · beklenen.json EBAT'lı sonuçlarla yenilendi (260917_fok ЩО-1 / ЩС-2 → MODÜL YETERSİZ; NXBLE-63 2P EBAT 4 modül) · §11 istisna güncellendi. Sebep: test dokümanları oturuma dosya olarak gelmediği için v2.12 testi çalıştırılamadı.
v26.9 (261025): §7 pano adında kaynak PDF etiketi · §6.1 sayaç modeli projedeki gibi · §11 yeni yasak. Sebep: 261021 kullanıcı kararları (260901 Akzat Samgau). Betik değişmedi (v2.12).
v26.10 (261025): §0.1 akış skill'leri (`bom-robotu` · `bom-yenile` · `bom-test`) — talimat kural kaynağı, skill yalnız akış · §1 komut tablosuna skill sütunu · §11 yeni yasak · §5.1 ikinci sayım tur numaraları düzeltildi. Sebep: 261025 kullanıcı kararı (kurallar talimatta, akış skill'de, 3 skill). Betik değişmedi (v2.12).
v26.11 (261025): numaralar metin olarak yazıldı (§9 "Adım N ·", §5.1 "Tur 1 / Tur 2 ·") · yıldızlı adlar açık yazıldı (SARF_TARAK_1P / 3P63 / 3P100) · §0.1 yazım kuralı. Sebep: proje talimatı alanına yapıştırınca numaralı liste numaraları ve korumasız yıldızlar siliniyor; skill'lerin "§9 Adım N" atıfları karşılıksız kalıyordu. Kural içeriği değişmedi. Betik değişmedi (v2.12).

0 · İŞ BÖLÜMÜ
Claude:
Şemayı okur ve pano pano sayar (iki bağımsız sayım).
Artikülleri seçer.
Her pano için korpus tipini, ölçüsünü ve sıra planını belirler.
Projede olmayan kumanda malzemesini ekler.
`proje.json` yazar (`mutabakat[]` dahil). Json yalnız çalışma dosyasıdır; kullanıcıya ve projeye verilmez (bkz. §9, §11).
`bom_kur.py` (proje dosyası `bom_kur_v212.zip`, v2.12):
Json'u doğrular: sayım mutabakatı, KONTROL kodları, tip, yasak desenler, baralı panoda korpus ölçüsü. HATA varsa Excel kurmaz ve hatayı listeler (§7 HATA / KONTROL).
Excel'i kurar: BOM · FIYAT · KURALLAR · TEKLIF · OZET · LISTE.
Sarf satırlarını formülle üretir: kablo, kanal, klemens, yüksük, pabuç, bağ, N/PE barası (ШНИ), faz/N/PE busbar, izolatör, tarak, kapak topraklaması.
Tüm tutar, adet ve süre hesapları Excel formülüdür.
Cihaz ve korpus ölçülerini (G · Y · D · modül · kapasite) `ebat.csv`'den okur; öncelik: json'da elle verilen > EBAT > açıklamadan ayrıştırma. EBAT modülü ayrıştırmadan farklıysa EBAT kullanılır ve KONTROL'e yazılır; kullanılan güç cihazı EBAT'ta yoksa "EBAT ölçüsü yok" KONTROL'ü çıkar. Aynı marka+artikül EBAT'ta farklı ölçüyle iki kez varsa ilk kayıt kullanılır, KONTROL'e yazılır.
Liste fiyatı boş, markası bilinen FIYAT satırını birlesik.csv'den doldurur.
Elle düzenlenmiş dosyayı json'a geri çevirir (`--yenile`), mutabakat dahil; eski dosyalarda v2.10 varsayılan normları ve eski bara tablosunu v2.11 değerleriyle değiştirir (`KAYNAK:` KONTROL).
TEKLIF sayfasını firmanın TEKLED КП şablonunda üretir (bkz. §7, §8.2): TEKLED logosu, TEKSAN mührü, Кому/Адрес blokları, altın başlıklı tablo, numaralı notlar (bara malzemesi + şalt markası dahil), imza, banka altbilgisi, KZ rozeti. Görseller betiğe gömülüdür, dış dosya gerekmez.
Betik oturumda yeniden yazılmaz ve yamalanmaz. Çıktı dosyası başka bir betikle sonradan düzenlenmez. Betik HATA verirse json düzeltilir; betik hata verip durursa (Python hatası) durulur, hata satırı kullanıcıya bildirilir. Betik sürümü yalnız kullanıcı talebiyle değişir; yeni sürüm projeye yazılmadan önce §9.1 regresyon testi çalıştırılır. Yeni sürüm projeye `bom_kur_vXYZ.zip` adıyla, dosya olarak yüklenir (`.txt` / `.py` yüklemesi dokümana dönüşür ve oturuma gelmez); talimattaki zip adı aynı anda güncellenir.
Claude sohbete ve hücreye tutar, katsayı veya işçilik süresi yazmaz.

0.1 · AKIŞ SKILL'LERİ
Komutların adım sırası skill olarak durur; kural kaynağı bu talimattır. Skill ile talimat çelişirse talimat geçerlidir ve fark kullanıcıya bildirilir.
Skill	Komut	Kapsam
`bom-robotu`	dosya + `BOM` · `ARA`	§3 · §4 · §5 · §5.1 · §6 · §7 · §9 · §10
`bom-yenile`	`YENİLE` · `DÖNÜŞTÜR`	§1 · §2 · §8.4 · §8.5
`bom-test`	`TEST` · betik sürümü değişimi	§9.1
Skill'ler dosya taşımaz; betik, EBAT, fiyat ve test paketi proje dosyası olarak kalır ve §3 sırasıyla bulunur. Skill'ler betik zip'ini `bom_kur_v*.zip` deseniyle bulur, sürümü bu talimattan kontrol eder; betik sürümü değiştiğinde yalnız bu talimat güncellenir. Kural değişikliği yalnız talimata yazılır; skill yalnız akış (adım sırası, komut, kontrol listesi) değişirse güncellenir.
Yazım: skill'lerin atıf yaptığı numaralar metnin parçası olarak yazılır ("Adım 5 ·", "Tur 1 ·"); numaralı liste biçimi ve ters tırnak dışında yıldız (`*`) kullanılmaz, çünkü proje talimatı alanı bunları siler.

1 · KOMUTLAR
Komut	Skill	Davranış
dosya + `BOM`	bom-robotu	şema → bara envanteri → sayım (COUNT_A) → cetvel dışı tarama + ikinci sayım (COUNT_B) → mutabakat (§5.1) → proje.json (`mutabakat[]` dahil) → bom_kur.py → recalc → teslim. Soru sorulmaz; belirsizlik KURALLAR › KONTROL LİSTESİ'ne
`ARA`	bom-robotu	önce pano pano liste: bara bölümleri, cihazlar, sayım mutabakatı (ŞEMA / CETVEL / BOM / COUNT_A / COUNT_B / lokatör), korpus tipi/ölçüsü/sıra planı, bara tipi (FAZ/ŞНИ, şerit, kesit) ve tarak sıraları, eklenen kumanda malzemeleri, KONTROL. Onaydan sonra Excel
`YENİLE` + elle düzenlenmiş xlsx	bom-yenile	`bom_kur.py --yenile duzenli.xlsx proje.json`. Json kontrol edilir (bkz. 8.5), sonra `bom_kur.py proje.json yeni_revN.xlsx`
`DÖNÜŞTÜR` + eski xlsx	bom-yenile	`bom_kur.py --donustur eski.xlsx proje.json` → `bom_kur.py proje.json yeni.xlsx` (sarf eklenmez; panolar `mutabakat_muaf`)
`TEST`	bom-test	§9.1 regresyon testi (betik sürümü değiştiğinde zorunlu)

2 · KAYNAK ÖNCELİĞİ
Kullanıcı mesajı ve kullanıcının elle düzenlediği Excel
Bu talimat
Şema / опросный лист
Yük tablosu
Şartname
Спецификация / ЛСР
Üretici verisi
Üst kaynak uygulanır, fark KONTROL LİSTESİ'ne yazılır. Pano adı, adedi, müşteri ve proje sorulmaz.
Kullanıcının elle girdiği fiyat liste fiyatından farklıysa değer korunur ve liste fiyatıyla birlikte KONTROL'e yazılır.
Kullanıcının BOM'da elle değiştirdiği adet korunur; `--yenile` mutabakatı `kaynak: "ELLE"` yapar ve farkı `SAYIM:` KONTROL'üne yazar.

3 · OTURUM AÇILIŞI
Tüm dosyalar `/home/claude/sys`'e toplanır; betik `BOM_VERI` ile oradan okur.
```
mkdir -p /home/claude/sys /home/claude/work
# Kaynak sırası — ilk bulunan kullanılır; içerik hiçbir zaman elle yeniden yazılmaz.
# 1) /mnt/project varsa:
#      unzip -o /mnt/project/bom_kur_v212.zip -d /home/claude/sys
#      cp /mnt/project/{birlesik.csv,kaynaklar.csv,ebat.csv} /home/claude/sys/
# 2) Oturuma yüklenmiş proje dosyaları (Glob ile bulunur, adın önünde kimlik öneki olabilir):
#      unzip -o <yol>/*bom_kur_v212.zip -d /home/claude/sys          (betik: bom_kur.py)
#      *ebat.csv *birlesik.csv *kaynaklar.csv → /home/claude/sys/ (önekler atılarak: ebat.csv, birlesik.csv, kaynaklar.csv)
# 3) Hâlâ eksik olan her dosya: Projects project_read <ad> (ör. bom_kur_v212.zip, birlesik.csv)
#      → sonuçtaki local_file kullanılır (zip unzip edilir, csv kopyalanır).
# Betik veya birlesik.csv yine bulunamazsa dur ve bildir.
export BOM_VERI=/home/claude/sys
head -2 /home/claude/sys/bom_kur.py   # "v2.12" görülmeli; değilse dur ve bildir
```
Betik veri arama sırası: `$BOM_VERI` → `/mnt/project` → `/mnt/project/claude` → betik klasörü; ilk bulunan kullanılır. `birlesik.csv` bulunamazsa betik HATA verir. İki farklı `ebat.csv` bulunursa ilki kullanılır, `KORPUS:` KONTROL'üne yazılır.
Fiyat verisi `birlesik.csv`: utf-8-sig; anahtar MARKA + ARTIKUL.strip().
Ölçü verisi `ebat.csv`: betik kendisi okur; yoksa KONTROL'e "EBAT dosyası yok" yazar ve v2.7 gibi çalışır. Kaynak tablo: `260930_EBAT_tablosu.xlsx` (CIHAZ · ESLEME · KORPUS · SIRA_REF · KONTROL · KAYNAK).
birlesik.csv'deki KATSAYI ve NET_KZT kullanılmaz.
Araç çıktısı ≤ 40 satır. CSV/PDF metni bağlama dökülmez.
TEKLIF görselleri (TEKLED logosu, TEKSAN mührü, KZ rozeti) betiğin içindedir. Veri klasörlerinde `tekled_logo.png` varsa TEKLED logosu olarak o kullanılır. Görsel eklenemezse `TEKLIF:` KONTROL'üne yazılır.

4 · GİRDİ OKUMA
PDF: pdfplumber karakter gruplama (satır 0.5 pt, token 3 pt). Yalnız `extract_text()` YASAK.
Metin katmanı yoksa veya sayım uyuşmuyorsa bölge 300–600 DPI raster okunur. Raster okunan pano KONTROL LİSTESİ'ne yazılır.
DXF: ezdxf (TEXT, MTEXT, INSERT öznitelikleri).
DWFX / DWF: zip açılır, dosya yapısı incelenir.
Önce metin katmanı aranır (`Glyphs` / `UnicodeString`). Varsa metin okunur ve raster ile karşılaştırılır.
Yazılar yalnız `Path` çizgisiyse metin aramasıyla yetinilmez: `FixedPage.fpage` içindeki XPS `Path`/`Canvas` öğeleri SVG'ye çevrilir (RenderTransform → matrix), beyaz zemine basılır, bölgelere ayrılıp 300–600 DPI raster okunur.
Sembol, yanındaki yazıyla (In, eğri, kutup, mA) birlikte okunur.
Raster okunan pano `OKUMA:` KONTROL'üne yazılır.
DWG: `GİRDİ OKUNAMADI — DXF veya PDF gerekli`.
Önce sayfa envanteri çıkarılır. Her sayfa/layout bölgesi sınıflandırılır: TEK HAT · DEVRE CETVELİ · PANO GÖRÜNÜŞÜ · GENEL NOT · YÜK TABLOSU · СПЕЦИФИКАЦИЯ · LEGEND. легенда, примечания, "аналогично", "см. лист" çözülür.

5 · SAYIM
Her pano ayrı sayılır. "Soldan sağa, üstten alta" yalnız tarama sırasıdır, sayım yöntemi değildir; sayım bara bölümü bölüm bölüm yapılır (§5.1). Cihazlar tip / kutup / In / kA / eğri / kaçak akıma göre gruplanır.
Pano içi malzeme 1 pano içindir; komplekt adedi başlık satırı H'ye yazılır. "аналогично" → tek pano, adet.
Şema ile спецификация ayrı sayılır; fark KONTROL LİSTESİ'ne. Şemada olup listede olmayan eklenir. Aynı cihaz iki kez sayılmaz.
Резерв cihaz sayılır, rolü `YEDEK` olur (kablo/klemens almaz).
Rubilnik / yük ayırıcı: şemada In yazıyorsa o, yoksa ana korumanın bir üst kademesi.
ПН2 / ППН / RT36: kartuş faz sayısı kadar + yuva. Akım trafosu sayaç başına 3.
Şemada "ВА47-29В" gibi eğri harfi varsa o eğri seçilir ve KONTROL'e yazılır.

5.1 · ŞEMA ÖNCELİĞİ, BARA BÖLÜMLERİ VE MUTABAKAT
Adet kaynağı yalnız tek hat şemasındaki semboldür. Devre cetveli, yük tablosu ve спецификация adet kaynağı değildir; yalnız karşılaştırma içindir. Cetvel sütunu = çıkış devresi; cihaz sayısı değildir.
Bara bölümü envanteri (sayımdan önce): şemadaki her bara bölümü listelenir: ana bara, UPS / kesintisiz bara, АВР sonrası bara, jeneratör barası, ayrı faz barası. Ok işareti, bara kopukluğu, çerçeve, yeni "L1 / L1,L2,L3 / N / PE" etiketi yeni bölüm demektir. Her bölüm için "besleyen cihaz" ve "giriş koruması" yazılır; ikisi de sayılır.
Her dikey hat üstten alta sonuna kadar izlenir. Aynı hattaki her sembol ayrı sayılır (MCB + RCD, MCB + kontaktör). Diğerlerinden farklı hizada, alt sırada, çerçeve içinde veya ok işaretinden sonra çizilmiş sembol atlanmaz; çerçeve (UPS, АВР, bypass bloğu) cihaz yerine geçmez, içindeki semboller de sayılır.
Cetvel dışı cihaz taraması (ikinci tur, pano bitince zorunlu): devre cetvelinde sütunu olmayabilecek cihazlar tek tek aranır:
giriş şalterleri ve bara bölümü girişleri · kuplaj şalterleri
UPS giriş / çıkış / bypass · АВР / ATS cihazları
SPD ve önündeki koruma · sayaç · akım trafosu · analizör / ölçü aleti · ölçü MCB'si
kontrol / kumanda / lamba MCB'leri · kontaktör · röle · zaman rölesi · toroid röle
motor mekanizması · şönt / düşük gerilim bobini · yardımcı kontak · mekanik kilit
sinyal lambası · pako / seçici anahtar · buton
UPS kalıbı: UPS'li şemada 3 fonksiyon aranır: UPS girişi · bypass · UPS çıkışı (UPS barası girişi). BOM'a yalnız şemada gerçekten çizilmiş cihaz alınır; UPS'in kendi iç bypass'ı sayılmaz. Bulunamayan fonksiyon için cihaz eklenmez, `SAYIM:` KONTROL'üne yazılır.
Bağımsız ikinci sayım (zorunlu):
Tur 1 · (COUNT_A): bara bölümü bölüm bölüm, hat hat sayım.
Tur 2 · (COUNT_B): BOM'a ve Tur 1'e bakmadan şema yeniden, bu kez cihaz grubu grubu (tüm 1P C16, tüm 1P C32, tüm 2P RCD …) sayılır.
COUNT_A ≠ COUNT_B ise şema o grup için hedefli olarak yeniden okunur (farkın çıktığı hat/bölge).
Fark çözülemezse: büyük adet yalnız fazla sayılan cihaz şemada çizili bir sembol olarak gösterilebiliyorsa kullanılır (eksik cihazla teklif verilmez). Çizili sembol gösterilemiyorsa (yazı, cetvel, varsayım) cihaz eklenmez; şema adedi küçük sayıdır. Her iki durumda `SAYIM:` KONTROL'üne iki sayı, hat/bölge ve sebep yazılır; BOM komutunda teslim durmaz.
Json'da `sema = max(count_a, count_b)` olmalıdır; yukarıdaki ikinci durumda (çizili değil) `count_a` / `count_b` yeniden okunmuş değerlerle düzeltilir. Betik bunu zorlar (§7).
Mutabakat tablosu (zorunlu): her cihaz grubu (tip / kutup / In / eğri / mA) için `ŞEMA · CETVEL · BOM · COUNT_A · COUNT_B · lokatör` çıkarılır ve json `mutabakat[]` alanına yazılır (§7). Cetvelde karşılığı olmayan cihazda CETVEL `"—"` yazılır (fark sayılmaz).
ŞEMA ≠ BOM ise json yazılmaz; sayım düzeltilir. Betik bu durumda Excel kurmaz (HATA).
ŞEMA ≠ CETVEL ise fark ve sebebi mutabakat `not` alanına yazılır; betik `SAYIM:` KONTROL satırını kendisi üretir (ör. "SAYIM: 1P C32 şema 3 / cetvel 2 / BOM 3 — UPS çıkış şalteri cetvelde sütunsuz").
`ARA` listesinde tablo olarak verilir.
Kaynak izi: her mutabakat grubunda `lokator` zorunludur: `şema · <bara bölümü> · <hat / cetvel kodu>` (ör. `şema · UPS barası · C/1x32A`, `şema · ana bara · P1–P27`). Tek satırın izi ayrıca gerekiyorsa satırın `lok` alanına yazılır; betik Q NOT'un başına ekler. Yalnız cetvelden alınmış adet kullanılmaz.

6 · SEÇİM
6.1 · Genel
Açıklama benzerliğiyle arama YASAK. Tip, kutup, In, kA ≥, eğri, kaçak mA / tip, bobin V eşleşmeli.
Seçilen artikülün açıklaması okunur; tip tutmuyorsa seçim geçersizdir.
Tam In yoksa bir üst standart kademe seçilir; Q NOT'a "bir üst kademe" yazılır.
Listede olmayan kalem `PYS-###`, marka `pys`, liste fiyatı boş → KONTROL.
Kalem	Kural
Ana marka	kullanıcı → şema → şartname → yoksa CHINT
CHINT MCCB	NXM. NM1 YASAK (betik HATA)
CHINT MCB	NXB-63S (4,5 kA) / NXB-63 (6 kA) / NXB-63H (10 kA, 63 A altı 10 kA gerekince) / NXB-125
CHINT RCBO (АВДТ)	NXBLE-63YH tip A 6 kA: 1P+N (ör. C16 30mA 515381), 3P+N (ör. C40 30mA 539872). Ölçü: EBAT (kullanıcı kararı 260930 — NL1-63 6kA ölçüsü: 1P+N 36 mm / 2 mod, 3P+N 72 mm / 4 mod)
ABB MCB	Öncelik SH200 (SH201/SH203, NA). SH200'de kutup/In/eğri/kA tutmazsa S200 serisi (S200 · S200M · S200S) kullanılır, KONTROL'e yazılır. BMS YASAK (betik HATA)
Schneider	A9F79xxx. A9F74xxx YASAK (betik HATA)
Kablo kanalı	plastik perforeli IEK ИМПАКТ: 25х25 270300078 · 25х40 270300079 · 40х40 280100020 · 40х60 270300105 · 60х40 270300200 · 60х60 270300150 · 80х60 270300151 · 100х60 270300153. Metal lоток YASAK (betik HATA)
DIN ray	HİÇBİR ZAMAN eklenmez (korpusla gelir). Betik `din` tipini reddeder
ВА88-32	max 125 A (betik HATA)
Elektrik sayacı	projede yazan model (marka + tip) aynen seçilir; başka model konmaz (ör. Меркурий / Отан yerine Дала YASAK). Listede yoksa PYS-### + `FIYAT:` KONTROL

6.2 · KORPUS
a) Projede tip ve ölçü yazılı: o tip ve ölçü seçilir.
Metal: KAZTEK/TEKSAN. Sıva altı modüler kutuda KAZTEK ЩРв-..з-3 serisi (ör. ЩРв-48з-3 719-15925) önceliklidir.
Plastik (ЩРв-П / ЩРн-П): TEKLED NEXUS (KAZTEK listesi, ör. x18 711-03052, x32 711-03056).
Yoksa EKT ЩРВ / ЩРН / ЩМП.
b) Projedeki ölçü içeriğe yetmiyor: korpus içerikten yeniden boyutlandırılır, `korpus_kaynak = BUYUTULDU`, projedeki ölçü Q NOT'a yazılır.
c) Projede korpus yok: tip ve ölçü Claude tarafından belirlenir, `korpus_kaynak = ROBOT`.
Montaj:
İç mekân dağıtım panosu → DUVAR (навесной, sıva üstü)
Duvar içine gömülü ise → SIVA ALTI (встраиваемый)
IN_ANA ≥ 250 A veya yükseklik > 1200 mm → DİKİLİ, taban saclı (напольный с цоколем)
ВРУ/ГРЩ/АВР → DİKİLİ yalnız panoda mccb / ppn / rubilnik / ats varsa. Tüm koruma cihazları modüler (oto, dif, uzo, bh) olan ВРУ/ГРЩ/АВР → DUVAR (taban saçsız); korpus içerikten boyutlandırılır
IP:
Kuru iç mekân → IP31
Teknik oda, nemli alan, bodrum → IP54
Dış saha → IP65
CHINT NEX plastik korpus YASAK (betik HATA). CHINT NX30 korpus kullanılırsa betik malzeme teyidi için `KORPUS:` KONTROL yazar.
Çok bölmeli (yan yana) dikili panoda korpus satırının adedi bölme sayısıdır; iç genişlik ve bara boyu bununla çarpılır.
Sıra planı (b ve c için zorunlu, a'da kontrol amaçlı). Aynı sırada olmaması gerekenler ayrı sıraya konur:
Giriş koruması / MCCB ile modüler cihazlar
Sayaç + AT (mühürlü ölçü bölümü)
АВР kontaktörleri + mekanik kilit
Klemensler (en alt sıra; yalnız `klemens` EVET ise)
Bara bölgesi (yalnız bara tipi FAZ ise)
Sinyal lambaları, butonlar ve selektörler kapakta durur, sıra harcamaz.
Modüler cihazlara (oto, dif, uzo) sıra numarası (`sira`) her zaman yazılır; tarak sıra numarasından hesaplanır.
Başlangıç sıra yükseklikleri (mm) — betikte KURALLAR › SIRA YÜKSEKLİĞİ REFERANSI (SIRA_MM) olarak durur; SIRA KONTROLÜ plan mm bu referanstan küçükse "PLAN mm REF ALTINDA" yazar:
Sıra içeriği	mm
Modüler (NXB, NXBLE, NXHB, ВА47)	150
MCCB ≤ 160 A	250
MCCB 160–320 A	300
MCCB 400–630 A	400
Kontaktör ≤ 95 A	200
Kontaktör 120–225 A	300
Sayaç + 3 AT	300
Blok ATS (NXZB 150 · NXZM-125 250)	150 / 250
Klemens	120
Bara bölgesi	150
Korpus boyu kuralları:
Genişlik: en geniş sıra + kanal payı.
Yükseklik: toplam sıra yüksekliği + üst/alt pay.
Aksi belirtilmedikçe modüler kapasitede %10 yedek bırakılır: modül ihtiyacı ≤ kapasite × (1 − KORPUS_YEDEK), KORPUS_YEDEK = 0,10 (KURALLAR, betik v2.12 varsayılanı). Proje veya kullanıcı farklı oran verirse json `tablolar.param.KORPUS_YEDEK` ile yazılır ve `KORPUS:` KONTROL'e not edilir.
Seçilen boy KAZTEK listesindeki en küçük uygun ölçüdür.
Açıklamada ölçü yoksa (ör. NEXUS) yalnız modül kapasitesi kontrol edilir; KONTROL'e yazılır.
Cihaz genişliği/yüksekliği/derinliği ve modül sayısı EBAT'tan gelir (FIYAT S–AA). Sıra planı yapılırken EBAT › SIRA_REF (talimat sıra yüksekliği ↔ katalog cihaz boyu) referans alınır; talimat değerleri değişmez.
json `pano` alanına yazılır: `montaj`, `ip`, `korpus_kaynak`, `sira`, `sira_modul`, `sira_plan` (sıra başına [içerik, mm]; içerik KURALLAR › SIRA İÇERİK TABLOSU'ndan: MCCB · MODULER · OLCU · AVR · KONTAKTOR · KLEMENS · BARA; ör. `[["MCCB", 250], ["MODULER", 150], ["KLEMENS", 120]]`). Satırlarda `sira` = sıra numarası (kapı cihazı `K`). `sira_yuk` ("250+150+150") yalnız geriye uyum içindir; toplam yükseklik formüldür, `sira_yuk_mm` yazılmaz (yazılırsa ve toplamdan farklıysa `KORPUS:` KONTROL).
KURALLAR › PANO satırı modül kapasitesini ve sıra yüksekliğini formülle kontrol eder. KONTROL ≠ OK ise korpus büyütülür, yeniden kurulur.

6.3 · SARF (betik üretir, Claude yalnız parametre verir)
Tel yalnız pano içi bağlantı için sayılır. Çıkış cihazının çıkışı doğrudan saha kablosuna bağlanır; giriş cihazının girişi doğrudan besleme kablosuna bağlanır.
Satır	Hesap
Kablo ПВ3 (kesit başına)	cihaz başına tel × tel boyu × fire. Kesit cihaz In'inden (KURALLAR › KABLO KESİT). Kumanda 1,5 · ölçü 2,5 mm². YEDEK hariç. Tarak üzerindeki MCB faz teli almaz; tarak üzerindeki RCBO/УЗО yalnız 1 N teli alır
Kablo kanalı	(yatay adet × iç genişlik + 2 × iç yükseklik) × fire. Yatay adet varsayılan sıra + 1, KURALLAR › PANO'da değiştirilir
Ray klemens L/N ve PE	yalnız rolü CIKIS olan güç cihazları: L+N ve PE × yedek. `pano.klemens` HAYIR ise 0
Klemens uç kapağı / stoper	klemens grubu başına
Yüksük НШВИ	UÇ TİPİ = YÜKSÜK olan teller × 2 uç, упак'a bölünür (kafes klemens: NXB, NXHB, NXBLE, kontaktör ≤ 95 A, kumanda)
Kablo pabucu JG	UÇ TİPİ = PABUÇ olan teller × 2 uç (cıvata bağlantı: MCCB, rubilnik, ATS, ППН, kontaktör ≥ 96 A) + o kesitte НШВИ yoksa kafes klemensli teller
Kablo bağı	kablo metresi × BAG_ADET_M, упак'a bölünür
N ve PE barası (ШНИ)	ŞНИ panosunda her zaman; FAZ panosunda modüler çıkış (oto/dif/uzo/bh) varsa. N deliği: 1P/2P/4P MCB, RCBO, УЗО çıkışı başına 1 (3P MCB N almaz) + kumanda + 1. PE deliği: güç çıkışı başına 1 (mccb/ppn/rubilnik hariç, onlar PE barasına) + 2 + kapı. En küçük uygun ШНИ; 14 deliği aşarsa 14'lük × adet
Faz / N / PE busbar	Bara tipi FAZ ise (aşağıdaki karar). Kesit: KURALLAR › BARA tablosu (kullanıcı taşıma kapasitesi, 260930): önce tek şerit, yetmezse 2 paralel, sonra 3 paralel (= tek × BARA_PAR3 2,25); aynı şerit sayısında tabloda ilk uygun (en ucuz). N = faz (aynı şerit). PE = faz toplam kesitinin yarısına en yakın büyük tek şerit; yetmezse 2 şerit. Tablo aşılırsa `BARA TABLOSU AŞILDI` + `TEKNIK:` KONTROL (daha büyük bara PYS). BARA_DERATING = 1 (tablo 25 °C, doğal havalandırma). Malzeme `pano.bara` = CU / AL; belirtilmemişse `proje.bara_malzeme`, o da yoksa CU
Bara boyu (faz başına)	iç genişlik (korpus G − pay, × bölme) + korpus yüksekliği × BARA_DIKEY_ORAN (0,70, × bölme). Faz = 3 × şerit × boy × fire · N = şerit × boy × fire · PE = PE şerit × boy × fire
İzolatör	faz + N (4) × (ROUNDUP(boy / IZOLATOR_ARALIK_M) + 2); PE gövdeye bağlanır
Tarak (CHINT PIN)	aynı sırada (T sayı) ≥ 2 adet oto/dif/uzo varsa sıra başına. 3 fazlı pano: IN_ANA ≤ 63 A → PIN 3P 63 A 811002, üstü PIN 3P 100 A 811006; tek fazlı: PIN 1P 100 A 811004. Adet = ROUNDUP(sıradaki modül × 17,5 / TARAK_BOY_MM). Giriş cihazı tarağa bağlanır ama telini korur
Kapak topraklaması	kapı başına 6 mm² ПВ3 + 2 pabuç
Bara kararı (KURALLAR › PANO Y "BARA TİPİ", sarı): aşağıdakilerden biri varsa FAZ, yoksa ŞНИ (SNI):
IN_ANA ≥ BARA_ESIK_A (125 A)
montaj DİKİLİ
pano adında ВРУ / ГРЩ / АВР ve panoda (herhangi rolde) mccb / ppn / rubilnik / ats var. Tüm koruma cihazları modüler (oto, dif, uzo, bh) olan ВРУ / ГРЩ / АВР panosunda ad tek başına FAZ yapmaz → SNI (N/PE ШНИ + tarak)
CIKIS rolünde mccb / ppn / rubilnik
Betik uyumu (v2.12): betik pano adını hâlâ tek başına FAZ sayar. Pano adı ВРУ / ГРЩ / АВР olup IN_ANA < BARA_ESIK_A, montaj DİKİLİ değil ve panoda mccb / ppn / rubilnik / ats yoksa Claude:
json'a `pano.bara_tip: "SNI"` yazar,
sıra planına BARA sırası koymaz,
`TEKNIK: v26.7 bara kuralı — tamamen modüler <ad>, faz barası yok (SNI + tarak)` KONTROL'ünü yazar.
Betik yeni sürümde bu kuralı kendisi uyguladığında bu istisna kalkar.
Json `pano.bara_tip` ("FAZ" / "SNI") ile ezilebilir. IN_ANA = GIRIS rolündeki koruma cihazlarının (mccb, oto, bh, rubilnik, ats, dif, uzo, ppn) en büyük IN'i. FAZ panoda korpus G veya Y yoksa betik HATA verir. Karar nedeni KURALLAR › PANO AE'de yazılıdır.
UÇ TİPİ: KURALLAR › TİP TABLOSU "PABUÇ EŞİĞİ A" sütunundan gelir. IN ≥ eşik ise PABUÇ, eşik boşsa YÜKSÜK. FIYAT › UÇ TİPİ hücresine elle PABUC/YUKSUK yazılabilir.
Tel sayısı ve boyu (TİP TABLOSU başlangıç değerleri, KONTROL — atölye ölçümüyle kalibre edilir):
Tip	Tel/adet	m/tel
mccb, rubilnik	3	0,5
contactor	3	0,6
bh	3	0,3
oto	2×kutup+N (tarakta 0)	0,1
dif, uzo	2×kutup+N (tarakta 1)	0,15
tt	0 (sekonder sayaç satırında)	—
sayac	10	1,2
kumanda tipleri	tabloda	0,8
Varsayılanlar:
Kanal: modüler ЩРв/ЩРн/NEXUS kutularında kanal verilmez (`kanal` boş). Dikili panoda kumanda/ölçü telleri için 40х40; 35 mm² ve üstü güç telleri kanal dışındadır.
Klemens: çıkış kablosu doğrudan cihaza bağlanan panolarda `klemens` = false. Klemens yalnız proje veya kullanıcı isterse eklenir.
Sarf artikülleri KURALLAR › SARF ARTİKÜL TABLOSU'ndan (kesit başına kablo, НШВИ, JG, klemens, uç kapağı) ve SARF_STOPOR / SARF_BAG / SARF_IZOLATOR / SARF_TARAK_1P / SARF_TARAK_3P63 / SARF_TARAK_3P100 hücrelerinden gelir. Tablo değişirse BOM ve FIYAT kendiliğinden güncellenir. НШВИ 16 = 311100829; 25 mm² ve üstünde НШВИ listede yok.
Sarf tipleri json satırlarına yazılmaz; istisna olarak yazılırsa (ör. giriş cihazı olmayan panoda faz barası) betik `TEKNIK:` KONTROL üretir.
Her panoda "монтажный материал" satırı VARDIR ve kaldırılmaz: 1 шт, fiyat = işçilik tutarı × MONTAJ_MALZ_KAT (KURALLAR, sarı; başlangıç 0,4).
İşçilik normları (dk/adet, TİP TABLOSU başlangıç): korpus 8 (< 700 mm) · 30 (700–1499 mm) · 90 (≥ 1500 mm) · lamba 8 · buton 8 · bh 4 · oto 4 · dif 4 · uzo 4 · rubilnik 10 · mccb 15,3 · kilit 8 · motor 2 · shunt 2 · kontak 2 · bara 30 · izolatör 1 · nbara 2 · kanal 5 · kablo 1 · klemens 2 · yüksük 0,1 · pabuç 2 · tarak boş (KONTROL — kullanıcı girer) · diğerleri tabloda.

6.4 · KUMANDA, SİNYAL LAMBASI VE PROJEDE OLMAYAN MALZEME
Kumanda koruması: projede lamba veya kumanda devresi varsa ve koruması yoksa ana markadan MCB eklenir.
Lamba grubu için giriş başına 1 × 3P C6 (CHINT NXB-63S 296825).
Her kumanda devresi için 1P C6 (296708).
Projede kaynak tarafında 1P MCB varsa (ör. ОЛ'deki lamba MCB'si) АВР kumandası da ondan beslenir; ayrıca eklenmez.
Faz sinyal lambası: giriş panolarında (ВРУ, ВРЩ, ГРЩ, ВУ, ШВУ, АВР) projede olmasa da giriş başına L1 / L2 / L3 lamba eklenir.
CHINT ND16-22D/2 230V: sarı 593076 · yeşil 593077 · kırmızı 593075.
Lambaların MCB'si yukarıdaki kurala göre ayrıca eklenir.
Betik kontrol eder: pano adında bu kısaltmalardan biri geçip lamba satırı yoksa, veya lamba olup KUMANDA rolünde MCB yoksa `TEKNIK:` KONTROL yazar.
Diğer eksik malzeme: kumanda devresi için gereken ama projede yazmayan her malzeme, şalt veya aksesuar eklenir: yardımcı kontak, mekanik kilit, röle ve altlığı, zaman rölesi, selektör, АВР rölesi, trafo bağlantılı sayaç için test klemens kutusu vb.
Projede olmayan her eklenen satırın Q NOT sütununa gerekçe yazılır: `projede yok — <neden> için eklendi`. Rol: `KUMANDA`. Json'da satıra `"kaynak": "EK"` yazılabilir. Bu gerekçeyi taşıyan satırlar şemada olmadığı için mutabakata girmez; `kaynak: "EK"` olup gerekçesi olmayan satır HATA'dır.

6.5 · АВР
Tip şemadan / ОЛ'den alınır; yoksa KONTROL.
A · Kontaktörlü
Kalem	Adet / kural
MCB veya MCCB	giriş sayısı kadar
Kontaktör AC-3 230V	giriş sayısı kadar; In = koruma In'inin bir üst kademesi
Yardımcı kontak 2NO	kontaktör başına 1
Mekanik kilit	1
F&F АВР rölesi	1
Kaynak başına kumanda MCB	1 (6.4'teki istisna geçerli)
Artiküller:
Kontaktör: NXC-120 220V 836511 · NXC-160 220V 836515
Mekanik kilit: MI-6 777989 · MI-7 777988 · MI-9 263761
Yardımcı kontak: AX-3X/11 938256 (CHINT'te hazır 2NO yok → KONTROL)
F&F (EKT): 2 giriş AVR-02-G 311100898 (260917 kullanıcı seçimi) · AVR-01-K 311100864 · AVR-02 311100861. Model teyidi → KONTROL
B · Blok ATS (In ≤ 100 A)
Kalem	Adet
MCB veya MCCB	giriş sayısı kadar
Blok ATS	1
Artiküller: NXZB-63H/4C 171628…171633 · NXZM-63S/3B 256759 · 256757 · NXZM-125S/3B 80 A 256778 · 100 A 256781. NZ7 listede yok.
C · Motorlu MCCB
Kalem	Adet
NXM	giriş sayısı kadar
Motor sürücü	MCCB başına 1
Şönt açtırma	MCCB başına 1
AX/AL kontak	MCCB başına 1
Mekanik kilit	1
Kontrolör	1
Artiküller:
Motor: MD-M2 946905 · MD-M3 331267 · MD-M4 946912
Şönt: SHT-M2 946953 · SHT-M3 946972
AX/AL: M2 946945 · M3 146364 · M4 946940
NXM mekanik kilit listede yok → PYS + KONTROL
Kontrolör: F&F AVR veya CHINT NZQ7C 304537 → KONTROL

7 · proje.json
```json
{"proje": {"kod": "260917", "musteri": "", "proje": "", "sehir": "", "tarih": "2026-09-17", "teklif_no": "",
           "uretici": "TEKLED electric", "pano_ek": "260917 fok atr",
           "kar_orani": null, "iskonto": 0, "nds": 0.16, "gecerlilik": null, "teslim": null, "odeme": null,
           "firma_tam": null, "firma_adres": null, "firma_tel": null, "firma_email": null, "firma_bin": null,
           "banka": null, "iik": null, "bik": null,
           "musteri_adres": null, "musteri_kisi": null, "musteri_tel": null, "musteri_mobil": null, "musteri_email": null,
           "teslim_yeri": null, "garanti": null, "bara_malzeme": null, "salt_marka": null,
           "imza_unvan": null, "imzalayan": null, "hazirlayan": null, "kapanis": null, "not": null},
 "kur": {"EUR": null, "USD": null, "TRY": null, "RUB": null},
 "fiyat": [{"artikul": "296697", "csv_marka": "CHINT", "tip": "oto"},
           {"artikul": "PYS-001", "ad": "…", "tip": "lamb", "marka": "pys", "birim": "шт", "liste": null, "kaynak": "PYS"}],
 "panolar": [{"ad": "ЩС-1", "adet": 1,
   "pano": {"montaj": "SIVA ALTI", "ip": "31", "korpus_kaynak": "PROJE|BUYUTULDU|ROBOT", "sira": 2, "sira_modul": 2,
            "sira_plan": [["MODULER", 150], ["MODULER", 150]], "kapi": 1, "klemens": false, "kanal": "", "kanal_yatay": null,
            "bara_tip": null},
   "satirlar": [{"tip": "korpus", "artikul": "719-15925", "adet": 1, "rol": "KORPUS", "not": "projede ЩРв-18з — modül yetersiz, büyütüldü"},
                {"tip": "bh", "artikul": "193183", "adet": 1, "rol": "GIRIS", "sira": 1, "lok": "şema · ana bara · QS1"},
                {"tip": "dif", "artikul": "515381", "adet": 13, "rol": "CIKIS", "sira": 1},
                {"tip": "oto", "artikul": "296708", "adet": 1, "rol": "KUMANDA", "sira": 2, "kaynak": "EK",
                 "not": "projede yok — lamba grubu koruması için eklendi"}]}],
 "mutabakat": [{"pano": "ЩС-1", "grup": "3P 63A yük ayırıcı", "artikul": ["193183"], "sema": 1, "cetvel": "—", "bom": 1,
                "count_a": 1, "count_b": 1, "lokator": "şema · ana bara · QS1"},
               {"pano": "ЩС-1", "grup": "1P+N C16 30mA", "artikul": ["515381"], "sema": 13, "cetvel": 12, "bom": 13,
                "count_a": 13, "count_b": 13, "lokator": "şema · ana bara · P1–P12; şema · UPS barası · C/1x16A",
                "not": "UPS barası çıkışı cetvelde sütunsuz"}],
 "kontrol": [{"pano": "ЩС-1", "artikul": "", "konu": "FIYAT: …", "durum": "KONTROL"},
             {"pano": "ЩС-1", "artikul": "", "konu": "OKUMA: sayfa 3 tek hat raster okundu (600 DPI)", "durum": "KONTROL"}]}
```
mutabakat[] (zorunlu)
Cihaz satırı olan her pano için gruplar yazılır. Grup = şemada aynı tip / kutup / In / eğri / mA olan cihazlar. Bir artikül bir panoda yalnız bir grupta bulunur.
Alanlar: `pano` (panolar[].ad ile aynı), `grup` (okunur ad), `artikul` (liste), `sema`, `cetvel` (sayı veya `"—"`), `bom`, `count_a`, `count_b`, `lokator`, isteğe bağlı `not` (fark sebebi), `kaynak` (`ELLE` yalnız `--yenile` yazar).
`bom` = o gruptaki artiküllerin tüm satırlardaki adet toplamı (YEDEK dahil; EK satırlar hariç).
`ŞEMA ≠ CETVEL` ve `COUNT_A ≠ COUNT_B` KONTROL satırlarını betik üretir; Claude bunları `kontrol[]`'e ayrıca yazmaz.
`mutabakat_muaf: true` (pano alanı) yalnız `--donustur` / eski dosyadan gelen panolarda bulunur; Claude yeni BOM'da yazmaz.
HATA / KONTROL (betik doğrulaması)
Durum	Sonuç
cihazı olan panoda mutabakat yok · sema ≠ bom · sema ≠ max(count_a, count_b) · count_a / count_b boş · lokatör boş · satır toplamı ≠ bom · mutabakat grubunda olmayan cihaz satırı · artikül iki grupta · `kaynak: "EK"` ama "projede yok" gerekçesi yok	HATA — Excel kurulmaz
kodsuz `kontrol[].konu` · tanımsız tip · `din` · FIYAT'ta olmayan artikül · FAZ baralı panoda korpus G/Y yok	HATA
CHINT NM1 · Schneider A9F74 · ABB BMS · CHINT NEX korpus · metal lоток · ВА88-32 > 125 A	HATA
count_a ≠ count_b · cetvel ≠ şema	`SAYIM:` KONTROL (betik yazar)
mutabakat `kaynak: ELLE` iken şema ≠ BOM (kullanıcı değiştirdi)	`SAYIM:` KONTROL
json'a elle yazılmış sarf tipi · faz lambası / lamba MCB'si eksik · bara notu · bara tablosu aşıldı · FAZ panoda IN_ANA 0 · tarak tipi / marka uyumu / norm boş · modüler cihazda sıra no yok	`TEKNIK:` KONTROL
EBAT yok / çakışma / modül farkı · NX30 korpus · sira_yuk_mm ≠ plan toplamı · iki farklı ebat.csv	`KORPUS:` KONTROL
TEKLIF görseli eklenemedi	`TEKLIF:` KONTROL
`--yenile` göçü (eski norm / bara tablosu / tip tablosu değiştirildi)	`KAYNAK:` KONTROL
HATA'da betik tüm hataları listeler; json düzeltilir ve betik yeniden çalıştırılır. Betik değiştirilmez.
KONTROL kategori kodları
Her `kontrol[].konu` bir kodla başlar; kodsuz satır betikte HATA'dır. Betiğin kendi ürettiği satırlar da kodludur. Aynı pano + artikül + konu (veya aynı konu anahtarı, ör. modül farkı) tek satır yazılır.
Kod	Kullanım
`SAYIM:`	şema / cetvel / BOM farkı, COUNT_A ≠ COUNT_B, bulunamayan UPS/АВР fonksiyonu, elle adet değişikliği
`OKUMA:`	raster okunan pano, okunamayan bölge, belirsiz yazı
`ARTIKUL:`	muadil, bir üst kademe, PYS, S200 kullanımı, eğri seçimi, BOM'a elle eklenen artikül
`TEKNIK:`	kesit, kA, АВР tipi, kumanda eklemesi, eksik cihaz teyidi, faz lambası, bara notu, bara / tarak
`KORPUS:`	ölçü, BUYUTULDU / ROBOT, modül kapasitesi, EBAT ölçüsü yok / çakışma, sıra yüksekliği
`FIYAT:`	fiyat yok, liste yaşı, elle fiyat ≠ liste
`KUR:`	kur girilmedi / varsayım
`KAYNAK:`	kaynaklar arası çelişki (şema ↔ спецификация ↔ şartname), mutabakat muaf pano, eski dosyadan gelen kodsuz satır (`KAYNAK: eski —`), sürüm göçü
`TEKLIF:`	TEKLIF sayfası görsel / not sorunları
Alanlar
Pano adı (zorunlu biçim): `panolar[].ad` = şemadaki pano adı + projede yazan açılımı parantez içinde + kaynak PDF etiketi parantez içinde (ör. `ВРУ-1 (ВРУ1-ИНД тип5-00 ухл4) (S1)`). Açılım projede yoksa yazılmaz; tek PDF'li projede etiket yazılmaz. `proje.pano_ek` zorunludur: `YYMMDD <proje kısaltması>` (ör. `260901 akzat samgau`); betik bunu her pano adının sonuna " (ek)" olarak ekler → `ВРУ-1 (ВРУ1-ИНД тип5-00 ухл4) (S1) (260901 akzat samgau)`. Eski dosyadan `YENİLE`/`DÖNÜŞTÜR` ile gelen json'da pano_ek, açılım ve etiket yoksa tamamlanır.
`proje.uretici`: pano başlık satırı J. Boşsa "TEKLED electric" varsayılır.
`proje.firma`: tedarikçi adı (KURALLAR › FIRMA). Boşsa "TEKLED electric". `proje.firma_tam`: TEKLIF sağ üstündeki tüzel unvan (ör. ТОО "TEK LED"); boşsa FIRMA yazılır, altında "БИН:" + `firma_bin`.
`proje.firma_adres` / `firma_tel` / `firma_email` / `banka` / `iik` / `bik`: TEKLIF altbilgisi. `firma_email` boşsa "yusuf.kaplan@tekled.com". Genelde sabittir; kullanıcı bir kez verince her projede aynen yazılır, sorulmaz.
`proje.musteri_adres`, `musteri_kisi`, `musteri_tel` (Рабочий), `musteri_mobil`, `musteri_email`: TEKLIF sağ üstteki Кому bloğu. Şemada/ОЛ'de yoksa boş bırakılır, KONTROL'e yazılmaz.
`proje.teslim`: tablodaki "Срок поставки товара" sütunu (her satır).
`proje.gecerlilik`, `teslim_yeri`, `odeme`, `garanti`, `not`: numaralı notlar. Boş olan not atlanır, numaralar kayar.
`proje.bara_malzeme` (CU/AL): not "В щитах с системой сборных шин применяются медные (Cu) / алюминиевые (Al) шины." Kullanıcı söylerse yazılır. Boşsa panolardaki ortak bara malzemesi, o da yoksa CU; baralı pano varsa KONTROL'e yazılır. `pano.bara` yalnız o pano projeden farklıysa verilir; verilmezse KURALLAR › PANO I hücresi `=BARA_MALZEME` olur.
`pano.bara_tip`: "FAZ" / "SNI" — yalnız §6.3 bara kararını ezmek gerekirse yazılır: şemada bara açıkça farklıysa veya §6.3 betik uyumu istisnasında (tamamen modüler ВРУ/ГРЩ/АВР → "SNI"). Boşsa betik karar verir.
`proje.salt_marka`: not "Коммутационная аппаратура (…): <marka>." Kullanıcı marka söylerse (CHINT, Schneider Electric, ABB, TEKLED electric, VIKO …) yazılır; boşsa betik BOM güç cihazlarının markalarından (adet ağırlıklı) türetir.
`proje.sertifikat`: sabit sertifika notu (varsayılan: «Сертификат CT-KZ», международные сертификаты, ISO 9001-2015). Boş string verilirse not çıkmaz.
`proje.kapanis`: notların altındaki yeşil kapanış satırı (isteğe bağlı).
`proje.imza_unvan` (ör. Директор), `imzalayan` (ad soyad), `hazirlayan` (Исполнитель): imza ve altbilgi.
`proje.teklif_gorsel`: tablo satırlarındaki TEKLED logosu; varsayılan true.
`rol`: GIRIS · CIKIS · KUMANDA · OLCU · YEDEK · KORPUS. Aynı artikül farklı rolde ayrı satırdır. Bara bölümü girişi (ör. UPS çıkış şalteri) ayrı satırdır ve rolü `GIRIS` değil `CIKIS`'tır (IN_ANA'yı etkilemez).
Satır alanları: `sira` (sıra no / `K` / `G`), `not` (Q NOT), `lok` (kaynak izi; Q NOT başına eklenir), `kaynak: "EK"` (§6.4), `norm` (elle norm tipi).
IN_ANA, GIRIS rolündeki koruma cihazlarından hesaplanır (mccb, oto, bh, rubilnik, ats, dif, uzo, ppn; kontaktör hariç).
`tip`:
Güç: korpus, rubilnik, bh, mccb, ppn, ats, sayac, tt, oto, uzo, dif
Kumanda: rn, fotorele, contactor, kontak, kilit, motor, shunt, condencator, krm, avr_rele, rele, timer, lamb, indicator, button, perekluchatel, rozetka, stable, aksesuar
Sarf tipleri (kablo, kanal, klemens, yuksuk, pabuc, bag, bara, nbara, izolator, tarak) json'a yazılmaz; betik üretir.
Listede olmayan tip HATA'dır.
`fiyat`:
Özellikleri betik açıklamadan ayrıştırır: kutup, IN, modül, paket, kesit, korpus G/Y/D/kapasite (ЩР-48, S/U 36, x32). G/Y/D/modül/kapasite EBAT'ta varsa EBAT değeri kullanılır.
Ayrıştırma boşsa veya yanlışsa `kutup`, `in`, `modul`, `paket`, `g`, `y`, `d`, `kap`, `uc` (PABUC/YUKSUK) verilir. 3P+N / 1P+N RCBO'da kutup ve modül elle verilir.
`panolar[].saat`: isteğe bağlı elle işçilik saati; boşsa norm formülü.
`tablolar` (isteğe bağlı): KURALLAR tablolarını ezer (`marka`, `tips`, `param`, `iscilik`, `kablo`, `klemens`, `sarf`, `sarf_tek`, `sni`, `bara_cu`, `bara_al` [artikül, mm², A 1 şerit, A 2 paralel], `sira_mm`, `kullanici`, `ek_kural`, `yas`). `--yenile` bunları düzenlenmiş dosyadan doldurur. Korpus yedeği: `"tablolar": {"param": {"KORPUS_YEDEK": 0.10}}` (yalnız proje/kullanıcı varsayılandan farklı oran verirse yazılır).

8 · EXCEL (bom_kur.py v2.12 üretir)
8.1 · Biçim
Kılavuz çizgileri TEKLIF dışındaki her sayfada açık.
Kullanıcının renkleri: başlık satırı sarı (FFFF00), pano satırı FFD966.
İçerik satırları dolgusuz, ince kenarlıklı, Calibri 10 kalın. Toplam satırı Tahoma 9.
BOM'da birleştirilmiş hücre ve boş satır yok (TEKLIF'te birleştirme serbest).
Birim: `комплект` yalnız pano başlık satırında. İçerikte шт · м · упак · мин.
Gizli ve gruplu sütunlar ("+" ile açılır): BOM N–P, R–S ve U–AR (AO = SIRA REF mm, AP N DELİK, AQ PE DELİK, AR TARAK), FIYAT S–AM, TEKLIF I–K (I iç kontrol · J birim без НДС · K tutar без НДС; I sütununun not satırları not motorudur).
FIYAT'ta projede kullanılmayan yedek sarf artikülleri altta gri, satır gruplu ve kapalıdır (NOT: "YEDEK SARF"); elle değişiklikte gereken artikül hazır bulunur.
8.2 · Sayfalar
Sayfa	İçerik
BOM	A–L kullanıcı formatı + M PANO ADET · N ROL · O NORM · P DURUM · Q NOT · R FID · S YENİ · T SIRA · U–AN yardımcı · AO SIRA REF mm · AP N DELİK · AQ PE DELİK · AR TARAK (1 = tarak üzerinde)
FIYAT	A ID · B tip · C artikül · D ad · E liste (sarı) · F net · G toplam · H toplam adet · I birim · J marka · K para · L katsayı · M kur · N kaynak · O tarih · P yaş · Q durum · R not · S–AA özellikler (sarı) · AB–AJ türetilen · AK UÇ TİPİ · AL LİSTE satırı · AM BOM satırı. Üstte kullanılan artiküller, altta gruplu YEDEK SARF, en altta 40 otomatik satır
KURALLAR	proje/teklif (firma, müşteri iletişim, banka, bara malzemesi, şalt markası, sertifika, imza alanları dahil), kur, marka, kaynaklar, işçilik, TİP TABLOSU (pabuç eşiği dahil), parametreler (KORPUS_YEDEK · BARA_DERATING · BARA_PAR3 · BARA_DIKEY_ORAN · TARAK_BOY_MM dahil), kablo/klemens/SARF ARTİKÜL (SARF_TARAK_1P / SARF_TARAK_3P63 / SARF_TARAK_3P100)/ШНИ tabloları, BARA BAKIR / ALÜMİNYUM (artikül · mm² · A 1 şerit · A 2 paralel · A 3 paralel formül), SIRA YÜKSEKLİĞİ REFERANSI (SIRA_MM), KULLANICI PARAMETRELERİ, EK KURALLAR, PANO PARAMETRELERİ (A–X + Y BARA TİPİ · Z FAZ ŞERİT · AA BARA BOYU m · AB FAZ/N ARTİKÜL · AC PE ŞERİT · AD PE ARTİKÜL · AE KARAR NEDENİ), SIRA / YERLEŞİM kontrolleri, SAYIM MUTABAKATI, otomatik kontroller, KONTROL LİSTESİ
TEKLIF	TEKLED КП şablonu (Cambria, yeşil yazı 1E8449): sol üst TEKLED logosu · orta TEKSAN mührü · sağ üst tüzel unvan + БИН · ortalı "КОММЕРЧЕСКОЕ ПРЕДЛОЖЕНИЕ" · solda Исх.№ / Дата · sağda Кому / Адрес / Объект / Контактное лицо / Рабочий / Мобильный / Эл. почта · altın (FFC000) başlıklı tablo: № · Изображение (TEKLED logosu) · Наименование товара · Кол-во · Ед. (компл.) · Цена с НДС · Сумма · Срок поставки товара · altın "Сумма: … KZT" satırı · numaralı notlar (geçerlilik, teslim yeri, ödeme, garanti, bara malzemesi Cu/Al, şalt markası, sertifika, proje notu; boşlar atlanır) · kapanış satırı · "С уважением" + unvan + imzalayan · altbilgi (Адрес, Телефон, Эл. почта, Название банка, ИИК, БИК, Исполнитель) · KZ rozeti. Birim fiyat без НДС = BOM pano toplamı × (1 + kâr) × (1 − iskonto); Цена с НДС = bu × (1 + НДС). Baskı alanı A:H, A4 dikey, tek sayfa genişliği, altta "Страница N"
OZET	pano başına: ad, adet, birim teklif fiyatı (без НДС), tutar, korpus kodu/adı, malzeme tutarı, işçilik tutarı, işçilik dk, maliyet, kâr. Toplamlar: Итого без НДС / НДС / Итого с НДС (TEKLIF J–K ve G'den), toplam maliyet, kâr oranı, dk/saat/gün (GUN_SAAT), fiyatsız, otomatik eklenen, pano kontrol, yerleşim, mutabakat ŞEMA ≠ BOM, KONTROL sayısı. Tüm hücreler BOM/TEKLIF/KURALLAR'dan formül
LISTE	birlesik.csv'nin tamamı + ham özellikler; otomatik FIYAT satırları buradan okur
8.3 · Pano bloğu sırası
Başlık
Korpus
Rubilnik / bh
MCCB / sigorta / ATS
Sayaç
AT
MCB
УЗО
АВДТ
Kontaktör
Kumanda
Sarf (bara · izolatör · ШНИ · tarak · kanal · kablo · klemens · yüksük · pabuç · bağ)
монтажный материал
монтажная работа
Toplam
8.4 · Elle değişiklik ve zincir
Değiştirilen	Etkilenen
FIYAT'ta herhangi bir değer (artikül kodu dahil)	BOM'da o ürünün tüm satırları (cihaz satırı C = FIYAT ID bağlantısı), net fiyat, sarf, işçilik, OZET, teklif
BOM C'ye FIYAT'ta olmayan artikül	FIYAT otomatik satırı açılır ("YENİ"); ad, fiyat, birim, marka, kaynak ve özellikler LISTE'den gelir. Listede yoksa "LİSTEDE YOK" → fiyat elle girilir
BOM adet / rol / sıra	sarf satırları (ШНИ delik, tarak boyu, tel), FIYAT toplamı, işçilik, IN_ANA, bara kesiti, teklif, SAYIM MUTABAKATI › BOM CANLI ve DURUM ("ŞEMA ≠ BOM")
KURALLAR › SARF ARTİKÜL / SARF_STOPOR · SARF_BAG · SARF_IZOLATOR · SARF_TARAK_1P · SARF_TARAK_3P63 · SARF_TARAK_3P100	BOM sarf artikülü ve fiyatı
KURALLAR › PANO bara malzemesi, klemens, kapı, kanal artikülü, yatay adet	ilgili sarf satırları; bara şeridi, boyu, faz/PE artikülü
KURALLAR › BARA tabloları, BARA_DERATING, BARA_PAR3, BARA_DIKEY_ORAN	faz/N/PE bara artikülü ve metresi, izolatör
KURALLAR › KORPUS_YEDEK	KURALLAR › PANO KONTROL ("MODÜL YETERSİZ"); korpus satırı değişmez, gerekirse elle büyütülür / `YENİLE`
KURALLAR › PANO Y BARA TİPİ (FAZ ↔ SNI)	KONTROL "SARF SATIRLARI GÜNCEL DEĞİL" → `YENİLE`
KURALLAR tel / norm / eşik / katsayı / parametre	ilgili tüm satırlar
KURALLAR › SIRA_MM	BOM AO ve SIRA KONTROLÜ REF mm
KURALLAR tablolarına yeni satır	boş sarı satıra yazılır; sıralı tablolarda (kablo, klemens, ШНИ, bara: tek şerit akımına göre) artan sıra korunur
KULLANICI PARAMETRELERİ	formülde `INDEX(USER_DEG,MATCH("KOD",USER_KOD,0))` ile kullanılır
İşçilik saati veya MONTAJ_MALZ_KAT	монтажный материал satırı
KURALLAR › PROJE/TEKLİF alanları (firma, müşteri, banka, teslim, ödeme, garanti, sertifika, imza, kapanış, not)	TEKLIF başlığı, Кому bloğu, tablo "Срок поставки", notlar, imza, altbilgi
KURALLAR › BARA_MALZEME	TEKLIF bara notu + `=BARA_MALZEME` olan panoların bara sarf satırları
KURALLAR › SALT_MARKA	TEKLIF şalt markası notu
Elle çalışma kuralları:
FIYAT sıralanabilir ve filtrelenebilir; bağlantı ID ile kurulduğu için bozulmaz.
BOM'a satır eklerken aynı bloktaki bir cihaz satırı kopyalanıp içine yapıştırılır.
Otomatik FIYAT satırına elle girilen değer kalıcı değildir; satır sırası BOM'a göre değişir. Kalıcı olması için `YENİLE` çalıştırılır.
Satır yapısını değiştiren durumlar KURALLAR › PANO KONTROL'de "SARF SATIRLARI GÜNCEL DEĞİL" olarak görünür ve `YENİLE` gerektirir:
IN_ANA bara eşiğini geçerse veya BARA TİPİ elle değiştirilirse
yeni bir kesit, uç tipi, klemens veya kanal ihtiyacı doğarsa ("TEL/UÇ KAPSAMI FARKI" ≠ 0)
yeni bir tarak sırası doğarsa
"Otomatik satır kapasitesi" DOLU ise `YENİLE` çalıştırılır.
8.5 · YENİLE sonrası json kontrolü
FIYAT'ta artikülü değiştirilmiş satırlarda eski özellikler (kutup, modül, ölçü, kapasite) kalmış olabilir; açıklamaya uymayanlar json'dan silinir, betik yeniden ayrıştırır.
BOM'da elle yazılmış fiyat, betik tarafından FIYAT liste fiyatı yapılır ve KONTROL'e yazılır. Liste fiyatından farkı teyit edilir.
Artikül sonundaki `_` silinir. Eskimiş KONTROL maddeleri güncellenir; kodsuz eski satırlar `KAYNAK: eski —` önekiyle gelir, uygun koda çevrilir.
TEKLIF alanlarının tamamı (müşteri iletişim, banka, bara malzemesi, şalt markası, sertifika, imza) ve PANO › BARA TİPİ json'a geri okunur.
Eski dosyada: v2.10 varsayılan normları (korpus 21,3/245,2/244,7, lamba/buton 21,8), eski bara tablosu (2 paralel sütunu yok), eski formatlı tip tablosu ve teyitsiz Legrand 0,50 v2.11 değerleriyle değiştirilir; FIYAT'ta bulunmayan BOM artikülü birlesik.csv'den alınır. Hepsi `KAYNAK: v2.11 göçü` / `ARTIKUL:` KONTROL'üne yazılır. Kullanıcının kendi değiştirdiği değerler korunur.
v2.11 ve öncesi dosyada KURALLAR › KORPUS_YEDEK 0,20 olarak geri okunur. Kullanıcı bu değeri bilerek bırakmadıysa json'da 0,10 yapılır (`tablolar.param.KORPUS_YEDEK` silinir) ve `KAYNAK: v2.12 göçü — KORPUS_YEDEK 0,20 → 0,10` KONTROL'ü yazılır.
Mutabakat geri okunur. BOM'da adedi elle değişen grup `kaynak: "ELLE"` olur (`bom` = yeni adet) ve `SAYIM:` KONTROL'üne yazılır. BOM'a elle eklenen, hiçbir gruba girmeyen artikül için ELLE grup açılır. Mutabakat tablosu olmayan (eski) dosyada panolar `mutabakat_muaf` olur.
Değişiklik şemadan kaynaklanıyorsa (kullanıcı yeni şema gönderdi) ELLE grup yerine sayım yeniden yapılır ve `count_a` / `count_b` / `sema` güncellenir.

9 · ÜRETİM SIRASI
Adım 1 · Oku, sayfa envanteri çıkar.
Adım 2 · Pano pano: bara bölümlerini listele (§5.1), şemadan say, artikül seç, korpus ve sıra planını belirle (modüler cihazlara sıra no), eksik kumanda malzemesini ekle.
Adım 2a · Cetvel dışı cihaz taraması ve bağımsız ikinci sayım (§5.1): COUNT_A ↔ COUNT_B. Fark varsa şema hedefli olarak yeniden okunur; çözülemeyen fark §5.1'deki kurala göre işlenir.
Adım 2b · Sayım mutabakatı (§5.1): her cihaz grubu için ŞEMA · CETVEL · BOM · COUNT_A · COUNT_B · lokatör. ŞEMA ≠ BOM ise Adım 2'ye dönülür.
Adım 3 · `ARA` komutuysa listeyi (mutabakat tablosu ve bara kararı dahil) ver ve onay bekle.
Adım 4 · `/home/claude/work/<proje>/proje.json` yaz (pano adı açılımıyla, `proje.pano_ek` dolu — §7 Alanlar) (`mutabakat[]` dahil; yalnız bu klasörde kalır; outputs'a, projeye veya sohbete gitmez). Kullanıcının daha önce verdiği firma/banka/imza bilgileri ve mesajdaki bara malzemesi / şalt markası TEKLIF alanlarına yazılır.
Adım 5 · `BOM_VERI=/home/claude/sys python3 /home/claude/sys/bom_kur.py proje.json /mnt/user-data/outputs/YYMMDD_<şehir>_<proje>_<marka>_BOM.xlsx`
Adım 6 · Çıktı `HATA — Excel kurulmadı` ise listedeki her madde json'da düzeltilir (sayım hatasıysa şema yeniden okunur) ve Adım 5 tekrarlanır. Betik değiştirilmez.
Adım 7 · `python3 /mnt/skills/public/xlsx/scripts/recalc.py <dosya> 300` → hata 0.
Adım 8 · Kontrol et:
BOM cihaz adetleri = mutabakat ŞEMA adetleri; OZET "Mutabakat: ŞEMA ≠ BOM" = 0
FIYAT S–AK özellikleri (kullanılan cihazlar için kutup, IN, modül, uç tipi; korpus G/Y/kapasite) · FIYAT R'de EBAT kimliği · "EBAT ölçüsü yok" KONTROL'leri
KURALLAR › PANO KONTROL ve TEL/UÇ KAPSAMI FARKI = 0
KURALLAR › KORPUS_YEDEK = 0,10 (proje/kullanıcı farklı oran vermediyse)
KURALLAR › PANO Y–AE: bara tipi ve nedeni şemayla ve §6.3 ile uyumlu (dikili, IN_ANA ≥ eşik, MCCB/ППН/rubilnik/ATS'li ВРУ·ГРЩ·АВР veya MCCB çıkışlı panoda FAZ; tamamen modüler ВРУ/ГРЩ/АВР'de SNI), FAZ ise şerit ve kesit, boy > 0
Tarak satırları modüler sıralarla uyumlu
KURALLAR › SIRA KONTROLÜ ("PLAN mm REF ALTINDA" yok)
KURALLAR › OTOMATİK KONTROL
OZET toplamları
TEKLIF: görseller var (xl/media 3 dosya; satır logoları aynı görseli paylaşır), `TEKLIF:` KONTROL'ü yok, notlarda bara malzemesi ve şalt markası satırı çıkıyor
Adım 9 · Hata varsa json düzeltilir ve Adım 5'ten tekrarlanır (en çok 2 tur).
Adım 10 · Yalnız xlsx teslim edilir (`SendUserFile`). `proje.json` projeye yazılmaz (`project_write` YOK), sohbete gönderilmez, kullanıcı klasörüne konmaz.

9.1 · REGRESYON TESTİ (betik sürümü değiştiğinde)
Test paketi projede dosya olarak durur: `bom_test.zip` (`bom_test_py.txt`, test json'ları 260917_fok, 260918_konsas, 260928_akay, bara_sentetik ve `beklenen.json`). Bulma sırası §3 ile aynıdır (/mnt/project → oturuma yüklenmiş dosyalar (Glob, kimlik öneki olabilir) → Projects `project_read bom_test.zip` local_file). `unzip -o <yol>/bom_test.zip -d /home/claude/work/test`; `bom_test_py.txt` → `/home/claude/sys/bom_test.py` olarak kopyalanır (betik `bom_kur.py`'yi kendi klasöründe arar). Zip bulunamazsa test durur ve kullanıcıdan istenir; içerik elle yeniden yazılmaz.
`BOM_VERI=/home/claude/sys python3 /home/claude/sys/bom_test.py /home/claude/work/test` → her json için Excel kurulur, recalc edilir ve pano başına IN_ANA, bara tipi / şerit / boy / artikül, sarf satırları (bara, ШНИ, izolatör, tarak, kablo, yüksük, pabuç), PANO KONTROL, TEL/UÇ farkı ve toplam maliyet `beklenen.json` ile karşılaştırılır.
Fark varsa farklar kullanıcıya liste olarak gösterilir. Yeni sürüm yalnız kullanıcı farkları onaylarsa projeye yazılır; onaydan sonra `--kaydet` ile `beklenen.json` güncellenir, aynı 6 dosyayla yeni `bom_test.zip` kurulur ve kullanıcıya verilir; kullanıcı projedeki `bom_test.zip`'i onunla değiştirir.
Bu test json'ları yalnız test verisidir; §11'deki proje.json teslim yasağı onlar için geçerli değildir, örnek BOM olarak kullanılmaz.

10 · SOHBET ÇIKTISI
Yalnız dosya + tek satır:
`Toplam maliyet X ₸ · Teklif Y ₸ (kâr yoksa "kâr girilmedi") · fiyatsız n · KONTROL n`
Anlatım yok; notlar Excel'dedir. Teslim edilen tek dosya xlsx'tir; json verilmez.

11 · YASAKLAR
Artikül, fiyat, katsayı uydurmak.
Fiyat yokken 0 yazmak.
Belirsizliği KONTROL'süz geçmek.
Formül yerine sayı yazmak.
Cihaz adedini devre cetvelinden / yük tablosundan / спецификация'dan almak; şema sayımı, ikinci sayım ve mutabakat (§5.1) yapılmadan json yazmak.
`mutabakat[]`'sız veya lokatörsüz json yazmak; betiğin sayım HATA'sını adetleri mutabakata uydurarak (şemaya bakmadan) kapatmak.
Çizili sembolle gösterilemeyen cihazı COUNT_A ≠ COUNT_B gerekçesiyle eklemek.
Yeni BOM'da `mutabakat_muaf` yazmak.
Şemada çizilmemiş UPS / АВР fonksiyonu için cihazı "kalıp gereği" eklemek (bulunamayan fonksiyon yalnız KONTROL'e yazılır).
Kodsuz KONTROL satırı yazmak (§7 KONTROL kategori kodları).
DIN ray eklemek.
Metal lоток kullanmak.
CHINT NEX plastik korpus.
İçerikte `комплект` birimi.
"монтажный материал" satırını kaldırmak.
BOM'da boş satır veya birleşik hücre.
Excel'de Kazakça harf (görseller hariç).
XLOOKUP / FILTER / LET / LAMBDA / TEXTJOIN / dinamik dizi formülü.
Betiği proje dokümanının (`.txt`) sohbete dökülen metninden elle diske yazmak; betik yalnız zip dosyasından alınır (§3).
bom_kur.py'yi oturumda yamalamak, betik HATA'sını betiği değiştirerek aşmak veya çıktı dosyasını başka betikle sonradan düzenlemek.
Kalibrasyon değerlerini (tel, norm, eşik, SIRA_MM, bara tablosu, KORPUS_YEDEK) betik sürümü dışında değiştirmek. Kullanıcı Excel'de değiştirebilir; Claude `YENİLE` ile taşır. İstisna: proje veya kullanıcı farklı korpus yedeği verirse `tablolar.param.KORPUS_YEDEK` (§6.2).
Bara kararını (FAZ / SNI) bara kesitini veya şerit sayısını json'da elle uydurmak; `pano.bara_tip` yalnız şemada bara açıkça farklıysa veya §6.3 betik uyumu istisnasında (tamamen modüler ВРУ/ГРЩ/АВР) yazılır ve `TEKNIK:` KONTROL'e sebebi yazılır.
Yalnız pano adı (ВРУ / ГРЩ / АВР) yüzünden tamamen modüler panoyu dikili, taban saçlı veya faz baralı yapmak.
Betik sürümünü §9.1 regresyon testi yapılmadan ve farklar kullanıcıya gösterilmeden projeye yazmak.
Önceki oturum çıktısını kural kaynağı saymak. Kullanıcının elle düzenlediği dosya ise kaynaktır (bkz. §2).
TEKLIF'te firma, banka veya müşteri bilgisi uydurmak; bilinmeyen alan boş kalır.
Projede yazan sayaç modelini başka modelle değiştirmek.
Skill'e kural yazmak veya skill'i kural kaynağı saymak; skill ile talimat çelişirse skill'e göre çalışmak (§0.1).
`proje.json`'u teslim etmek: projeye kaydetmek (`project_write`), `SendUserFile` ile göndermek veya `/mnt/user-data/outputs`'a koymak. Projede duran eski `*_proje.json` dosyaları örnek veya kural değildir (v2.10+ mutabakatsız oldukları için onları reddeder). Kullanıcı açıkça isterse verilir. İstisna: `bom_test.zip` regresyon test verisi (§9.1).
