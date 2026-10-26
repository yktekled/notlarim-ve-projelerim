# Beni oku — Claude için not

## Kullanıcıyla konuşma
- Her zaman Türkçe konuş. Çok basit anlat (10 yaşındaki bir çocuğa anlatır gibi). Kullanıcı kod bilmiyor.
- Seçenek sorarken seçenekleri rakamla numaralandır; kullanıcı yalnız rakam yazar.
- Ekran kartı (GPU) kullanan hiçbir şeyi sormadan çalıştırma.
- Her yeni proje (BOM veya üretim onay) ayrı sohbette yapılır.

## Klasörler
| Klasör | İçerik |
|---|---|
| `bom-robotu/talimat/BOM_ROBOTU_talimat.md` | BOM robotu kuralları (tek kural kaynağı) |
| `bom-robotu/skills/` | akış skill'leri: bom-robotu · bom-yenile · bom-test |
| `bom-robotu/program/bom_kur.py` | BOM kurucu betik (v2.12); `bom_kur_v212.zip` aynısı; `eski/` geçmiş sürümler |
| `bom-robotu/test/bom_test.zip` | regresyon test paketi |
| `bom-robotu/veri/` | `birlesik.csv` (fiyat), `ebat.csv` (ölçü), `tekled_logo.webp` |
| `uretim-onay-robotu/` | üretim onay talimatı (`00_PROJE_TALIMATI.md` + 01–06), `bom_oku.py`, `uretim_onay.py`, `ornek_proje.json` |
| `musteri-bom/` | müşteri BOM Excel'leri (Drive klasör yapısıyla) |
| `musteri-uretim-onay/` | müşteri üretim onay dosyaları |

## Komutlar
- Şema + `BOM` / `ARA` → BOM robotu talimatı ve `bom-robotu/skills/bom-robotu/SKILL.md` akışı.
- Excel + `YENİLE` / `DÖNÜŞTÜR` → `bom-robotu/skills/bom-yenile/SKILL.md`.
- `TEST` → `bom-robotu/skills/bom-test/SKILL.md`.
- BOM Excel + `üretim onay` → `uretim-onay-robotu/00_PROJE_TALIMATI.md`.

## Bulutta çalışma (talimatlar Cowork/claude.ai projesi için yazıldı)
Talimattaki yollar burada yok; karşılıkları:
| Talimatta | Bulutta |
|---|---|
| `/mnt/project/...`, `project_read` | bu repodaki dosyalar (yukarıdaki tablo) |
| `/home/claude/sys`, `/home/claude/work` | oturumun scratchpad klasöründe `sys/` ve `work/` |
| `/mnt/user-data/outputs/` + `SendUserFile` | çıktıyı scratchpad'e yaz, `SendUserFile` ile gönder |
| `/mnt/skills/public/xlsx/scripts/recalc.py` | aynı yolda var (LibreOffice kurulu) |
| `tools/`, `docs/` (üretim onay) | `uretim-onay-robotu/` klasörü |

BOM robotu hazırlığı (§3 yerine):
```
SYS=<scratchpad>/sys; mkdir -p $SYS <scratchpad>/work
cp bom-robotu/program/bom_kur.py bom-robotu/veri/birlesik.csv bom-robotu/veri/ebat.csv $SYS/
export BOM_VERI=$SYS
head -2 $SYS/bom_kur.py   # "v2.12" görülmeli
```
Regresyon testi: `bom_test.zip` → `work/test`'e aç, `bom_test_py.txt` → `$SYS/bom_test.py`, sonra `python3 $SYS/bom_test.py <work>/test`.
261026: v2.12 ile bulutta test 4/4 GEÇTİ.

## Bilinen eksikler
- `kaynaklar.csv` repoda yok (zorunlu değil).
- `ezdxf` kurulu değil: DXF şema gelirse önce `pip install ezdxf`.
- Üretim onay robotu: `uretim_onay.py` DejaVuSansCondensed fontu ister; yoksa önce `apt-get install -y fonts-dejavu-extra`.
- `bom_oku.py` hem eski "fiyat burda" şablonunu hem BOM robotu (bom_kur.py) çıktısını okur; `--sayfa` verilmezse "fiyat burda" → "BOM" → ilk sayfa.
- `proje.json` çalışma dosyasıdır; repoya commit edilmez, kullanıcıya gönderilmez (BOM talimatı §11).
