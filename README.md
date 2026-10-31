# Duplicate Audit Studio · Dinamik Formüllü Excel

**Excel LTSC 2024 içinde, kendi tablo ve sütunlarınızla çok kriterli mükerrer analizi.**

[![Çalışma kitabı bütünlüğü](https://github.com/Blacksidemre/mukerrer-analiz/actions/workflows/integrity.yml/badge.svg)](https://github.com/Blacksidemre/mukerrer-analiz/actions/workflows/integrity.yml)

[Dinamik Excel'i indir](deliverables/Duplicate_Audit_Studio.xlsx) · [Kullanım rehberi](README.tr.md) · [Deneme verisi](deliverables/Ornek_Ham_Veri.xlsx) · [Doğrulama](docs/DOGRULAMA.md)

Ham veri bir **Excel tablosudur**. Tabloya eklenen satırlar ve sütunlar formüllerin kapsamına girer. Kriter sayfasında mevcut başlıklar otomatik listelenir. Ad, tutar, icra dairesi, dosya numarası veya başka alanları seçip 1/0 değerleriyle analize dahil edebilirsiniz. Orijinaller yeşil, fazladan mükerrerler kırmızı gösterilir; kaynak kayıtlar silinmez.

> **Doğrulama durumu:** Dosya yapısı, formüllerin sözdizimi ve bağımsız eşleştirme örnekleri kontrol edildi. Bu ortamda masaüstü Excel bulunmadığından Excel LTSC 2024 üzerinde yeniden hesaplama ve kullanıcı akışı henüz çalıştırılmadı. Formül önbellekleri Excel tarafından hesaplanmış test sonuçları değildir.

## Hemen kullanın

1. `Duplicate_Audit_Studio.xlsx` dosyasını Excel LTSC 2024'te açın. **Hesaplama Seçenekleri → Otomatik** seçin; Ctrl+Alt+F9 ile tam hesaplatın.
2. `01_Ham_Veri` sayfasındaki **HamVeri** tablosuna başlıklarınızla birlikte veriyi yerleştirin. Eski kayıtları önce temizleyin. **Tablo Tasarımı → Tabloyu Yeniden Boyutlandır** ile sınırın yeni verinizi tam kapsadığını doğrulayın.
3. `02_Kriterler` sayfasının solunda güncel başlıkları görün. Sağdaki sarı **Sütun / Etkin / Tür** seçim tablosunda başlığı açılır listeden seçin, Etkin değerini **1** veya **0** yapın.
4. `03_Kontrol` sayfasını inceleyin. Durum, grup kimliği ve metin aramasıyla filtreleyin; bir sütun seçerek sıralayın.
5. `04_Ozet` sayfasında **Toplam = Tekil + Orijinal + Fazladan mükerrer** mutabakatını kontrol edin.

**Kriter seçimi neden sütun adına bağlı?** Sütunların yeri değişince seçimlerin başka alanlara kaymaması için. Yeni başlıklar soldaki listede otomatik görünür; sağdaki seçim tablosuna eklenene kadar pasiftir. Daha önce seçilmiş bir başlık silinir veya değişirse analiz durur ve seçim düzeltmenizi ister. Başlığı aynı kalan bir sütunun seçimi korunur; yeni veri yüklerken önceki seçimleri yine gözden geçirin.

## Örnek seçimler

| Amaç | Etkin yapılacak alanlar | Önerilen tür |
|---|---|---|
| Aynı adlı kayıtlar | Ad | Kesin; gerekirse boşluk ve harf seçenekleri |
| Aynı tutarlı kayıtlar | Tutar | Gerçek sayılarda Kesin; Türkçe sayı metinlerinde Türkçe tutar |
| Aynı icra dosyaları | İcra Dairesi + Dosya Numarası | Kesin |
| Aynı kişi ve tutar | Ad + Tutar | Her alanın gerçek türüne göre |
| Aynı dosya numarası | Dosya Numarası | Kesin; kimlik alanlarını kaynakta metin tutun |

## Veri türleri ve normalizasyon

- **Kesin:** Metin ve sayı farklıdır. Metin `001`, metin `1` ve sayı `1` ayrı anahtarlar üretir. Harf ve boşluklar varsayılan olarak korunur.
- **Metin:** Sayı ve metin, Excel'in metinsel gösterimleri üzerinden karşılaştırılır. Sayı `1` ile metin `1` eşleşir; metin `001` ayrı kalır.
- **Türkçe tutar:** Sayısal hücreleri korur; `1.234,50` gibi metinleri virgül ondalık, nokta binlik ayracıyla sayıya çevirir. Yabancı sayı biçimlerinde kullanmayın. Dönüştürülemeyen değer analizi durdurur. Para birimi yazılarını önce temizleyin. Tutarlar yuvarlanmaz.
- **Fazla boşlukları temizle:** `02_Kriterler!L7=1`; metinlerin kenarlarındaki normal boşlukları kaldırır ve içteki tekrarlı normal boşlukları teke indirir. Sekme ve bölünemez boşlukları temizlemez.
- **Türkçe harf farkını kaldır:** `L8=1`; `I → ı`, `İ → i` dönüşümleri uygulanır ve küçük harfe çevrilir.
- **Tarih:** Excel'in gerçek tarih hücreleri seri sayıdır. Aynı seri sayı ile eşdeğerdir. Tarih gibi görünen metinler kendiliğinden tarihe çevrilmez. Sonuç sayfası kaynak hücre biçimlerini dinamik kopyalayamadığından tarihleri seri sayı olarak gösterebilir; ham verideki tarih biçimi korunur.

Boş hücre ve boş metin eşleşir. Tamamen boş ara satırlar tablo içindeyse kayıt sayılır; seçilen alanları boş olan kayıtlar aynı grupta toplanabilir. Excel hata değerleri analizi durdurur. Unicode bileşim normalizasyonu yapılmaz. Excel'in 15 anlamlı basamak sınırı nedeniyle önceden kaybolmuş kimlik basamakları geri getirilemez.

Anahtarlar tür + uzunluk + değer bileşenleriyle oluşturulur; `AB + C` ile `A + BC` eşleşmez. Karşılaştırmada joker karakterler normal karakterdir. `TEKİL`, grubunda tek kayıttır; `ORİJİNAL`, tekrarlı grubun ham tablodaki ilk kaydı; `MÜKERRER` sonraki kayıtlarıdır. Grup kimliği ilk kaynak satırından türetilir. Sonuç görünümünün filtre/sıralama ayarları bu referansı değiştirmez. Ham tabloyu sıralamak ise kaynak sırasını değiştirdiği için orijinali yeniden belirler.

## Dinamik aralıkların sınırları

- En fazla **2.000 veri satırı ve 100 sütun** desteklenir. Aşımda açık mesajla analiz durur. Bu sınır ölçülmüş hız garantisi değildir; tam karşılaştırmaların yükü satır sayısının karesiyle artar.
- Dinamik kapsam, **HamVeri tablosunun içidir**. Tablo dışına yapıştırılmış veriler dahil olmaz. Özellikle daha küçük veriyle değiştirirken eski hücreleri temizleyip tabloyu küçültün.
- Kriterler için 100 seçim satırı vardır. Aynı başlığı iki satırda kullanmak hatadır. Etkin bir seçimde bulunmayan veya boş başlık kullanmak hatadır.
- Başlıklar Excel tablosunda benzersiz ve dolu olmak zorundadır. Excel boş/tekrarlı başlıkları kendisi adlandırır; analiz bu son adları kullanır.
- Excel'in dinamik dizi sonuçları **taşma alanı** kullanır. Kriter listesinin ve sonuçların üzerine veri yazmayın; engellenirse `#TAŞMA!` hatası oluşur.
- Tek bir satırın eşleştirme anahtarı 32.767 karakteri aşarsa analiz durur. Kaynak hücre değerleri değiştirilmez.
- Sonuçlar hücre formülleridir. Görünüm kontrollerini kullanın; taşan sonuç aralığını elle düzenlemeyin veya Excel'in yerleşik sıralamasıyla değiştirmeyin.

## Sayfalar

| Sayfa | İçerik |
|---|---|
| `00_Baslangic` | Kullanım adımları ve sınırlar |
| `01_Ham_Veri` | HamVeri adlı kaynak Excel tablosu; canlı satır renkleri |
| `02_Kriterler` | Otomatik başlık listesi, sütun adına bağlı seçimler ve tür seçenekleri |
| `03_Kontrol` | Dinamik sonuçlar, renkler, arama, filtre ve sıralama |
| `04_Ozet` | Filtrelerden bağımsız metrikler ve mutabakat |
| `05_Testler` | Kontrol adımları ve masaüstü test durumu |
| `_Hesaplama` | Gizli yardımcı formüller |

[Formül dökümü](docs/FORMULLER.md) teknik inceleme içindir. Çalışma kitabında LET, FILTER, XLOOKUP, XMATCH, BYROW, LAMBDA, CHOOSECOLS, HSTACK, SORT ve dinamik diziler kullanılır. Hedef Excel LTSC 2024'tür; eski Excel sürümleri desteklenmez.

## Raporu kaydetme ve VBA seçeneği

Dinamik sürümü normal XLSX olarak kaydetmeniz yeterlidir. Bu dosyada makro veya otomatik statik rapor düğmesi yoktur. Statik rapor gerekiyorsa sonuçları ayrı bir kitaba **Özel Yapıştır → Değerler** ile aktarın; bu kopya kriter değişikliklerini takip etmez.

Önceki VBA akışı ayrı tutulmuştur: [VBA taban kitabı](deliverables/Duplicate_Audit_Studio_VBA_Taban.xlsx). Makroları **yalnızca bu taban kitaba** kurun; dinamik formüllü kitap farklı hücre düzeni kullanır ve eski makrolarla uyumlu değildir. VBA tabanını `.xlsm` kaydedip `vba/DuplicateAudit.bas` ile `vba/AuditTests.bas` modüllerini içe aktarın; `Kurulum` çalıştırın. Veri/kriter değişince `AnalizEt` gerekir. VBA motorundaki 50.000 satır sınırı ve testler de masaüstü Excel'de doğrulanmayı bekler. VBA tabanındaki eski formüllü örnek 370 satır/12 sütuna sabittir.

## Test, güvenlik ve lisans

[Doğrulama raporu](docs/DOGRULAMA.md), [formül kabul listesi](tests/formul/KABUL_KONTROLLERI.md) ve [VBA kabul listesi](tests/KABUL_KONTROLLERI.md) yapılan ve bekleyen kontrolleri ayırır. GitHub Actions, çalışma kitaplarının ZIP/XML yapısını kontrol eder; Excel veya VBA çalıştırmaz.

Uygulama Excel içinde çalışır; analiz sunucusu, telemetri ve harici AI servisi yoktur. OneDrive eşitlemesi ayrı bir Excel/Windows ayarıdır. Yerel işleme otomatik mevzuat uyumluluğu sağlamaz. Örnek dosyalar sentetiktir. Kullanıcının yüklediği dosya depoya eklenmez.

Henüz açık kaynak lisansı verilmemiştir; telif hakları depo sahibine aittir. Teknik dosya/prosedür adları ve Excel'in dosya içinde kullandığı İngilizce işlev adları uyumluluk için korunur.
