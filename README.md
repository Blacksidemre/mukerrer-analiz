# Duplicate Audit Studio · Excel Sürümü

**Microsoft Excel LTSC 2024 içinde çok kriterli, kesin ve izlenebilir mükerrer kayıt analizi.**

[![Çalışma kitabı bütünlüğü](https://github.com/Blacksidemre/m-kerrer-analiz/actions/workflows/integrity.yml/badge.svg)](https://github.com/Blacksidemre/m-kerrer-analiz/actions/workflows/integrity.yml)

[Kullanım rehberi](README.tr.md) · [Excel kitabını indir](deliverables/Duplicate_Audit_Studio.xlsx) · [Doğrulama raporu](docs/DOGRULAMA.md)

Veri analistleri için hazırlanan bu çözüm, hiçbir kaydı silmeden mükerrerleri belirler. Sütunları 1/0 değerleriyle seçin, satır renklerinden kayıt durumlarını inceleyin, her grubun ilk kaynak kaydına ulaşın ve denetlenebilir bir Excel raporu oluşturun. Uygulama yalnızca Excel formülleri ve VBA ile çalışır.

> **Sürüm durumu:** Formüllü çalışma kitabı ve VBA kaynakları teslim edilmiştir. Geliştirme ortamında masaüstü Excel bulunmadığı için VBA çalıştırma, Excel formüllerinin yeniden hesaplanması, uçtan uca kullanım ve 50.000 satır performansı Excel LTSC 2024 üzerinde doğrulanmayı beklemektedir. Bu bir **inceleme sürümüdür**. XLSX dosyası gömülü makro içermez; XLSM oluşturmak için iki VBA modülünü içe aktarın.

## Çalışma biçimleri

| Çalışma biçimi | Dosya veya işlem | Yeniden hesaplama | Kapsam |
|---|---|---|---|
| Formüllü örnek | `deliverables/Duplicate_Audit_Studio.xlsx` | Kriter değişince Excel formülleri yeniden hesaplanır | Sabit 370 satır ve 12 sütun |
| VBA analizi | `vba/DuplicateAudit.bas` | Veri/kriter değişince `AnalizEt` çalıştırılır | En fazla 50.000 satır / 100 sütun; performans ölçülmedi |
| VBA testleri | `vba/AuditTests.bas` | Excel içinde ayrıca çalıştırılır | Doğrulama kontrolleri ve dört veri büyüklüğünde ölçüm |
| Dışa aktarma | `ExcelAktar` | Statik değerler; yeniden analiz yapılmaz | Analiz edilen bütün kayıtlar |

Formül modunda kesin karşılaştırma sayısı satır sayısının karesiyle artar. Bu örneği 50.000 satıra genişletmeyin. VBA motoru sözlük yapısıyla çalışır; beklenen işlem miktarı satır sayısı × seçili sütun sayısıyla orantılıdır. Bu algoritmanın özelliğidir; ölçülmüş hız iddiası değildir.

## Hızlı başlangıç

1. XLSX dosyasını Excel LTSC 2024 ile açın. **Formüller → Hesaplama Seçenekleri → Otomatik** seçin. **Ctrl+Alt+F9** ile tam hesaplatın.
2. `02_Kriterler` sayfasında `B7:B18` hücrelerini 1 veya 0 yapın. Başlangıç seçimi: dosya kodu, borçlu kodu, ürün türü ve anapara.
3. `03_Kontrol` sayfasında başlık filtreleriyle durum veya grup seçin. Sıralarken **bütün tabloyu** sıralayın; tek sütunu bağımsız sıralamayın.
4. `04_Ozet` sonucunu kontrol edin: **370 = 180 tekil + 90 orijinal + 100 fazladan mükerrer**. Mükerrer grup sayısı **90**.

## VBA kurulumu

Yeni dosyaları yüklemek, seçenekleri uygulamak, büyük veri kümelerini analiz etmek ve rapor almak için:

1. Çalışma kitabını yerel bir klasöre **Excel Makro Etkin Çalışma Kitabı (.xlsm)** olarak kaydedin.
2. **Alt+F11 → Dosya → Dosya Al** yoluyla `vba/` klasöründeki iki `.bas` dosyasını içe aktarın. VBA düzenleyicisi İngilizceyse menü adları **File → Import File** şeklindedir.
3. VBA düzenleyicisinde **Hata Ayıkla → VBAProject'i Derle** komutunu çalıştırın; İngilizce menü karşılığı **Debug → Compile VBAProject**. Excel'e dönüp **Alt+F8 → Kurulum** çalıştırın.
4. Yalnızca incelediğiniz makroları kurumunuzun politikalarına göre etkinleştirin. Uygulama, VBA proje nesne modeline güven erişimi gerektirmez.
5. Önce **TestleriCalistir**, ardından **PerformansOlc** çalıştırın. Sonuçlar `05_Testler` sayfasına yazılır.

Kurulum makrosu altı işlevsel düğme oluşturur. Diğer makrolara Alt+F8 ile ulaşılır. Kurulumdan sonra dosyanın adını değiştirirseniz düğme bağlantılarını güncellemek için `Kurulum` işlemini yeniden çalıştırın.

## Dosya yükle → kriter seç → analiz et → incele → dışa aktar

- **DosyaYukle:** Bir XLSX dosyası ve sayfa numarası seçin. İlk satır başlık olmalıdır. Mevcut veri onayınızdan sonra değiştirilir; eski sonuçlar temizlenir. Yeni veri içindeki boş ara satırlar dahil hiçbir kayıt silinmez. Boş veya tekrarlı başlıklara benzersiz adlar verilir.
- **CSV:** Excel'in **Veri → Metin/CSV'den** komutunu kullanın. UTF-8 veya Windows-1254 kodlamasını ve virgül/noktalı virgül/sekme ayracını seçin. Kod sütunlarını **Metin** olarak yükleyin, XLSX olarak kaydedip içe aktarın. Makro kodlamayı tahmin etmez ve kod sütunlarını kendiliğinden sayıya dönüştürmez.
- **02_Kriterler:** Yüklenen her sütunun yanında 1/0 seçimi bulunur. Yalnızca VBA için `F7`, metnin kenarlarındaki ASCII boşluklarını temizler; `F8`, Türkçe büyük/küçük harf farkını kaldırır. Varsayılan kesin eşleşmedir.
- **AnalizEt:** Bütün kayıtları etkin kriterlerle analiz eder. Hiç kriter seçilmemesi hatadır. Her kayıtta durum, grup, eşleşme adedi, orijinal referansı ve kaynak satır numarası yer alır. Kaynak satırları başlıktan sonra, 2'den başlar.
- **GrubuIncele:** Tam grup kimliğiyle filtreler. Boş giriş filtreyi kaldırır. Başlık menüleri durum filtrelemeyi ve metin aramasını da destekler.
- **YanYanaKarsilastir:** Seçilen grubun en fazla 20 kaydını `06_Karsilastirma` sayfasında yan yana gösterir. Grubun tamamı sonuç tablosunda bulunur.
- **ExcelAktar:** Dört analiz sayfasını yalnızca değerlerden oluşan anlık rapor olarak kaydeder. Veri/kriter değişikliğinden sonra analizi yeniden çalıştırdığınızı doğrulayın. Filtreler ve koşullu renkler korunur. Dışa aktarılan kriterlerin değiştirilmesi raporu yeniden hesaplamaz.

## Eşleştirme kuralları

**TEKİL**, grubunda tek olan kayıttır. **ORİJİNAL**, tekrarlı grubun kaynak dosyadaki ilk kaydıdır. **MÜKERRER**, aynı grubun sonraki kayıtlarıdır. Orijinal belirlenirken kaynak sırası kullanılır.

Anahtarlar veri türünü ve her değerin uzunluğunu içerir. `AB + C` ile `A + BC` farklıdır. Metin ve sayı ayrıdır; metin `001`, metin `1` ile eşleşmez. Boş hücre ile boş metin eşleşir. Joker karakterler normal karakter olarak değerlendirilir. Varsayılan karşılaştırma harf ve boşluk farklarını korur. Yaklaşık eşleştirme, tutar yuvarlama veya Unicode bileşim normalizasyonu yapılmaz.

Excel tarihleri seri sayılardır; tarih biçimi ayrı bir veri türü oluşturmaz. `0000` biçimiyle gösterilen sayı `1`, yine sayı `1` değeridir. Kimlik alanlarını içe aktarmadan önce metin olarak tutun. Excel'in 15 anlamlı basamak sınırı nedeniyle kaynakta kaybolmuş sayısal basamaklar geri getirilemez. Formüllü hücreler mevcut değerleriyle alınır; kaynak kitabı önceden hesaplatıp kaydedin. Kaynak formülleri, açıklamalar, grafikler, birleştirilmiş düzenler ve bağlantılar çalıştırılabilir nesneler olarak arşivlenmez.

VBA sözlükleri ikili karşılaştırma kullanır. İsteğe bağlı Türkçe harf dönüşümü `I → ı` ve `İ → i` eşleştirmelerini açıkça uygular. Formüllü örnek, VBA seçeneklerinden bağımsız olarak her zaman harfe duyarlı `EXACT` karşılaştırması kullanır. Grup kimliği ilk kaynak satırını içerir (`GRP-000002`); tek bir analiz içinde benzersizdir, farklı dosyalar arasında genel bir kimlik değildir.

VBA çıktısındaki orijinal referanslar sıralamayla değişmez. Formül modu sabit kaynak satır numaralarını ve bu numaralarla kaynak değerlerini okuyan formülleri kullanır. Bütün tabloyu sıralayın ve kontrol sütunlarını koruyun. Ham veriyi sıraladıktan veya değiştirdikten sonra analizi yeniden çalıştırın.

## Çalışma kitabı düzeni

| Sayfa | Görevi |
|---|---|
| `00_Baslangic` | Gezinme ve kurulumdan sonra makro düğmeleri |
| `01_Ham_Veri` | Kaynak değerler, başlıklar ve hücre biçimleri |
| `02_Kriterler` | Sütun bazında seçimler ve VBA seçenekleri |
| `03_Kontrol` | Bütün kaynak kayıtları ve kontrol alanları |
| `04_Ozet` | Sayılar, mükerrer oranı ve mutabakat |
| `05_Testler` | Masaüstü test durumu ve performans sonuçları |
| `06_Karsilastirma` | Gerektiğinde oluşturulan karşılaştırma sayfası |

Çalışma kitabında yeşil ağırlıklı sade bir tasarım, sabitlenmiş başlıklar, otomatik filtreler ve bütün satırı renklendiren yeşil/kırmızı koşullu biçimlendirme kullanılır. Geliştirme ortamında masaüstü Excel görüntüsü alınamadığından uygulama ekran görüntüsü sunulmamıştır.

## Test ve doğrulama

[Doğrulama raporunu](docs/DOGRULAMA.md) ve [masaüstü kabul kontrollerini](tests/KABUL_KONTROLLERI.md) inceleyin. GitHub Actions iş akışı ZIP/XML yapısını, beklenen sayfaları, formülleri, dış bağlantı ve makro ikilisi bulunmamasını ve modül dosyalarını kontrol eder. **Bu işlem Excel'i çalıştırmaz, VBA'yı derlemez ve uygulama davranışının doğruluğunu kanıtlamaz.**

Eşleştirme kontrolleri için `TestleriCalistir` çalıştırın. `PerformansOlc`, motoru bilgisayarınızda 500, 5.000, 20.000 ve 50.000 satırla ölçüp süreleri saniye olarak kaydeder. İçe aktarma, biçimlendirme, Excel arayüzü ve dışa aktarma bu ölçüme dahil değildir; bunları kabul listesiyle ayrıca ölçün. Ölçülmüş bir hız veya arayüz yanıt süresi garantisi verilmez.

## Veri güvenliği

Uygulama işlemleri Excel içinde gerçekleşir. VBA ağ çağrısı, telemetri veya harici yapay zekâ servisi kullanmaz. Kaynak dosyalar salt okunur olarak, makro çalıştırma ve bağlantı güncelleme kapalı biçimde açılır. İçe aktarma kaynak dosyayı kaydetmez. Güvenilmeyen çalışma kitapları için kurumunuzun politikalarını izleyin; uygulama bir güvenli çalıştırma ortamı sağlamaz. Formül gibi görünen metinler düz metin olarak yazılır.

Depodaki örnek bağımsız üretilmiş sentetik veridir; kullanıcının yüklediği çalışma kitabı değildir. Gerçek kullanıcı verisi depoya eklenmez. Excel/OneDrive otomatik kaydının dosyaları eşitlemesini istemiyorsanız yerel klasör kullanın. Yerel işleme otomatik mevzuat uyumluluğu anlamına gelmez.

## Sınırlamalar ve sonraki geliştirmeler

- Bu ortamda Excel/VBA derleyicisi bulunmadığı için makro kurulumu elle yapılır; çalıştırılmaya hazır XLSM sunulmamıştır.
- Formüllü örnek sabit aralıklara bağlıdır. Farklı veri büyüklükleri VBA gerektirir. VBA sonuçları her değişiklikten sonra açıkça yenilenmelidir.
- Excel/VBA işlemleri eşzamanlıdır. 50.000 satır çalıştırma ve iptal davranışı doğrulanmamıştır; 64 bit Excel önerilir.
- İçe aktarma değerleri ve hücre biçimlerini korur. Kaynak kitabın tüm özelliklerini kopyalamaz. Sondaki yalnızca biçimlendirilmiş satırlar alınmaz; iç boş satırlar korunur.
- Formül anahtarları Excel'in hücre başına 32.767 karakter sınırına tabidir. VBA sözlük anahtarları hücrelere yazılmaz.
- CSV yükleme Excel'in yerleşik içe aktarma arayüzünü kullanır. Bu sürüm masaüstü Excel içindir; web demosu bulunmaz.
- Doğrulama sonrasında planlananlar: imzalı VBA sürümü, değişikliklerde otomatik geçersizleştirme, ayrıntılı veri önizlemesi ve test edilmiş XLSM paketi.

## Depo yapısı

```text
vba/                    VBA uygulaması ve motor testleri
deliverables/           Formüllü çalışma kitabı ve sentetik örnek
docs/                   Doğrulama raporu ve dosya kontrol sonuçları
tests/                  Excel kabul kontrol listesi
.github/workflows/     Dosya bütünlüğü iş akışı
```

Dosya yolları, VBA modül/prosedür adları ve Excel'in dosya içinde kullandığı İngilizce işlev adları teknik uyumluluk için korunur. VBA dosyaları farklı Windows kod sayfalarında içe aktarılabilsin diye ASCII biçimindedir; kullanıcı mesajlarındaki Türkçe harfler `ChrW` ile oluşturulur. Kod açıklamaları bu nedenle Türkçe, ASCII yazımlıdır.

## Lisans

Henüz açık kaynak lisansı verilmemiştir. Telif hakları depo sahibine aittir. Kamuya dağıtımdan önce uygun bir lisans seçilebilir.
