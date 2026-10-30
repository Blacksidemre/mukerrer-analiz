# Doğrulama kaydı — 31 Ekim 2026

## Geliştirme ortamında yapılan kontroller

- Üretilen iki XLSX dosyası açıldı; ZIP bütünlüğü ve XML ayrıştırma kontrolleri geçti.
- Beklenen altı sayfa, filtreler, koşullu satır renkleri ve kaynak satır bilgileri yapısal olarak doğrulandı.
- Her dinamik kitapta 6.300 formül sözdizimsel olarak ayrıştırıldı. Bu kontrol Excel'in formülleri hesaplaması anlamına gelmez.
- 370 kaynak kaydı korundu; önbellekteki kontrol değerleri bütün kaynak sütunlarıyla karşılaştırıldı ve fark bulunmadı.
- Paylaşılan sentetik kayıtlar başlangıçtaki dört kriterle bağımsız gruplandı: **180 tekil, 90 orijinal, 100 fazladan mükerrer, 90 tekrarlı grup**. Mutabakat: **370 = 180 + 90 + 100**.
- Formül önbellekleri beklenen sonuçları içerir. Bu değerler dosya oluşturulurken yazılmıştır; **Excel'in formülleri çalıştırdığının kanıtı değildir**.
- Dört sayfalı statik çıktı açıldı; hücre formülü kalmadığı doğrulandı. Bu dosya dinamik değildir; analizin anlık görüntüsüdür.
- Teslim edilen XLSX dosyalarında dış bağlantı parçaları veya VBA ikilisi bulunmaz.
- Dosya özetleri ve formül sayıları `dogrulama-sonuclari.json` içindedir. Yüklenen kaynak dosya ve analizi GitHub'a eklenmemiştir.

## Çalıştırılmayan veya doğrulanmayan kontroller

- Excel LTSC 2024'te yeniden hesaplama ve yerleşik sıralama davranışı.
- VBA derlemesi ve `TestleriCalistir` içindeki 17 doğrulama kontrolü.
- Masaüstü Excel'de içe aktarma, kriter değiştirme, filtreleme, sıralama, yan yana karşılaştırma ve dışa aktarma akışı.
- 500 / 5.000 / 20.000 / 50.000 satırlık VBA performans ölçümleri. Hız veya arayüz yanıt süresi iddiası yoktur.
- Yayın sırasında PowerShell tabanlı GitHub Actions işinin çalıştığı gözlemlenmedi. Bu iş yalnızca çalışma kitabının dosya yapısını kontrol eder.
- Gerçek Excel ekran görüntüsü, imzalı makro ikilisi veya derlenmiş XLSM üretilmedi.

## Excel üzerinde tamamlanacak doğrulama

Windows ve Excel LTSC 2024 üzerinde `README.tr.md` rehberini izleyin. İki modülü içe aktarın, derleyin, testleri ve performans ölçümlerini çalıştırın; `tests/KABUL_KONTROLLERI.md` listesini tamamlayın. Bu kontroller geçmeden inceleme sürümünü üretim ortamında test edilmiş olarak tanımlamayın. Dosya oluşturma ve inceleme araçları teslim edilen çözümün çalışması için gerekmez; uygulama mantığı yalnızca Excel formülleri ve VBA'dan oluşur.

## Türkçeleştirme kontrolü

Belgeler, iş akışı mesajları, çalışma kitabının açıklamaları ve VBA kullanıcı metinleri Türkçeleştirildi. Çalışma kitabındaki formüller, hesaplama önbellekleri ve veri sayfaları önceki sürümle karşılaştırıldı. Türkçe metinlerin dosya içindeki karakter bütünlüğü kontrol edildi. Bu kontroller VBA'nın Excel içinde çalıştırıldığı anlamına gelmez.
