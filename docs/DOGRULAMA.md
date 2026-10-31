# Dinamik formüllü sürüm — doğrulama kaydı

## Yapılan kontroller

- Üç teslimat dosyası ZIP ve XML bütünlüğü kontrollerinden geçti; dış bağlantı ve gömülü VBA ikilisi yok.
- Yeni dinamik kitapta **7 sayfa, 2 gerçek Excel tablosu ve 24 formül** bulundu. Formüllerin sözdizimi, parantez dengesi ve temel işlevlerin bağımsız değişken sayıları kontrol edildi.
- Taşma başvuruları OOXML içinde `ANCHORARRAY` olarak kaydedildi. LAMBDA parametreleri `_xlpm` önekiyle yazıldı. Türkçe metinlerin karakter bütünlüğü kontrol edildi.
- HamVeri tablosu, başlangıçta `A1:L371`; KriterSecimleri tablosu `G6:I106`. Ham verinin 370 satır ve 12 sütundaki bütün hücre değerleri önceki sürümle karşılaştırıldı; fark yok.
- Ham veri ve sonuç sayfalarında koşullu renklendirme kuralları bulundu. Gizli yardımcı sayfa ve dinamik başlıklar doğrulandı.
- **15 bağımsız referans senaryosu** geçti: tek/çoklu kriter, üçlü grup, sayı/metin ayrımı, Türkçe tutar, metne dönüştürme, boş değerler, anahtar çakışması, Türkçe harf, boşluk, joker karakterler, icra/dosya eşleşmesi, sıfır kriter reddi, sütun sırası ve seri tarih.
- Bu senaryolar formül tasarımının beklenen sonuçlarını bağımsız hesaplar; **Excel motorunu çalıştırmaz ve gerçek formüllerin doğru hesaplandığını kanıtlamaz**.
- Örnek_Ham_Veri.xlsx: Ad, Tutar, İcra Dairesi, Dosya Numarası ve Tarih alanlarında 10 sentetik kayıt. Beklenen sonuçlar `tests/formul/BEKLENEN_SONUCLAR.md` dosyasında.
- Önceki çalışma kitabı VBA için ayrı `Duplicate_Audit_Studio_VBA_Taban.xlsx` adıyla korundu: 6 sayfa, 6.300 formül. VBA kaynak kodu bu revizyonda değiştirilmedi.
- Dosya özetleri ve sayılar `dogrulama-sonuclari.json` içindedir. Formüllerin tam dökümü `FORMULLER.md` dosyasındadır.

## Çalıştırılmayan kontroller

- Masaüstü Excel LTSC 2024 üzerinde formül hesaplaması, otomatik tablo genişlemesi, koşullu renklerin çizilmesi ve kullanıcı akışı.
- 500 veya 2.000 kayıtla gerçek hesaplama süresi ve arayüz yanıtı.
- GitHub Actions işinin uzak sunucudaki sonucu.
- Eski VBA sürümünün derlemesi, motor testleri, içe/dışa aktarma ve 50.000 satır ölçümleri.

Bu ortamda masaüstü Excel yoktur. Dosyalardaki formül önbellekleri kitap hazırlanırken yazılmış örnek değerlerdir. Dinamik dizi sonuçları açılışta hesaplanır; önbelleğe bakarak Excel testi geçtiği sonucu çıkarılamaz. Yeni kitap için `tests/formul/KABUL_KONTROLLERI.md`, eski VBA akışı için `tests/KABUL_KONTROLLERI.md` uygulanmalıdır.

## Kapsam ve bilinen sınırlar

Formüllü sürüm 2.000 satır / 100 sütun sınırında analizi açıkça durdurur; kısmi sonuç üretmek amaçlanmaz. Bu sınır performans ölçümü değildir. Sütun başlıkları otomatik listelenir; kullanıcı seçimi ayrı, başlığa bağlı giriş tablosundadır. Yeni başlıklar kullanıcı tarafından seçilene kadar pasiftir. Excel tabloları dışındaki hücreler kapsama girmez. Tarih hücrelerinin görüntü biçimi dinamik sonuçlara otomatik kopyalanmaz.
