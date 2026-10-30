# Duplicate Audit Studio — Excel LTSC 2024

Uygulama yalnızca Excel formülleri ve VBA kullanır. Web uygulaması veya sunucu gerektirmez.

**Teslimat durumu:** Formüllü XLSX ve içe aktarılabilir VBA kaynakları hazırdır. Bu ortamda Excel bulunmadığı için makrolar çalıştırılmadı; Excel yeniden hesaplama, kullanıcı akışı ve 50.000 satır performansı henüz doğrulanmadı. Dosyalar inceleme sürümüdür.

## Hemen deneyin

1. `deliverables/Duplicate_Audit_Studio.xlsx` dosyasını Excel LTSC 2024 ile açın.
2. Formüller → Hesaplama Seçenekleri → Otomatik seçin; Ctrl+Alt+F9 ile tam hesaplatın.
3. `02_Kriterler` sayfasında B7:B18 hücrelerini 1/0 olarak değiştirin.
4. `03_Kontrol` sayfasında satır renklerini ve grup filtrelerini inceleyin.
5. `04_Ozet` sayfasını kontrol edin. Başlangıç sonucu: **370 = 180 tekil + 90 orijinal + 100 mükerrer**, 90 grup.

Bu formüllü örnek 370 satır/12 sütuna sabittir. Yeni boyutlarda veri yüklemek için VBA kullanın. Büyük dosyalar için formülleri aşağıya kopyalamayın.

## Makroları kurun

1. XLSX dosyasını **Farklı Kaydet → Excel Makro Etkin Çalışma Kitabı (.xlsm)** ile kaydedin.
2. **Alt+F11 → Dosya → Dosya Al** üzerinden `vba/DuplicateAudit.bas` ve `vba/AuditTests.bas` dosyalarını alın.
3. **Debug / Hata Ayıkla → Compile VBAProject** ile derleyin.
4. Excel'e dönün; **Alt+F8 → Kurulum → Çalıştır** seçin. Ana sayfada altı işlem düğmesi oluşur.
5. Önce **TestleriCalistir**, sonra **PerformansOlc** çalıştırın. Gerçek sonuçlar `05_Testler` sayfasına yazılır.

Kod incelemesi ve kurumunuzun makro politikası doğrultusunda etkinleştirin. Genel makro güvenliğini düşürmeyin. Uygulama VBA proje nesne modeline güven erişimi istemez. Dosyanın adını değiştirirseniz düğme bağlantıları için Kurulum'u yeniden çalıştırın.

## Günlük kullanım

**DosyaYukle → Kriterleri seç → AnalizEt → GrubuIncele → ExcelAktar**

- Kaynak `.xlsx` dosyasını ve veri sayfasının numarasını seçin. İlk satır başlık olmalıdır.
- CSV için Excel **Veri → Metin/CSV'den** kullanın. UTF-8/Windows-1254 kodlamasını ve ayracı seçin; kod sütunlarını **Metin** yapın. XLSX olarak kaydedip yükleyin.
- Yükleme mevcut ham veriyi değiştirir; önceki analiz temizlenir. Yeni verideki kayıtlar silinmez.
- Her sütun için 1/0 seçin. En az bir sütun zorunludur.
- VBA'da F7=1 kenar boşluklarını temizler; F8=1 Türkçe büyük/küçük harf farkını kaldırır. Formül modu bu iki ayarı uygulamaz; kesin eşleşir.
- AnalizEt bütün kayıtları işler. Yeşil **ORİJİNAL**, kırmızı **MÜKERRER**, renksiz **TEKİL**.
- Başlık filtrelerinden durum/grup filtreleyin. Bütün tabloyu sıralayın; tek sütunu bağımsız sıralamayın. Orijinal referans kaynak sırasına dayanır.
- Alt+F8 → **YanYanaKarsilastir** ile bir grubun ilk 20 kaydını karşılaştırın. Bütün grubu GrubuIncele ile görüntüleyin.
- Veri/kriter/normalizasyon değiştiğinde **AnalizEt'i tekrar çalıştırın**. VBA çıktısı anlık görüntüdür.
- **ExcelAktar**, dört sayfalı, değerlerden oluşan statik XLSX çıkarır. Çıktı içindeki kriterleri değiştirmek yeniden analiz yapmaz.

## Veri kuralları

Metin `001`, metin `1` ve sayı `1` farklıdır. Ancak Excel'de sayı `1` üzerine `0000` biçimi uygulanması yeni bir kimlik oluşturmaz. Kodları baştan metin olarak tutun. Tarihler Excel seri sayısıdır. Boş hücre ve boş metin eşleşir. Türkçe harfler korunur; varsayılan karşılaştırma harfe duyarlıdır. Sayılar yuvarlanmaz. Unicode bileşim normalizasyonu yapılmaz.

Formüllü kaynakları yüklemeden önce Excel'de hesaplayıp kaydedin. İçe aktarma değerleri ve biçimleri korur; kaynak makroları, grafikler, yorumlar ve canlı formüller arşivlenmez. Başlıkları boş/tekrarlı sütunlar benzersiz adlandırılır. İç boş satırlar korunur.

## Güvenlik ve test durumu

Dosyalar bir analiz sunucusuna gönderilmez. VBA ağ çağrısı yapmaz. OneDrive eşitlemesi açıksa Excel dosyayı eşitleyebilir; yerel klasör tercih edebilirsiniz. Bu mimari otomatik mevzuat uyumluluğu sağlamaz.

[Doğrulama raporu](docs/VALIDATION.md) ve [Excel kabul kontrolü](tests/ACCEPTANCE.md) henüz çalıştırılmamış masaüstü kontrollerini belirtir. Test edilmemiş bir XLSM veya canlı demo teslim edilmiş gibi sunulmaz.
