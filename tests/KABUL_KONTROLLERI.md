# Excel LTSC 2024 masaüstü kabul kontrolleri — ÇALIŞTIRILMADI

Bu liste yalnızca `../deliverables/Duplicate_Audit_Studio_VBA_Taban.xlsx` dosyasından oluşturulan VBA kitabı içindir. Dinamik formüllü sürüm için `formul/KABUL_KONTROLLERI.md` listesini kullanın.

Excel sürümünü ve derleme numarasını, 32/64 bit bilgisini, işletim sistemini, işlemciyi, belleği ve test tarihini kaydedin. Çalışma kitabının bir kopyasını kullanın.

- [ ] XLSX dosyasını tam hesaplatın: ad, değer, başvuru veya sıfıra bölme hatası olmamalı; sonuçlar 370/180/90/100/90 olmalı.
- [ ] Kriter seçimlerini değiştirin; grupları elle doğrulanmış örnekle karşılaştırın. Sıfır kriter seçiminde geçersiz analiz durumu açıkça gösterilmeli.
- [ ] Bütün tabloyu artan/azalan sıralayın: kaynak değerleri, grup ve orijinal referansları tutarlı kalmalı.
- [ ] Her kayıt durumunu ve grubu filtreleyin; grup üyelerini yan yana karşılaştırın.
- [ ] İki modülü içe aktarın; VBA projesini derleyin. Kurulum iki kez çalıştırıldığında tam altı düğme bulunmalı.
- [ ] TestleriCalistir çalıştırın; geçen kontrol sayısını ve varsa hata ayrıntılarını kaydedin.
- [ ] Bir sütun/bir veri satırı içeren XLSX ve üç sayfalı başka bir XLSX yükleyin; ilk sayfa dışında bir sayfa seçin.
- [ ] Boş, tekrarlı ve ad çakışması oluşturabilecek başlıklar yükleyin; benzersiz başlıkları ve tüm sütunların korunmasını doğrulayın.
- [ ] Yeni bir dosya yükleyin: eski sonuçlar ve özet geçersizleşmeli.
- [ ] Sıfır kriter reddedilmeli; tek/çoklu kriter ve kriterlerle birlikte sütun sırasını değiştirme denenmeli; kayıt silinmemeli.
- [ ] Türkçe harfleri, başında sıfır bulunan kodları, metin/sayı ayrımını, seri tarihleri, boşlukları, uzun anahtarları ve joker karakterleri kontrol edin.
- [ ] =, +, -, @ ile başlayan metinler düz metin kalmalı; kaynak çalışma kitabı değiştirilmemeli.
- [ ] CSV içe aktarmada UTF-8 BOM / Windows-1254; virgül / noktalı virgül / sekme; tırnak içindeki ayraçlar ve çok satırlı alanlar denenmeli.
- [ ] Anlık raporu dışa aktarın ve XLSX'i yeniden açın; satır sayısı, bütün kaynak değerleri, tarih biçimleri, filtreler, renkler ve referanslar eşleşmeli.
- [ ] Dışa aktarılan dosyada 1/0 kriterlerini değiştirin: sonuçlar statik olduğu için aynı kalmalı.
- [ ] PerformansOlc işlemini 500/5.000/20.000/50.000 satırla çalıştırın; motor süresini saniye olarak kaydedin.
- [ ] Dört veri büyüklüğünde içe aktar → analiz et → filtrele → sırala → dışa aktar akışını ayrıca deneyin; toplam süreyi, bellek kullanımını, arayüz yanıtını, mutabakatı ve çıktı satır sayısını kaydedin.
- [ ] Düğmeler, hata mesajları ve test sonuçlarında Türkçe harfler doğru görünmeli; farklı Windows dil ayarlarında modüller içe aktarılabilmeli.

Dosya bütünlüğü iş akışının başarılı olması bu masaüstü kontrollerinin geçtiğini göstermez.
