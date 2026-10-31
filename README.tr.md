# Kendi verinizle dinamik mükerrer analizi

1. [Duplicate_Audit_Studio.xlsx](deliverables/Duplicate_Audit_Studio.xlsx) dosyasını Excel LTSC 2024 ile açın. Makro gerekmez.
2. **Formüller → Hesaplama Seçenekleri → Otomatik** seçin. İlk açılışta **Ctrl+Alt+F9** ile tam hesaplatın.
3. `01_Ham_Veri` sayfasındaki eski tablo verilerini temizleyin. Kendi verinizi, başlıklar ilk satırda olacak şekilde aynı alana yapıştırın.
4. **Tablo Tasarımı → Tabloyu Yeniden Boyutlandır** ile **HamVeri** tablosunun yeni verinin bütün satır ve sütunlarını tam kapsadığını doğrulayın. Daha küçük veri koyduysanız eski sütun ve satırları kapsamdan çıkarın. Tablo adını değiştirmeyin.
5. `02_Kriterler` sayfasında güncel başlıklar solda otomatik listelenir. Sağdaki sarı seçim tablosunda **Sütun** alanından istediğiniz başlığı seçin, **Etkin** alanına **1** yazın. Hariç tutmak için **0** yazın. Önceki dosyadan kalan etkin başlıkları temizleyin veya 0 yapın.
6. **Tür** varsayılan olarak **Kesin** kalsın. Metin biçimindeki Türkçe tutarları sayılarla eşleştirmek için **Türkçe tutar** seçebilirsiniz. Dosya numarası gibi kod alanlarında bu dönüşümü kullanmayın.
7. `03_Kontrol` ve ham veri satırları kendiliğinden renklenir: **ORİJİNAL yeşil**, **MÜKERRER kırmızı**, **TEKİL nötr**. Sonuç sayfasındaki sarı kontrollerle filtreleyin, arayın ve sıralayın.
8. `04_Ozet` sayfasındaki mutabakat farkı geçerli analizde **0** olmalıdır.

## Sık kullanılan seçimler

- Yalnızca ada göre: **Ad = 1**, diğerleri 0.
- Aynı ad ve tutar: **Ad = 1, Tutar = 1**.
- Aynı icra dosyası: **İcra Dairesi = 1, Dosya Numarası = 1**.
- Yalnızca tutara göre: **Tutar = 1**. Bu durumda farklı kişilerin aynı tutarı da aynı gruba girer; bu, seçtiğiniz kuralın sonucudur.

[Örnek ham veri dosyası](deliverables/Ornek_Ham_Veri.xlsx) bu alanlarla deneme yapmanız içindir.

## Neden seçim tablosunda başlık seçiyorum?

Otomatik başlık listesinin yanındaki hücreleri doğrudan değiştirirsek sütun yerleri değiştiğinde önceki 1/0 seçimleri yanlış alanlara taşınabilir. Bu yüzden girişler ayrı bir **sütun adına bağlı seçim tablosunda** tutulur. Yeni başlıklar otomatik listelenir fakat siz seçene kadar pasiftir. Aynı adın seçimi sütun sırası değişse bile korunur. Başlık adı değişirse yeni adı seçmeniz gerekir.

## Dikkat edilecekler

- Desteklenen üst sınır **2.000 kayıt / 100 sütun**; aşımda analiz durur. Bu bir hız garantisi değildir.
- Dinamik kapsam tablonun içidir; tablo dışındaki veri analize girmez.
- İçteki tamamen boş tablo satırları da kayıt sayılır. Boş seçili değerler birbirleriyle eşleşebilir.
- Başında sıfır bulunan dosya numaralarını kaynakta **metin** tutun. Excel'de önceden kaybolmuş sıfırlar veya basamaklar geri getirilemez.
- Tarihler kaynakta gerçek Excel tarihi olmalıdır; tarih metinleri otomatik dönüştürülmez. Sonuç sayfası tarihi seri sayı olarak gösterebilir.
- Sonuç ve otomatik kriter listesi taşma alanıdır; bu alanlara yazmayın. `#TAŞMA!` görülürse engelleyen hücreleri temizleyin.
- Sonuçların sıralama kontrolü referansı değiştirmez. Ham tabloyu sıralarsanız kaynak sırası değişir ve orijinal kayıt yeniden belirlenir.
- Genel boşluk/harf seçenekleri `02_Kriterler!L7:L8` alanındadır. Boşluk temizleme içteki tekrarlı normal boşlukları da teke indirir.

## VBA kullanmak isterseniz

Dinamik kitaba eski VBA modüllerini kurmayın. Bunun için ayrı [VBA taban kitabını](deliverables/Duplicate_Audit_Studio_VBA_Taban.xlsx) kullanın: `.xlsm` kaydedin, Alt+F11 → Dosya Al ile iki `.bas` dosyasını ekleyin, projeyi derleyin ve Alt+F8 → Kurulum çalıştırın. Bu ayrı akışta kriter değişiminden sonra AnalizEt gerekir.

**Doğrulama sınırı:** Bu ortamda masaüstü Excel yoktur. Dosya/formül kontrolleri tamamlanmış olsa da Excel LTSC 2024 hesaplaması ve kullanıcı akışı çalıştırılmadı. [Kabul listesini](tests/formul/KABUL_KONTROLLERI.md) önce bir kopya üzerinde deneyin. Ayrıntılar [ana README](README.md) ve [doğrulama raporundadır](docs/DOGRULAMA.md).
