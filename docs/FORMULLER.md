# Dinamik kitabın formül dökümü

Bu döküm dosyada saklanan OOXML formüllerini gösterir. İşlev adları dosya standardına göre İngilizcedir. Excel arayüzü bunları yerelleştirebilir. `ANCHORARRAY`, taşma başvurusunun (`#`) dosya içindeki gösterimidir.

## 02_Kriterler!A7

```excel
=TRANSPOSE(HamVeri[#Headers])
```

## 02_Kriterler!B7

```excel
=_xlfn.XLOOKUP(_xlfn.ANCHORARRAY(A7),KriterSecimleri[Sütun],KriterSecimleri[Etkin],0)
```

## 02_Kriterler!C7

```excel
=_xlfn.XLOOKUP(_xlfn.ANCHORARRAY(A7),KriterSecimleri[Sütun],KriterSecimleri[Tür],"Kesin")
```

## 03_Kontrol!A7

```excel
=_xlfn.HSTACK(HamVeri[#Headers],"Kayıt Durumu","Eşleşme Grubu Kimliği","Eşleşme Adedi","Orijinal Kayıt Referansı","Kaynak Satır Numarası")
```

## 03_Kontrol!A8

```excel
=IF('_Hesaplama'!B2<>"HAZIR",'_Hesaplama'!B2,_xlfn.LET(veri,_xlfn.ANCHORARRAY('_Hesaplama'!DC4),kriter,IF(B3="Tümü",1,_xlfn.ANCHORARRAY('_Hesaplama'!DA4)=B3)*IF(E3="",1,_xlfn.ANCHORARRAY('_Hesaplama'!DB4)=E3)*_xlfn.BYROW(_xlfn.ANCHORARRAY('_Hesaplama'!A4),_xlfn.LAMBDA(_xlpm.satir,IF(H3="",TRUE,SUM(--ISNUMBER(SEARCH(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(H3,"~","~~"),"*","~*"),"?","~?"),_xlpm.satir&"")))>0))),sira,IFERROR(_xlfn.XMATCH(B4,HamVeri[#Headers],0),COLUMNS(HamVeri[#Headers])+5),IF(SUM(--kriter)=0,"Görünüm filtresiyle eşleşen kayıt yok.",_xlfn._xlws.SORT(_xlfn._xlws.FILTER(veri,kriter),sira,IF(E4="Azalan",-1,1)))))
```

## 04_Ozet!B4

```excel
='_Hesaplama'!B2
```

## 04_Ozet!B5

```excel
=IFERROR(ROWS(HamVeri[#Data]),0)
```

## 04_Ozet!B6

```excel
=IF('_Hesaplama'!B2="HAZIR",SUM(--(_xlfn.ANCHORARRAY('_Hesaplama'!DA4)="TEKİL")),0)
```

## 04_Ozet!B7

```excel
=IF('_Hesaplama'!B2="HAZIR",SUM(--(_xlfn.ANCHORARRAY('_Hesaplama'!DA4)="ORİJİNAL")),0)
```

## 04_Ozet!B8

```excel
=IF('_Hesaplama'!B2="HAZIR",SUM(--(_xlfn.ANCHORARRAY('_Hesaplama'!DA4)="MÜKERRER")),0)
```

## 04_Ozet!B9

```excel
=B7
```

## 04_Ozet!B10

```excel
=SUM(KriterSecimleri[Etkin])
```

## 04_Ozet!B11

```excel
=IFERROR(B8/B5,0)
```

## 04_Ozet!B12

```excel
=IF(B4="HAZIR",B5-B6-B7-B8,"Analiz geçersiz")
```

## _Hesaplama!B1

```excel
=IFERROR(_xlfn.LET(baslik,HamVeri[#Headers],secim,KriterSecimleri[Sütun],etkin,KriterSecimleri[Etkin],tur,KriterSecimleri[Tür],adet,IFERROR(ROWS(HamVeri[#Data]),0),IF(OR(adet>2000,COLUMNS(baslik)>100),"SINIR: En fazla 2.000 satır ve 100 sütun.",IF(SUM(--ISERROR(HamVeri[#Data]))>0,"HATA: Ham veride Excel hücre hatası var.",IF(COUNTA(HamVeri[#Data])=0,"HAM VERİ BOŞ",IF(SUM(--ISERROR(etkin))+SUM(--ISERROR(secim))+SUM(--ISERROR(tur))>0,"HATA: Kriter hücrelerini düzeltin.",IF(SUM((etkin<>0)*(etkin<>1))+SUM((etkin=1)*(secim=""))>0,"HATA: Sütun seçin; Etkin yalnızca 0/1 olmalı.",IF(SUM((etkin=1)*ISNA(_xlfn.XMATCH(secim,baslik,0)))>0,"HATA: Etkin bir sütun ham tabloda bulunmuyor.",IF(SUM((secim<>"")*(COUNTIF(secim,SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(secim,"~","~~"),"*","~*"),"?","~?"))>1))>0,"HATA: Aynı sütunu seçim tablosunda iki kez kullanmayın.",IF(SUM((etkin=1)*(tur<>"Kesin")*(tur<>"Metin")*(tur<>"Türkçe tutar"))>0,"HATA: Karşılaştırma türünü seçin.",IF(OR(AND('02_Kriterler'!L7<>0,'02_Kriterler'!L7<>1),AND('02_Kriterler'!L8<>0,'02_Kriterler'!L8<>1)),"HATA: Genel seçenekler 0/1 olmalı.",IF(SUM(etkin)=0,"KRİTER SEÇİN","HAZIR"))))))))))),"HATA: Tablo ve kriter hücrelerini kontrol edin.")
```

## _Hesaplama!B2

```excel
=IF(B1<>"HAZIR",B1,IF(SUM(--(_xlfn.ANCHORARRAY(CW4)="#DÖNÜŞÜM HATASI#"))>0,"HATA: Seçili türle dönüştürülemeyen değer veya 32.767 karakteri aşan anahtar var.","HAZIR"))
```

## _Hesaplama!A4

```excel
=IF($B$1="HAZIR",IF(ISBLANK(HamVeri[#Data]),"",HamVeri[#Data]),"")
```

## _Hesaplama!CW4

```excel
=IF($B$1<>"HAZIR","",IFERROR(_xlfn.LET(secim,_xlfn._xlws.FILTER(KriterSecimleri[Sütun],KriterSecimleri[Etkin]=1),tur,TRANSPOSE(_xlfn._xlws.FILTER(KriterSecimleri[Tür],KriterSecimleri[Etkin]=1)),veri,_xlfn.CHOOSECOLS(_xlfn.ANCHORARRAY(A4),TRANSPOSE(_xlfn.XMATCH(secim,HamVeri[#Headers],0))),_xlfn.BYROW(veri,_xlfn.LAMBDA(_xlpm.satir,_xlfn.LET(bos,_xlpm.satir="",ham,IF(tur="Türkçe tutar",IF(bos,"",IF(ISNUMBER(_xlpm.satir),_xlpm.satir,_xlfn.NUMBERVALUE(_xlpm.satir,",","."))),IF(tur="Metin",_xlpm.satir&"",_xlpm.satir)),metin,IF('02_Kriterler'!$L$7=1,TRIM(ham&""),ham&""),duzgun,IF('02_Kriterler'!$L$8=1,LOWER(SUBSTITUTE(SUBSTITUTE(metin,"I","ı"),"İ","i")),metin),deger,IF(ISNUMBER(ham),TEXT(ham,"0.00000000000000E+00"),duzgun),tip,IF(deger="","E",IF(ISNUMBER(ham),"N",IF(ISLOGICAL(ham),"B","S"))),_xlfn.TEXTJOIN("",FALSE,tip&LEN(deger)&":"&deger&";"))))),"#DÖNÜŞÜM HATASI#"))
```

## _Hesaplama!CX4

```excel
=IF($B$2<>"HAZIR","",_xlfn.BYROW(_xlfn.ANCHORARRAY(CW4),_xlfn.LAMBDA(_xlpm.anahtar,SUM(--EXACT(_xlfn.ANCHORARRAY(CW4),_xlpm.anahtar)))))
```

## _Hesaplama!CY4

```excel
=IF($B$2<>"HAZIR","",_xlfn.BYROW(_xlfn.ANCHORARRAY(CW4),_xlfn.LAMBDA(_xlpm.anahtar,_xlfn.XMATCH(TRUE,EXACT(_xlfn.ANCHORARRAY(CW4),_xlpm.anahtar),0)+1)))
```

## _Hesaplama!CZ4

```excel
=IF($B$2<>"HAZIR","",_xlfn.SEQUENCE(ROWS(_xlfn.ANCHORARRAY(A4)),1,2))
```

## _Hesaplama!DA4

```excel
=IF($B$2<>"HAZIR","",IF(_xlfn.ANCHORARRAY(CX4)=1,"TEKİL",IF(_xlfn.ANCHORARRAY(CY4)=_xlfn.ANCHORARRAY(CZ4),"ORİJİNAL","MÜKERRER")))
```

## _Hesaplama!DB4

```excel
=IF($B$2<>"HAZIR","","GRP-"&TEXT(_xlfn.ANCHORARRAY(CY4),"000000"))
```

## _Hesaplama!DC4

```excel
=IF($B$2<>"HAZIR","",_xlfn.HSTACK(_xlfn.ANCHORARRAY(A4),_xlfn.ANCHORARRAY(DA4),_xlfn.ANCHORARRAY(DB4),_xlfn.ANCHORARRAY(CX4),_xlfn.ANCHORARRAY(CY4),_xlfn.ANCHORARRAY(CZ4)))
```
