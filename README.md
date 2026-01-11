# Word Rapor Birleştirme Programı

Anketteki sıraya göre Word formatındaki rapor dosyalarını birleştiren Python programı.

## Kurulum

1. Python 3.6 veya üzeri yüklü olmalı
2. Gerekli kütüphaneyi yükleyin:

```bash
pip install -r requirements.txt
```

## Kullanım

### 1. Dosya Hazırlığı

Rapor dosyalarınızı bu dizine ekleyin. Dosya isimleri anketteki sıraya göre numaralandırılmalı:

```
1_rapor.docx
2_rapor.docx
3_rapor.docx
...
```

veya

```
1-anket_sonuclari.docx
2-demografik_bilgiler.docx
3-analiz.docx
...
```

### 2. Programı Çalıştırma

```bash
python rapor_birlestir.py
```

### 3. Adımlar

1. Program otomatik olarak dizindeki `.docx` dosyalarını bulur
2. Dosyaları numaraya göre sıralar ve listeler
3. Onayınızı ister
4. Çıktı dosyası adını sorar (varsayılan: `birlesmis_rapor.docx`)
5. Dosyaları birleştirir ve kaydeder

## Özellikler

- ✅ Otomatik dosya bulma ve sıralama
- ✅ Dosyalar arası sayfa sonu ekleme
- ✅ Formatları koruma (paragraflar, stiller, tablolar)
- ✅ Türkçe karakter desteği
- ✅ Hata kontrolü

## Önemli Notlar

- Dosya isimleri sayı ile başlamalı (örn: `1_`, `2_`, `3_`)
- Geçici Word dosyaları (`~$` ile başlayanlar) otomatik filtrelenir
- Her dosya arasına sayfa sonu eklenir
- Orijinal dosyalar değiştirilmez

## Sorun Giderme

### "ModuleNotFoundError: No module named 'docx'"
```bash
pip install python-docx
```

### Dosyalar bulunamıyor
- Dosya uzantılarının `.docx` olduğundan emin olun
- Dosya isimlerinin sayı ile başladığından emin olun

### Birleştirme hatası
- Word dosyalarının bozuk olmadığından emin olun
- Dosyaların başka bir program tarafından açık olmadığından emin olun

## Örnek Kullanım

```bash
$ python rapor_birlestir.py

============================================================
Word Rapor Birleştirme Programı
============================================================

📋 Bulunan dosyalar (5 adet):

  1. 1_giris.docx
  2. 2_metodoloji.docx
  3. 3_bulgular.docx
  4. 4_analiz.docx
  5. 5_sonuc.docx

============================================================
Bu sırayla birleştirmek ister misiniz? (E/H): E

Çıktı dosyası adı (varsayılan: birlesmis_rapor.docx): final_rapor.docx

5 dosya birleştiriliyor...

[1/5] 1_giris.docx ekleniyor...
[2/5] 2_metodoloji.docx ekleniyor...
[3/5] 3_bulgular.docx ekleniyor...
[4/5] 4_analiz.docx ekleniyor...
[5/5] 5_sonuc.docx ekleniyor...

✅ Başarılı! Birleştirilmiş rapor: final_rapor.docx
```
