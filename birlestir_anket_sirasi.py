#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Anket kategorilerine göre raporları birleştirir
"""

from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
import os

# Anketteki sıraya göre dosya listesi
DOSYA_SIRASI = [
    # 1. Demografik Bilgiler
    "01 Demografik_Bilgiler_Analiz_Raporu.docx",

    # 2. Politika ve Program Tasarımı
    "02 Politika_Program_Tasarimi_Analiz_Raporu_Revize (1).docx",

    # 3. Kaynaklar ve Altyapı
    "03 Kaynak_Altyapi_LDA_Analizi_Detayli.docx",
    "Kaynak_Altyapi_Phi_Korelasyon_Raporu.docx",

    # 4. Personel Kapasitesi ve Eğitim
    "Personel_Kapasitesi_Egitim_Sorunlari_Raporu.docx",

    # 5. Ölçme Araçları, Veri, Kayıt, İzleme ve Değerlendirme
    "Olcme_Araclari_Veri_Yonetimi_Raporu.docx",

    # 6. Ebeveyn ve Toplum Katılımı
    "Ebeveyn_Toplum_Katilimi_Raporu.docx",
    "Ebeveyn_Toplum_Katilimi_Metin_Analizi_Raporu.docx",

    # 7. Erişim, Eşitlik ve Kapsayıcılık
    "Erisim_Esitlik_Kapsayicilik_Raporu.docx",
    "Erisim_Esitlik_Kapsayicilik_Raporu (1).docx",

    # 8. Paydaş Katılımı, İş Dünyası Bağlantıları
    "Paydas_Katilimi_Is_Dunyasi_Baglantilari_Raporu.docx",
    "Paydas_Katilimi_Is_Dunyasi_Baglantilari_Raporu (1).docx",

    # Ek Analizler
    "LDA_Topic_Analizi_Raporu.docx",
    "Acik_Uclu_Sorular_Analiz_Raporu.docx",
    "Acik_Uclu_Sorular_LDA_Analiz_Raporu.docx",
]


def merge_documents(file_list, output_file):
    """Dosyaları birleştirir"""
    print(f"\n{'='*70}")
    print("Yaşam Boyu Mesleki Gelişim Anketi - Rapor Birleştirme")
    print(f"{'='*70}\n")
    print(f"Toplam {len(file_list)} dosya birleştiriliyor...\n")

    # Yeni boş belge oluştur
    merged_doc = Document()

    for idx, file_path in enumerate(file_list, 1):
        if not os.path.exists(file_path):
            print(f"⚠️  [{idx}/{len(file_list)}] {file_path} - BULUNAMADI, ATLANIYOR")
            continue

        print(f"✓  [{idx}/{len(file_list)}] {os.path.basename(file_path)}")

        try:
            doc = Document(file_path)

            # İlk dosya değilse sayfa sonu ekle
            if idx > 1:
                merged_doc.add_page_break()

            # Tüm içeriği kopyala
            for element in doc.element.body:
                merged_doc.element.body.append(element)

        except Exception as e:
            print(f"❌ HATA: {file_path} - {str(e)}")
            continue

    # Kaydet
    merged_doc.save(output_file)
    print(f"\n{'='*70}")
    print(f"✅ TAMAMLANDI!")
    print(f"Çıktı dosyası: {output_file}")
    print(f"{'='*70}\n")


if __name__ == "__main__":
    output_filename = "Yasam_Boyu_Mesleki_Gelisim_Anketi_Tam_Rapor.docx"
    merge_documents(DOSYA_SIRASI, output_filename)
