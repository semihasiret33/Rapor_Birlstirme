#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Gelişmiş Word Birleştirme - Resimler ve tüm içerik korunur
"""

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from copy import deepcopy
import os

# Anketteki sıraya göre dosya listesi
DOSYA_SIRASI = [
    "01 Demografik_Bilgiler_Analiz_Raporu.docx",
    "02 Politika_Program_Tasarimi_Analiz_Raporu_Revize (1).docx",
    "03 Kaynak_Altyapi_LDA_Analizi_Detayli.docx",
    "Kaynak_Altyapi_Phi_Korelasyon_Raporu.docx",
    "Personel_Kapasitesi_Egitim_Sorunlari_Raporu.docx",
    "Olcme_Araclari_Veri_Yonetimi_Raporu.docx",
    "Ebeveyn_Toplum_Katilimi_Raporu.docx",
    "Ebeveyn_Toplum_Katilimi_Metin_Analizi_Raporu.docx",
    "Erisim_Esitlik_Kapsayicilik_Raporu.docx",
    "Erisim_Esitlik_Kapsayicilik_Raporu (1).docx",
    "Paydas_Katilimi_Is_Dunyasi_Baglantilari_Raporu.docx",
    "Paydas_Katilimi_Is_Dunyasi_Baglantilari_Raporu (1).docx",
    "LDA_Topic_Analizi_Raporu.docx",
    "Acik_Uclu_Sorular_Analiz_Raporu.docx",
    "Acik_Uclu_Sorular_LDA_Analiz_Raporu.docx",
]


def add_page_break(doc):
    """Sayfa sonu ekle"""
    doc.add_page_break()


def merge_with_images(master, other_doc):
    """
    İki belgeyi birleştirir - resimler ve tüm içerik korunur
    """
    # Relationships'leri kopyala (resimler için gerekli)
    for rel in other_doc.part.rels.values():
        if "image" in rel.reltype:
            # Resmi master belgeye ekle
            image_part = rel.target_part
            new_rel = master.part.relate_to(image_part, rel.reltype)

    # Tüm elementleri kopyala
    for element in other_doc.element.body:
        # Deep copy yaparak orijinali bozmadan kopyala
        new_element = deepcopy(element)

        # rId'leri güncelle (resimler için)
        for desc in new_element.iter():
            # Resim referanslarını kontrol et
            if desc.tag == qn('a:blip'):
                embed_attr = qn('r:embed')
                if embed_attr in desc.attrib:
                    old_rid = desc.attrib[embed_attr]
                    # Eski relationship'i bul
                    if old_rid in other_doc.part.rels:
                        old_rel = other_doc.part.rels[old_rid]
                        # Yeni relationship ekle
                        image_part = old_rel.target_part
                        new_rid = master.part.relate_to(image_part, old_rel.reltype)
                        # rId'yi güncelle
                        desc.attrib[embed_attr] = new_rid

            # Drawing ve picture referanslarını kontrol et
            elif desc.tag in [qn('w:drawing'), qn('pic:pic')]:
                for blip in desc.iter():
                    if blip.tag == qn('a:blip'):
                        embed_attr = qn('r:embed')
                        if embed_attr in blip.attrib:
                            old_rid = blip.attrib[embed_attr]
                            if old_rid in other_doc.part.rels:
                                old_rel = other_doc.part.rels[old_rid]
                                image_part = old_rel.target_part
                                new_rid = master.part.relate_to(image_part, old_rel.reltype)
                                blip.attrib[embed_attr] = new_rid

        master.element.body.append(new_element)


def merge_documents_advanced(file_list, output_file):
    """Dosyaları gelişmiş yöntemle birleştirir"""
    print(f"\n{'='*70}")
    print("Yaşam Boyu Mesleki Gelişim Anketi - Gelişmiş Rapor Birleştirme")
    print(f"{'='*70}\n")
    print(f"Toplam {len(file_list)} dosya birleştiriliyor...\n")

    master_doc = None

    for idx, file_path in enumerate(file_list, 1):
        if not os.path.exists(file_path):
            print(f"⚠️  [{idx}/{len(file_list)}] {file_path} - BULUNAMADI, ATLANIYOR")
            continue

        print(f"✓  [{idx}/{len(file_list)}] {os.path.basename(file_path)}")

        try:
            if master_doc is None:
                # İlk dosyayı master olarak kullan
                master_doc = Document(file_path)
            else:
                # Sayfa sonu ekle
                add_page_break(master_doc)

                # Diğer dosyayı ekle
                other_doc = Document(file_path)
                merge_with_images(master_doc, other_doc)

        except Exception as e:
            print(f"❌ HATA: {file_path} - {str(e)}")
            import traceback
            traceback.print_exc()
            continue

    if master_doc:
        master_doc.save(output_file)
        print(f"\n{'='*70}")
        print(f"✅ TAMAMLANDI!")
        print(f"Çıktı dosyası: {output_file}")
        print(f"{'='*70}\n")
    else:
        print("\n❌ Hiç dosya birleştirilemedi!")


if __name__ == "__main__":
    output_filename = "Yasam_Boyu_Mesleki_Gelisim_Anketi_Tam_Rapor.docx"
    merge_documents_advanced(DOSYA_SIRASI, output_filename)
