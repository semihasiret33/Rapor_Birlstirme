#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Word Rapor Birleştirme Programı
Anketteki sıraya göre Word dosyalarını birleştirir
"""

from docx import Document
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import os
import sys
from pathlib import Path


def add_page_break(document):
    """Belgeye sayfa sonu ekler"""
    document.add_page_break()


def merge_documents(file_list, output_file):
    """
    Word dosyalarını birleştirir

    Args:
        file_list: Birleştirilecek dosya listesi (sıralı)
        output_file: Çıktı dosyasının adı
    """
    # İlk belgeyi temel al
    merged_doc = Document()

    print(f"\n{len(file_list)} dosya birleştiriliyor...\n")

    for idx, file_path in enumerate(file_list, 1):
        if not os.path.exists(file_path):
            print(f"⚠️  Uyarı: {file_path} bulunamadı, atlanıyor...")
            continue

        print(f"[{idx}/{len(file_list)}] {os.path.basename(file_path)} ekleniyor...")

        try:
            doc = Document(file_path)

            # İlk dosya değilse sayfa sonu ekle
            if idx > 1:
                add_page_break(merged_doc)

            # Her paragrafı kopyala
            for element in doc.element.body:
                merged_doc.element.body.append(element)

        except Exception as e:
            print(f"❌ Hata: {file_path} işlenirken sorun oluştu: {e}")
            continue

    # Birleştirilmiş belgeyi kaydet
    merged_doc.save(output_file)
    print(f"\n✅ Başarılı! Birleştirilmiş rapor: {output_file}")


def get_sorted_word_files(directory="."):
    """
    Dizindeki Word dosyalarını sayısal sıraya göre döndürür
    Dosya formatı: 1_rapor.docx, 2_rapor.docx gibi olmalı
    """
    word_files = list(Path(directory).glob("*.docx"))
    # Geçici dosyaları filtrele
    word_files = [f for f in word_files if not f.name.startswith("~$")]

    # Dosya adındaki sayıya göre sırala
    def extract_number(filename):
        try:
            # Dosya adının başındaki sayıyı al
            num_str = filename.stem.split('_')[0].split('-')[0]
            return int(num_str)
        except (ValueError, IndexError):
            return float('inf')  # Sayı yoksa sona koy

    sorted_files = sorted(word_files, key=lambda x: extract_number(x))
    return [str(f) for f in sorted_files]


def main():
    """Ana program"""
    print("=" * 60)
    print("Word Rapor Birleştirme Programı")
    print("=" * 60)

    # Mevcut dizindeki Word dosyalarını bul
    word_files = get_sorted_word_files()

    if not word_files:
        print("\n❌ Hata: Hiç Word dosyası bulunamadı!")
        print("Lütfen .docx uzantılı dosyaları bu dizine ekleyin.")
        print("Dosya isimleri şu formatta olmalı: 1_rapor.docx, 2_rapor.docx, vb.")
        return

    print(f"\n📋 Bulunan dosyalar ({len(word_files)} adet):\n")
    for idx, file in enumerate(word_files, 1):
        print(f"  {idx}. {os.path.basename(file)}")

    # Kullanıcıdan onay al
    print("\n" + "=" * 60)
    response = input("Bu sırayla birleştirmek ister misiniz? (E/H): ").strip().upper()

    if response != 'E':
        print("\nİptal edildi.")
        return

    # Çıktı dosyası adını al
    output_file = input("\nÇıktı dosyası adı (varsayılan: birlesmis_rapor.docx): ").strip()
    if not output_file:
        output_file = "birlesmis_rapor.docx"

    if not output_file.endswith('.docx'):
        output_file += '.docx'

    # Birleştirme işlemini yap
    merge_documents(word_files, output_file)


if __name__ == "__main__":
    main()
