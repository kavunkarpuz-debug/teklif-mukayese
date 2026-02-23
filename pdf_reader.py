"""
pdf_reader.py — Teklif Mukayese Yardımcısı
==========================================
Klasördeki tüm PDF ve Excel dosyalarını okur, ham metni ekrana basar.
Claude bu çıktıyı ana konuşmada analiz eder.

Kullanım:
    python pdf_reader.py "C:/Klasor/Yolu"
"""

import sys
import os
import glob

# Windows terminal encoding sorunu: UTF-8 zorla
if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8', errors='replace')

def main():
    if len(sys.argv) < 2:
        print("Kullanim: python pdf_reader.py \"KLASOR_YOLU\"")
        sys.exit(1)

    klasor = sys.argv[1]

    if not os.path.isdir(klasor):
        print(f"HATA: Klasor bulunamadi: {klasor}")
        sys.exit(1)

    # ── PDF'leri oku ────────────────────────────────────────────────────────
    pdf_dosyalari = sorted(glob.glob(os.path.join(klasor, "*.pdf")))
    if pdf_dosyalari:
        try:
            import pdfplumber
        except ImportError:
            print("HATA: pdfplumber yuklu degil. Yuklemek icin: pip install pdfplumber")
            sys.exit(1)

        for f in pdf_dosyalari:
            print(f"\n{'='*70}")
            print(f"PDF: {os.path.basename(f)}")
            print('='*70)
            try:
                with pdfplumber.open(f) as pdf:
                    metin_var = False
                    for i, page in enumerate(pdf.pages):
                        txt = page.extract_text()
                        if txt and txt.strip():
                            metin_var = True
                            print(f"\n--- Sayfa {i+1} ---")
                            print(txt)
                        # Tablo varsa ayrıca yazdır
                        tablolar = page.extract_tables()
                        if tablolar:
                            for j, tablo in enumerate(tablolar):
                                print(f"\n[TABLO {i+1}.{j+1}]")
                                for satir in tablo:
                                    print('\t'.join(str(h or '') for h in satir))
                    if not metin_var:
                        print("UYARI: Metin cikarilamadi — gorsel/taranmis PDF olabilir.")
                        print("       Read tool ile direkt acin (Claude multimodal gorebilir).")
            except Exception as e:
                print(f"HATA: {e}")
    else:
        print("(PDF dosyasi bulunamadi)")

    # ── Excel / XLS dosyalarını oku ─────────────────────────────────────────
    excel_dosyalari = sorted(
        glob.glob(os.path.join(klasor, "*.xlsx")) +
        glob.glob(os.path.join(klasor, "*.xls"))
    )
    if excel_dosyalari:
        try:
            import openpyxl
        except ImportError:
            print("HATA: openpyxl yuklu degil. Yuklemek icin: pip install openpyxl")
            sys.exit(1)

        for f in excel_dosyalari:
            print(f"\n{'='*70}")
            print(f"EXCEL: {os.path.basename(f)}")
            print('='*70)
            try:
                wb = openpyxl.load_workbook(f, data_only=True)
                for sn in wb.sheetnames:
                    ws = wb[sn]
                    print(f"\n-- Sheet: {sn} --")
                    for satir in ws.iter_rows(values_only=True):
                        if any(h is not None for h in satir):
                            print(satir)
            except Exception as e:
                print(f"HATA: {e}")
    else:
        print("(Excel dosyasi bulunamadi)")

    print(f"\n{'='*70}")
    print("OKUMA TAMAMLANDI — Yukardaki ciktiyi analiz edin.")
    print('='*70)

if __name__ == "__main__":
    main()
