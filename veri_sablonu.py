"""
veri.py — Mukayese Veri Dosyasi
================================
Claude bu sablonu kullanarak proje klasorune veri.py yazar.
excel_generator.py bu dosyayi arguman olarak alir.

Kullanim:
    python excel_generator.py "KLASOR_YOLU/veri.py"

CIKTI yolunda Turkce karakter varsa unicode escape kullan:
    u -> \u00fc   s -> \u015f   g -> \u011f
    Ornek: "C:/Users/.../Masa\u00fcst\u00fc/..."
"""

CIKTI      = r"KLASOR_YOLU\Mukayese_TARIH.xlsx"   # <- DEGISTIR
KUR_TARIHI = "GG.AA.YYYY"                          # <- DEGISTIR (veya "Tum teklifler USD")
RFQ_ADI    = "RFQ Adi"                             # <- DEGISTIR

# Kurlar: 1 birim yabanci para = kac USD
# Tum teklifler USD ise sadece {"USD": 1.0} yeter
KUR = {
    "USD": 1.0,
    # "EUR": 1.08,
    # "TRY": 0.028,
}

# RFQ kalemleri — PDF/Excel'den cikarilan siraya gore
rfq_items = [
    # {"item": 1,  "spec": "Aciklama (P/N varsa dahil et)", "qty": 4,  "unit": "EA"},
    # {"item": 2,  "spec": "Aciklama",                      "qty": 10, "unit": "SET"},
]

# Tedarikçiler
# color sirasi: D9E1F2 -> E2EFDA -> FFF2CC -> FCE4D6 -> EAD1DC
# prices: {rfq_item_no: birim_fiyat_orijinal_para} — teklif verilmeyenler None
suppliers = [
    # {
    #     "name":      "KisaAd",           # sutun basliginda gorununur
    #     "full_name": "Tam Firma Adi",
    #     "color":     "D9E1F2",
    #     "currency":  "USD",
    #     "prices": {
    #         1: 100.00,
    #         2: None,    # N/A — teklif verilmemis
    #     },
    #     "delivery": "6 to 8 Weeks",       # varsayilan teslim suresi (tum kalemler)
    #     # Opsiyonel: kalem bazli farkli teslim sureleri
    #     # "delivery_times": {
    #     #     3: "Stokta / 1-2 Gun",     # kalem 3 icin ozel sure
    #     #     7: "8-10 Hafta",           # kalem 7 icin ozel sure
    #     # },                              # eksik kalemler "delivery"'e duser
    #     "payment":  "100% T/T Advance",
    #     "incoterm": "EXW China",
    #     "location": "China",
    # },
]

# ── Opsiyonel: Anomali Notlari ────────────────────────────────────────────────
# Bu degisken yoksa Excel'de "Anomaliler & Notlar" blogu basılmaz.
# Para birimi farklıysa excel_generator.py otomatik not ekler — buraya yazma.
# Satir basinda sembol: ⚠ anomali/uyari   ℹ bilgi
NOTLAR = [
    # "⚠ FIRMA – Kalem N: X adet teklif verilmis; RFQ miktari (Y adet) esas alindi.",
    # "⚠ FIRMA – Kalem N (aciklama): RFQ disi kalem, toplamdan cikarildi.",
    # "⚠ FIRMA – Ambalaj ucreti $XXX; Grand Total'a dahil edilmistir.",
    # "⚠ FIRMA – Firma toplami $X,XXX; hesaplanan $X,XXX. Hesaplanan deger kullanildi.",
    # "ℹ FIRMA – Fiyatlar XX gun icin gecerlidir.",
    # "ℹ FIRMA – MOQ: N adet; RFQ miktari altinda.",
]

# ── Opsiyonel: Yapay Zeka Analizi ─────────────────────────────────────────────
# Bu degisken yoksa "Yapay Zeka Analizi" blogu basılmaz.
# Her madde ayri satirda • ile gosterilir.
# Icermesi gerekenler: fiyat araliği + %, incoterm farki (EXW maliyeti),
# odeme riski, teslimat suresi karsilastirmasi, N/A analizi, genel tavsiye.
AI_ANALIZ = [
    # "Fiyat araligi: $X,XXX (FirmaA) → $X,XXX (FirmaB), fark %XX.",
    # "EXW teklif veren firmalar nakliye maliyetini icermiyor; nakliye farki goz onunde bulundurulmali.",
    # "Odeme riski: FirmaA pesin odeme istiyor.",
    # "Teslimat: FirmaB en hizli (X is gunu); FirmaC en uzun (X hafta).",
    # "Tavsiye: FirmaX fiyat-sart-sure dengesi acisindan one cikiyor.",
]
