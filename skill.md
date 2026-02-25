---
name: teklif-mukayese
description: Birden fazla tedarikçi teklifini analiz edip renkli, formüllü, tek para birimli Excel mukayese tablosu oluşturur. PDF, Excel, Word formatlarını destekler. Günlük kur ile otomatik dönüşüm yapar.
---

# /teklif-mukayese Skill

Güney Yıldızı Petrol (GYP) için tedarikçi tekliflerini karşılaştırır.

**Script dosyaları (skill klasöründe):**
- `pdf_reader.py` — tüm PDF/Excel'leri okur, çıktısı ham metin
- `excel_generator.py` — Excel üretir; Claude bu dosyaya DOKUNMAZ
- `veri_sablonu.py` — veri dosyası şablonu; Claude bunu referans alır

---

## ⚡ ALTIN KURAL

**SUBAGENT / TASK KULLANMA. Her şey ana conversation'da, doğrudan Bash ile yapılır.**

| Yanlış ❌ | Doğru ✅ |
|-----------|---------|
| Her PDF için Task agent | 1 Bash: `python pdf_reader.py` |
| excel_generator.py'yi kopyala/düzenle | Write ile sadece veri.py yaz |
| 7 subagent → 40 dakika | 4 Bash → 4 dakika |

---

## AKIŞ — 5 ADIM

### Adım 1 — Klasörü tara (Bash)

```bash
ls "KLASOR_YOLU"
```

- RFQ: adında "RFQ", "Talep", "List", "Inquiry" geçen dosya
- Teklifler: geri kalan PDF/Excel/Word dosyaları
- Atla: görsel dosyalar (50KB altı), `desktop.ini`, `.DS_Store`

---

### Adım 2 — Tüm dosyaları oku (1 Bash komutu)

```bash
python3 "C:\Users\tugrademirors\.claude\skills\teklif mukayese\pdf_reader.py" "KLASOR_YOLU"
```

Çıktıyı **ana konuşmada** oku ve analiz et.

**Görsel PDF:** pdfplumber "Metin çıkarılamadı" uyarısı verirse →
Read tool ile PDF'i direkt aç (Claude multimodal görür). Subagent AÇMA.

---

### Adım 3 — Veri çıkar (ana konuşmada)

Script çıktısından şunları çıkar:

**rfq_items** — RFQ dosyasından, sıra korunur:
```python
{"item": 1, "spec": "P/N 619501 Diaphragm...", "qty": 8, "unit": "EA"}
```

**suppliers** — Her tedarikçi için:
```python
{
    "name":      "KısaAd",         # sütun başlığı
    "full_name": "Tam Firma Adı",
    "color":     "D9E1F2",         # renk sırası aşağıda
    "currency":  "USD",
    "prices":    {1: 1530.10, 2: None, ...},  # None = N/A
    "delivery":  "6 to 8 Weeks",        # varsayilan teslim suresi
    # Opsiyonel — kalem bazli farkli teslim suresi varsa:
    # "delivery_times": {3: "Stokta", 7: "8-10 Hafta"},  # eksikler "delivery"'e duser
    "payment":   "100% T/T Advance",
    "incoterm":  "EXW China",
    "location":  "China",
}
```

**Renk sırası:** `D9E1F2` → `E2EFDA` → `FFF2CC` → `FCE4D6` → `EAD1DC`

**Eşleştirme kuralı:**
1. P/N kodu varsa → direkt eşleştir
2. Yoksa → açıklama benzerliği %70+
3. Eşleşme yok veya fiyat `$ -` veya `NO QUOTE` → `None`

**Anomali Tespiti — NOTLAR listesi için aktif ara:**

| Durum | NOTLAR'a ekle |
|-------|---------------|
| Supplier qty ≠ RFQ qty | `"⚠ FIRMA – Kalem N: X adet teklif verilmiş; RFQ miktarı (Y adet) esas alındı."` |
| RFQ dışı ekstra kalem | `"⚠ FIRMA – Kalem N (açıklama): RFQ dışı kalem, toplamdan çıkarıldı."` |
| Ambalaj/belge/nakliye ücreti | `"⚠ FIRMA – Ek ücret: açıklama ($XXX); Grand Total'a dahil edildi/edilmedi."` |
| Firma toplamı ≠ hesaplanan | `"⚠ FIRMA – Firma toplamı $X,XXX; hesaplanan $X,XXX. Hesaplanan değer kullanıldı."` |
| Fiyat geçerlilik şartı | `"ℹ FIRMA – Fiyatlar XX gün için geçerlidir."` |
| MOQ şartı | `"ℹ FIRMA – MOQ: N adet; RFQ miktarı altında."` |
| Para birimi farkı | **Yazma** — excel_generator.py otomatik ekler |

**AI_ANALIZ için yazılacaklar (liste — her madde ayrı satırda • ile gösterilir):**
- Fiyat aralığı ve fark %'si (en ucuz → en pahalı)
- Incoterm farkı: EXW tekliflerin nakliye maliyeti içermediğine dikkat çek
- Ödeme riski: 100% avans talep eden firmalar
- Teslimat süresi karşılaştırması
- N/A analizi: hangi firma kaç kalem atladı
- Genel tavsiye: fiyat–şart–süre dengesi

---

### Adım 4 — Kur verisi (SADECE yabancı para birimi varsa)

Veri çıkardıktan sonra kontrol et:

```
Tüm supplier["currency"] == "USD" ?
  → EVET: KUR = {"USD": 1.0} yaz, kur çekme, devam et
  → HAYIR: WebFetch → https://api.exchangerate-api.com/v4/latest/USD
```

Kur çekilemezse dur, tahmin yapma.

---

### Adım 5 — Excel üret (Write + 1 Bash)

**ESKI YÖNTEM ❌:** excel_generator.py'yi kopyalayıp temp dosyaya yaz (500 satır)

**YENİ YÖNTEM ✅:** Sadece veri.py dosyasını proje klasörüne yaz (~40 satır)

```
1. Write tool → "KLASOR_YOLU/veri.py" (bkz. veri_sablonu.py şablonu)
2. Bash ile çalıştır:
   python "C:\Users\tugrademirors\.claude\skills\teklif mukayese\excel_generator.py" "KLASOR_YOLU/veri.py"
3. Excel otomatik oluşur
4. __pycache__ klasörünü sil — veri.py'yi SILME (proje verisi, yeni teklif gelince güncellenir)
```

**Yeni teklif geldiğinde:** veri.py'yi yeniden yazmak yerine `suppliers` listesine yeni blok ekle → generator'ı tekrar çalıştır → Excel otomatik genişler.

**CIKTI yolu yazarken dikkat:** Python string içinde Türkçe karakter varsa
unicode escape kullan: `ü` → `\u00fc`, `ş` → `\u015f`, `ğ` → `\u011f`
Örnek: `"C:/Users/.../Masa\u00fcst\u00fc/..."`

**veri.py içeriği** (sadece değişkenler, hiç logic yok):
```python
CIKTI      = "C:/Proje/Mukayese_20260222.xlsx"
KUR_TARIHI = "22.02.2026"
RFQ_ADI    = "Safa Clutch Talep 1"
KUR        = {"USD": 1.0}
rfq_items  = [
    {"item": 1, "spec": "...", "qty": 4, "unit": "EA"},
    ...
]
suppliers  = [
    {"name": "ADT", "full_name": "...", "color": "D9E1F2", ...},
    ...
]

# Opsiyonel — yoksa Excel'de bu bloklar basılmaz
NOTLAR = [
    "⚠ FIRMA – Kalem N: açıklama...",
    "ℹ FIRMA – Bilgi notu...",
]
AI_ANALIZ = [
    "Fiyat aralığı: $X,XXX (FirmaA) → $X,XXX (FirmaB), fark %XX.",
    "EXW teklifler nakliye içermiyor; bu fark gözetilmeli.",
    "Ödeme riski: FirmaX 100% avans talep ediyor.",
    "Teslimat: FirmaB en hızlı (X iş günü); FirmaC en uzun (X hafta).",
    "Tavsiye: FirmaY fiyat-şart-süre dengesi açısından öne çıkıyor.",
]
```

---

## ÇIKTI YAPISI

```
Satır 1  │ [GYP + RFQ Adı]  │ Firma1 ─── │ Firma2 ─── │ ...
Satır 2  │ Item│Spec│Qty│Unit│ Birim(USD)│Toplam(USD)│Teslim│ ...
─────────┤─────────────────────────────────────────────────
Satır 3+ │ veri satırları (N/A gri, min yeşil, max kırmızı)
...      │
GT satır │ GRAND TOTAL (USD)  │ ──│──│SUM│──  │ ...  lacivert
GT+1     │ Payment Terms       │ her tedarikçi için    │ ...  koyu gri
GT+2     │ Incoterm            │                       │
GT+3     │ Delivery Location   │                       │
GT+5     │ [dipnot — kur tarihi]
         │  (boş satır)
GT+7     │ ANOMALILER & NOTLAR [turuncu başlık]  (NOTLAR varsa)
GT+8..n  │   ⚠ / ℹ  her not — ayrı satır, sarı arka plan, sınır yok
         │  (boş satır)
GT+n+2   │ YAPAY ZEKA ANALIZI  [mavi başlık]     (AI_ANALIZ varsa)
GT+n+3   │   analiz metni — wrap, açık mavi arka plan, sınır yok
```

**3 Sheet:** Mukayese | Kur Bilgisi | Ham Veri

**Renk referansı:**

| Öğe | Hex |
|-----|-----|
| Başlık | `#1F3864` |
| Min birim fiyat (yeşil) | `#C6EFCE` |
| Max birim fiyat (kırmızı) | `#FFC7CE` |
| N/A | `#D3D3D3` |
| Grand Total / Ticari bilgi | `#1F3864` / `#2F4F4F` |

**Border (otomatik):**

| Çizgi | Kalınlık |
|-------|---------|
| Dış çerçeve | thick |
| Satır 1/2 altı, D sütunu sağı, tedarikçi ayırıcı, GT üstü | medium |
| Veri hücreleri | thin |

---

## ÖZET (işlem sonunda chat'e yaz)

```
İşlem Tamamlandı
────────────────────────────────────────
Tedarikçi    : N
İşlenen kalem: N
N/A kalem    : Firma:sayı | ...
Kur tarihi   : GG.AA.YYYY  (veya "Tüm teklifler USD")

En düşük     : FirmaX — $XXX,XXX
En yüksek    : FirmaY — $XXX,XXX
────────────────────────────────────────
Dosya: Mukayese_TARIH.xlsx
⚠ [varsa uyarılar]
```

---

## ALTERNATİF TEKLİF (Aynı Firmadan İkinci Opsiyon)

Kullanıcı "alternatif teklif", "ikinci opsiyon", "aynı firmadan farklı teklif" derse
veya dosya adında "Alt", "Alternatif", "Alternative", "Option" geçiyorsa → bu akışı uygula.

**Revizeden farkı:** Mevcut sütunu güncellemez — yanına YENİ sütun olarak ekler.

### Kurallar

| Alan | Kural |
|------|-------|
| `name` | `FirmaAdi-Alt` (örn. `ADT-Alt`) |
| `full_name` | `Firma Adı (Alternatif / Quote#)` |
| `color` | Ana firmayla görsel bağ için benzer ton — standart 5 renkten bağımsız seç (örn. ADT mavi `D9E1F2` → Alt için `BDD7EE`) |
| Konum | `suppliers` listesinde orijinal firmanın hemen arkasına ekle |
| NOTLAR | Fiyat farklılıkları, kapsam farkı, teslimat farkı zorunlu olarak belgelenir |

### Renk önerileri (ana renge göre)
| Ana renk | Alternatif tonu |
|----------|-----------------|
| `D9E1F2` (mavi) | `BDD7EE` |
| `E2EFDA` (yeşil) | `C6E0B4` |
| `FFF2CC` (sarı) | `FFE699` |
| `FCE4D6` (turuncu) | `F4B183` |
| `EAD1DC` (pembe) | `D9A6BA` |

### NOTLAR'a zorunlu eklenecekler
```
"ℹ FirmaAdi-Alt – Aynı firmadan alternatif teklif; orijinal sütunun yanında gösterilmiştir."
"⚠ FirmaAdi-Alt – Kalem N: $X (Alt) vs $Y (orijinal), fark %Z. [açıklama / doğrulama gerekli]"
```

---

## REVİZE TEKLİF AKIŞI

Kullanıcı "güncelle", "revize geldi", "teklifi değişti" gibi bir ifade kullanırsa
veya klasörde adında "Rev", "Revize", "Revised", "v2", "güncel", "update" geçen
yeni bir dosya varsa → bu akışı uygula.

### Adımlar

1. **Revize dosyayı oku** — pdf_reader.py ile veya Read tool ile
2. **Tedarikçiyi tespit et** — dosya adı veya içerik → veri.py'deki `name`/`full_name` ile eşleştir
3. **Eşleşme kararı:**

| Durum | Eylem |
|-------|-------|
| veri.py'de aynı firma var | Mevcut bloğu güncelle — YENİ FIRMA EKLEME |
| veri.py'de bu firma yok | Normal akış — yeni supplier olarak ekle |
| Emin değilsen | Kullanıcıya sor: "Bu X firmasının revizyonu mu, yoksa yeni firma mı?" |

4. **veri.py'de değişen alanları güncelle:**
   - `prices`: sadece değişen kalemleri güncelle, değişmeyenlere dokunma
   - `delivery`, `delivery_times`, `payment`, `incoterm`: değiştiyse güncelle
   - `currency`: değiştiyse güncelle

5. **NOTLAR'a revizyon notu ekle** (başa, diğer notların üstüne):
   ```
   "ℹ FİRMA – Teklif GG.AA.YYYY tarihinde revize edildi. "
   "Değişen kalemler: Kalem N: ESKİ → YENİ EUR, Kalem M: ESKİ → YENİ EUR."
   ```

6. **AI_ANALIZ'i güncelle** — revizyon fiyat sıralamasını değiştiriyorsa yeniden yaz

7. **Excel'i yeniden üret** → `__pycache__` sil, veri.py koru

### Revizyon notu örneği
```python
NOTLAR = [
    "ℹ Göver – Teklif 25.02.2026 tarihinde revize edildi. Kalem 3: 75 EUR → 65 EUR, Kalem 5: 122 EUR → 110 EUR.",
    # ... önceki notlar
]
```

---

## ÖZEL DURUMLAR

**Para birimi dönüşümü:** `currency: "EUR"` yaz → script otomatik dönüştürür.

**Ekstra ücretler:** Tedarikçi belge/nakliye ücreti eklemişse Grand Total'dan önce
satır ekle. Arka plan `#FFF2CC`, italik. Grand Total formülüne dahil et.

**Matematik kontrolü:** Tedarikçi kendi toplamını yazmışsa `birim × qty` ile karşılaştır.
Fark varsa özette uyar, tabloda hesaplanan değeri kullan.

**Tedarikçi qty ≠ RFQ qty:** Tedarikçi farklı adet üzerinden teklif vermişse birim fiyatı
doğrula, RFQ miktarını esas al, özette uyar.

---

## SKILL OTOMATİK GÜNCELLEME KURALI

Her execution sonunda şu soruyu sor:

> "Bu mukayesede mevcut skill yapısıyla karşılanamayan, gerçekten yeni bir mimari durum var mıydı?"

### Güncelleme GEREKTİREN durumlar (yeni mimari pattern)
- Mevcut `suppliers` / `rfq_items` yapısıyla ifade edilemeyen yeni bir alan türü
  (örn. miktar bazlı kademeli fiyat, koşullu indirim, çoklu lot)
- pdf_reader.py'nin çözemediği yeni bir input formatı
- excel_generator.py'ye yeni bir davranış eklenmesini gerektiren çıktı ihtiyacı
- Daha önce hiç karşılaşılmamış ve tekrar karşılaşılabilecek bir anomali tipi

### Güncelleme GEREKTİRMEYEN durumlar (normal varyasyon)
- Farklı fiyat, para birimi, firma sayısı, kalem sayısı
- Farklı teslim süreleri, ödeme koşulları, incoterm
- delivery_times, NOTLAR, AI_ANALIZ içeriği — bunlar zaten esnektir
- Firma adı, lokasyon, renk farklılıkları

### Güncelleme Akışı
1. Yeni pattern tespit et
2. **Önce kullanıcıya sor:** "X durumu yeni bir pattern, skill.md'ye şunu eklemek istiyorum: [özet]. Onaylıyor musun?"
3. Onay alınınca → skill.md ve gerekirse veri_sablonu.py güncelle
4. `excel_generator.py`'e dokunmak gerekiyorsa → ayrıca onay al, mevcut formatı bozma
5. GitHub'a push et

---

## GÜVENLİK
- Orijinal dosyalara dokunma
- Kur çekilemezse dur, tahmin yapma
- `$ -` veya boş fiyat veya `NO QUOTE` → `None`

---

## TEKNİK NOT — MergedCell Border Kısıtlaması

openpyxl, non-topleft MergedCell hücrelerine border yazmıyor (kaydetmiyor).
**Çözüm (excel_generator.py'de uygulandı):** Her section'ın son sütunu merge'e dahil
edilmez; regular cell olarak bırakılır, aynı fill rengi verilir. Böylece thick/medium
borderlar tüm tabloda tutarlı görünür.
