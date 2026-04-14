# MEBBİS – Taşıma Servis İşlemleri Otomasyonu

MEBBİS (Milli Eğitim Bakanlığı Bilgi İşlem Sistemi) üzerindeki **Taşıma Servis İşlemleri** ekranında öğrenci–servis eşleştirme işlemini otomatik olarak gerçekleştiren Python betiği.

---

## 📋 Ne Yapar?

- Excel dosyasından şoför, plaka ve öğrenci bilgilerini okur
- MEBBİS'teki servis listesini tarayarak **Şoför Başlangıç / Bitiş Tarihi** kontrolü yapar
- Tarihi bugünün tarihine uyan servislerin düzenleme ekranını açar
- Excel'deki öğrencileri ilgili servise otomatik olarak işaretler ve kaydeder
- İşlem sonunda **TXT ve Excel formatında detaylı rapor** oluşturur

---

## 🗂️ Excel Dosyası Formatı

Program bir Excel dosyası (.xlsx) okur. Başlık satırı olmaksızın **2. satırdan itibaren** veri girilmelidir.

| A Sütunu | B Sütunu | C Sütunu |
|----------|----------|----------|
| Şoför Adı Soyadı | Araç Plaka No | Öğrenci Adı Soyadı |
| İSMAİL ALAOĞLU | 74M2080 | ELANUR DURMUŞOĞLU |
| İSMAİL ALAOĞLU | 74M2080 | MUHAMMET ALİ ALAOĞLU |
| ALİM ATALAY | 74M2081 | AHMET YILMAZ |

> **Not:** Aynı şoför ve plaka için birden fazla öğrenci varsa her öğrenci ayrı satıra yazılır.

---

## ⚙️ Kurulum

### 1. Python Gereksinimleri

Python 3.10 veya üzeri gereklidir.

```bash
pip install selenium openpyxl
```

### 2. ChromeDriver

Selenium'un Chrome'u kontrol edebilmesi için bilgisayarınızdaki Chrome sürümüyle uyumlu **ChromeDriver** kurulu olmalıdır.

- Chrome sürümünüzü öğrenmek için: `chrome://settings/help`
- ChromeDriver indirme: https://googlechromelabs.github.io/chrome-for-testing/

İndirilen `chromedriver.exe` dosyasını Python betiğiyle aynı klasöre veya sistem PATH'ine koyun.

---

## 🚀 Kullanım

```bash
python mebbis_otomasyon.py
```

Program adım adım sizi yönlendirir:

**Adım 1 –** Excel dosyası seçimi (dosya seçim penceresi açılır)

**Adım 2 –** ENTER'a basınca Chrome otomatik olarak açılır ve MEBBİS ana sayfasına gider

**Adım 3 –** Siz manuel olarak MEBBİS'e giriş yapın, ardından:
  - Sol menüden **Taşıma Servis İşlemleri**'ne tıklayın
  - **Sorgula** butonuna basın ve listenin yüklenmesini bekleyin
  - Konsola dönüp **ENTER**'a basın

**Adım 4 –** Otomasyon başlar, tüm işlemleri otomatik tamamlar

---

## 🔄 Çalışma Mantığı

```
Excel okunur
    └── Her servis için öğrenci listesi oluşturulur

Web tablosu taranır (her turda baştan)
    ├── Şoför Başlangıç Tarihi ≤ Bugün ≤ Şoför Bitiş Tarihi?
    │       ✗ → Bu satır atlanır (aynı şoför/plaka aşağıda tekrar olabilir, kontrol edilir)
    │       ✓ → Excel'de eşleşme aranır
    │               ✗ → "Excel eşleşme yok" olarak raporlanır
    │               ✓ → Düzenle butonuna tıklanır
    │                       └── Popup açılır
    │                               ├── Her öğrenci kontrol edilir
    │                               │     okay_d.png (gri)   → Seçilir ✅
    │                               │     okay_e.png (yeşil) → Zaten seçili, atlanır
    │                               │     okay_r.png (kırmızı) → Başka araçta, atlanır
    │                               └── Servisi Kaydet → X ile popup kapatılır

Excel'deki tüm servisler tamamlanınca döngü sona erer
```

---

## 📊 Rapor

Program tamamlandığında Excel dosyasının bulunduğu klasörde otomatik olarak **`Raporlar/`** klasörü oluşturulur. İçine tarih damgalı iki dosya kaydedilir:

```
Raporlar/
├── mebbis_rapor_20260413_1430.txt
└── mebbis_rapor_20260413_1430.xlsx
```

### Excel Raporu – Servis Raporu Sayfası

| Servis | Durum | Açıklama |
|--------|-------|----------|
| İSMAİL ALAOĞLU / 74M2080 | TAMAMLANDI | 3 öğrenci yerleştirildi |
| SADİK DÖNMEZ / 74M2150 | ATLANDI | Tarih aralığı dışında |
| ALİM ATALAY / 74M2081 | EXCEL_ESLESME_YOK | Tarih uygun ancak Excel'de karşılık bulunamadı |

### Excel Raporu – Öğrenci Raporu Sayfası

| Öğrenci Adı | Durum | Servis |
|-------------|-------|--------|
| ELANUR DURMUŞOĞLU | Servise yerleştirildi | İSMAİL ALAOĞLU / 74M2080 |
| AHMET YILMAZ | Yerleştirilmedi | ALİM ATALAY / 74M2081 |

### Durum Kodları

| Kod | Açıklama |
|-----|----------|
| `TAMAMLANDI` | Popup açıldı, öğrenciler seçildi, kaydedildi |
| `ATLANDI` | Şoför tarihi bugünün dışında |
| `EXCEL_ESLESME_YOK` | Web'de tarih uygun ama Excel'de bu servis tanımlı değil |
| `WEB_TABLOSUNDA_YOK` | Excel'de tanımlı ama web tablosunda tarih uygun satır bulunamadı |

---

## ⚠️ Önemli Notlar

- Program Selenium'un **kendi açtığı Chrome penceresinde** çalışır. Giriş işlemini bu pencerede yapın.
- Aynı ad-soyada sahip birden fazla öğrenci varsa program sizi uyarır; bu öğrencileri **manuel olarak kontrol** etmeniz gerekir.
- Sistemde aynı şoföre ait birden fazla dönem kaydı olabilir (plaka/tarih değişikliği). Program tarih kontrolü yaparak yalnızca aktif dönemi işler.
- Kırmızı tik (başka araçta kayıtlı) öğrencilere dokunulmaz, raporda belirtilmez; bu durumun çözümü manuel işlem gerektirir.

---

## 🛠️ Sorun Giderme

| Sorun | Çözüm |
|-------|-------|
| `ChromeDriver` hatası | Chrome sürümünüzle uyumlu ChromeDriver indirin |
| Tablo bulunamadı hatası | Sorgula butonuna basıp liste yüklendikten sonra ENTER'a basın |
| Popup açılmıyor | Sayfanın tamamen yüklendiğinden emin olun |
| Öğrenci seçilemiyor | Öğrenci başka bir araçta kayıtlı olabilir (kırmızı tik) |

---

## 📁 Proje Yapısı

```
proje/
├── mebbis_otomasyon.py   # Ana program
├── ogrenci_listesi.xlsx  # Girdi: şoför–plaka–öğrenci eşleştirme tablosu
└── Raporlar/             # Çıktı: otomatik oluşturulur
    ├── mebbis_rapor_YYYYMMDD_HHMM.txt
    └── mebbis_rapor_YYYYMMDD_HHMM.xlsx
```

---

## 📄 Lisans

Bu proje kişisel/kurumsal kullanım için serbesttir.
