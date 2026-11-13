# 📑 XLS Sayfa Ayırıcı (Excel Sheet Splitter)

Çok sayfalı Excel (`.xls`) çalışma kitaplarındaki her sayfayı, orijinal hücre biçimlendirmelerini, renklerini, fontlarını ve kenarlıklarını koruyarak ayrı ayrı `.xls` dosyalarına bölen Windows uyumlu pratik bir araç.

Windows sağ tık menüsüne entegre olur; herhangi bir `.xls` dosyasına sağ tıklayıp tek tıkla sayfaları ayırabilirsiniz!

---

## 🚀 Özellikler

- **Hücre Stilleri ve Renk Koruması:** `xlrd` (`formatting_info=True`) ve `xlutils` sayesinde hücre arka plan renkleri, yazı tipleri, kenarlıklar ve hizalamalar birebir korunur.
- **Windows Sağ Tık Entegrasyonu:** Herhangi bir `.xls` dosyasına sağ tıklayıp **"XLS Sayfalarını Ayır"** diyerek anında çalıştırabilirsiniz.
- **Otomatik Klasörleme:** Ayrılan sayfalar, orijinal dosyanın bulunduğu dizinde `Ayrılmış` adlı bir klasör açılarak oraya kaydedilir (`01_Sayfa1.xls`, `02_Sayfa2.xls`...).
- **Otomatik Klasör Açma:** İşlem bittiğinde ayrılan dosyaların bulunduğu klasör otomatik olarak Windows Gezgini'nde açılır.
- **Esnek Kullanım:**
  - Sağ tık menüsü ile tek tıkla ayırma
  - Komut satırından parametre vererek çalıştırma (`python ayir.py dosya.xls`)
  - Çift tıklama durumunda grafiksel dosya seçme penceresi (GUI File Dialog)

---

## 🛠️ Kurulum ve Sağ Tık Menüsünü Aktif Etme

### 1. Bağımlılıkları Yükleme
```bash
pip install -r requirements.txt
```

### 2. Sağ Tık Menüsüne Ekleme
Proje dizinindeki `kurulum.bat` dosyasına çift tıklayın veya terminalden çalıştırın:
```bash
python kurulum_sag_tik.py
```
> Kurulum tamamlandığında Windows kayıt defterine kullanıcı bazlı (`HKEY_CURRENT_USER`) olarak eklenir, yönetici (Admin) izni gerektirmez.

### 3. Sağ Tık Menüsünden Kaldırma (İsteğe Bağlı)
İleride menüden kaldırmak isterseniz `kaldir.bat` dosyasına çift tıklayın veya terminalden:
```bash
python kaldir_sag_tik.py
```

---

## 💻 Kullanım Yöntemleri

### Yöntem 1: Sağ Tık (En Pratik)
1. Bilgisayarınızdaki herhangi bir `.xls` dosyasına sağ tıklayın.
2. Açılan menüden **"XLS Sayfalarını Ayır"** seçeneğine tıklayın.
3. Konsol penceresinde sayfaların ayırma durumu saniyeler içinde gösterilecek ve `Ayrılmış` klasörü otomatik açılacaktır.

### Yöntem 2: Komut Satırı (Terminal / CLI)
```bash
python ayir.py "C:\Raporlar\YillikVeriler.xls"
```

### Yöntem 3: Çift Tıklama (GUI)
`ayir.py` dosyasına çift tıklayarak çalıştırırsanız, karşınıza bir dosya seçim penceresi gelecektir. Ayırmak istediğiniz `.xls` dosyasını seçmeniz yeterlidir.

---

## 📂 Proje Yapısı

```
xls ayır/
├── ayir.py              # Ana ayırıcı script (Stil korumalı & argüman destekli)
├── kurulum_sag_tik.py   # Windows sağ tık menüsüne ekleme scripti
├── kaldir_sag_tik.py    # Windows sağ tık menüsünden temizleme scripti
├── kurulum.bat          # Tek tıkla sağ tık menüsünü kuran batch dosyası
├── kaldir.bat           # Tek tıkla sağ tık menüsünü kaldıran batch dosyası
├── requirements.txt     # Gerekli Python kütüphaneleri
└── README.md            # Dokümantasyon
```

---

## 🔮 Gelecek Özellikler (Roadmap)

- [ ] `.xlsx` (modern Excel) dosyaları için biçimlendirme korumalı ayırma desteği
- [ ] Belirli sayfaları seçerek ayırma (Filtreleme desteği)
- [ ] Sayfaları doğrudan PDF veya HTML formatına aktarma
- [ ] Çoklu dosya seçerek toplu ayırma (Batch processing)
