# Ürün Fiyat Hesaplayıcı

Excel dosyasındaki ürün verilerini okuyarak fiyat hesaplamaları yapan Windows uygulaması.

## 🚀 Hızlı Başlangıç

### 1. Kurulum
```bash
# Basit kurulum
setup.bat

# Gelişmiş kurulum (önerilen)
python setup_advanced.py
```

### 2. Programı Çalıştırma
```bash
# Windows'ta
run_app.bat

# Veya Python ile
python main.py
```

### 3. Windows Executable Oluşturma
```bash
python build_advanced.py
```

## 📋 Gereksinimler

- Python 3.7+
- sturmmm.xlsx dosyası (program ile aynı klasörde)
- Windows 10/11

## 🎯 Özellikler

- ✅ Ürün arama (stok kodu ve ürün adına göre)
- ✅ Fiyat hesaplama (KDV hariç)
- ✅ Çoklu fiyat seviyeleri (1.10x, 1.05x, 1.15x, 1.20x)
- ✅ Alış fiyatı hesaplama (kur ve katsayı ile)
- ✅ Tek ürün güncelleme
- ✅ Çoklu ürün güncelleme
- ✅ Excel çıktısı
- ✅ Güzel ve kullanıcı dostu arayüz

## 📁 Dosya Yapısı

```
project/
├── main.py                 # Ana uygulama
├── setup.bat              # Basit kurulum
├── setup_advanced.py      # Gelişmiş kurulum
├── run_app.bat            # Program başlatma
├── build_advanced.py      # Executable oluşturma
├── requirements.txt       # Python bağımlılıkları
├── sturmmm.xlsx          # Excel veri dosyası
└── README.md             # Bu dosya
```

## 🔧 Sorun Giderme

- **Python bulunamıyor**: Python'u yükleyin ve PATH'e ekleyin
- **Excel dosyası bulunamıyor**: sturmmm.xlsx dosyasının program ile aynı klasörde olduğundan emin olun
- **Paket yükleme hatası**: İnternet bağlantınızı kontrol edin

## 📞 Destek

Herhangi bir sorun yaşarsanız, setup_advanced.py dosyasını çalıştırarak detaylı hata mesajlarını görebilirsiniz.