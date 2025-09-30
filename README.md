# Delta Görev Formu Sistemi

Delta Proje ekibinin saha görevlerini planlayıp kaydetmesini sağlayan modern bir masaüstü uygulaması. Tkinter ile geliştirilen arayüz, görev formunu adım adım doldurmanıza yardımcı olur ve verileri şirket formatına uygun Excel dosyalarına aktarır.

## Son Güncellemeler
- **Çok adımlı sihirbaz:** Form bilgileri, personel listesi, avans ve taşeron detayları, görev tanımı, görev yeri, zaman çizelgesi, araç ve hazırlayan adımları tek tek takip ediliyor. Her adımda veriler güncellenip sonraki adıma aktarılıyor.
- **Zengin özet ekranı:** Kaydedilmeden önce tüm veriler, görev durumu ile birlikte renkli ve yazdırmaya hazır bir tabloda gösteriliyor.
- **Durum kontrolü:** Gidiş/dönüş ve çalışma saatleri tamamlanmadan form "YARIM" olarak işaretleniyor; eksiksiz doldurulduğunda "TAMAMLANDI" durumuna geçiyor.
- **Dosya numaralandırması:** `form_config.json` dosyası üzerinden form numarası otomatik artırılıyor. Excel çıktıları `gorev_formu_XXXXX.xlsx` adıyla saklanıyor.
- **Kısmi kaydetme desteği:** Görev sahada devam ederken formlar "Görev Formu Çağır" seçeneğiyle tekrar açılıp tamamlanabiliyor.

## Proje Yapısı
```
core/                # Form verilerini işleyen servis katmanı
  form_service.py    # Excel okuma/yazma, durum hesaplama, numaralandırma
web_app/             # (Gelecek web sürümü için) başlangıç modülü
gorev_formu_app.py   # Tkinter arayüzü ve çok adımlı form sihirbazı
ss                   # Test amaçlı dosya (örnek içerik)
tests/               # Pytest senaryoları
```

## Kurulum
1. Python 3.10 veya üzeri bir sürüm kurulu olduğundan emin olun.
2. Gerekli paketleri yükleyin:
   ```bash
   pip install -U tkcalendar openpyxl pytest
   ```

## Çalıştırma
```bash
python gorev_formu_app.py
```

Uygulama ilk açılışta ana menüyü gösterir:
1. **Yeni Görev Oluştur:** Yeni bir form numarası üretir ve adım adım doldurmanızı sağlar.
2. **Görev Formu Çağır:** Daha önce kaydettiğiniz bir Excel dosyasını seçip düzenlemeye devam edebilirsiniz.

Form kaydedildiğinde Excel dosyası mevcut klasöre kaydedilir. `form_config.json` dosyasını silmek, numaralandırmayı sıfırlar (bir dahaki kaydetmede yeniden oluşturulur).

## Testler
Pytest senaryoları servis katmanının tamamlanma durumunu, numaralandırmayı ve Excel kaydını doğrular. Testleri çalıştırmak için:
```bash
pytest
```

## İpuçları
- "Kaydet" butonu her adımda mevcut verileri saklar; özet ekranında **Kaydet** derseniz formu tamamlanmış olarak kaydedersiniz.
- Tarih alanlarında takvim, saat alanlarında ise saat/dakika seçim kutuları bulunur; boş bırakılan kritik alanlar form durumunu "YARIM" yapar.
- Excel dosyaları başka bir klasöre taşınacaksa uygulamayı o klasörde çalıştırmanız yeterlidir.
