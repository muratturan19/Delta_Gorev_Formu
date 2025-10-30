# Delta Görev Formu Sistemi

Delta Proje ekibinin saha görevlerini planlayıp kaydetmesini sağlayan Flask tabanlı bir web uygulaması.
Çok adımlı form sihirbazı sayesinde görev bilgilerini, personel listelerini, taşeron ve maliyet
bilgilerini düzenleyip şirket formatına uygun Excel dosyalarına aktarabilirsiniz.

## Öne Çıkan Özellikler
- **Web tabanlı arayüz:** Tüm form adımlarını tarayıcı üzerinden takip edin, dilediğiniz adımda
geri dönüp verileri güncelleyin.
- **Dinamik özet ekranı:** Kaydetmeden önce formun son durumunu ve eksik alanları tek ekranda kontrol edin.
- **Durum takibi:** Zaman bilgileri tamamlanmamış formlar "YARIM" olarak işaretlenir; bütün alanlar
başarıyla doldurulduğunda durum "TAMAMLANDI" olur.
- **Excel çıktısı:** Her kayıt `gorev_formu_XXXXX.xlsx` formatında saklanır, numaralandırma otomatik yönetilir.

## Proje Yapısı
```
core/                # Form verilerini işleyen servis katmanı (Excel, durum hesaplama, numaralandırma)
web_app/             # Flask uygulaması, rotalar ve Jinja şablonları
  static/            # Stil dosyaları, görseller
  templates/         # Çok adımlı form ve özet ekranı şablonları
gorev_formu_app.py   # Tkinter tabanlı eski istemci (artık varsayılan değil)
ss/                  # Örnek Excel dosyası
requirements.txt     # Uygulamanın bağımlılık listesi
tests/               # Pytest senaryoları
```

## Kurulum
1. Python 3.10 veya üzeri bir sürüm kurulu olduğundan emin olun.
2. (Önerilir) Sanal ortam oluşturup etkinleştirin.
3. Bağımlılıkları yükleyin:
   ```bash
   pip install -r requirements.txt
   ```

## Web Uygulamasını Çalıştırma
1. Proje klasöründe aşağıdaki komutla geliştirme sunucusunu başlatın:
   ```bash
   flask --app web_app run --debug --host=0.0.0.0
   ```
   Alternatif olarak `python -m flask --app web_app run --host=0.0.0.0`
   kullanabilirsiniz.
2. Tarayıcınızda [http://localhost:5000](http://localhost:5000) adresine gidin.
3. Ana ekrandan yeni form başlatabilir veya mevcut bir form numarasını girerek düzenlemeye devam edebilirsiniz.

> Varsayılan olarak oturum verileri Flask'in yerleşik imzalı çerez mekanizmasıyla yönetilir.
> Kalıcı bir gizli anahtar tanımlamak için `FLASK_SECRET_KEY` ortam değişkenini ayarlayabilirsiniz.

## Kullanıcı Rolleri ve Oturum Akışı

- Uygulama açıldığında kullanıcı seçimi için bir karşılama penceresi görünür. Tüm kullanıcılar
  açılır listeden kendisini seçer; admin ve görev atama yetkilileri için ek olarak şifre girişi
  istenir.
- Roller: `admin`, `atayan` (görev atama yetkilisi) ve `calisan`. Admin ve atayan hesaplarının
  şifreleri zorunludur; çalışan hesapları şifresizdir ve yalnızca kendilerine atanan görevlerin
  6. adımını (görev raporu ve harcamalar) düzenleyebilir.
- Yan menü ve ana sayfa kartları rol bazlı olarak şekillenir. Örneğin adminler admin paneli ve
  raporlama bölümlerine erişirken, çalışanlar "Görevlerim" listesine yönlendirilir.
- Çıkış işlemi üst menüdeki "Çıkış Yap" bağlantısıyla yapılır; işlem tamamlandığında karşılama
  penceresi yeniden açılır.


## Form Dosyaları
- Kayıtlı formlar proje kökünde `gorev_formu_XXXXX.xlsx` adıyla oluşur.
- Numara sıralaması `form_config.json` dosyasından takip edilir; dosyayı silmek numaralandırmayı sıfırlar
  (bir sonraki kayıtta otomatik yeniden oluşturulur).

## Testler
Servis katmanının tamamlanma durumunu, numaralandırmayı ve Excel kaydını doğrulamak için pytest
senaryoları mevcuttur. Testleri çalıştırmak için:
```bash
pytest
```

## Render (ve benzeri platformlarda) Yayınlama

Render gibi barındırma platformları uygulamanın `0.0.0.0` adresine bağlanmasını
ve kendilerinin tanımladığı `PORT` ortam değişkenini dinlemesini bekler. Render
dashboard'ında **Start Command** alanına aşağıdaki komutu yazdığınızdan emin olun:

```bash
python -m web_app
```

Bu komut, `web_app` paketindeki yeni komut satırı giriş noktasını kullanarak Flask
uygulamasını doğru host ve port ayarlarıyla başlatır. Geliştirme ortamında da
aynı komutu kullanabilirsiniz; `FLASK_DEBUG=1` gibi bir değişken tanımladığınızda
otomatik olarak debug modu açılır.
