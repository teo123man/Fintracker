# Harcamalarım

Banka ekstrelerini ve hesap dökümlerini içe aktaran, harcamaları kategorilere ayıran ve yerel bir panel ile Excel defteri oluşturan Python uygulaması.

Uygulama bilgisayarınızda çalışır. Paneldeki dosya yükleme işlemi dosyaları yerel sunucuya aktarır. İçe aktarma sırasında eksik EUR/USD kurları için TCMB'ye istek yapılabilir; ekstreler bu isteğe eklenmez.

## Özellikler

- Garanti ve QNB kredi kartı PDF ekstreleri.
- Garanti ve Ziraat XLSX hesap dökümleri, QNB PDF hesap dökümleri.
- Dosya özetiyle tekrar içe aktarmayı engelleme ve örtüşen hesap hareketlerini ayıklama.
- İşyeri kuralları ve işlem bazında kategori düzeltmeleri.
- Aylık bütçeler, dönem seçimi, veri kapsamı kontrolleri ve Excel çıktısı.
- Gerçek veri gerektirmeyen yapay demo.

Banka ayrıştırıcıları belirli belge düzenleri için yazılmıştır. Farklı düzenler ve taranmış PDF'ler desteklenmeyebilir; yeni bir biçim için kişisel bilgileri temizlenmiş örnek gerekir. Arayüzün önemli bir bölümü İngilizcedir.

## Kurulum

### Windows: arayüzle başlatma

1. Python 3.12 veya sonraki bir sürümü kurun. Python kurulumunda **Add python.exe to PATH** ve **tcl/tk** bileşenini etkinleştirin.
2. Projeyi indirip ZIP dosyasını bir klasöre çıkarın.
3. **Harcamalarim.bat** dosyasına çift tıklayın.
4. Başlangıç penceresinde **Önce demoyu dene** veya bir veri klasörü seçerek **Kendi verilerimle başla** düğmesine basın.

Başlangıç penceresi uygulama ortamını oluşturur, gerekli bileşenleri kurar ve tarayıcıyı açar. İlk kurulum internet gerektirir; sonraki açılışlarda bileşen listesi değişmedikçe ve kurulum sağlamsa yeniden indirme yapılmaz. Kurulum başarısız olursa hata penceresini kapatıp yeniden deneyebilirsiniz.

Kendi verilerinizle ilk açılışta PDF/XLSX belgelerinizi seçebileceğiniz bir karşılama ekranı gelir. Panel için işlem içeren, desteklenen bir kredi kartı ekstresi gerekir. Dosyalar seçtiğiniz klasörün `Belgeler/` altına kopyalanır; asıl dosyalarınız yerinde kalır. Ayar dosyasını elle düzenlemeniz gerekmez.

Başlangıç penceresini kullanım boyunca açık tutun veya küçültün. **Paneli aç** tarayıcıyı açar; **Durdur ve geri dön** sunucuyu kapatır. Pencereyi kapatırken de sunucunun duracağına ilişkin onay istenir. Arayüz Windows'un yüksek ekran ölçeklerine uyum sağlar.

Dosyaları başka klasörlerden seçebilirsiniz. Yükleme ekranı dosyayı okuyamama, yerel uygulamaya bağlanamama ve eski sekme oturumu durumlarını ayrı mesajlarla belirtir. Bağlantı hatasında başlangıç penceresinden paneli yeniden açın. Dosya okuma hatasında dosyanın bilgisayarınıza indirilmiş olduğundan emin olup yeniden seçin.

Python 3.12 ile doğrulanmıştır. Python kurulu değilse başlatıcı kurulum adresini gösterir; Python'u otomatik kurmaz. Windows başlatıcısı ve isteğe bağlı Excel hesaplatma adımı Windows içindir. Demo ve test akışı Excel gerektirmez.

### Elle kurulum (isteğe bağlı)

Proje klasöründe PowerShell açın:

```powershell
python -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
```

macOS/Linux üzerinde sanal ortamın Python yolu `.venv/bin/python` olur.

## Önce demoyu deneyin

```powershell
.\.venv\Scripts\python.exe app.py --demo
```

İlk açılışta `demo/` altında yapay veriler üretilir. Panel tarayıcıda açılır; varsayılan adres `http://127.0.0.1:8765` olur. Port doluysa 8784'e kadar boş bir port aranır. Durdurmak için terminalde `Ctrl+C` kullanın.

Demoyu sıfırlamak için:

```powershell
.\.venv\Scripts\python.exe demo.py
```

Bu komut `demo/` klasörünü silip yeniden üretir. Bu klasöre gerçek veri koymayın.

## Kendi verilerinizle kullanım

Başlangıç arayüzü klasör seçimini ve kaydetmeyi otomatik yapar. Komut satırından kullanmak isterseniz:

1. `home.example.txt` dosyasını `home.txt` adıyla kopyalayın.
2. İçindeki yolu kod klasörünün dışındaki özel veri klasörünüz olarak ayarlayın. Windows örneği `%USERPROFILE%\Harcamalarim` yolunu kullanır.
3. Banka dosyalarınızı bu özel klasörün `Belgeler/` alt klasörüne koyun.
4. `Harcamalarim.bat` dosyasını açın veya aşağıdaki komutu çalıştırın; panelde içe aktarma düğmesini kullanın.

```powershell
.\.venv\Scripts\python.exe app.py
```

Komut satırında veri klasörü önceliği: `HARCAMA_HOME` ortam değişkeni → `home.txt` → kod klasörü. Başlangıç arayüzü mevcut ayarı gösterir; seçtiğiniz klasörü kaydeder ve o açılış için kullanır. Kendi verileriniz için kod klasörünün dışında bir klasör seçin. Uygulama `.env` dosyalarını otomatik okumaz.

CSV kayıtları özel veri klasöründeki `data/` altında tutulur. `Harcamalarim.html` ve `Harcamalarim.xlsx` bunlardan üretilir. **HTML ve Excel çıktıları da kişisel finansal veri içerir.** Yedeklerinizi özel bir yerde tutun.

İlk ekstre döneminden önce de kullanımda olan kart/hesaplar varsa `coverage_sources.example.txt` dosyasını `coverage_sources.txt` olarak kopyalayıp kaynak adlarını ekleyin. Bu ayar veri kapsamı hesabını etkiler; dosya Git dışında tutulur. Kart rotasyonu analizi şu an kodda tanımlı Garanti/QNB takvimini kullanır (`ROTATION`, `ROTATION_START`); kişisel takviminize göre uyarlayın.

## Komut satırı

Aşağıdaki komutlarda da sanal ortamın Python'unu kullanın:

```text
python harcama.py                       # yeni belgeleri içe aktar ve çıktıları üret
python harcama.py --build --no-recalc   # mevcut CSV'lerden çıktıları yeniden üret
python app.py --demo --no-browser      # tarayıcı açmadan demo sunucusu
python -m unittest discover -v         # kapsam ve başlangıç akışı testleri
```

`--no-recalc` Microsoft Excel üzerinden formül hesaplatmayı atlar. Formüller dosyada kalır; Excel açıldığında yeniden hesaplanır. Gerçek veri olmadan test etmek için `HARCAMA_HOME` değerini `demo` yapın. GitHub Actions Windows ve Linux üzerinde birim testlerini ve demo üretimini çalıştırır; bu kontroller tüm banka biçimlerinin doğruluğunu kapsamaz.

## Dosya yapısı

| Dosya | Görev |
| --- | --- |
| `app.py` | Yerel web sunucusu ve panel işlemleri |
| `launcher.py` | Kurulum, veri klasörü seçimi ve uygulama başlatma penceresi |
| `first_run.html` | İlk ekstreleri ekleme ekranı |
| `harcama.py` | Belge ayrıştırma, kategorilendirme ve Excel üretimi |
| `dashboard.py` | Panel verisi ve HTML üretimi |
| `dashboard_template.html` | Panel arayüzü |
| `demo.py` | Tamamen yapay veri üretimi |
| `test_coverage.py` | Veri kapsamı hesaplama testleri |

## Katkı ve gizlilik

Geliştirmeyi demo üzerinde yapın. Hata bildirimlerine gerçek ekstre, hesap numarası, IBAN, işlem dökümü veya kişisel panel görüntüsü eklemeyin. `.gitignore` yerel ayarları, finansal dosyaları ve üretilen çıktıları dışarıda tutar; farklı adlarla oluşturduğunuz kopyaları ayrıca kontrol edin.

GitHub'a ilk gönderimden önce `git status --short` ve `git diff --cached` ile gönderilecek içeriği inceleyin. `.gitignore` daha önce takip edilmiş dosyaları veya Git geçmişini temizlemez.

## Lisans

Bu proje [MIT lisansı](LICENSE) ile dağıtılır.
