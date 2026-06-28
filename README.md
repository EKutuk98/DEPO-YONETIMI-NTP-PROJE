# 🏪 Depo & Stok Yönetim Sistemi

Profesyonel stok takibi, sipariş yönetimi, tedarikçi ilişkileri ve gelişmiş stok analizini kapsayan kapsamlı bir depo yönetim platformu. PyQt5 ile geliştirilmiş, SQLite veritabanı destekli, QPainter tabanlı grafikler ve ABC analizi içermektedir.

**🔐 Giriş Bilgileri:**
- Admin: `admin` / `admin123`
- Personel: `depo` / `depo123`

---

## 📸 Ekran Görüntüleri

<img width="225" height="185" alt="image" src="https://github.com/user-attachments/assets/1ab1231d-8e53-46ef-a481-44238042486c" />

---

## 📋 Özellikler

---

### 📊 Dashboard

- Toplam ürün sayısı, toplam stok adedi, toplam stok değeri KPI kartları
- Aktif sipariş sayısı ve düşük stok uyarısı KPI kartları
- Gerçek zamanlı güncelleme (3 saniyede bir)
- QPainter grafik ile stok durumu özeti

<img width="1279" height="748" alt="image" src="https://github.com/user-attachments/assets/739ec6f4-56e1-4910-a395-516566dc19e9" />


---

### 📦 Ürünler

- Ürün ekleme, düzenleme, silme
- Ürün kodu, ad, kategori, stok adedi, birim fiyat, minimum stok seviyesi
- Stok arttır / Stok azalt butonları (hareket kaydı ile)
- Düşük stok uyarısı (minimum seviyenin altındaki ürünler kırmızı vurgulanır)
- Kategori bazlı filtreleme
- Anlık arama (ürün adı, kodu veya kategoriye göre)

<img width="323" height="445" alt="image" src="https://github.com/user-attachments/assets/db58f66e-915c-477c-9445-c83ecc2cc4db" />


<img width="1277" height="747" alt="image" src="https://github.com/user-attachments/assets/696b74ea-d534-4444-9851-b34baf23f270" />



---

### 🛒 Siparişler

- Sipariş oluşturma (ürün seçimi + adet + müşteri bilgisi)
- Otomatik fiyat hesaplama (adet × birim fiyat)
- Sipariş onaylama (stoktan otomatik düşüş)
- Sipariş teslim etme
- Sipariş iptal (stok iadesi)
- Müşteri adı ve telefon kaydı
- Durum bazlı filtreleme (Beklemede / Hazırlanıyor / Teslim Edildi / İptal)
- Anlık arama

<img width="1279" height="743" alt="image" src="https://github.com/user-attachments/assets/26febe4c-3a32-4022-b642-471c21fe2fec" />

---

### 🏢 Tedarikçiler

- Tedarikçi ekleme, düzenleme, silme
- Ad, iletişim kişisi, telefon, email, adres
- Tedarikçi bazlı ürün listesi
- Anlık arama

<img width="1279" height="742" alt="image" src="https://github.com/user-attachments/assets/cd69e143-784b-4dfe-ada7-b47f723d6731" />

<img width="1277" height="749" alt="image" src="https://github.com/user-attachments/assets/f2b07eda-ce29-420d-86bf-ca1e0037507f" />

---

### 📋 Stok Hareketleri

- Tüm stok giriş/çıkış işlemlerinin tam geçmişi
- Her işlemde önceki stok → işlem miktarı → sonraki stok gösterimi
- İşlem tipi renklendirmesi: Giriş (yeşil) / Çıkış (kırmızı)
- Tarih bazlı sıralama (en son önce)
- Kullanıcı ve açıklama bilgisi

<img width="1279" height="755" alt="image" src="https://github.com/user-attachments/assets/cfcf6f53-2453-4f9e-be1e-f2e2418fd15a" />

---

### 📊 İstatistikler

- Kategori dağılımı bar grafiği (QPainter)
- Stok durumu pasta grafiği: Normal stok vs Düşük stok (QPainter)
- Sistem istatistikleri karşılaştırma grafiği (QPainter)
- İleri İstatistik diyaloğu ile detaylı analizler

<img width="1279" height="742" alt="image" src="https://github.com/user-attachments/assets/3c88f126-f0da-47ee-96d8-5a659e607637" />

<img width="1279" height="593" alt="image" src="https://github.com/user-attachments/assets/87a8d15d-bddb-4b9b-87b3-95d8e5a9d1d7" />


---

### 📈 Grafikler

- Aylık stok giriş/çıkış trendi çizgi grafiği (QPainter)
- Kategori bazlı ciro grafiği
- Ürün bazlı hareket yoğunluğu

<img width="449" height="402" alt="image" src="https://github.com/user-attachments/assets/da71414c-65a9-4d49-a286-958ae32eaa8e" />

---

### 💾 Export

- Excel raporu (Ürünler, Siparişler, Hareketler — çok sayfalı)
- CSV dışa aktarma
- Düşük stok raporu export
<img width="476" height="408" alt="image" src="https://github.com/user-attachments/assets/9ceb4031-64e5-4b9e-b8c7-847269da7261" />


---

### 🔴 Uyarılar (İleri İstatistik)

- Minimum stok seviyesinin altındaki ürünlerin listesi
- Kritik stok uyarısı (sıfıra yakın ürünler)
- Yeniden sipariş önerileri
<img width="417" height="406" alt="image" src="https://github.com/user-attachments/assets/3fe85411-d06c-4175-a623-6203e0d9c904" />


---

### 📊 ABC Analizi

- A sınıfı (yüksek değerli), B sınıfı (orta), C sınıfı (düşük) ürün sınıflandırması
- Ciro bazlı analiz
- Stok optimizasyon önerileri

<img width="434" height="410" alt="image" src="https://github.com/user-attachments/assets/6ba9db66-6e9a-4bbd-8a3c-cefb98e8ed05" />


---

### 📦 EOQ & Öneriler

- Ekonomik Sipariş Miktarı (EOQ) hesaplama
- Ürün bazlı optimal sipariş miktarı önerisi
- Yeniden sipariş noktası hesaplama

<img width="433" height="402" alt="image" src="https://github.com/user-attachments/assets/c7215131-1382-4298-a703-934ffeb38892" />


---

### 🎖️ Gamification

- Personel rozet sistemi (Badge'ler)
- Liderlik tablosu (en çok sipariş işleyen personel)
- Mood Tracker (günlük durum takibi)

<img width="468" height="463" alt="image" src="https://github.com/user-attachments/assets/0c9f7cfb-02ba-4e32-b52f-4a663615468e" />


<img width="481" height="474" alt="image" src="https://github.com/user-attachments/assets/2bb01b75-f8b6-44eb-a56c-0a287dd32ccb" />


<img width="449" height="470" alt="image" src="https://github.com/user-attachments/assets/4084817f-f989-444c-a21c-ebdc0cca0bc2" />


---


- Serbest metin not alma
- Ürün veya sipariş bazlı not ekleme


<img width="1277" height="749" alt="image" src="https://github.com/user-attachments/assets/b69fef92-240c-4e63-bdc6-012c32d89760" />


---

## ❓ Final Soruları

### Sistemde hangi kullanıcılar veya nesneler vardır?

**Kullanıcı Türleri:**
- **Admin** — Tam yetkili; kullanıcı yönetimi, sistem ayarları, tüm modüller
- **Personel** — Ürün, sipariş, stok ve tedarikçi işlemleri

**Ana Nesneler / Varlıklar:**
- **Ürün** — urun_id, ürün kodu, ad, kategori, stok, minimum stok, birim fiyat, tedarikçi, durum
- **Sipariş** — siparis_id, ürün, adet, birim fiyat, toplam, müşteri adı, telefon, durum, tarih
- **Tedarikçi** — tedarikci_id, ad, iletişim kişisi, telefon, email, adres
- **Stok Hareketi** — hareket_id, ürün, tür (Giriş/Çıkış), miktar, önceki stok, sonraki stok, açıklama, kullanıcı, tarih
- **Kategori** — kategori_id, ad, açıklama
- **Not** — not_id, başlık, içerik, ilgili tür, ilgili id, tarih
- **Badge (Rozet)** — badge_id, ad, açıklama, kriter
- **Kullanıcı** — kullanici_id, ad, soyad, kullanıcı adı, SHA256 parola, rol

---

### Kullanıcı sistemde hangi işlemleri gerçekleştirebilir?

**Admin:**
- Kullanıcı ekleyebilir, pasife alabilir
- Sistem ayarlarını güncelleyebilir, DB yedeği alabilir
- Aşağıdaki tüm personel işlemlerini yapabilir

**Personel:**
- Ürün ekleyebilir, düzenleyebilir, silebilir
- Stok arttırabilir ve azaltabilir (hareket kaydı ile)
- Kategori ve tedarikçi ekleyebilir, düzenleyebilir
- Sipariş oluşturabilir, onaylayabilir, teslim edebilir, iptal edebilir
- Stok hareketlerini görüntüleyebilir
- İstatistik ve grafikleri görüntüleyebilir
- ABC analizi ve EOQ önerilerini inceleyebilir
- Düşük stok ve kritik uyarıları görüntüleyebilir
- Excel ve CSV raporu export edebilir
- Not ekleyebilir ve görüntüleyebilir

---

## 🖥️ Teknolojiler

| Teknoloji | Kullanım Alanı |
|-----------|----------------|
| Python 3.9+ | Ana programlama dili |
| PyQt5 | GUI Framework |
| SQLite3 | Veritabanı yönetimi (`@contextmanager`) |
| QPainter | Tüm grafikler (Bar, Pie, Line) |
| hashlib (SHA256) | Şifre güvenliği |
| openpyxl / csv | Rapor export |
