# Kırlı Oto Kurtarma — Nakliye Takip Sistemi Kurulum Rehberi

Bu paket, Claude ile geliştirdiğimiz prototipin **gerçek, yayına alınabilir**
halidir. Aşağıdaki adımları sırayla takip edin (yaklaşık 20-30 dakika sürer).
Teknik bilginiz sınırlıysa bu dosyayı bir geliştiriciye de gönderebilirsiniz.

## 1) Supabase hesabı açın (veritabanı)

1. https://supabase.com adresine gidin, ücretsiz hesap açın.
2. "New Project" ile yeni bir proje oluşturun (bölge olarak Frankfurt/EU
   seçmeniz Türkiye'ye en yakın ve hızlı seçenektir).
3. Proje oluşunca sol menüden **SQL Editor**'e girin, bu paketteki
   `supabase-schema.sql` dosyasının içeriğini yapıştırıp **Run** deyin.
4. Sol menüden **Project Settings > API**'ye girin. Şu iki değeri
   not edin:
   - **Project URL**
   - **anon public key**

## 2) Ortam değişkenlerini ayarlayın

1. Bu paketteki `.env.example` dosyasını `.env` olarak kopyalayın.
2. İçindeki iki değeri, 1. adımda aldığınız Project URL ve anon key ile
   değiştirin.

## 3) Kodu GitHub'a yükleyin

1. https://github.com üzerinde ücretsiz hesap açın (yoksa).
2. Yeni bir repo oluşturun (örn. `nakliye-takip`).
3. Bu paketteki tüm dosyaları (`.env` HARİÇ — onu yüklemeyin, gizli
   kalmalı) o repoya yükleyin.

## 4) Vercel'e yayına alın (hosting)

1. https://vercel.com adresine gidin, GitHub hesabınızla giriş yapın
   (ücretsiz).
2. "New Project" ile az önce oluşturduğunuz GitHub reposunu seçin.
3. "Environment Variables" bölümüne, `.env` dosyanızdaki iki değeri
   (VITE_SUPABASE_URL ve VITE_SUPABASE_ANON_KEY) tek tek girin.
4. **Deploy** butonuna basın. 1-2 dakika içinde bir
   `xxxx.vercel.app` adresi verecek — bu adresten uygulamanız zaten
   çalışır durumda olacak.

## 5) Kendi alan adınıza bağlayın

1. Vercel projenizde **Settings > Domains** kısmına girin.
2. `calisan.kirliotokurtarma.com` gibi bir alt alan adı yazıp ekleyin
   (istediğiniz ismi seçebilirsiniz — örn. `takip.kirliotokurtarma.com`).
3. Vercel size bir **CNAME kaydı** verecek (örn. `cname.vercel-dns.com`).
4. Alan adınızı yönettiğiniz yere gidin (domain'i nereden aldıysanız —
   Natro, isimtescil, GoDaddy vb. ya da WordPress barındırıcınızın DNS
   paneli) ve bu CNAME kaydını ekleyin.
5. DNS değişikliği genelde 10 dakika - birkaç saat içinde aktif olur.

## 6) Çalışan hesaplarını oluşturun (Kullanıcı Adı + Şifre Girişi)

Bu sistemde her çalışan **kendi e-posta ve şifresiyle** giriş yapar. Hesapları
siz (yönetici) Supabase panelinden oluşturuyorsunuz — çalışanlar kendi
kendine kayıt olamaz, bu güvenlik açısından doğru olan yöntem.

1. Supabase panelinde sol menüden **Authentication > Users** kısmına gidin.
2. **Add user** (veya "Invite user") butonuna tıklayın.
3. Çalışanın e-postasını yazın (gerçek bir e-posta olması şart değil, örn.
   `ahmet@kirliotokurtarma.com` gibi kurumsal bir format kullanabilirsiniz —
   önemli olan benzersiz olması) ve bir şifre belirleyin.
4. **"Auto Confirm User"** seçeneğini işaretli bırakın (e-posta doğrulaması
   istemesin).
5. Kullanıcıyı oluşturduktan sonra, listede o kullanıcıya tıklayın.
   **"User Metadata"** (veya "Raw User Meta Data") alanına şunu yapıştırın
   (çalışan için):
   ```json
   { "role": "calisan", "name": "Ahmet Yılmaz" }
   ```
   Yönetici hesabı için:
   ```json
   { "role": "yonetici", "name": "Mehmet Kırlı" }
   ```
   **Önemli:** Buradaki `"name"` değeri, uygulamanın "Çalışanlar" listesinde
   göstereceği isimle birebir aynı olmalı — nakliye atama, belge, izin gibi
   her şey bu isme göre eşleşiyor.
6. Her çalışan ve yönetici için bu adımları tekrarlayın.
7. Ayrıca uygulamanın kendi içindeki **"Çalışanlar"** sekmesinden de aynı
   ismi ekleyin (bu, dropdown listelerinde ve filtrelerde görünmesi için
   ayrı bir adımdır — Supabase hesabı ile uygulamanın kendi çalışan
   listesi şu an için birbirinden bağımsız iki kayıt).

Artık çalışanlar `calisan.kirliotokurtarma.com` adresine girip kendi
e-posta/şifreleriyle giriş yapabilir.

## 7) Sitenize link ekleyin

WordPress'te menüye veya bir sayfaya
`https://calisan.kirliotokurtarma.com` adresine giden bir buton/link
eklemeniz yeterli — "Çalışan Girişi" gibi.

---

## Önemli notlar

- **Güvenlik:** Artık gerçek e-posta+şifre girişi var (Supabase Auth). Sadece
  giriş yapmış kullanıcılar veriye erişebiliyor (RLS policy bunu zorluyor).
  Hesapları sadece siz (Supabase panelinden) oluşturabiliyorsunuz — bu,
  11-50 kişilik bir ekip için sağlam bir güvenlik seviyesi.
- **Şifre sıfırlama:** Bir çalışan şifresini unutursa, Supabase panelinde
  Authentication > Users'tan o kullanıcıyı bulup "Reset Password" veya
  yeni şifre atayabilirsiniz.
- **Maliyet:** Supabase ve Vercel'in ücretsiz katmanları, sizin
  ölçeğiniz (11-50 kişi, günde 20-30 nakliye) için başlangıçta
  yeterlidir. Belge/fotoğraf birikimi arttıkça (tahminen 4-6 ay sonra)
  Supabase'de ücretli katmana geçmeniz gerekebilir (~$25/ay) — bunu
  daha önce konuşmuştuk.
- Herhangi bir adımda takılırsanız, hata mesajını bana gösterin, birlikte
  çözelim.
