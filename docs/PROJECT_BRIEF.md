# Proje Özeti — AI Procurement Intelligence

> Durum: Taslak (v0.3) · Kapsam: İlk sürüm (MVP) · Dil: Türkçe arayüz, Türkçe dokümantasyon

---

## 1. Amaç

Bursa'daki imalat ve sanayi şirketlerinde satın alma ekiplerinin, tedarikçilerden gelen teklifleri
**tek ve karşılaştırılabilir bir maliyet tabanında** değerlendirmesini sağlayan bir masaüstü/web
uygulaması geliştirmek.

Temel problem: teklifler farklı birimlerde, farklı ambalaj ve teslim koşullarında, farklı para
birimlerinde ve eksik bilgilerle geliyor. Satın alma uzmanı bugün bu karşılaştırmayı çoğu zaman
Excel'de, kişiye bağımlı ve tekrarlanamaz biçimde yapıyor. Uygulamanın hedefi, karşılaştırmayı
standartlaştırmak, hataları görünür kılmak ve kararın gerekçesini belgelenebilir hale getirmek.

**Başarı ölçütü:** bir teklif setinin yüklenmesinden karşılaştırma raporuna kadar geçen süreyi
kısaltmak ve aynı veriyle herkesin aynı sonucu almasını sağlamak.

---

## 2. Kapsam

### 2.1 İlk sürümde olanlar

| # | İşlev | Not |
|---|---|---|
| 1 | Standart Excel teklif şablonu yükleme | Şablon uygulama tarafından sağlanır, sürüm numarası taşır |
| 2 | Birden fazla tedarikçi teklifinin karşılaştırılması | Aynı şablon yapısı, tedarikçi başına bir teklif |
| 3 | Veri doğrulama ve eksik bilgi uyarıları | Hatalı/eksik alanlar karşılaştırmaya girmeden önce işaretlenir |
| 4 | Ağırlık ve nakliye maliyetini içeren toplam maliyet hesabı | Hesap tamamen Python tarafında, deterministik |
| 5 | Fiyat, termin ve ödeme koşullarının yan yana gösterimi | Keyfî ağırlıklı puan veya otomatik kazanan önerisi yok |
| 6 | İsteğe bağlı AI destekli değerlendirme | Yalnızca doğrulanmış sonuçların yorumu |
| 7 | Excel (XLSX) ve PDF rapor çıktısı | Karar gerekçesini içerir |

### 2.2 İlk sürümde olmayanlar

ERP/SAP entegrasyonu, tedarikçi portalı, otomatik fiyat tahmini, lojistik optimizasyonu, gelişmiş
teknik özellik yönetimi, kalıcı teklif geçmişi (veritabanı), çoklu para birimi dönüşümü, kullanıcı
yönetimi/yetkilendirme. Bu başlıklar arayüzde "yakında" olarak gösterilebilir; arkasında işlev
bulunmaz ve MVP kapsamına alınmaz.

---

## 3. Ürün Kategorisi

İlk sürümde ürün **kesinleşmiştir**; tek ürün ve tek para birimi ile çalışılır.

| Alan | Değer |
|---|---|
| Ürün | Galvanizli kare çelik kutu profil |
| Kesit | 40 × 40 mm |
| Et kalınlığı | 2 mm |
| Profil uzunluğu | 6 m (sabit) |
| Sipariş miktarı | Adet (kullanıcı girer) |
| Ağırlık | kg/m üzerinden değerlendirilir (tedarikçi bildirir) |
| Referans teorik ağırlık | ≈ 2,39 kg/m · tolerans ±%5 (2,27–2,51 kg/m) — yalnızca kontrol değeri |
| Birim fiyat | TL/kg |
| Para birimi | TL (tek para birimi) |

Bu değerler şablonun sabit alanlarıdır. Şablonda farklı kesit, farklı et kalınlığı veya farklı
uzunluk girilirse doğrulama hatası üretilir (bkz. Bölüm 5 ve 6); farklı ürün özellikleri
karşılaştırılmaz.

> **Not:** Ürün sabitlendiği için kategori seçimi açık kararı kapanmıştır. Ürün mimarisi
> kategoriye bağımlı değildir; ileride yeni bir ürün eklenmesi yeni bir şablon/doğrulama seti
> gerektirir, yeni özellik değildir.

> **Referans ağırlık uyarısı:** 2,39 kg/m değeri **teorik bir kontrol değeridir**, üretim
> gerçeği değildir. Gerçek üretim ağırlığı; profil geometrisi (köşe yuvarlatma yarıçapı, kaynak
> dikişi), haddelenmiş sac kalınlık toleransları ve galvaniz kaplama kalınlığı nedeniyle
> değişebilir. Bu nedenle referans değer yalnızca olağan dışı beyanları yakalamak için kullanılır;
> hesaplamalarda **tedarikçinin beyan ettiği kg/m** esas alınır.

---

## 4. Veri Modeli (kavramsal)

Tek bir ürün talebi (RFQ) etrafında toplanan yapı:

**Talep (RFQ)**
- Ürün: 40 × 40 × 2 mm galvanizli kare çelik kutu profil (sabit)
- Profil uzunluğu: 6 m (sabit)
- Sipariş miktarı (adet)
- Teslim yeri
- Talep tarihi, teklif geçerlilik son tarihi

**Tedarikçi**
- Firma adı, ülke/şehir
- İletişim ve teklif referans numarası
- Termin taahhüdü ve geçmiş performans notu (MVP'de elle girilen tek alan)

**Teklif (tedarikçi başına bir teklif)**
- Kesit, et kalınlığı, uzunluk (doğrulama için okunur; sabit değerle eşleşmelidir)
- Bildirilen ağırlık: **kg/m**
- Birim fiyat: **TL/kg**
- Nakliye bedeli (TL) ve teslim şekli (teslim yeri dahil/hariç)
- Teslimat süresi (termin, gün) ve teslim tarihi
- Ödeme koşulu — **orijinal metin korunur** (ör. "Peşin", "Net 60", "%50 peşin, %50 teslimatta")
- Vade günü — ayrı sayısal alan (peşin = 0; net vadede gün sayısı; karma koşullarda boş)
- Teklif geçerlilik tarihi, notlar

**Hesaplanan alanlar (salt okunur, kullanıcı değiştiremez)**
- Toplam uzunluk (m)
- Toplam ağırlık (kg)
- Ürün bedeli (TL)
- Nakliye bedeli (TL) — teklifteki tutar veya kullanıcı varsayımı
- Kg başına nakliye maliyeti (TL/kg)
- Toplam tedarik maliyeti (TL)

Hesaplanan alanların hiçbiri Excel'den okunmaz; her zaman yeniden üretilir. Böylece yüklenen
dosyadaki eski veya elle değiştirilmiş sonuçlar karara sızamaz.

**Vade günü türetme kuralı**

| Ödeme koşulu metni | Vade günü alanı |
|---|---|
| Peşin / peşin ödeme | 0 |
| Net 30, Net 45, Net 60, "60 gün vadeli" | 30 / 45 / 60 / 60 |
| Karma (ör. %50 peşin, %50 teslimatta) | **boş bırakılır**, tek sayıya indirgenmez |

Orijinal metin her zaman korunur ve olduğu gibi gösterilir. Vade günü, metnin yerine geçmez;
yanında ayrı bir bilgi alanıdır. Metinden güvenilir bir gün değeri çıkarılamıyorsa alan boş kalır
ve arayüzde "karma / belirsiz" olarak işaretlenir.

---

## 5. Temel İş Akışı

1. **Şablon indirme** — kullanıcı güncel Excel şablonunu indirir.
2. **Doldurma** — tedarikçi teklifleri şablona işlenir (tedarikçi başına bir teklif).
3. **Yükleme** — dosya yüklenir, şablon sürümü ve sütun yapısı kontrol edilir.
4. **Doğrulama** — satır bazında hata/uyarı listesi üretilir.
   - **Hata (bloklayıcı):** karşılaştırma yapılamaz (kesit/et kalınlığı/uzunluk uyuşmazlığı,
     negatif veya sıfır fiyat/miktar/ağırlık, sayı olmayan değer, aynı tedarikçinin mükerrer
     kaydı, farklı para birimi).
   - **Uyarı (bloklamaz):** sonucu etkiler ama devam edilebilir (beyan edilen kg/m değerinin
     2,27–2,51 kg/m aralığının dışında olması, teklif geçerlilik tarihi geçmiş, aykırı değer).
   - **Bilgi talebi:** nakliye tutarı eksikse sıfır kabul edilmez; kullanıcıdan tutar veya açık
     varsayım istenir. Bilgi verilmeden karşılaştırma tamamlanmış sayılmaz.
5. **Düzeltme** — kullanıcı arayüzde düzeltir veya dosyayı yeniden yükler.
6. **Karşılaştırma** — hesaplanan alanlar, termin, ödeme koşulu metni ve vade günü yan yana
   gösterilir; en düşük ürün bedeli ve en düşük toplam maliyet ayrı ayrı işaretlenir.
7. **AI yorumu (opsiyonel)** — kullanıcı isterse, hesaplanmış sonuçlar yorumlanır.
8. **Rapor** — Excel ve/veya PDF çıktısı alınır.

---

## 6. Hesaplama Kuralları

Hesaplamalar **yalnızca Python tarafında** yapılır. Hiçbir maliyet AI tarafından üretilmez veya
tahmin edilmez.

**Girdiler (tedarikçi başına)**

    miktar (adet)        = kullanıcı girişi
    profil uzunluğu      = 6 m (sabit)
    kg/m                 = tedarikçinin bildirdiği ağırlık
    birim fiyat          = TL/kg
    nakliye bedeli       = teklifteki tutar veya kullanıcı varsayımı

**Formüller**

1. Toplam uzunluk (m) = adet × 6
2. Toplam ağırlık (kg) = adet × 6 × kg/m
3. Ürün bedeli (TL) = toplam ağırlık × birim fiyat (TL/kg)
4. Nakliye bedeli (TL) = teklif edilen nakliye tutarı **veya** kullanıcının açıkça girdiği varsayım
5. Toplam tedarik maliyeti (TL) = ürün bedeli + nakliye bedeli
6. Kg başına nakliye maliyeti (TL/kg) = nakliye bedeli ÷ toplam ağırlık

**Ağırlık ve nakliye kuralları**
- Toplam ağırlık, tedarikçinin bildirdiği kg/m üzerinden hesaplanır; kg/m tedarikçiden gelir,
  uygulama tahmin etmez ve beyan edilen değeri düzeltmez.
- **kg/m kontrol aralığı:** referans teorik ağırlık 2,39 kg/m, tolerans ±%5 → **2,27–2,51 kg/m**.
  - 2,27 kg/m altı veya 2,51 kg/m üstü beyanlar **uyarı** üretir; bloklayıcı hata değildir.
  - Uyarıya rağmen hesaplama tedarikçinin beyan ettiği kg/m ile yapılır.
  - Referans değer teorik bir kontrol değeridir; gerçek ağırlık profil geometrisi (köşe
    yuvarlatma, kaynak dikişi), sac kalınlık toleransları ve galvaniz kaplama nedeniyle
    değişebilir.
  - Uyarının gerekçesi arayüzde ve raporda görünür; sessizce yutulmaz.
- Nakliye tutarı teklifte yoksa **sıfır kabul edilmez**; kullanıcıdan tutar veya açık varsayım
  istenir. Girilen varsayım arayüzde ve raporda "varsayım" olarak etiketlenir. Gizli varsayım
  kullanılmaz.
- Toplam ağırlık sıfır veya negatif olamaz (kg/m ve miktar için de aynı kural).

**Karşılaştırma kuralları**
- Tedarikçi başına tek teklif değerlendirilir.
- Gösterilecek alanlar: kg/m, toplam ağırlık, TL/kg, ürün bedeli, nakliye, kg başına nakliye,
  toplam tedarik maliyeti, termin, ödeme koşulu (orijinal metin) ve vade günü.
- **En düşük ürün bedeli** ve **en düşük toplam tedarik maliyeti** ayrı ayrı işaretlenir. Bu iki
  alanın farklı tedarikçileri göstermesi normaldir ve açıkça belirtilir.
- İlk sürümde keyfî ağırlıklı puan, bileşik skor veya otomatik kazanan önerisi **yoktur**.
  Uygulama veriyi ve sıralamayı gösterir; seçimi kullanıcı yapar.
- **Ödeme koşulları maliyet puanına dahil edilmez.** Vade günü yalnızca karşılaştırma bilgisidir;
  toplam tedarik maliyetini değiştirmez, sıralamayı etkilemez.
- Karma ödeme koşulları metin olarak korunur ve tek bir vade gününe indirgenmez; bu satırlarda
  vade günü alanı boş bırakılır.
- Farklı ürün özellikleri veya farklı para birimleri karşılaştırılmaz. Tek ürün (40 × 40 × 2 mm,
  6 m) ve tek para birimi (TL) geçerlidir; aykırı kayıt doğrulama hatasıdır.

---

## 7. AI Kullanımı — Sınırlar

Bu proje için değişmez kural: **AI karar verici değildir.**

- AI yalnızca Python tarafından hesaplanmış, doğrulanmış sonuç tablosunu **yorumlar**.
- AI'a ham Excel içeriği gönderilmez; yalnızca hesaplanmış ve doğrulanmış özet gönderilir.
- AI'ın ürettiği metin; öneri, risk ve dikkat noktaları içerebilir; sayısal sonuç üretemez,
  hesaplama yapmaz, kazanan ilan etmez.
- AI çıktısı raporda ayrı ve etiketli bir bölümde yer alır ("AI Yorumu"), karşılaştırma tablosunun
  veya sıralamanın yerine geçmez.
- AI servisi kullanılamıyorsa uygulama tüm işlevlerini AI olmadan sürdürür.
- API anahtarı kod içinde tutulmaz; ortam değişkeni üzerinden okunur (`DEPLOY_API_KEY`).
- AI'a gönderilen veride tedarikçi adı gibi kimlik bilgilerinin maskelenmesi değerlendirilecek
  (varsayılan: maskeleme açık).

---

## 8. Teknoloji

| Katman | Seçim |
|---|---|
| Dil / çalışma ortamı | Python 3.11+ (mevcut `.venv` ile uyumlu) |
| Arayüz | Streamlit — yerel web uygulaması |
| Veri işleme | Pandas |
| Excel okuma/yazma | OpenPyXL |
| Grafikler | Plotly |
| PDF raporu | ReportLab |
| AI | Deploy üzerinden DeepSeek V4.1 Flash (OpenAI uyumlu API) |

**Mimari ilke:** hesaplama mantığı (saf Python/Pandas fonksiyonları) arayüzden ayrı tutulur.
Streamlit yalnızca giriş/çıkış katmanıdır. Hesaplama katmanı arayüzsüz test edilebilir olmalıdır.

Bağımlılıklar `requirements.txt` ile sabitlenir. MVP'de veritabanı yoktur; oturum verisi
uygulama belleğinde tutulur, kullanıcı raporu indirerek saklar. Kalıcı teklif geçmişi yoktur.

---

## 9. Raporlama

**Excel raporu**
- Özet sayfası: tedarikçi karşılaştırma tablosu (kg/m, toplam ağırlık, TL/kg, ürün bedeli,
  nakliye, kg başına nakliye, toplam tedarik maliyeti, termin, ödeme koşulu metni, vade günü),
  en düşük ürün bedeli ve en düşük toplam maliyet işaretleri, farklar
- Detay sayfası: satır bazında hesaplama dökümü (denetlenebilir)
- Doğrulama sayfası: hata/uyarı listesi ve varsayımlar (kg/m kontrol aralığı dışı uyarıları dahil)

**PDF raporu**
- Başlık, talep bilgisi (ürün, kesit, et kalınlığı, uzunluk, miktar), tarih
- Karşılaştırma tablosu ve grafik (Plotly → görsel)
- Sonuç ve gerekçe, varsayımlar, doğrulama özeti
- Ödeme koşulları (orijinal metin + vade günü) — bilgi amaçlı, maliyet hesabına dahil değil
- Kullanıldıysa ayrı etiketli AI yorumu

Her iki raporda da **hangi varsayımlarla** hesaplandığı ve doğrulama uyarıları yer alır; raporsuz
karar, belgesiz karar sayılır.

---

## 10. MVP Dışı (Yol Haritası)

Arayüzde "gelecek sürüm" olarak gösterilebilir:

- ERP/SAP entegrasyonu
- Tedarikçi portalı (teklif girişi tedarikçi tarafında)
- Otomatik fiyat tahmini (geçmiş veriye dayalı)
- Lojistik optimizasyonu (mod/rota seçimi)
- Gelişmiş teknik özellik yönetimi ve revizyon takibi
- Kalıcı teklif geçmişi (veritabanı, karşılaştırma trendleri)
- Çoklu para birimi dönüşümü (kur tarihi bazlı)

Bu başlıklar için MVP'de veri modeli, arayüz alanı veya bağımlılık eklenmez; yalnızca görsel
bilgilendirme yapılır.

---

## 11. Varsayımlar ve Açık Sorular

**Varsayımlar**
- Kullanıcılar satın alma uzmanı/şef düzeyinde, Excel ile çalışmaya alışkın.
- Teklif başına tedarikçi sayısı tipik olarak 3–8.
- Nakliye bedeli teklifte bulunur; bulunmadığında kullanıcı açık varsayım girer.
- Uygulama tek kullanıcı/tek oturum, yerel makinede çalışır.
- kg/m değeri tedarikçinin beyanıdır; uygulama maliyet hesabı için teorik ağırlık kullanmaz,
  yalnızca uyarı üretmek için 2,39 kg/m referansını kullanır.

**Karara bağlanan konular**
1. Ürün kategorisi: 40 × 40 × 2 mm galvanizli kare çelik kutu profil, 6 m.
2. Nakliye teklifte yoksa: sıfır kabul edilmez, kullanıcıdan bilgi/varsayım istenir.
3. Değerlendirme: keyfî ağırlıklı puan yok; en düşük ürün bedeli ve en düşük toplam maliyet ayrı
   gösterilir.
4. Geçerlilik tarihi geçmiş teklif: karşılaştırmaya dahil edilir, uyarı verilir.
5. AI yorumu: varsayılan kapalı, kullanıcı açar.
6. **kg/m uyarı eşiği (kapandı):** referans teorik ağırlık 2,39 kg/m, tolerans ±%5 → 2,27–2,51
   kg/m. Aralık dışı beyan uyarı üretir; bloklayıcı değildir; hesaplama tedarikçi beyanıyla yapılır.
7. **Ödeme koşulları (kapandı):** orijinal metin korunur, vade günü ayrı alanda gösterilir
   (peşin = 0, net vadede sayısal gün), karma koşullar metin olarak korunur ve tek güne
   indirgenmez. Ödeme koşulları ilk sürümde maliyet puanına dahil edilmez; yalnızca karşılaştırma
   bilgisidir.

**Kalan açık nokta**
Yok. İlk sürüm kapsamındaki kararlar kapatılmıştır; yeni kararlar kapsam değişikliği sayılır.

---

## 12. Sözlük

| Terim | Anlam |
|---|---|
| RFQ | Teklif talebi (Request for Quotation) |
| Kutu profil | Kare/dikdörtgen kesitli, içi boş çelik profil |
| Et kalınlığı | Profil cidar kalınlığı (bu üründe 2 mm) |
| kg/m | Profilin metre başına ağırlığı; tedarikçi bildirir |
| Referans teorik ağırlık | 40 × 40 × 2 mm profil için ≈ 2,39 kg/m; yalnızca kontrol değeri |
| kg/m kontrol aralığı | 2,27–2,51 kg/m (±%5); dışındaki beyanlar uyarı üretir |
| TL/kg | Birim fiyat: kilogram başına Türk lirası |
| Toplam ağırlık | adet × 6 m × kg/m |
| Ürün bedeli | Toplam ağırlık × TL/kg |
| Nakliye bedeli | Teklifteki tutar veya kullanıcı varsayımı |
| Toplam tedarik maliyeti | Ürün bedeli + nakliye bedeli |
| Kg başına nakliye | Nakliye bedeli ÷ toplam ağırlık |
| Termin | Sipariş onayından teslimata kadar geçen süre |
| Vade günü | Ödeme koşulunun sayısal gün karşılığı; peşin = 0, karma koşullarda boş |
| Karma ödeme | Birden fazla ödeme dilimi içeren koşul (ör. %50 peşin, %50 teslimatta); tek güne indirgenmez |
| Teklif geçerliliği | Teklifin bağlayıcı olduğu son tarih |
| Doğrulama | Verinin biçim, tutarlılık ve bütünlük kontrolü |
| AI yorumu | Doğrulanmış sonuçların AI tarafından açıklanması; karar değildir |
