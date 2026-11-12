# Excel Teklif Şablonu — Teknik Spesifikasyon

> Belge sürümü: 1.2 · Şablon sürümü: 1.0 · Bağlı belge: `docs/PROJECT_BRIEF.md` v0.3
> Durum: Uygulama öncesi spesifikasyon. Bu aşamada `.xlsx` dosyası üretilmez, uygulama kodu yazılmaz.

**Belge sürüm geçmişi**

| Belge sürümü | Değişiklik |
|---|---|
| 1.0 | İlk spesifikasyon taslağı (14 sütun; `Payment Days` Excel girişi olarak) |
| 1.1 | KDV hariç fiyat standardı; `Quote Date` sütunu eklendi; `Payment Days` Excel'den kaldırıldı (Python türetir); RFQ hücre korumaları tek tek tanımlandı; sürüm uyumluluğunda varsayılan doldurma yasağı |
| 1.2 | `Notes` alanındaki otomatik KDV kelime taraması kaldırıldı; kelime taramasına bağlı hata kodu katalogdan çıkarıldı; KDV standardı Instructions sayfasında zorunlu kullanıcı beyanı (`B4`, `E-DECL-01`) hâline getirildi |

**Şablon sürümü neden 1.0 kalıyor:** Henüz hiçbir `.xlsx` şablonu yayınlanmadı; 1.0 yalnızca bir belge
taslağıydı. Kullanıcıya dağıtılan **ilk gerçek şablon** bu belgedeki 14 sütunlu yapıdır ve sürümü
`1.0`'dır. Belge revizyonu (1.0 → 1.1 → 1.2) ile şablon sürümü ayrı şeylerdir; belge revizyonu
şablon dosyasının sürüm numarasını kendiliğinden artırmaz. 1.2 revizyonu yalnızca doğrulama
yaklaşımını değiştirir (kelime taraması yerine zorunlu beyan), sütun setine dokunmaz; bu yüzden
şablon sürümü `1.0` kalır.

---

## 1. Amaç ve Kapsam

Bu belge, satın alma kullanıcısının tedarikçi tekliflerini doldurduğu **tek Excel dosyasının**
yapısını tanımlar. Dosya, uygulamanın tek veri giriş kanalıdır.

Temel ilkeler:

- Excel **hesaplama yapmaz**. Dosyada maliyet formülü, toplam satırı, KDV hesabı veya hesaplanmış
  sütun yoktur.
- Kullanıcı yalnızca ham veriyi girer: teklif tarihi, miktar, kg/m, TL/kg, nakliye, termin, ödeme koşulu.
- Tüm hesaplamalar Python tarafında yapılır ve arayüz/raporda üretilir.
- Ürün sabittir: 40 × 40 × 2 mm galvanizli kare çelik kutu profil, 6 m, adet, TL/kg.
- Dosya tek para birimi (TRY) ile çalışır.
- **Tüm fiyatlar KDV hariçtir.** Şablonda KDV oranı, KDV tutarı veya KDV dahil toplam alanı yoktur;
  uygulama KDV hesaplamaz (bkz. Bölüm 2.1).
- KDV hariç fiyat standardı, `Instructions` sayfasındaki **zorunlu kullanıcı beyanı** ile onaylanır
  (bkz. Bölüm 4.1). Beyan olmadan dosya işleme alınmaz.

Kapsam dışı: çoklu ürün, çoklu para birimi, KDV hesaplama, kalıcı geçmiş, ERP alanları, otomatik
fiyat tahmini, otomatik kazanan puanı.

---

## 2. Dosya, Sürüm ve KDV Standardı

### 2.1 KDV standardı (kesin)

| Kural | Tanım |
|---|---|
| Birim fiyat | **KDV hariç** TRY/kg |
| Nakliye maliyeti | **KDV hariç** TRY |
| Şablondaki KDV alanları | Yok — oran, tutar, KDV dahil toplam alanı bulunmaz |
| Uygulamanın KDV hesabı | Yok — hiçbir aşamada KDV hesaplanmaz veya eklenmez |
| Karşılaştırma tabanı | Yalnızca KDV hariç tutarlar |
| Rapor | KDV hariç tutarlar; "tüm tutarlar KDV hariçtir" ibaresi zorunlu |
| Standart nasıl güvenceye alınır? | `Instructions` sayfasındaki **zorunlu kullanıcı beyanı** (bkz. 2.1.1) |
| `Notes` alanı | Serbest metin; **otomatik kelime taraması yapılmaz** |

#### 2.1.1 Zorunlu kullanıcı beyanı

KDV hariç fiyat standardı, dosya işleme alınmadan önce kullanıcının açık beyanıyla onaylanır.
Beyan hücresi `Instructions!B4`'tür (bkz. Bölüm 4.1).

| Öğe | Tanım |
|---|---|
| Hücre | `Instructions!B4` |
| Etiket | `Instructions!A4` = `VAT Declaration` |
| İzinli değerler | `EVET` / `HAYIR` (Excel liste doğrulaması) |
| Varsayılan | `HAYIR` |
| Beyan metni | `Girdiğim tüm birim fiyat ve nakliye tutarları KDV hariçtir. Tedarikçi yalnızca KDV dahil fiyat verdiyse KDV hariç tutarı yazılı olarak teyit ettim ve teyit açıklamasını Notes alanına yazdım.` |
| Zorunlu | Evet |
| Onaylanmadıysa | Dosya işleme alınmaz — `E-DECL-01` |

Beyan, standartla ilgili sorumluluğu **örten bir onay değil, görünür bir taahhüttür**: kullanıcı
beyanı işaretlemeden hiçbir satır karşılaştırılmaz. Beyan `EVET` olsa bile KDV hesaplanmaz,
dönüştürülmez ve doğrulanmaz (bkz. aşağıdaki sınır).

#### 2.1.2 KDV dahil fiyat verilmesi durumu

Tedarikçi yalnızca KDV dahil fiyat verdiyse kullanıcı:

1. Tedarikçiden KDV hariç tutarı **yazılı olarak teyit eder**,
2. Teyit açıklamasını `Notes` alanına **serbest metin** olarak yazar
   (ör. `Teklif 34,20 KDV dahil idi; 28,50 KDV hariç yazılı teyit edildi`),
3. Ancak bundan sonra fiyatı girer.

**`Notes` alanı kuralı.** `Notes` serbest metindir ve **taranmaz**. İçinde `KDV dahil`, `KDVli`,
`incl. VAT`, `VAT included`, `brüt` gibi ifadelerin geçmesi tek başına hata üretmez; bu ifadeler
teyit açıklamasının doğal parçasıdır. Kelime listesine dayalı otomatik hata kodu **yoktur**;
`Notes` içeriği üzerinden KDV ile ilgili hiçbir hata veya uyarı üretilmez. `Notes` ile ilişkili tek
kural `W-SRC-01`'dir; o kural KDV ile ilgili değildir, yalnızca `Customer Assumption` seçilen
nakliye tutarının gerekçesinin yazılmasını ister (bkz. Bölüm 6).

**Karşılaştırma kapısı.** KDV dahil fiyat veren tedarikçi varsa, bu durumun ve KDV hariç tutarın
yazılı teyidinin beyan edilmesi `Instructions!B4` üzerinden yapılır. Beyan `EVET` değilse **hiçbir
teklif** karşılaştırmaya alınmaz (`E-DECL-01`); yani teyit edilmemiş KDV dahil tutarın
karşılaştırmaya sızması dosya düzeyinde engellenir.

**Yanlış giriş nasıl önlenir?** Otomatik tahmin yerine üç görünür kontrol kullanılır:

- **Talimat:** `Instructions` Blok 5, KDV kuralını "zorunlu okuma" olarak yazar; KDV dahil tutarın
  girilmemesi gerektiğini açıkça belirtir.
- **Beyan:** `Instructions!B4` onaylanmadan dosya işleme alınmaz (`E-DECL-01`); kullanıcı standardı
  okuyup kabul ettiğini aktif olarak işaretler.
- **Etiket:** Sütun başlıkları `excl. VAT` ifadesini taşır (`Unit Price (TRY/kg, excl. VAT)`,
  `Shipping Cost (TRY, excl. VAT)`); rapor satırlarında "KDV hariç" ibaresi ve varsa teyit notu
  birlikte gösterilir.

Uygulama, girilen tutarın KDV içerip içermediğini **tahmin etmez**, kelime aramaz ve tutarı
dönüştürmez; bu sınır arayüzde ve raporda açıkça yazılır. Böylece yanlış giriş sessizce geçmez:
beyan, etiket ve rapordaki teyit notu aynı satırda birlikte görünür.

### 2.2 Dosya bilgileri

| Özellik | Değer |
|---|---|
| Dosya biçimi | `.xlsx` (Office Open XML) |
| Reddedilen biçimler | `.xls`, `.xlsm`, `.csv`, `.ods` |
| Maksimum dosya boyutu | 5 MB |
| Sayfa sayısı | 3 (Instructions, RFQ, Supplier Quotes) |
| Şablon sürümü | `1.0` |
| Şablon kodu | `RFQ-KUTUPROFIL-40X40X2` |
| Uyumluluk kuralı | Aynı şablon kodu + aynı major sürüm zorunlu |
| Önerilen dosya adı | `TEKLIF_<RFQNo>_v1.xlsx` (zorunlu değil) |

**Sürüm alanları** Instructions sayfasında, `B2` (sürüm) ve `B3` (şablon kodu) hücrelerindedir.
Beyan alanı `B4`'tür. Uygulama `B4`, `B2` ve `B3` hücrelerini okumadan dosyayı işleme almaz.

### 2.3 Sürüm uyumluluğu — güvenli kurallar

Sürüm numarası `MAJOR.MINOR` biçimindedir. Sütun ekleme/çıkarma, başlık metni veya anlam
değişikliği **major**; yalnızca açıklama/örnek/biçim değişikliği **minor** artışıdır.

| Durum | Davranış | Kod |
|---|---|---|
| Şablon kodu farklı veya boş | Dosya reddedilir | `E-TPL-01` |
| Major sürüm farklı | Dosya reddedilir | `E-TPL-02` |
| Aynı major, zorunlu sütunlardan biri yok | Dosya reddedilir — **varsayılan değer doldurulmaz** | `E-TPL-03` |
| Aynı major, tüm zorunlu sütunlar var, minor farklı | Okunur, bilgilendirme uyarısı | `W-TPL-01` |
| Beklenmeyen/bilinmeyen sütun var | Sütun okunmaz, veri olarak kullanılmaz | `W-HDR-01` |
| İsteğe bağlı alan boş | Boş kabul edilir; alanın kendi kuralı geçerlidir | `W-`/`I-` kodları |

Zorunlu sütunların varlığı **sürüm numarasından bağımsız olarak** ayrıca kontrol edilir; bu yüzden
elle düzenlenmiş veya sütunu silinmiş bir dosya, sürümü doğru yazsa bile reddedilir.

**Kontrol sırası:** (0) `Instructions!B4` beyanı → `E-DECL-01`, (1) şablon kodu → `E-TPL-01`,
(2) major sürüm → `E-TPL-02`, (3) zorunlu sütunların varlığı → `E-TPL-03`, (4) minor sürüm farkı →
`W-TPL-01`. İlk hata dosyayı reddeder; sonraki adımlar çalıştırılmaz ve kullanıcıya tek, açık bir
mesaj gösterilir.

**Değişmez kurallar**

1. **Eksik zorunlu alan asla tahmin edilmez veya varsayılanla doldurulmaz.** Hata gösterilir ve
   satır (gerekirse dosya) işlenmez.
2. **Eksik isteğe bağlı alan boş kalabilir.** Boş alan "0", "N/A" veya tahmini değerle doldurulmaz.
3. **Bilinmeyen sütunlar sessizce veri olarak kullanılmaz.** Yalnızca Bölüm 6'da tanımlı 14
   başlık okunur; diğerleri yok sayılır ve `W-HDR-01` üretilir.
4. Sürüm yükseltmesi hiçbir zaman değer dönüştürmez, birim çevirmez veya alan uydurmaz.

---

## 3. Genel Yapı

### 3.1 Sayfalar

| Sıra | Sayfa adı (tam) | İçerik | Zorunlu |
|---|---|---|---|
| 1 | `Instructions` | Kullanım talimatı + sürüm bilgisi + zorunlu KDV beyanı | Evet |
| 2 | `RFQ` | Talep başlığı (alan–değer) | Evet |
| 3 | `Supplier Quotes` | Tedarikçi teklifleri (satır = tedarikçi) | Evet |

Sayfalar **adıyla** bulunur; sıra değişse de dosya okunur. Sayfa adları değiştirilemez
(büyük/küçük harf duyarsız eşleşir, ama görünen ad korunmalıdır).

Fazladan sayfa varsa okuma sürer, kullanıcıya "fazladan sayfa yok sayıldı" uyarısı verilir
(`W-SHEET-01`).

### 3.2 Yerleşim kuralı

Her iki veri sayfasında da:

- Satır 1: başlık bandı (birleştirilmiş, yalnızca görsel — ayrıştırıcı yok sayar)
- Satır 2: sütun başlıkları (sabit, korumalı)
- Satır 3 ve sonrası: veri

Başlık satırı **2. satırdır** ve hücre değeri birebir eşleşmelidir (bkz. Bölüm 5, 6). Sütun sırası
korunmalıdır; ayrıştırıcı yine de başlık adına göre eşleştirir, ama sıra değişikliği `W-HDR-02`
uyarısı üretir.

### 3.3 Renk ve biçim göstergeleri

| Gösterge | Anlam |
|---|---|
| Gri dolgu | Sabit / değiştirilemez değer |
| Beyaz dolgu | Zorunlu kullanıcı girişi |
| Açık sarı dolgu | İsteğe bağlı giriş |
| Kalın başlık + koyu mavi metin | Sütun başlığı (korumalı) |
| Kenarlık | Veri alanı sınırı |

Veri alanında **birleştirilmiş hücre, gizli satır/sütun, filtre, alt toplam, formül** bulunmaz.
Yalnızca başlık bandında birleştirme kullanılır. KDV'ye ait hiçbir sütun, renk kodu veya ayrı alan
yoktur; KDV bilgisi başlık metnindeki `excl. VAT` ifadesi ve `Instructions!B4` beyanıyla taşınır.

### 3.4 Sayı ve tarih biçimi

| Tür | Hücre biçimi | Not |
|---|---|---|
| Adet | `#,##0` | Tam sayı |
| kg/m | `0.000` | 3 ondalık |
| TRY tutar (KDV hariç) | `#,##0.00` | 2 ondalık |
| Gün | `0` | Tam sayı |
| Tarih | `yyyy-mm-dd` | ISO; yerel ayar bağımsız |

**Yerel ayar kuralı:** Hücreye `28,50` gibi virgüllü ondalık yazılırsa ve değer sayı olarak
yorumlanabiliyorsa kabul edilir. `1.234,56` gibi binlik/ondalık karışık yazım veya para simgesi
içeren metin **hata** üretir (`E-FMT-01`); sessiz tahmin yapılmaz.

Tarih hücresi metin olarak girilmişse `yyyy-mm-dd` veya `dd.mm.yyyy` kabul edilir; diğer biçimler
`E-FMT-02` hatasıdır.

---

## 4. `Instructions` Sayfası

### 4.1 Sürüm ve beyan bloğu (A1–B4)

| Hücre | Değer |
|---|---|
| `A1` | `AI Procurement Intelligence — Tedarikçi Teklif Şablonu` |
| `A2` | `Template Version` |
| `B2` | `1.0` (metin) |
| `A3` | `Template Code` |
| `B3` | `RFQ-KUTUPROFIL-40X40X2` |
| `A4` | `VAT Declaration` |
| `B4` | `EVET` / `HAYIR` (kullanıcı seçimi; varsayılan `HAYIR`) |

`B4`, **kullanıcının doldurması gereken tek Instructions hücresidir** ve bu yüzden kilitsizdir.
Beyan cümlesi `A5` hücresinde sabit metin olarak yer alır (birleştirme yapılmaz; metin sağa taşar) —
bkz. Bölüm 2.1.1:

> Girdiğim tüm birim fiyat ve nakliye tutarları KDV hariçtir. Tedarikçi yalnızca KDV dahil fiyat
> verdiyse KDV hariç tutarı yazılı olarak teyit ettim ve teyit açıklamasını Notes alanına yazdım.

`B4` değeri `EVET` değilse dosya işleme alınmaz (`E-DECL-01`). Beyan metni silinemez; yalnızca
hücre değeri seçilir.

### 4.2 İçerik blokları (A6 ve sonrası)

| Blok | Başlık | Zorunlu mesaj |
|---|---|---|
| 1 | Ne satın alıyoruz? | 40 × 40 × 2 mm galvanizli kare çelik kutu profil, 6 m sabit, adet, TL/kg |
| 2 | Nasıl doldurulur? | `RFQ` sayfasını bir kez, `Supplier Quotes` sayfasını her tedarikçi için bir satır doldurun |
| 3 | Zorunlu alanlar | Bölüm 5 ve 6'daki "Zorunlu = Evet" alanlar. RFQ'da eksik zorunlu alan dosyayı reddeder; `Supplier Quotes` satırında eksik zorunlu alan yalnızca o satırı devre dışı bırakır |
| 4 | İsteğe bağlı alanlar | Termin, teklif geçerlilik tarihi, notlar, istenen teslimat süresi, açıklama |
| **5** | **KDV kuralı (zorunlu okuma)** | **Tüm fiyatlar KDV HARİÇ TRY/kg, nakliye KDV HARİÇ TRY'dir. Şablonda KDV alanı yoktur ve uygulama KDV hesaplamaz. Tedarikçi yalnızca KDV dahil fiyat verdiyse önce KDV hariç tutarı yazılı teyit edin, teyit açıklamasını Notes alanına serbest metin olarak yazın, sonra KDV hariç tutarı girin. KDV dahil tutarı fiyat alanına yazmayın. B4 hücresindeki `VAT Declaration` alanını `EVET` olarak işaretlemeden dosya işleme alınmaz.** |
| 6 | Hesaplamalar | Tüm maliyet hesapları uygulama tarafından yapılır; Excel'de formül yoktur |
| 7 | Yapılmaması gerekenler | Hesaplanmış sütun eklemeyin, satır silmeyin, başlıkları değiştirmeyin, birim değiştirmeyin, KDV sütunu eklemeyin |
| 8 | Nakliye kuralı | Nakliye tutarı yoksa boş bırakın; sıfır yazmayın. Uygulama sizden tutar veya açık varsayım ister |
| 9 | Ödeme koşulu | Yalnızca metni yazın (ör. `Peşin`, `Net 30`, `%50 peşin, %50 teslimatta`). Vade günü **yazılmaz**; uygulama tek anlamlı koşullardan türetir, karma koşullarda boş bırakır |
| 10 | kg/m notu | kg/m tedarikçinin beyanıdır; 2,27–2,51 aralığı dışındaki değerler yalnızca uyarı üretir |
| 11 | Tarihler | `Quote Date` zorunludur; `Quote Valid Until` teklif tarihinden önce olamaz |

Talimatlar kısa cümlelerle yazılır, örnek değer içerir, 45 satırı geçmez. KDV bloğu (Blok 5) kalın
yazı ve kırmızı kenarlıkla vurgulanır. Beyan bloğu (Blok 5 ile birlikte `A4:B4`) aynı vurguyu taşır.
Sayfada formül, makro, harici bağlantı bulunmaz.

**Instructions hücreleri ve kilit durumu**

| Hücre | İçerik | Kilit |
|---|---|---|
| `A1:B3` | Başlık, sürüm ve şablon kodu | Kilitli |
| `A4` | `VAT Declaration` etiketi | Kilitli |
| `B4` | `EVET` / `HAYIR` beyanı (zorunlu) | **Açık** |
| `A5:B50` | Beyan cümlesi + içerik blokları | Kilitli |

---

## 5. `RFQ` Sayfası

Yerleşim: A = alan adı (korumalı), B = değer, C = açıklama/örnek (korumalı, gri).
Başlık bandı A1:C1. Başlık satırı A2:C2 = `Field` / `Value` / `Guidance`.

| # | Alan (TR) | Başlık (EN, birebir) | Hücre | Tip | Zorunlu | Kilit | Örnek | Doğrulama | Hata/Uyarı |
|---|---|---|---|---|---|---|---|---|---|
| 1 | Talep numarası | `RFQ Number` | B3 | Metin (≤ 32) | Evet | **Açık** | `RFQ-2026-0042` | Boş olamaz; `^RFQ-[0-9]{4}-[0-9]{3,4}$` önerilir | `E-RFQ-01` boş; `W-RFQ-01` desen dışı |
| 2 | Talep tarihi | `RFQ Date` | B4 | Tarih | Evet | **Açık** | `2026-11-12` | Geçerli tarih; bugünden 1 günden fazla ileri tarih uyarı üretir | `E-RFQ-01` boş; `E-FMT-02` geçersiz; `W-RFQ-02` ileri tarih |
| 3 | Ürün kodu | `Product Code` | B5 | Metin (sabit) | Evet | Kilitli | `KUTUPROFIL-40X40X2-GALV` | Birebir eşleşmeli | `E-SPEC-04` farklı |
| 4 | Profil ölçüsü | `Profile Specification` | B6 | Metin (sabit) | Evet | Kilitli | `40x40x2` | Normalize sonrası birebir eşleşmeli | `E-SPEC-03` farklı |
| 5 | Profil uzunluğu | `Length (m)` | B7 | Sayı (sabit) | Evet | Kilitli | `6` | Tam olarak 6 | `E-LEN-02` farklı |
| 6 | Et kalınlığı | `Wall Thickness (mm)` | B8 | Sayı (sabit) | Evet | Kilitli | `2` | Tam olarak 2 | `E-SPEC-02` farklı |
| 7 | Sipariş miktarı | `Quantity (pcs)` | B9 | Tam sayı | Evet | **Açık** | `500` | Pozitif, 1–1.000.000 | `E-RFQ-02` geçersiz |
| 8 | Teslimat yeri | `Delivery Location` | B10 | Metin (≤ 120) | Evet | **Açık** | `Bursa / Nilüfer OSB` | Boş olamaz | `E-RFQ-01` boş |
| 9 | İstenen teslimat süresi | `Requested Delivery Days` | B11 | Tam sayı | Hayır | **Açık** | `20` | ≥ 0; boş bırakılabilir | `W-DLV-03` termin eksik bilgi |
| 10 | Para birimi | `Currency` | B12 | Metin (sabit) | Evet | Kilitli | `TRY` | Tam olarak `TRY` | `E-CUR-01` farklı |
| 11 | Açıklama | `Description` | B13 | Metin (≤ 500) | Hayır | **Açık** | `Projeler için, 6 m boy, galvanizli` | Serbest metin | — |

Sabit alanlar (3, 4, 5, 6, 10) gri dolgulu ve **kilitli**dir; programatik olarak da yazılamaz.
Hücrelerin boş bırakılması hata sayılır (sabit değer beklenir). Kullanıcı girişi gereken tüm
hücreler (B3, B4, B9, B10, B11, B13) **kilitsizdir** — `Delivery Location` (`B10`) dahil.

**RFQ sayfası kilit doğrulaması (tek tek)**

| Hücre | İçerik | Kilit | Gerekçe |
|---|---|---|---|
| `B3` | RFQ Number | Açık | Kullanıcı girişi (zorunlu) |
| `B4` | RFQ Date | Açık | Kullanıcı girişi (zorunlu) |
| `B5` | Product Code | Kilitli | Sabit ürün kimliği |
| `B6` | Profile Specification | Kilitli | Sabit ürün kimliği |
| `B7` | Length (m) | Kilitli | Sabit: 6 m |
| `B8` | Wall Thickness (mm) | Kilitli | Sabit: 2 mm |
| `B9` | Quantity (pcs) | Açık | Kullanıcı girişi (zorunlu) |
| `B10` | Delivery Location | Açık | Kullanıcı girişi (zorunlu) |
| `B11` | Requested Delivery Days | Açık | Kullanıcı girişi (isteğe bağlı) |
| `B12` | Currency | Kilitli | Sabit: TRY |
| `B13` | Description | Açık | Kullanıcı girişi (isteğe bağlı) |

A sütunu (alan adları) ve C sütunu (yönlendirme) tamamen kilitlidir.

**RFQ zorunlu / isteğe bağlı alanlar**

| Sınıf | Alanlar |
|---|---|
| Zorunlu (boş → dosya hatası) | B3 `RFQ Number`, B4 `RFQ Date`, B5 `Product Code`, B6 `Profile Specification`, B7 `Length (m)`, B8 `Wall Thickness (mm)`, B9 `Quantity (pcs)`, B10 `Delivery Location`, B12 `Currency` |
| İsteğe bağlı (boş kalabilir) | B11 `Requested Delivery Days`, B13 `Description` |

Sabit alanlar boş bırakılamaz; kullanıcı bunları değiştiremez ve uygulama boş değeri varsayılanla
tamamlamaz.

**RFQ miktarı ile satır miktarı ilişkisi:** `Supplier Quotes` sayfasındaki `Quantity (pcs)` her
satırda zorunludur ve RFQ miktarıyla **eşit olmalıdır** (`E-Q-01`, hata). Böylece her teklifin
hangi sipariş büyüklüğü için verildiği satır düzeyinde belgelenir ve sessiz sapma engellenir.

---

## 6. `Supplier Quotes` Sayfası

Yerleşim: başlık satırı `2`, veri `3:102` (en fazla 100 tedarikçi). Tamamen boş satırlar atlanır.
Sütun genişlikleri: metin 22, sayı 14, not 40; satır yüksekliği 20; başlık satırı donmuş.
Toplam **14 sütun (A–N)**.

| # | Alan (TR) | Başlık (EN, birebir) | Sütun | Tip | Zorunlu | Örnek | Doğrulama | Hata/Uyarı |
|---|---|---|---|---|---|---|---|---|
| 1 | Tedarikçi adı | `Supplier Name` | A | Metin (≤ 120) | Evet | `Bursa Metal A.Ş.` | Boş olamaz; büyük/küçük harf ve boşluk duyarsız tekilleştirme | `E-NAME-01` boş; `E-DUP-01` mükerrer |
| 2 | Teklif numarası | `Supplier Quote Number` | B | Metin (≤ 60) | Evet | `BM-2026-1187` | Boş olamaz | `E-NAME-02` boş; `W-DUP-02` aynı tedarikçide tekrar |
| 3 | Profil ölçüsü | `Profile Specification` | C | Metin (sabit) | Evet | `40x40x2` | Normalize sonrası `40x40x2` | `E-SPEC-01` farklı → satır devre dışı |
| 4 | Profil uzunluğu | `Length (m)` | D | Sayı (sabit) | Evet | `6` | Tam olarak 6 | `E-LEN-01` farklı → satır devre dışı |
| 5 | Miktar | `Quantity (pcs)` | E | Tam sayı | Evet | `500` | Pozitif, 1–1.000.000, RFQ ile eşit | `E-QTY-01` geçersiz; `E-Q-01` uyuşmaz |
| 6 | Birim ağırlık | `Weight (kg/m)` | F | Sayı | Evet | `2.39` | > 0; 2,27–2,51 aralığı dışı uyarı | `E-WT-01` ≤ 0; `W-WT-01` aralık dışı |
| 7 | Birim fiyat (KDV hariç) | `Unit Price (TRY/kg, excl. VAT)` | G | Sayı | Evet | `28.50` | > 0; makul üst sınır 1.000.000; **KDV hariç** | `E-PRICE-01` ≤ 0; `W-PRICE-01` aşırı değer |
| 8 | Nakliye maliyeti (KDV hariç) | `Shipping Cost (TRY, excl. VAT)` | H | Sayı | Evet* | `7500.00` | ≥ 0; boş bırakılabilir ama sıfır sayılmaz; **KDV hariç** | `E-SHIP-01` negatif; `I-SHIP-01` boş → bilgi talebi |
| 9 | Nakliye kaynağı | `Shipping Cost Source` | I | Enum | Evet* | `Supplier Quote` | `Supplier Quote` veya `Customer Assumption` | `E-SRC-01` eksik/geçersiz; `W-SRC-01` varsayım ama not yok |
| 10 | Teslimat süresi | `Delivery Days` | J | Tam sayı | Hayır | `15` | ≥ 0 | `W-DLV-01` boş; `W-DLV-02` > 365 |
| 11 | Ödeme koşulu | `Payment Terms` | K | Metin (≤ 120) | Evet | `Net 30` | Boş olamaz; metin aynen korunur | `E-PAY-01` boş; `W-PAY-01` karma/belirsiz |
| 12 | Teklif tarihi | `Quote Date` | L | Tarih | Evet | `2026-11-12` | Geçerli tarih; bugünden sonrası uyarı | `E-QD-01` boş; `E-FMT-02` okunamayan tarih; `W-QD-01` gelecek tarih |
| 13 | Teklif geçerlilik tarihi | `Quote Valid Until` | M | Tarih | Hayır | `2026-12-31` | Geçerli tarih; `Quote Date`'ten önce olamaz | `E-VAL-01` önce; `E-FMT-02` geçersiz; `W-VAL-01` geçmiş tarih |
| 14 | Notlar | `Notes` | N | Metin (≤ 500) | Hayır | `Nakliye dahil; fiyat KDV hariç teyitli` | Serbest metin; otomatik tarama yok | — |

`*` **Nakliye alanları koşullu zorunludur:** `Shipping Cost (TRY, excl. VAT)` boşsa
`Shipping Cost Source` da boş bırakılır ve uygulama **bilgi talebi** üretir. `Shipping Cost` doluysa
kaynak alanı zorunludur. Boş nakliye hiçbir koşulda sıfır kabul edilmez.

**Zorunlu / isteğe bağlı alan özeti**

| Sınıf | Alanlar |
|---|---|
| Zorunlu (boş → satır hatası) | A `Supplier Name`, B `Supplier Quote Number`, C `Profile Specification`, D `Length (m)`, E `Quantity (pcs)`, F `Weight (kg/m)`, G `Unit Price (TRY/kg, excl. VAT)`, K `Payment Terms`, L `Quote Date` |
| Koşullu zorunlu | H `Shipping Cost` (boş → `I-SHIP-01` bilgi talebi; yanıtlanmadan karşılaştırma tamamlanmaz), I `Shipping Cost Source` (H doluysa zorunlu) |
| İsteğe bağlı (boş kalabilir) | J `Delivery Days`, M `Quote Valid Until`, N `Notes` |

Hiçbir zorunlu alan için varsayılan değer, tahmin, ortalama veya komşu satırdan kopyalama yapılmaz.

**Kaldırılan alan:** `Payment Days` artık Excel'de **yoktur**. Vade günü yalnızca `Payment Terms`
metninden Python tarafında türetilir (bkz. Bölüm 8) ve arayüzde gösterilir. Şablonda bu alan için
sütun, veri doğrulaması, koşullu biçimlendirme veya koruma aralığı bulunmaz.

**Sütun sırasının gerekçesi:** Tarih alanları (`Quote Date`, `Quote Valid Until`) ve `Notes` sonda
toplanmıştır; `Payment Terms` metninden sonra sayısal bir vade sütunu gelmediği için satır, metin →
tarih → serbest not akışında okunur.

**`Shipping Cost Source` değerleri**

| Değer | Anlam | Ek koşul |
|---|---|---|
| `Supplier Quote` | Tutar tedarikçi teklifinde yazıyor | — |
| `Customer Assumption` | Tutar satın alma tarafından varsayıldı | `Notes` alanında kısa gerekçe zorunlu (yoksa `W-SRC-01`) |

**Başlık benzersizliği:** 14 başlık birbirinden farklıdır. Para birimi ve KDV durumu ayrı sütun
değildir; başlık adının içinde (`TRY`, `TRY/kg`, `excl. VAT`) taşınır.

**Satır sınırı:** 3–102 arası dolu satırlar okunur. 100'den fazla tedarikçi gerekirse dosya
bölünür; şablon genişletilmez.

---

## 7. Veri Kuralları ve Doğrulama

### 7.1 Şiddet düzeyleri

| Düzey | Kod öneki | Etki |
|---|---|---|
| Hata | `E-` | İlgili satır karşılaştırmaya alınmaz (dosya düzeyindeyse dosya reddedilir) |
| Bilgi talebi | `I-` | Karşılaştırma tamamlanmış sayılmaz; kullanıcı bilgi girene kadar bekler |
| Uyarı | `W-` | Karşılaştırma sürer; uyarı arayüzde ve raporda görünür |

Dosya düzeyi hatalar: KDV beyanı (`E-DECL-01`), şablon kodu/sürümü, eksik zorunlu sütun, eksik
sayfa, bozuk başlık satırı, eksik zorunlu RFQ alanı. Satır düzeyi hatalar: yalnızca o tedarikçiyi
etkiler; diğerleriyle karşılaştırma yapılabilir.

### 7.2 Hata ve uyarı kataloğu

Katalogda **46 kod** bulunur: **29 hata (`E-`)**, **16 uyarı (`W-`)**, **1 bilgi talebi (`I-`)**.
Kodlar tektir; her kod belgede yalnızca bu katalogda tanımlanır ve başka bölümlerde yalnızca
referans olarak geçer.

`Notes` alanına özgü **KDV kodu yoktur**; alan serbest metindir (bkz. Bölüm 2.1.2). Alanla ilişkili
tek kod `W-SRC-01`'dir ve o da KDV ile değil, `Customer Assumption` nakliyesinin gerekçesiyle
ilgilidir.

| Kod | Alan | Koşul | Düzey |
|---|---|---|---|
| `E-DECL-01` | Instructions!B4 | `VAT Declaration` = `HAYIR` veya boş | Hata (dosya) |
| `E-TPL-01` | Instructions!B3 | Şablon kodu farklı/boş | Hata (dosya) |
| `E-TPL-02` | Instructions!B2 | Major sürüm uyumsuz | Hata (dosya) |
| `E-TPL-03` | Instructions!B2 | Aynı major sürümde zorunlu sütun eksik | Hata (dosya) |
| `W-TPL-01` | Instructions!B2 | Minor sürüm farklı, zorunlu sütunların tamamı var | Uyarı |
| `E-SHEET-01` | — | Zorunlu sayfa yok | Hata (dosya) |
| `W-SHEET-01` | — | Fazladan sayfa var | Uyarı |
| `E-HDR-01` | Başlık satırı | Beklenen başlık eksik veya yazımı farklı | Hata (dosya) |
| `W-HDR-01` | Başlık satırı | Beklenmeyen/bilinmeyen sütun var (okunmaz) | Uyarı |
| `W-HDR-02` | Başlık satırı | Sütun sırası değişmiş | Uyarı |
| `E-RFQ-01` | RFQ B3/B4/B10 | Zorunlu alan boş (talep no, talep tarihi, teslimat yeri) | Hata (dosya) |
| `E-RFQ-02` | RFQ B9 | Sipariş miktarı geçersiz (pozitif tam sayı değil) | Hata (dosya) |
| `W-RFQ-01` | RFQ B3 | Önerilen desen dışı | Uyarı |
| `W-RFQ-02` | RFQ B4 | İleri tarih | Uyarı |
| `E-SPEC-01` | C | Profil ölçüsü `40x40x2` değil | Hata (satır) |
| `E-SPEC-02` | RFQ B8 | Et kalınlığı 2 mm değil | Hata (dosya) |
| `E-SPEC-03` | RFQ B6 | Profil ölçüsü `40x40x2` değil | Hata (dosya) |
| `E-SPEC-04` | RFQ B5 | Ürün kodu `KUTUPROFIL-40X40X2-GALV` değil | Hata (dosya) |
| `E-LEN-01` | D | Satır uzunluğu 6 m değil | Hata (satır) |
| `E-LEN-02` | RFQ B7 | RFQ uzunluğu 6 m değil | Hata (dosya) |
| `E-CUR-01` | RFQ B12 | Para birimi `TRY` değil | Hata (dosya) |
| `E-QTY-01` | E | Pozitif tam sayı değil / aralık dışı | Hata (satır) |
| `E-Q-01` | E | RFQ miktarından farklı | Hata (satır) |
| `E-WT-01` | F | kg/m ≤ 0 veya sayı değil | Hata (satır) |
| `W-WT-01` | F | kg/m < 2,27 veya > 2,51 | Uyarı |
| `E-PRICE-01` | G | TL/kg ≤ 0 veya sayı değil | Hata (satır) |
| `W-PRICE-01` | G | TL/kg > 1.000.000 (olası birim hatası) | Uyarı |
| `E-SHIP-01` | H | Negatif nakliye | Hata (satır) |
| `I-SHIP-01` | H | Nakliye boş | Bilgi talebi |
| `E-SRC-01` | I | H dolu iken kaynak boş/geçersiz | Hata (satır) |
| `W-SRC-01` | I + N | `Customer Assumption` ama not boş | Uyarı |
| `W-DLV-01` | J | Termin boş | Uyarı |
| `W-DLV-02` | J | Termin > 365 gün | Uyarı |
| `W-DLV-03` | RFQ B11 | İstenen teslimat süresi boş | Uyarı |
| `E-PAY-01` | K | Ödeme koşulu boş | Hata (satır) |
| `W-PAY-01` | K | Karma/belirsiz koşul — vade günü türetilemedi | Uyarı |
| `E-QD-01` | L | Teklif tarihi boş | Hata (satır) |
| `W-QD-01` | L | Teklif tarihi gelecekte | Uyarı |
| `E-VAL-01` | M | Geçerlilik tarihi `Quote Date`'ten önce | Hata (satır) |
| `W-VAL-01` | M | Geçerlilik tarihi geçmiş | Uyarı |
| `E-NAME-01` | A | Tedarikçi adı boş | Hata (satır) |
| `E-NAME-02` | B | Teklif numarası boş | Hata (satır) |
| `E-DUP-01` | A | Aynı tedarikçi birden fazla satırda | Hata (satır) |
| `W-DUP-02` | A + B | Aynı tedarikçi + aynı teklif no tekrar | Uyarı |
| `E-FMT-01` | Sayısal alanlar | Sayı okunamadı (para simgesi, karışık ayırıcı) | Hata (satır) |
| `E-FMT-02` | Tarih alanları | Tarih okunamadı | Hata (satır) |

### 7.3 Kural notları

- **Miktar:** pozitif tam sayı. Ondalık girilirse `E-QTY-01`. `0` ve negatif reddedilir.
- **kg/m:** pozitif sayı. `0` reddedilir (`E-WT-01`), çünkü sıfır ağırlık toplam maliyeti
  anlamsız kılar. 2,27–2,51 dışı değer **bloklamaz**, yalnızca uyarır; hesaplama yine de beyan
  edilen değerle yapılır.
- **TL/kg:** pozitif sayı ve **KDV hariç**. Fiyat `0` ise satır karşılaştırmaya alınmaz.
- **KDV:** hesaplanmaz, eklenmez, çıkarılmaz, dönüştürülmez. Standart iki görünür mekanizmayla
  korunur: `Instructions!B4` zorunlu beyanı (`E-DECL-01`) ve sütun başlıklarındaki `excl. VAT`
  etiketi. `Notes` alanı serbest metindir; içindeki KDV ifadeleri **taranmaz ve hata üretmez**
  (otomatik kelime taraması yoktur). Uygulama tutarın KDV içerip içermediğini tahmin etmez.
- **Nakliye:** negatif olamaz ve **KDV hariçtir**. Boş bırakmak serbesttir ama sıfır anlamına
  gelmez; uygulama `I-SHIP-01` üretir ve kullanıcıdan tutar veya açık varsayım ister.
- **Termin:** boş bırakılabilir; karşılaştırmada "—" olarak gösterilir ve `W-DLV-01` üretilir.
- **Tarihler:** `Quote Date` (L) zorunludur; boşsa `E-QD-01`, okunamıyorsa `E-FMT-02`. Gelecek
  tarih `W-QD-01` uyarısı üretir ama satırı düşürmez. `Quote Valid Until` (M) isteğe bağlıdır;
  doluysa `Quote Date`'ten önce olamaz (`E-VAL-01`).
- **Mükerrer tedarikçi:** ad, baştaki/sondaki boşluklar silinip büyük/küçük harf duyarsız
  karşılaştırılır (`"bursa metal a.ş."` = `"Bursa Metal A.Ş."`). Aynı tedarikçi iki satırda ise
  her iki satır da karşılaştırmaya alınmaz; kullanıcı tek satır bırakmalıdır.
- **Farklı ürün:** `Profile Specification` veya `Length (m)` uyuşmayan satır karşılaştırmaya
  alınmaz (`E-SPEC-01`, `E-LEN-01`). Uyarı ile geçilmez; ürün kimliği pazarlık konusu değildir.
- **Normalizasyon:** `×` → `x`, boşluk silme, `mm` sonekini atma, Türkçe `İ/I` katlama yalnızca
  karşılaştırma içindir; raporda kullanıcının yazdığı değer gösterilir.
- **Termin kaynağı:** şablon yalnızca `Delivery Days` taşır; teslim tarihi ayrı sütun değildir.
  Brief'teki "teslim tarihi" alanı, `RFQ Date` + `Delivery Days` üzerinden **uygulama tarafında**
  türetilir ve yalnızca okunur biçimde gösterilir. Şablona teslim tarihi sütunu eklenmez.
- **Vade günü kaynağı:** şablonda vade günü sütunu **yoktur**; brief'teki vade günü alanı
  türetilmiş bir görüntü alanıdır (bkz. Bölüm 8).

---

## 8. Ödeme Koşulu — Vade Günü Python Tarafında Türetilir

`Payment Terms` (K) tek giriş alanıdır ve **orijinal metin** olarak saklanıp raporda aynen
gösterilir. Vade günü Excel'de sorulmaz; aşağıdaki kuralla uygulama tarafında türetilir ve yalnızca
arayüzde/raporda bilgi olarak gösterilir.

| `Payment Terms` örneği | Türetilen vade günü | Kural |
|---|---|---|
| `Peşin` / `Peşin ödeme` / `Cash` | `0` | Peşin = 0 gün |
| `Net 30` / `30 gün vadeli` | `30` | Rakam doğrudan alınır |
| `Net 45` | `45` | — |
| `Net 60` | `60` | — |
| `%50 peşin, %50 teslimatta` | **boş** | Karma; tek güne indirgenemez |
| `30 gün çek + 30 gün vade` | **boş** | Karma/çok dilimli |
| Metinden gün çıkarılamıyor | **boş** | Belirsiz; arayüzde "karma / belirsiz" |

**Değişmez kurallar**

1. Sistem belirsiz veya karma koşulda **gün uydurmaz**; alan boş kalır ve "karma / belirsiz" olarak
   işaretlenir (`W-PAY-01`, uyarı düzeyinde; karşılaştırmayı durdurmaz).
2. Türetilen değer kullanıcı tarafından düzeltilebilir bir Excel alanı değildir; yalnızca okunur
   gösterim ve rapordur.
3. Ödeme koşulu ve vade günü **maliyet hesabına ve sıralamaya dahil edilmez**; yalnızca
   karşılaştırma bilgisidir.
4. Şablonda ödeme koşuluna bağlı hiçbir hesaplama alanı, veri doğrulaması veya koşullu
   biçimlendirme yoktur.

---

## 9. Excel İçi Doğrulama ve Korumalar

Şablon, kullanıcı hatasını dosya yüklenmeden önce azaltmak için şu Excel özelliklerini taşır.
Bunlar **uygulama doğrulamasının yerine geçmez**; ikinci savunma hattıdır.

| Hücre/alan | Excel veri doğrulaması | Stil | Mesaj |
|---|---|---|---|
| `A3:A102` | Uzunluk ≤ 120, boş değil | Durdur | "Tedarikçi adı zorunludur" |
| `B3:B102` | Uzunluk ≤ 60, boş değil | Durdur | "Teklif numarası zorunludur" |
| `C3:C102` | Özel, `=EXACT(TRIM(C3),"40x40x2")` | Durdur | "Ürün sabittir: 40x40x2" |
| `D3:D102` | Ondalık, `= 6` | Durdur | "Profil uzunluğu sabittir: 6 m" |
| `E3:E102` | Tam sayı, `1`–`1000000` | Durdur | "Miktar pozitif tam sayı olmalıdır" |
| `F3:F102` | Ondalık, `> 0` | Durdur | "kg/m pozitif olmalıdır" |
| `F3:F102` | Ondalık, `2,27`–`2,51` | **Uyarı** | "Beyan edilen kg/m referans aralığın dışında; devam edilebilir" |
| `G3:G102` | Ondalık, `> 0` | Durdur | "Birim fiyat KDV HARİÇ TRY/kg olarak pozitif olmalıdır" |
| `H3:H102` | Ondalık, `>= 0` | Durdur | "Nakliye KDV HARİÇ TRY'dir ve negatif olamaz; bilinmiyorsa boş bırakın" |
| `I3:I102` | Liste: `Supplier Quote,Customer Assumption` | Durdur | "Kaynak seçin" |
| `J3:J102` | Tam sayı, `0`–`3650` | Bilgi | "Termin gün cinsindendir" |
| `K3:K102` | Uzunluk ≤ 120, boş değil | Durdur | "Ödeme koşulu metni zorunludur" |
| `L3:L102` | Tarih | Durdur | "Teklif tarihi zorunludur (yyyy-mm-dd)" |
| `M3:M102` | Özel, `=OR(M3="",M3>=L3)` | Durdur | "Geçerlilik tarihi teklif tarihinden önce olamaz" |
| `N3:N102` | Uzunluk ≤ 500 | Bilgi | "Serbest not alanı; KDV teyit açıklamanızı buraya yazabilirsiniz" |
| `RFQ!B3` | Uzunluk ≤ 32, boş değil | Durdur | "Talep numarası zorunludur" |
| `RFQ!B4` | Tarih | Durdur | "Talep tarihi (yyyy-mm-dd)" |
| `RFQ!B9` | Tam sayı, `1`–`1000000` | Durdur | "Sipariş miktarı pozitif tam sayı olmalıdır" |
| `RFQ!B10` | Uzunluk ≤ 120, boş değil | Durdur | "Teslimat yeri zorunludur" |
| `RFQ!B11` | Tam sayı, `0`–`3650` | Bilgi | "İstenen teslim süresi gün cinsindendir" |
| `Instructions!B4` | Liste: `EVET,HAYIR` | Durdur | "KDV hariç fiyat beyanını onaylamadan dosya işleme alınmaz" |

`Payment Terms` (K) serbest metin alanıdır; Excel veri doğrulaması yoktur. `Payment Days` alanı
tamamen kaldırıldığı için bu alana ait veri doğrulaması, koşullu biçimlendirme veya ayrı koruma
aralığı bulunmaz (K sütunu genel veri girişi aralığı `A3:N102` içindedir).

**Koşullu biçimlendirme**

- `F3:F102`: 2,27–2,51 dışı → açık sarı zemin (`W-WT-01`)
- `H3:H102`: boş → açık mavi zemin (bilgi talebi göstergesi)
- `M3:M102`: `M < L` → açık kırmızı zemin (`E-VAL-01`)
- `G3:G102` ve `H3:H102`: başlık bandında `excl. VAT` vurgusu; hücre düzeyinde KDV rengi yoktur

**Koruma**

| Aralık | Durum | İçerik |
|---|---|---|
| `Instructions!A1:B3` | Kilitli | Başlık, sürüm ve şablon kodu |
| `Instructions!A4` | Kilitli | `VAT Declaration` etiketi |
| `Instructions!B4` | **Açık** | `EVET` / `HAYIR` beyanı (zorunlu) |
| `Instructions!A5:B50` | Kilitli | Beyan cümlesi + içerik blokları (en çok 45 satır) |
| `RFQ!A1:C1` | Kilitli | Başlık bandı |
| `RFQ!A2:A13` | Kilitli | Alan adları |
| `RFQ!C2:C13` | Kilitli | Yönlendirme metinleri |
| `RFQ!B2` | Kilitli | Sütun başlığı (`Value`) |
| `RFQ!B3` | **Açık** | RFQ Number (zorunlu) |
| `RFQ!B4` | **Açık** | RFQ Date (zorunlu) |
| `RFQ!B5` | Kilitli | Product Code (sabit) |
| `RFQ!B6` | Kilitli | Profile Specification (sabit) |
| `RFQ!B7` | Kilitli | Length (sabit) |
| `RFQ!B8` | Kilitli | Wall Thickness (sabit) |
| `RFQ!B9` | **Açık** | Quantity (zorunlu) |
| `RFQ!B10` | **Açık** | Delivery Location (zorunlu) |
| `RFQ!B11` | **Açık** | Requested Delivery Days (isteğe bağlı) |
| `RFQ!B12` | Kilitli | Currency (sabit) |
| `RFQ!B13` | **Açık** | Description (isteğe bağlı) |
| `Supplier Quotes!A1:N1` | Kilitli | Başlık bandı |
| `Supplier Quotes!A2:N2` | Kilitli | Sütun başlıkları |
| `Supplier Quotes!A3:N102` | **Açık** | Veri girişi |

- Parola kullanılmaz (kullanıcı gerekirse korumayı kaldırabilir); amaç yanlışlıkla başlık/satır
  silmeyi engellemektir.
- Kullanıcının doldurması gereken hiçbir hücre kilitli değildir; kilitli her hücre sabit değer
  taşır. Bu kontrol Bölüm 5'teki hücre tablosuyla tek tek eşleşir.
- Dosyada makro, harici bağlantı, gizli sayfa, tanımlı ad bulunmaz. Koşullu biçimlendirme ve veri
  doğrulaması dışında hiçbir hücrede formül yoktur.

---

## 10. Hesaplanan Alanlar — Excel'de Yer Almaz

Şablonda hesaplanmış hiçbir sütun bulunmaz. Aşağıdaki değerler yalnızca Python tarafında üretilir
ve arayüz/raporda gösterilir. **Tümü KDV hariç tutarlarla çalışır; hiçbir adımda KDV eklenmez.**

1. Toplam uzunluk (m) = `Quantity (pcs)` × 6
2. Toplam ağırlık (kg) = `Quantity (pcs)` × 6 × `Weight (kg/m)`
3. Ürün bedeli (TRY, KDV hariç) = toplam ağırlık × `Unit Price (TRY/kg, excl. VAT)`
4. Toplam tedarik maliyeti (TRY, KDV hariç) = ürün bedeli + nakliye
5. Kg başına nakliye (TRY/kg, KDV hariç) = nakliye ÷ toplam ağırlık

Nakliye değeri `Shipping Cost (TRY, excl. VAT)` alanından veya kullanıcının arayüzde girdiği açık
varsayımdan gelir; varsayım raporda "varsayım" olarak etiketlenir.

Şablona kullanıcı tarafından eklenen hesaplanmış sütunlar okunmaz ve yok sayılır; fazladan sütun
`W-HDR-01` uyarısıyla, sütun sırası değişikliği `W-HDR-02` uyarısıyla bildirilir. `E-HDR-01`
yalnızca beklenen 14 başlığın eksik veya bozuk olması durumunda üretilir.

---

## 11. Dosya Okuma Kuralları (ayrıştırıcı sözleşmesi)

1. Dosya `.xlsx` olarak açılır; makro içeren dosyalar reddedilir.
2. `Instructions!B4` beyanı okunur; `EVET` değilse dosya reddedilir (`E-DECL-01`) — **hiçbir alan
   varsayılanla doldurulmaz**.
3. `Instructions!B3` şablon kodu, `Instructions!B2` sürüm okunur; Bölüm 2.3 uygulanır.
   Zorunlu sütun eksikse dosya reddedilir — **hiçbir alan varsayılanla doldurulmaz**.
4. `RFQ` sayfasında `B3:B13` alanları adlarıyla eşleştirilir; boş zorunlu alan `E-RFQ-01`/`E-RFQ-02`.
5. `Supplier Quotes` sayfasında `2.` satır başlık kabul edilir; başlıklar normalize edilip
   (boşluk kırpma, çoklu boşluk tekilleştirme) karşılaştırılır.
6. Veri `3.` satırdan itibaren okunur; tamamen boş satırlar atlanır.
7. Tamamen boş olmayan ama `Supplier Name` boş olan satır `E-NAME-01` üretir.
8. Yalnızca Bölüm 6'daki 14 başlık okunur. Tanımsız sütunlar **veri olarak kullanılmaz**
   (`W-HDR-01`).
9. Zorunlu alan boşsa ilgili satır hatası üretilir; tahmin, ortalama, komşu değer veya varsayılan
   atama yapılmaz. İsteğe bağlı alan boşsa olduğu gibi boş bırakılır.
10. `Notes` alanı serbest metindir; içeriği taranmaz, anahtar kelime aranmaz ve bu alandan dolayı
    hata üretilmez. KDV teyit açıklaması burada saklanır ve rapora bilgi olarak taşınır.
11. Tüm satırlar okunduktan sonra mükerrer kontrolü ve RFQ miktar eşleşmesi yapılır.
12. Sonuç: satır başına hesaplanan değerler + hata/uyarı/bilgi listesi. Hiçbir hata yoksa (ve bilgi
    talebi yoksa) karşılaştırma tamamlanır.
13. Vade günü dosyadan okunmaz; Bölüm 8'deki kuralla bu aşamada türetilir.

Okuma sırasında hiçbir hücreye yazılmaz; yüklenen dosya değiştirilmez.

---

## 12. Örnek Doldurulmuş Satırlar

`Instructions` sayfası (beyan):

| Hücre | Örnek değer |
|---|---|
| `B2` | `1.0` |
| `B3` | `RFQ-KUTUPROFIL-40X40X2` |
| `B4` | `EVET` |

`RFQ` sayfası:

| Alan | Değer |
|---|---|
| RFQ Number | `RFQ-2026-0042` |
| RFQ Date | `2026-11-12` |
| Product Code | `KUTUPROFIL-40X40X2-GALV` |
| Profile Specification | `40x40x2` |
| Length (m) | `6` |
| Wall Thickness (mm) | `2` |
| Quantity (pcs) | `500` |
| Delivery Location | `Bursa / Nilüfer OSB` |
| Requested Delivery Days | `20` |
| Currency | `TRY` |
| Description | `6 m boy, sıcak daldırma galvaniz; tüm fiyatlar KDV hariç` |

`Supplier Quotes` sayfası (3 tedarikçi, 14 sütun):

| Supplier Name | Supplier Quote Number | Profile Specification | Length (m) | Quantity (pcs) | Weight (kg/m) | Unit Price (TRY/kg, excl. VAT) | Shipping Cost (TRY, excl. VAT) | Shipping Cost Source | Delivery Days | Payment Terms | Quote Date | Quote Valid Until | Notes |
|---|---|---|---|---|---|---|---|---|---|---|---|---|---|
| Bursa Metal A.Ş. | BM-2026-1187 | 40x40x2 | 6 | 500 | 2.39 | 28.50 | 7500.00 | Supplier Quote | 15 | Net 30 | 2026-11-12 | 2026-12-31 | Nakliye dahil; fiyat KDV hariç teyitli |
| Gemlik Profil Ltd. | GP-8841 | 40x40x2 | 6 | 500 | 2.31 | 30.10 | 5200.00 | Supplier Quote | 25 | %50 peşin, %50 teslimatta | 2026-11-10 | 2026-12-15 | Tedarikçi 36,12 KDV dahil fiyat verdi; 30,10 KDV hariç yazılı teyit edildi |
| İnegöl Çelik San. | IC-2026-77 | 40x40x2 | 6 | 500 | 2.58 | 27.90 | 9100.00 | Customer Assumption | 18 | Peşin | 2026-11-11 | 2026-11-30 | Nakliye tahmini; kg/m beyanı yüksek |

Bu örnekte beklenen uygulama davranışı:

- 3. satırda `W-WT-01` (2,58 > 2,51) uyarısı; satır yine karşılaştırmaya girer.
- 2. satırda `W-PAY-01` bilgilendirmesi (karma koşul → vade günü boş); maliyet etkilenmez.
- 1. ve 2. satırda vade günü sırasıyla 30 ve boş olarak türetilir; 3. satırda 0.
- 3. satırda `Customer Assumption` seçildiği ve not dolu olduğu için `W-SRC-01` yok.
- Hiçbir satırda `E-` hatası yok; karşılaştırma tamamlanır.

**Hatalı örnek (reddedilir):** `Quote Date` = `2026-12-20`, `Quote Valid Until` = `2026-12-15`
ise `E-VAL-01` üretilir.

**Hatalı örnek (reddedilir):** `Instructions!B4` = `HAYIR` veya boşsa `E-DECL-01` üretilir ve dosya
işleme alınmaz; kullanıcı beyanı `EVET` yapmalıdır.

**Hata üretmeyen örnek:** `Notes` alanında `Teklif 34,20 KDV dahil idi; 28,50 KDV hariç yazılı teyit
edildi` yazması **hata üretmez**. Not yalnızca bilgi olarak saklanır ve raporda gösterilir.

---

## 13. Gelecekte Genişleme (MVP Dışı)

Şablon tek ürün için tasarlanmıştır ve MVP'de genişletilmez. İleride yeni bir ürün kategorisi
eklenirse yaklaşım şudur — **bu sürümde uygulanmaz**:

- Ürün kimliği (`Product Code`, `Profile Specification`) parametreleştirilir.
- Şablon kodu ürün başına ayrı olur (ör. `RFQ-PROFIL-...`, `RFQ-SAC-...`).
- Sütun seti ürün tipine göre değişebilir; ortak çekirdek alanlar korunur.
- Sürüm kodu major artırılır; eski dosyalar reddedilir veya dönüştürülür.

KDV hesaplama, çoklu para birimi, veritabanı ve otomatik kazanan puanı bu genişlemenin de parçası
değildir; ayrı kapsam kararı gerektirir. Bu genişleme için bugünkü şablona ek sütun, ek sayfa veya
ek alan **konmaz**.

---

## 14. Açık Nokta

Yok. Bu spesifikasyon `PROJECT_BRIEF.md` v0.3'teki kapatılmış kararlarla uyumludur; tek fark,
brief'te "ayrı sayısal alan" olarak geçen vade gününün Excel girişi değil **Python tarafında
türetilen görüntü alanı** olmasıdır (Bölüm 8).

KDV standardı **içerik taramasıyla değil**, `Instructions!B4` zorunlu beyanı (`E-DECL-01`),
`Instructions` Blok 5 talimatı ve `excl. VAT` başlık etiketiyle güvenceye alınır. `Notes` alanı
serbest metindir ve taranmaz; bu bilinçli bir tasarım kararıdır (Bölüm 2.1.1–2.1.2). Yeni alan,
yeni sayfa veya yeni kural eklenmesi kapsam değişikliği sayılır.
