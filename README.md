# is-yatirim-fetcher

## Excel gereksinimleri

- Çalışma kitabında **`Main`** sayfası zorunludur: `A1` = `Ticker`, hisse kodları `A2` ve aşağısı. Betik **Main** sayfasına yazmaz; tickler yalnızca buradan okunur.
- **Sinyal** sayfası Python ile hesaplanıp değer olarak yazılır; Microsoft 365 gerekmez. `A` sütunu `Main!A` ile senkron formül; `B` (Sektör) Yahoo Finance’ten gelen metin değeridir.
- **Config** sayfası: `analiz_donemi` (**B1**) — hem çekilecek **çeyrek sayısını** hem Sinyal’deki pencere / `pct` paydasını belirler; `ucuz_esik` (B2), `pahali_esik` (B3). Eski kitaplarda kalan **F sütunu** dönem etiketleri betik tarafından temizlenir (kullanılmaz).

## Sayfa yapısı

- **Raw_data**: Uzun tablo — `Ticker`, `Metrik`, `Dönem`, `Değer`. Önce çeyreklik geçmiş (`PDDD`, `FD_FAVOK`, `BORC_OZK`), ardından özet snapshot satırları (TTM `FD_FAVOK` vb.; çeyreklik `FD_FAVOK` yalnızca geçmiş blokta, güncel Yabancı fiyatı ile).
- **Yabancı Oranı**: Ticker, fiyat, dönem yabancı oranları ve değişim.

### Raw_data `Metrik` sütunu — kodların anlamları

| Kod | Anlam |
|-----|--------|
| `PDDD` | PD/DD — piyasa değeri ÷ özkaynaklar (özsermaye defter değeri). Yalnızca **çeyreklik geçmiş** blokta (özet snapshot’ta tekrar yazılmaz; aynı dönem için çift satır/çelişen değer oluşmasın diye). |
| `FD_FAVOK` | FD/FAVÖK — (güncel fiyat × `item_2oa` + net borç) ÷ FAVÖK. **Çeyreklik geçmiş:** her çeyrek için o çeyreğin bilanço kırılımları + aynı güncel fiyat (Yabancı Oranı); payda çeyrek FAVÖK × 4. **Snapshot:** yalnızca `… (TTM)` satırı — payda `favok_ttm`. Bankalarda çeyreklik blokta değer yoktur. |
| `BORC_OZK` | Borç / özsermaye (`debt_equity`); yalnızca **çeyreklik geçmiş**te her dönem ayrı satır. (Eski `BORC` / çift `BORC_OZK` snapshot kaldırıldı.) |
| `EBIT` | Faaliyet kârı (Faiz ve Vergi Öncesi Kâr) — çeyrek (`Q`) ve yuvarlanmış dört çeyrek (`TTM`) için ayrı satırlar. |
| `NET_OP_INC` | Net Faaliyet Kârı (Net Operating Income - EBIT + İştirak ve Yatırım Gelir/Giderleri) — çeyrek (`Q`) ve yuvarlanmış dört çeyrek (`TTM`) için ayrı satırlar. |
| `NCF` | İşletme faaliyetlerinden nakit akışı (`op_cash_flow`) — Q ve TTM. |
| `NET_FX` | Net yabancı para (döviz) pozisyonu; Q/TTM satırlarında aynı bilanço kırılımı. |
| `EXP_RATIO` | İhracatın satışlara oranı (yurtdışı satışlar ÷ satış gelirleri); yüzde olarak. |
| `YAB_END` | Yabancı yatırımcı payı — seçilen dönem sonu tarihindeki oran (İş Yatırım verisi). |
| `YAB_CHG` | Yabancı oranındaki mutlak değişim (aynı API’den gelen `Change`). |
| `PD` | Piyasa değeri — hisse fiyatı × dolaşımdaki pay adedi (`item_2oa` mantığı). |
| `FAVOK` | FAVÖK (faaliyet kârı + amortisman) — Q ve TTM. |

Not: Eski notlarda geçen `DDDD` ifadesi bu projede **`PDDD`** (PD/DD) koduna karşılık gelir; `BORC_OZH` yazımı kodda **`BORC_OZK`** (borç/özkaynak) olarak kullanılır.

- **Config**: **B1** çekilecek çeyrek sayısı + Sinyal penceresi / `pct` paydası; **B2–B3** eşikler.
- **Sinyal**: Güncel değer, medyan, yüzdelik dilim, ucuz/pahalı/nötr ve skor — bellekteki `matrix_rows` + Config ile Python’da üretilir; Excel’deki **Raw_data** ile aynı koşuda doldurulur.

### Sinyal — `%ile` (`pct`) ve UCUZ / PAHALI / NÖTR

Son **analiz_donemi** kadar çeyreğe bakılır; **güncel** değer bu geçmişle kıyaslanır.

**`pct` (tablodaki `%ile`):** Geçmişteki her değer için “**bu, güncelden küçük mü veya eşit mi?**” diye sayılır; çıkan adet, **Config B1**’deki sayıya bölünür. Excel’de çoğunlukla `0,32` gibi 0–1 arası görünür. *(Payda her zaman B1’dir; elde gerçekte daha az çeyrek olsa bile.)*

**Örnek:** B1 = 4, penceredeki PD/DD: `6, 8, 10, 12`, güncel `9`. Dokuzdan küçük veya eşit: `6` ve `8` → 2 adet → **`pct = 2÷4 = 0,50`**.

**Sayı hissi:**

- **`0,76`** örneği: payda 25 ise `19÷25` — 25 gözlemden 19’u güncelin altında veya eşit → güncel geçmişe göre **yüksek** tarafta.
- **`0,32`** örneği: payda 25 ise `8÷25` — az gözlem güncelin altında → güncel geçmişe göre **düşük** tarafta.

PD/DD, FD/FAVÖK ve borç için aynı mantık; her biri kendi `%ile` sütununu üretir.

**Satır sinyali (UCUZ / PAHALI / NÖTR):** Config **B2** (ucuz, varsayılan 25) ve **B3** (pahalı, 75) yüzdeye çevrilir: `pct` **ucuz eşiğinin altındaysa** UCUZ, **pahalı eşiğinin üstündeyse** PAHALI, ikisi arasında NÖTR.

**Genel sinyal:** Üç metrikten bir **skor** hesaplanır; burada “yüksek skor = ucuz” gibi **ters** bir özet kullanılır (tek satırlardaki `pct` kuralının aynısı değildir). Bankalarda FD/FAVÖK ve borç boş; skor yalnızca PD/DD’ye dayanır.

Eski **Facts** sayfası varsa betik çalışırken kaldırılır. **TR_TTM** / **Main_TR_Q** artık yazılmaz; veri **Raw_data** içindedir.

## Çalıştırma

```bash
python isyat.py
```

Çekilecek çeyrek sayısı komut satırından verilmez; **Config `analiz_donemi` (B1)** değerine göre belirlenir (yeni kitapta varsayılan 12). B1’i değiştirip kaydettikten sonra betiği çalıştırın.

Her koşuda betik, **Main** dışında `Raw_data`, `Yabancı Oranı` ve `Sinyal` sayfalarında başlık hariç eski veri ve hücre dolgularını temizleyip yeniden yazar.


İşte kodlarken veya analiz yaparken her iki metriği neden yan yana kullanman gerektiğini hatırlatacak en özet rehber:

**1. `3DF` (Esas Faaliyet Kârı / EBIT)**

* **Ne Ölçer?** Sadece şirketin **ana dükkanının/fabrikasının** kârını (Saf operasyon gücü).
* **Hangi Şirketlerde Önemli?** Üretim ve perakende (Örn: Ford, BİM).

**2. `3H` (Net Faaliyet Kârı / Net Operating Income)**

* **Ne Ölçer?** Ana iş + **İştirak (alt şirket) ve yatırım** kârları.
* **Hangi Şirketlerde Önemli?** Alt şirketlerinden kâr sağlayan holdingler (Örn: Koç Holding, Sabancı Holding).

---

### 💡 Altın Kural: Neden İkisi de Lazım?

İki metriği yan yana koyarak şu iki büyük hatadan kurtulursun:

1. **Holdingleri Yanlış Okumamak:** Holdinglerin ana geliri iştirakleridir. Sadece `3DF` (ana dükkan) bakarsan dev holdingleri zayıf sanırsın. Onları `3H` ile tartmalısın.
2. **Makyajlı Bilançolara Kanmamak:** Bir tekstil şirketinin kendi işi zarar ediyordur (`3DF` negatiftir), ama tek seferlik bir arsa satmıştır (`3H` pozitife fırlar). İkisine birden bakmazsan, o şirketi "kârlı" sanıp tuzağa düşersin.

**Özetin Özeti:** Ana işin ne kadar sağlıklı olduğunu görmek için **`3DF`**'ye, şirketin tüm iştirak yapısıyla cebine koyduğu operasyonel kârı görmek için **`3H`**'ye bak!