# is-yatirim-fetcher

## Excel gereksinimleri

- **Microsoft 365** (veya FILTER / LET / XLOOKUP destekleyen Excel): `Sinyal` sayfasındaki formüller bu işlevlere bağlıdır.
- Çeyrek listesi `Config` sayfasında **F sütununa** yazılır; `analiz_donemi` (B1) pencere genişliğini belirler.

## Sayfa yapısı

- **Facts**: Ticker, Metrik (`PDDD`, `FDFAVOK`, `BORC`), Dönem, Değer (uzun tablo).
- **Config**: Eşikler + F sütununda dönem etiketleri.
- **Sinyal**: `Facts` ve `Config` üzerinden güncel değer, medyan ve yüzdelik dilim.

Eski **Matris_PDDD** / **Matris_FDFAVOK** / **Matris_BORC** sayfaları artık kullanılmıyorsa çalışma kitabından manuel olarak silebilirsiniz; betik bu sayfaları otomatik silmez.

## Çalıştırma

```bash
python isyat.py [çeyrek_sayısı]
```

Varsayılan çeyrek sayısı 12’dir.
