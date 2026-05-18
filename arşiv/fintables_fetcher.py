import cloudscraper
from datetime import datetime, timedelta, timezone
import xlwings as xw # xlwings'i import ediyoruz
import time

# ---------- Fintables Scraper ve Yardımcılar ----------
# Bu fonksiyonları Excel'in çağırmasına gerek yok, 
# ana fonksiyonumuz tarafından dahili olarak kullanılacaklar.

# Scraper'ı modül seviyesinde bir kez oluşturmak daha verimlidir.
scraper = cloudscraper.create_scraper(browser={
    "browser": "chrome", "platform": "windows", "mobile": False
})

# Oturumu başlat
try:
    print("Fintables oturumu başlatılıyor...")
    base_url = "https://fintables.com/radar"
    scraper.get(base_url).raise_for_status()
    print("Oturum başarılı.")
except Exception as e:
    print(f"Oturum başlatılamadı: {e}")
    scraper = None

def ts_utc_now():
    return int(datetime.now(timezone.utc).timestamp())

def fetch_history(scraper, symbol, resolution, from_ts, to_ts):
    data_url = "https://gate.fintables.com/barbar/udf/history"
    params = {
        "symbol": symbol,
        "resolution": str(resolution),
        "from": str(from_ts),
        "to": str(to_ts)
    }
    try:
        r = scraper.get(data_url, params=params)
        r.raise_for_status()
        return r.json()
    except Exception as e:
        print(f"[HATA] fetch_history {symbol} için başarısız: {e}")
        return {"s": "error", "errmsg": str(e)}

def get_last_price_for_symbol(symbol, resolution="5"):
    """
    Belirli bir sembol için son fiyatı Fintables'tan çeker.
    (Sizin genişletme mantığınızı kullanır)
    """
    if not scraper:
        print("Scraper başlatılamadığı için fiyat çekilemiyor.")
        return "SC_HATA"

    to_ts = ts_utc_now()
    
    # 1) İlk deneme: son 30 dk
    from_ts = to_ts - 30 * 60
    data = fetch_history(scraper, symbol, resolution, from_ts, to_ts)
    if data.get("s") == "ok" and data.get("o"):
        print(f"-> {symbol} (30dk) bulundu.")
        return data["o"][-1]

    # 2) Veri yoksa aralığı genişlet
    for widen_hours in (48, 7*24):
        from_ts_wide = to_ts - widen_hours * 3600
        data = fetch_history(scraper, symbol, resolution, from_ts_wide, to_ts)
        if data.get("s") == "ok" and data.get("o"):
            print(f"-> {symbol} ({widen_hours}sa) bulundu.")
            return data["o"][-1]
    
    print(f"-> {symbol} için 7 günde veri bulunamadı.")
    return "BULUNAMADI"

# ---------- EXCEL TARAFINDAN ÇAĞRILACAK ANA FONKSİYON ----------

def guncelle_portfolyo():
    """
    Bu fonksiyon Excel'deki VBA tarafından çağrılacak.
    Aktif sayfadaki A sütununu okur, B sütununu yazar.
    """
    try:
        # 1. Çağıran Excel kitabını ve sayfasını bul
        wb = xw.Book.caller()
        sheet = wb.sheets["source fintables"]
        
        print("Portfolyo güncelleme tetiklendi...")
        
        # 2. Hangi satırdan başlayacağımızı ve sütunları tanımla
        START_ROW = 2
        SYMBOL_COL = 'A'
        PRICE_COL = 'B'

        # 3. A sütunundaki tüm sembolleri bul
        # A2 hücresinden başlayıp aşağı doğru dolu olan son hücreye kadar alır
        range_str = f"{SYMBOL_COL}{START_ROW}"
        semboller = sheet.range(range_str).expand('down').value

        # Eğer sadece tek bir sembol varsa (A2'de), 'semboller' string olur.
        # Bunu bir listeye çevirmemiz gerekir.
        if isinstance(semboller, str):
            semboller = [semboller]
            
        print(f"Bulunan semboller: {semboller}")

        # 4. Her sembol için döngüye gir
        for i, symbol in enumerate(semboller):
            if not symbol:  # Boş hücreyi atla
                continue
                
            print(f"--- İşleniyor: {symbol} ---")
            
            # Güncellenecek fiyat hücresini belirle
            price_cell = sheet.range(f"{PRICE_COL}{START_ROW + i}")
            
            # Hücreye "Yükleniyor..." yaz (Kullanıcı görsün)
            price_cell.value = "Yükleniyor..."
            
            # Fiyatı çek
            last_price = get_last_price_for_symbol(symbol)
            
            # Fiyatı hücreye yaz
            price_cell.value = last_price
            print(f"-> {symbol} fiyatı güncellendi: {last_price}")
            
            # (API'yi yormamak için çok kısa bir bekleme eklenebilir)
            # time.sleep(0.1) 

        print("Güncelleme tamamlandı.")
        
    except Exception as e:
        # Hata olursa Excel'deki C1 hücresine yaz (opsiyonel)
        try:
            wb = xw.Book.caller()
            wb.sheets.active.range('C1').value = f"PYTHON HATA: {e}"
        except:
            pass # Hata raporlama bile başarısız olabilir
        print(f"Bir hata oluştu: {e}")

# Bu if bloğu, script'in doğrudan çalıştırılması (test için)
# ile Excel tarafından 'import' edilmesi arasındaki farkı yönetir.
if __name__ == "__main__":
    # Test için xlwings'in bir Excel'e bağlanmasını sağla
    # Bu satır, VBA olmadan test yapmanızı sağlar
    # Çalıştırmadan önce 'portfolio.xlsm' dosyanızın açık olduğundan emin olun
    try:
        xw.Book("portfolio.xlsm").set_mock_caller()
        guncelle_portfolyo()
    except Exception as e:
        print(f"Test çalıştırması başarısız. Excel dosyanızın açık olduğundan emin misiniz? Hata: {e}")