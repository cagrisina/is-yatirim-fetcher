import requests
import json

def test_isyatirim_api_v2():
    url = "https://www.isyatirim.com.tr/_layouts/15/IsYatirim.Website/Common/Data.aspx/MaliTablo"
    
    # Sunucunun çökmemesi için tam 4 dönemi eksiksiz gönderiyoruz
    params = {
        "companyCode": "KCHOL",
        "exchange": "TRY",
        "financialGroup": "XI_29",
        "year1": "2026",
        "period1": "3",
        "year2": "2025",
        "period2": "12",
        "year3": "2025",
        "period3": "9",
        "year4": "2025",
        "period4": "6"
    }
    
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36",
        "Accept": "application/json, text/javascript, */*; q=0.01",
        "X-Requested-With": "XMLHttpRequest"
    }

    print("İş Yatırım API'sine eksiksiz (4 çeyrek) test isteği atılıyor...")
    
    try:
        response = requests.get(url, params=params, headers=headers, timeout=10)
        print(f"Sunucu Yanıt Kodu: {response.status_code}")
        
        if response.status_code == 200:
            data = response.json()
            
            # Gelen yanıtı incelemek için kaydediyoruz
            with open("isyatirim_test_sonuc.json", "w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False, indent=4)
                
            print("\n✅ Veri başarıyla çekildi ve 'isyatirim_test_sonuc.json' dosyasına kaydedildi.\n")
            
            # Ekrana ilk 3 kalemi yazdıralım ki yapıya bakalım
            if "value" in data and len(data["value"]) > 0:
                print("--- GELEN VERİNİN YAPISI (İLK 3 SATIR) ---")
                for item in data["value"][:3]:
                    print(item)
                print("------------------------------------------")
            else:
                print("Hata: JSON döndü ama 'value' anahtarı boş!")
        else:
            print(f"Hata detayı: {response.text}")
            
    except Exception as e:
        print(f"Bağlantı hatası: {e}")

if __name__ == "__main__":
    test_isyatirim_api_v2()