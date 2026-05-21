import pandas as pd
import requests
import time
import concurrent.futures
import re

MAX_WORKERS = 50

def get_matriks_token():
    """
    Belirtilen sayfaya tek bir GET isteği atarak HTML kaynağını çeker, 
    içerisindeki 'var token' satırını Regex ile bulur ve JWT formatında döndürür.
    """
    url = "https://www.matriksdata.com/website/temel-analiz-raporu"
    
    headers = {
        "accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8",
        "accept-language": "en-TR,en-US;q=0.9,en-GB;q=0.8,en;q=0.7,tr;q=0.6",
        "cache-control": "no-cache",
        "pragma": "no-cache",
        "user-agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/148.0.0.0 Safari/537.36",
        "upgrade-insecure-requests": "1"
    }
    
    try:
        response = requests.get(url, headers=headers, timeout=15)
        response.raise_for_status()
        
        # HTML kaynağı içinde "var token = 'eyJ...';" yapısını arıyoruz
        match = re.search(r"var token\s*=\s*'([^']+)'", response.text)
        
        if match:
            # Token'ı alıp başına 'jwt ' ekliyoruz
            extracted_token = match.group(1)
            return f"jwt {extracted_token}"
        else:
            print("Hata: Kaynak kodda token bulunamadı.")
            return None
            
    except requests.exceptions.RequestException as e:
        print(f"Token alınırken ağ hatası oluştu: {e}")
        return None

def fetch_financial_data(symbol, headers):
    """
    Belirtilen sembol için API'den finansal verileri çeker.
    Headers argümanı, ana fonksiyonda oluşturulan dinamik token'ı içerir.
    """
    url = f"https://api.matriksdata.com/dumrul/v1/fundamental-dashboard?symbol={symbol}&unadjusted=false&currency=TRY"
    
    try:
        response = requests.get(url, headers=headers, timeout=10)
        response.raise_for_status()
        return response.json()
    except requests.exceptions.RequestException as e:
        return {"error": str(e)}

def parse_and_flatten_data(symbol, raw_response):
    if not raw_response:
        return {"Sembol": symbol, "Durum": "Boş Yanıt"}
        
    if "error" in raw_response:
        return {"Sembol": symbol, "Durum": f"Ağ Hatası: {raw_response['error']}"}
        
    row_data = {"Sembol": symbol, "Durum": "Başarılı"}
    
    for key, value in raw_response.items():
        if key == "symbol":
            continue
            
        if key == "periodicData" and isinstance(value, list):
            for period_item in value:
                period_val = period_item.get("period", "BilinmeyenDonem")
                for sub_key, sub_value in period_item.items():
                    if sub_key != "period":
                        column_name = f"{period_val}_{sub_key}"
                        row_data[column_name] = sub_value
        else:
            row_data[key] = value
            
    return row_data

def process_single_symbol(symbol, headers):
    """Her thread'in çalıştıracağı ana birim."""
    raw_data = fetch_financial_data(symbol, headers)
    parsed_data = parse_and_flatten_data(symbol, raw_data)
    return parsed_data

def main():
    input_file = "a.xlsx"
    output_file = "finansal_sonuclar_matriks.xlsx"
    
    print(f"'{input_file}' dosyası okunuyor...")
    try:
        df_input = pd.read_excel(input_file)
        symbols = df_input.iloc[:, 0].dropna().astype(str).str.strip().unique().tolist()
    except Exception as e:
        print(f"Excel okuma hatası: {e}")
        return

    total_symbols = len(symbols)
    print(f"Toplam {total_symbols} adet sembol bulundu.")
    
    # 1. Aşama: Sadece BİR KERE yetkilendirme token'ını al
    print("\n[1/2] Site üzerinden güncel token alınıyor...")
    auth_token = get_matriks_token()
    
    if not auth_token:
        print("İşlem iptal ediliyor. Token alınamadığı için API'ye erişilemez.")
        return
        
    print(f"Token başarıyla alındı! (Önizleme: {auth_token[:20]}...)\n")
    
    # Tüm API isteklerinde kullanılacak olan dinamik başlıklar
    api_headers = {
        "accept": "*/*",
        "accept-language": "en-TR,en-US;q=0.9,en-GB;q=0.8,en;q=0.7,tr;q=0.6",
        "authorization": auth_token,  # Yakaladığımız token'ı buraya ekledik
        "cache-control": "no-cache",
        "pragma": "no-cache",
        "user-agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/148.0.0.0 Safari/537.36",
    }
    
    # 2. Aşama: Multithread ile API'den verileri çek
    print(f"[2/2] {MAX_WORKERS} thread ile veri çekimi başlıyor...\n")
    
    all_results = []
    start_time = time.time()
    
    with concurrent.futures.ThreadPoolExecutor(max_workers=MAX_WORKERS) as executor:
        # process_single_symbol fonksiyonuna symbol ve api_headers'ı yolluyoruz
        futures = {executor.submit(process_single_symbol, symbol, api_headers): symbol for symbol in symbols}
        
        completed_count = 0
        for future in concurrent.futures.as_completed(futures):
            completed_count += 1
            symbol = futures[future]
            
            try:
                result = future.result()
                all_results.append(result)
                print(f"İlerleme: [{completed_count}/{total_symbols}] - {symbol} çekildi. Durum: {result['Durum']}")
            except Exception as e:
                print(f"\nBeklenmeyen Hata ({symbol}): {e}")
                all_results.append({"Sembol": symbol, "Durum": f"Beklenmeyen Hata: {e}"})

    df_output = pd.DataFrame(all_results)
    
    cols = ['Sembol', 'Durum'] + [c for c in df_output.columns if c not in ['Sembol', 'Durum']]
    df_output = df_output.reindex(columns=cols)
    
    df_output.to_excel(output_file, index=False)
    
    elapsed_time = time.time() - start_time
    print(f"\nİşlem tamamlandı! Toplam süre: {elapsed_time:.2f} saniye.")
    print(f"Veriler '{output_file}' dosyasına kaydedildi.")

if __name__ == "__main__":
    main()