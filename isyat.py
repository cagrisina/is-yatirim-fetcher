import os
import re
import sys
import time
import requests
from requests.adapters import HTTPAdapter
import pandas as pd
from pandas.tseries.offsets import BDay
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
import concurrent.futures
import datetime
import yfinance as yf

# ==============================================================================
# 0. PAYLAŞIMLI HTTP SESSION (BÜYÜK CONNECTION POOL)
# ==============================================================================
# Varsayılan requests.get() host başına sadece 10 bağlantı açar.
# 50 thread ile paralel çalışabilmek için pool boyutunu büyütüyoruz.
_session = requests.Session()
_adapter = HTTPAdapter(pool_connections=200, pool_maxsize=200, max_retries=1)
_session.mount("https://", _adapter)
_session.mount("http://", _adapter)

# ==============================================================================
# 1. ORTAK YARDIMCI FONKSİYONLAR (TARİH & DÖNEM)
# ==============================================================================

def build_ttm_periods(y, m):
    if m == 12: return [(y, 12), (y, 9), (y, 6), (y, 3)], f"{y}/12", False 
    else:
        filler_m = m - 3 if m > 3 else 12
        filler_y = y if m > 3 else y - 1
        return [(y, m), (y-1, 12), (y-1, m), (filler_y, filler_m)], f"{y}/{m:02d}", True

def get_isyatirim_periods(period_str):
    p_str = str(period_str).strip().upper()
    if p_str.isdigit() and len(p_str) == 4: return build_ttm_periods(int(p_str), 12)
    if "Q" in p_str:
        try:
            parts = p_str.split("Q")
            y, m = int(re.sub(r'\D', '', parts[0])), int(re.sub(r'\D', '', parts[1])) * 3
            return build_ttm_periods(y, m)
        except: pass
    if "/" in p_str:
        try:
            parts = p_str.split("/")
            y, m = int(re.sub(r'\D', '', parts[0])), int(re.sub(r'\D', '', parts[1].split()[0]))
            if m in [3, 6, 9, 12]: return build_ttm_periods(y, m)
        except: pass
    return None, None, False


# ==============================================================================
# 2. TÜRKİYE (TR) MOTORU (SAF SAYILARLA)
# ==============================================================================

def fetch_tr_raw_data(symbol, period_val=None):
    clean_symbol = symbol.split('.')[0].strip().upper()
    url = "https://www.isyatirim.com.tr/_layouts/15/IsYatirim.Website/Common/Data.aspx/MaliTablo"
    headers = {"User-Agent": "Mozilla/5.0", "Accept": "application/json", "X-Requested-With": "XMLHttpRequest"}
    fin_groups = ["XI_29", "UFRS_K", "UFRS"] 

    if period_val is not None:
        y, q = period_val
        attempts = 1
    else:
        now = datetime.datetime.now()
        y, m = now.year, now.month
        q = 9 if m >= 11 else 6 if m >= 8 else 3 if m >= 5 else 12
        if q == 12: y -= 1
        attempts = 4
    
    for attempt in range(attempts):
        periods_list, resolved_name, is_partial = build_ttm_periods(y, q)
        for group in fin_groups:
            params = {"companyCode": clean_symbol, "exchange": "TRY", "financialGroup": group}
            for i, (py, pm) in enumerate(periods_list, start=1):
                params[f"year{i}"] = str(py); params[f"period{i}"] = str(pm)
            try:
                res = _session.get(url, params=params, headers=headers, timeout=10)
                if res.status_code == 200:
                    data = res.json()
                    if "value" in data and data["value"]:
                        df = pd.DataFrame(data["value"])
                        if not df[df['itemDescTr'].str.contains("Özkaynak|Sermaye", case=False, na=False)].empty:
                            ret_name = resolved_name if period_val is not None else f"{resolved_name} (En Güncel)"
                            return df, ret_name, is_partial, group
            except: pass
        if period_val is None:
            q -= 3
            if q <= 0: q = 12; y -= 1
    return None, "Bulunamadı", False, None

def extract_tr_value(df, keyword_contains, regex=False):
    v1, v2, v3, v4 = 0.0, 0.0, 0.0, 0.0
    mask = df['itemDescTr'].fillna('').str.contains(keyword_contains, case=False, na=False, regex=regex)
    if not df[mask].empty:
        row = df[mask].iloc[0]
        try: v1 = float(row.get('value1') or 0)
        except: pass
        try: v2 = float(row.get('value2') or 0)
        except: pass
        try: v3 = float(row.get('value3') or 0)
        except: pass
        try: v4 = float(row.get('value4') or 0)
        except: pass
    return v1, v2, v3, v4

def parse_tr_data(df, resolved_period_name, is_partial, group):
    if df is None or df.empty: return None
    is_bank = (group in ["UFRS", "UFRS_K"])

    try: period_num = int(resolved_period_name.split('/')[1].split()[0])
    except: period_num = 12

    def calc_ttm_and_q(v1, v2, v3, v4):
        ttm = (v1 + v2 - v3) if is_partial else v1
        if period_num == 12: q = v1 - v2
        elif period_num == 3: q = v1
        else: q = v1 - v4
        return ttm, q

    item_2oa_v1 = 0.0
    if 'itemCode' in df.columns:
        mask_2oa = df['itemCode'].fillna('').str.upper() == '2OA'
        if not df[mask_2oa].empty:
            try: item_2oa_v1 = float(df[mask_2oa].iloc[0].get('value1') or 0)
            except: pass

    ozk_v1, _, _, _ = extract_tr_value(df, "Ana Ortaklığa Ait Özkaynaklar")
    if ozk_v1 == 0: ozk_v1, _, _, _ = extract_tr_value(df, "Özkaynaklar")

    if is_bank:
        op_inc_v1, op_inc_v2, op_inc_v3, op_inc_v4 = extract_tr_value(df, "NET FAALİYET KARI/ZARARI")
        op_inc_ttm, op_inc_q = calc_ttm_and_q(op_inc_v1, op_inc_v2, op_inc_v3, op_inc_v4)
        return {
            "debt_equity": "", 
            "op_income_ttm": op_inc_ttm,
            "op_income_q": op_inc_q,
            "op_cash_flow_ttm": "", 
            "op_cash_flow_q": "", 
            "net_fx": "", 
            "export_ratio_ttm": "", 
            "export_ratio_q": "", 
            "item_2oa": item_2oa_v1,
            "ozk": ozk_v1,
            "net_borc": "",
            "favok_ttm": "",
            "favok_q": "",
            "resolved_period": resolved_period_name, "is_bank": True
        }
    else:
        def get_val_by_code(code_str):
            v1, v2, v3, v4 = 0.0, 0.0, 0.0, 0.0
            if 'itemCode' not in df.columns: return v1, v2, v3, v4
            mask = df['itemCode'].fillna('').astype(str).str.strip().str.upper() == code_str
            if not df[mask].empty:
                row = df[mask].iloc[0]
                try: v1 = float(row.get('value1') or 0)
                except: pass
                try: v2 = float(row.get('value2') or 0)
                except: pass
                try: v3 = float(row.get('value3') or 0)
                except: pass
                try: v4 = float(row.get('value4') or 0)
                except: pass
            return v1, v2, v3, v4

        kv_borc_v1, _, _, _ = get_val_by_code("2AA")
        uv_borc_v1, _, _, _ = get_val_by_code("2BA")
        borc_v1 = kv_borc_v1 + uv_borc_v1
        
        nakit_v1, _, _, _ = get_val_by_code("1AA")
        fin_yat_kisa_v1, _, _, _ = get_val_by_code("1AB")
        fin_yat_uzun_v1, _, _, _ = get_val_by_code("1BC")
        
        net_borc = borc_v1 - (nakit_v1 + fin_yat_kisa_v1 + fin_yat_uzun_v1)
        
        op_inc_v1, op_inc_v2, op_inc_v3, op_inc_v4 = get_val_by_code("3DF")
        if op_inc_v1 == 0 and op_inc_v2 == 0 and op_inc_v3 == 0:
            op_inc_v1, op_inc_v2, op_inc_v3, op_inc_v4 = extract_tr_value(df, r"^\s*FAALİYET KARI\s*(?:\(ZARARI\))?$", regex=True)
        op_inc_ttm, op_inc_q = calc_ttm_and_q(op_inc_v1, op_inc_v2, op_inc_v3, op_inc_v4)
        
        amort_v1, amort_v2, amort_v3, amort_v4 = get_val_by_code("4B")
        if amort_v1 == 0 and amort_v2 == 0 and amort_v3 == 0:
            amort_v1, amort_v2, amort_v3, amort_v4 = get_val_by_code("4CAB")
        amort_ttm, amort_q = calc_ttm_and_q(amort_v1, amort_v2, amort_v3, amort_v4)
        
        favok_ttm = op_inc_ttm + amort_ttm
        favok_q = op_inc_q + amort_q
        
        net_fx, _, _, _ = extract_tr_value(df, "Net Yabancı Para Pozisyonu")
        
        cf_v1, cf_v2, cf_v3, cf_v4 = extract_tr_value(df, "İşletme Faaliyetlerinden Kaynaklanan")
        cf_ttm, cf_q = calc_ttm_and_q(cf_v1, cf_v2, cf_v3, cf_v4)
        
        ihracat_v1, ihracat_v2, ihracat_v3, ihracat_v4 = extract_tr_value(df, "Yurtdışı Satışlar")
        ihracat_ttm, ihracat_q = calc_ttm_and_q(ihracat_v1, ihracat_v2, ihracat_v3, ihracat_v4)
        
        hasilat_v1, hasilat_v2, hasilat_v3, hasilat_v4 = extract_tr_value(df, "Satış Gelirleri")
        hasilat_ttm, hasilat_q = calc_ttm_and_q(hasilat_v1, hasilat_v2, hasilat_v3, hasilat_v4)
        
        export_ratio_ttm = (ihracat_ttm / hasilat_ttm) if hasilat_ttm != 0 else 0.0
        export_ratio_q = (ihracat_q / hasilat_q) if hasilat_q != 0 else 0.0

        return {
            "debt_equity": (borc_v1 / ozk_v1) if ozk_v1 != 0 else "N/A",
            "op_income_ttm": op_inc_ttm, 
            "op_income_q": op_inc_q, 
            "op_cash_flow_ttm": cf_ttm, 
            "op_cash_flow_q": cf_q, 
            "net_fx": net_fx,
            "export_ratio_ttm": export_ratio_ttm,
            "export_ratio_q": export_ratio_q,
            "item_2oa": item_2oa_v1,
            "ozk": ozk_v1,
            "net_borc": net_borc,
            "favok_ttm": favok_ttm,
            "favok_q": favok_q,
            "resolved_period": resolved_period_name, "is_bank": False
        }



# ==============================================================================
# 4. YABANCI TAKAS ORANI MOTORU
# ==============================================================================

def get_yabanci_oran_dates(test_symbol):
    today = pd.Timestamp.today().normalize()
    
    valid_end = None
    for i in range(6):
        test_end = today - pd.Timedelta(days=i)
        end_str = test_end.strftime("%d-%m-%Y")
        res = fetch_yabanci_oran(test_symbol, end_str, end_str)
        if res.get("end", "") != "":
            valid_end = test_end
            break
            
    if valid_end is None:
        valid_end = today - BDay(1)
        
    valid_start = None
    base_start = today - pd.Timedelta(days=30)
    for i in range(6):
        test_start = base_start + pd.Timedelta(days=i)
        start_str = test_start.strftime("%d-%m-%Y")
        res = fetch_yabanci_oran(test_symbol, start_str, start_str)
        if res.get("end", "") != "":
            valid_start = test_start
            break
            
    if valid_start is None:
        valid_start = base_start
        
    return valid_start.strftime("%d-%m-%Y"), valid_end.strftime("%d-%m-%Y")

def fetch_yabanci_oran(symbol, start_date, end_date):
    clean_symbol = symbol.split('.')[0].strip().upper()
    url = "https://www.isyatirim.com.tr/_layouts/15/IsYatirim.Website/StockInfo/CompanyInfoAjax.aspx/GetYabanciOranlarXHR"
    payload = {"baslangicTarih": start_date, "bitisTarihi": end_date, "sektor": None, "endeks": "09", "hisse": clean_symbol}
    headers = {"accept": "application/json", "content-type": "application/json; charset=UTF-8", "x-requested-with": "XMLHttpRequest"}

    try:
        res = _session.post(url, json=payload, headers=headers, timeout=10)
        if res.status_code == 200 and "d" in res.json() and len(res.json()["d"]) > 0:
            item = res.json()["d"][0]
            # Değerleri % formatında Excel'de gösterebilmek için 100'e bölüyoruz
            return {"start": item.get("YAB_ORAN_START", 0)/100, "end": item.get("YAB_ORAN_END", 0)/100, "change": item.get("DEGISIM", 0)/100, "effect": item.get("ETKI", 0)/100, "price": item.get("PRICE_TL", 0)}
    except: pass
    return {"start": "", "end": "", "change": "", "effect": "", "price": ""}


# ==============================================================================
# 5. MULTITHREAD GÖREV YÖNETİCİLERİ
# ==============================================================================

def task_tr_stock(row_idx, symbol):
    df_raw, resolved_p, is_partial, group = fetch_tr_raw_data(symbol) # period_val ignored
    result = parse_tr_data(df_raw, resolved_p, is_partial, group)
    if result is None: result = {k: "" for k in ['debt_equity', 'op_income_ttm', 'op_income_q', 'op_cash_flow_ttm', 'op_cash_flow_q', 'net_fx', 'export_ratio_ttm', 'export_ratio_q', 'resolved_period', 'item_2oa', 'ozk', 'net_borc', 'favok_ttm', 'favok_q']}
    
    try:
        info = yf.Ticker(symbol + ".IS").info
        result['industry'] = info.get('industry', '')
    except Exception:
        result['industry'] = ''
        
    return {'row_idx': row_idx, 'symbol': symbol, 'data': result}

def task_yabanci_oran(row_idx, symbol, date_a, date_b):
    return {'row_idx': row_idx, 'symbol': symbol, 'data': fetch_yabanci_oran(symbol, date_a, date_b)}

def task_historical_q(row_idx, symbol, periods):
    results = []
    df_raw, resolved_p, is_partial, group = fetch_tr_raw_data(symbol)
    if df_raw is None:
        return {'row_idx': row_idx, 'symbol': symbol, 'data': []}
    
    def get_hist_price(t_sym, year, month):
        try:
            next_month = month + 1 if month < 12 else 1
            next_year = year if month < 12 else year + 1
            data = yf.Ticker(f"{t_sym}.IS").history(start=f"{year}-{month:02d}-01", end=f"{next_year}-{next_month:02d}-01")
            if not data.empty: return float(data['Close'].iloc[-1])
        except: pass
        return ""

    def get_hist_yabanci(t_sym, year, month):
        import pandas as pd
        if month in [1, 3, 5, 7, 8, 10, 12]: day = 31
        elif month in [4, 6, 9, 11]: day = 30
        else: day = 29 if year % 4 == 0 else 28
        
        base_date = pd.Timestamp(year=year, month=month, day=day)
        
        for i in range(7):
            test_date = base_date - pd.Timedelta(days=i)
            d_str = test_date.strftime("%d-%m-%Y")
            res = fetch_yabanci_oran(t_sym, d_str, d_str)
            if res.get("end", "") != "":
                return res.get("end")
        return ""

    first_res = parse_tr_data(df_raw, resolved_p, is_partial, group)
    
    try:
        parts = resolved_p.split('/')[0:2]
        y = int(re.sub(r'\D', '', parts[0]))
        q = int(re.sub(r'\D', '', parts[1].split()[0]))
    except:
        if first_res: results.append(first_res)
        return {'row_idx': row_idx, 'symbol': symbol, 'data': results}
    
    if first_res:
        first_res['historical_price'] = get_hist_price(symbol, y, q)
        first_res['historical_yabanci'] = get_hist_yabanci(symbol, y, q)
        results.append(first_res)
    
    for _ in range(periods - 1):
        q -= 3
        if q <= 0:
            q = 12
            y -= 1
        
        df_hist, resolved_h, partial_h, group_h = fetch_tr_raw_data(symbol, period_val=(y, q))
        if df_hist is not None:
            res_hist = parse_tr_data(df_hist, resolved_h, partial_h, group_h)
            if res_hist:
                res_hist['historical_price'] = get_hist_price(symbol, y, q)
                res_hist['historical_yabanci'] = get_hist_yabanci(symbol, y, q)
                results.append(res_hist)
    
    return {'row_idx': row_idx, 'symbol': symbol, 'data': results}


# ==============================================================================
# 6. ORKESTRASYON VE EXCEL ANA MOTORU (BELLEK ODAKLI)
# ==============================================================================

def main_automation():
    file_name = "analysis.xlsx"
    sheet_tr_name, sheet_yab_name = "Main_TR", "Yabancı Oranı"
    if not os.path.exists(file_name): print(f"Hata: '{file_name}' dosyası bulunamadı!"); return
    wb = load_workbook(file_name)

    # İşlemlere başlamadan önce tüm hücrelerin arka plan rengini (highlight) temizle
    clear_fill = PatternFill(fill_type=None)
    for sheet_name in wb.sheetnames:
        for row in wb[sheet_name].iter_rows():
            for cell in row:
                cell.fill = clear_fill

    tr_tasks = []
    hist_q_tasks = []
    date_a, date_b = "01-01-2024", "01-01-2024" # Varsayılan (eğer sayfa yoksa)
    
    if "analyze_input" in wb.sheetnames:
        ws_input = wb["analyze_input"]
        for row in range(2, ws_input.max_row + 1):
            hedef = str(ws_input.cell(row=row, column=1).value or "").strip()
            ticker = str(ws_input.cell(row=row, column=2).value or "").strip()
            donem = ws_input.cell(row=row, column=3).value
            if hedef == "Historical_Q" and ticker and donem:
                try: donem_int = int(donem)
                except: donem_int = 0
                if donem_int > 0:
                    hist_q_tasks.append((row, ticker, donem_int))

    # TR Hisse Listesi
    sheet_tr_name = "Main_TR_Q" if "Main_TR_Q" in wb.sheetnames else "Main_TR"
    if sheet_tr_name in wb.sheetnames:
        ws_tr = wb[sheet_tr_name]
        for row in range(2, ws_tr.max_row + 1):
            symbol = str(ws_tr.cell(row=row, column=1).value or "").strip()
            if symbol: tr_tasks.append((row, symbol))
        
        if sheet_yab_name not in wb.sheetnames: wb.create_sheet(sheet_yab_name)
        ws_yab = wb[sheet_yab_name]
        if tr_tasks:
            test_symbol = tr_tasks[0][1]
            print(f"Yabancı oranı için geçerli iş günleri aranıyor (Referans: {test_symbol})...")
            date_a, date_b = get_yabanci_oran_dates(test_symbol)
        else:
            date_a, date_b = "01-01-2024", "01-01-2024"

    print(f"\n🚀 Belleğe Alınıyor... TR: {len(tr_tasks)} şirket")

    # ==============================================================
    # 2 AYRI MOTOR — HER API KENDİ THREAD HAVUZUNU KULLANIR
    # ==============================================================
    
    _t_start = time.perf_counter()

    executor_tr  = concurrent.futures.ThreadPoolExecutor(max_workers=50, thread_name_prefix="TR-Mali")
    executor_yab = concurrent.futures.ThreadPoolExecutor(max_workers=50, thread_name_prefix="Yabanci")
    executor_hist = concurrent.futures.ThreadPoolExecutor(max_workers=10, thread_name_prefix="Hist-Q")

    # Her motor kendi görevlerini alır
    futures_tr  = {executor_tr.submit(task_tr_stock, r, sym): ('TR', sym) for r, sym in tr_tasks}
    futures_yab = {executor_yab.submit(task_yabanci_oran, r, sym, date_a, date_b): ('YAB', sym) for r, sym in tr_tasks}
    futures_hist = {executor_hist.submit(task_historical_q, r, sym, p): ('HIST_Q', sym) for r, sym, p in hist_q_tasks}

    all_futures = {}
    all_futures.update(futures_tr)
    all_futures.update(futures_yab)
    all_futures.update(futures_hist)

    total_tr, total_yab, total_hist = len(futures_tr), len(futures_yab), len(futures_hist)
    done_tr, done_yab, done_hist = 0, 0, 0
    total_all = total_tr + total_yab + total_hist
    done_all = 0
    last_sym = ""

    def _bar(done, total, width=15):
        filled = int(width * done / total) if total > 0 else width
        return f"[{'█' * filled}{'░' * (width - filled)}]"

    def _print_progress():
        elapsed = time.perf_counter() - _t_start
        pct = (done_all / total_all * 100) if total_all > 0 else 100
        line = (
            f"\r  ⏳ {pct:5.1f}% {_bar(done_all, total_all, 20)} "
            f"│ TR Mali: {done_tr}/{total_tr} "
            f"│ Yabancı: {done_yab}/{total_yab} "
            f"│ Hist-Q: {done_hist}/{total_hist} "
            f"│ ⏱ {elapsed:.1f}s "
            f"│ Son: {last_sym:<10}"
        )
        sys.stdout.write(line)
        sys.stdout.flush()

    _print_progress()

    tr_fin_results_map, tr_yab_results_map, hist_q_results_map = {}, {}, {}

    for future in concurrent.futures.as_completed(all_futures):
        task_type, sym = all_futures[future]
        result = future.result()
        last_sym = sym

        if task_type == 'TR':
            tr_fin_results_map[sym] = result
            done_tr += 1
        elif task_type == 'YAB':
            tr_yab_results_map[sym] = result
            done_yab += 1
        elif task_type == 'HIST_Q':
            hist_q_results_map[sym] = result
            done_hist += 1

        done_all += 1
        _print_progress()

    executor_tr.shutdown(wait=False)
    executor_yab.shutdown(wait=False)
    executor_hist.shutdown(wait=False)

    tr_fin_results = [tr_fin_results_map[sym] for _, sym in tr_tasks]
    tr_yab_results = [tr_yab_results_map[sym] for _, sym in tr_tasks]
    hist_q_results = [hist_q_results_map[sym] for _, sym, _ in hist_q_tasks]

    elapsed_total = time.perf_counter() - _t_start
    print(f"\n\n✅ Tüm veriler RAM'e çekildi ({elapsed_total:.1f}s). Excel'e formatlanarak yazılıyor...\n")

    # ==============================================================
    # FORMATLAYARAK EXCEL'E TEK SEFERDE YAZ (WRITE-LOCK)
    # ==============================================================
    green_fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
    no_fill = PatternFill(fill_type=None)
    
    yab_dict = {res['symbol']: res['data'] for res in tr_yab_results}

    sheet_tr_ttm_name = "TR_TTM"
    sheet_tr_q_name = "Main_TR_Q"

    if sheet_tr_ttm_name not in wb.sheetnames: wb.create_sheet(sheet_tr_ttm_name)
    if sheet_tr_q_name not in wb.sheetnames: wb.create_sheet(sheet_tr_q_name)

    ws_ttm = wb[sheet_tr_ttm_name]
    ws_q = wb[sheet_tr_q_name]

    headers_ttm = {
        1: "Ticker", 2: "Sektör", 3: "Dönem", 4: "Borç / Özsermaye", 5: "Faaliyet Kârı (TTM)", 
        6: "Nakit Akışı (TTM)", 7: "Net Döviz Pozisyonu", 8: "İhracat Geliri (TTM)", 
        9: "Yabancı Oranı", 10: "Yabancı Oranı Değişim",
        11: "Piyasa Değeri", 12: "PD/DD",
        13: "FAVÖK (TTM)", 14: "FD / FAVÖK (TTM)"
    }
    headers_q = {
        1: "Ticker", 2: "Sektör", 3: "Dönem", 4: "Borç / Özsermaye", 5: "Faaliyet Kârı (Q)", 
        6: "Nakit Akışı (Q)", 7: "Net Döviz Pozisyonu", 8: "İhracat Geliri (Q)", 
        9: "Yabancı Oranı", 10: "Yabancı Oranı Değişim",
        11: "Piyasa Değeri", 12: "PD/DD",
        13: "FAVÖK (Q)", 14: "FD / FAVÖK (Yıllıklandırılmış Q)"
    }

    for col_num, text in headers_ttm.items(): ws_ttm.cell(row=1, column=col_num, value=text)
    for col_num, text in headers_q.items(): ws_q.cell(row=1, column=col_num, value=text)

    for res in tr_fin_results:
        r, sym, d = res['row_idx'], res['symbol'], res['data']
        is_b = d.get('is_bank', False)
        
        yab_end = yab_dict[sym]['end'] if sym in yab_dict else ""
        yab_change = yab_dict[sym]['change'] if sym in yab_dict else ""
        price = yab_dict[sym]['price'] if sym in yab_dict else ""
        item_2oa = d.get('item_2oa', "")
        ozk = d.get('ozk', "")
        net_borc = d.get('net_borc', "")
        
        piyasa_degeri = ""
        pddd = ""
        if isinstance(price, (int, float)) and isinstance(item_2oa, (int, float)):
            piyasa_degeri = price * item_2oa
            if isinstance(ozk, (int, float)) and ozk != 0:
                pddd = piyasa_degeri / ozk

        def write_row(ws, suffix, is_q=False):
            ws.cell(row=r, column=1, value=sym)
            ws.cell(row=r, column=2, value=d.get('industry', ''))
            ws.cell(row=r, column=3, value=d['resolved_period'])
            
            c4 = ws.cell(row=r, column=4, value=d['debt_equity'])
            if isinstance(c4.value, (int, float)): c4.number_format = '0.00'
            
            c5 = ws.cell(row=r, column=5, value=d.get(f'op_income_{suffix}', ''))
            if isinstance(c5.value, (int, float)): c5.number_format = '#,##0'
            
            c6 = ws.cell(row=r, column=6, value=d.get(f'op_cash_flow_{suffix}', ''))
            if isinstance(c6.value, (int, float)): c6.number_format = '#,##0'
            
            c7 = ws.cell(row=r, column=7, value=d.get('net_fx', ''))
            if isinstance(c7.value, (int, float)): c7.number_format = '#,##0'
            
            c8 = ws.cell(row=r, column=8, value=d.get(f'export_ratio_{suffix}', ''))
            if isinstance(c8.value, (int, float)): c8.number_format = '0.00%'
            
            c9 = ws.cell(row=r, column=9, value=yab_end)
            if isinstance(c9.value, (int, float)): c9.number_format = '0.00%'
            
            c10 = ws.cell(row=r, column=10, value=yab_change)
            if isinstance(c10.value, (int, float)): c10.number_format = '0.00%'

            c11 = ws.cell(row=r, column=11, value=piyasa_degeri)
            if isinstance(c11.value, (int, float)): c11.number_format = '#,##0'

            c12 = ws.cell(row=r, column=12, value=pddd)
            if isinstance(c12.value, (int, float)): c12.number_format = '0.00'

            favok = d.get(f'favok_{suffix}', "")
            c13 = ws.cell(row=r, column=13, value=favok)
            if isinstance(c13.value, (int, float)): c13.number_format = '#,##0'

            fd_favok = ""
            if isinstance(piyasa_degeri, (int, float)) and isinstance(net_borc, (int, float)):
                fd = piyasa_degeri + net_borc
                if isinstance(favok, (int, float)) and favok != 0:
                    if is_q:
                        fd_favok = fd / (favok * 4)
                    else:
                        fd_favok = fd / favok
            
            c14 = ws.cell(row=r, column=14, value=fd_favok)
            if isinstance(c14.value, (int, float)): c14.number_format = '0.00'

            c4.fill = green_fill if not is_b and isinstance(c4.value, (int, float)) and c4.value < 1.0 else no_fill
            c5.fill = green_fill if isinstance(c5.value, (int, float)) and c5.value > 0 else no_fill
            c6.fill = green_fill if isinstance(c6.value, (int, float)) and c6.value > 0 else no_fill
            c7.fill = green_fill if not is_b and isinstance(c7.value, (int, float)) and c7.value > 0 else no_fill

        write_row(ws_ttm, 'ttm', is_q=False)
        write_row(ws_q, 'q', is_q=True)

    # Yabancı Oranı Sayfasını Yazma
    if sheet_yab_name in wb.sheetnames:
        ws_yab = wb[sheet_yab_name]
        headers_yab = {1: "Ticker", 2: "Fiyat", 3: date_a, 4: date_b, 5: "Change", 6: "Effect"}
        for col_num, text in headers_yab.items(): ws_yab.cell(row=1, column=col_num, value=text)
        
        for idx, res in enumerate(tr_yab_results, start=2):
            sym, yd = res['symbol'], res['data']
            ws_yab.cell(row=idx, column=1, value=sym)
            c_price = ws_yab.cell(row=idx, column=2, value=yd['price'])
            c_start = ws_yab.cell(row=idx, column=3, value=yd['start'])
            c_end = ws_yab.cell(row=idx, column=4, value=yd['end'])
            c_change = ws_yab.cell(row=idx, column=5, value=yd['change'])
            c_effect = ws_yab.cell(row=idx, column=6, value=yd['effect'])
            
            if isinstance(c_price.value, (int, float)): c_price.number_format = '#,##0.00'
            for c in [c_start, c_end, c_change, c_effect]:
                if isinstance(c.value, (int, float)): c.number_format = '0.00%'

    # Historical_Q Sayfasını Yazma
    if hist_q_tasks:
        sheet_hist_name = "Historical_Q"
        if sheet_hist_name not in wb.sheetnames: wb.create_sheet(sheet_hist_name)
        ws_hist = wb[sheet_hist_name]
        
        headers_hist = {
            1: "Ticker", 2: "Sektör", 3: "Dönem", 4: "Borç / Özsermaye", 5: "Faaliyet Kârı (Q)", 
            6: "Nakit Akışı (Q)", 7: "Net Döviz Pozisyonu", 8: "İhracat Geliri (Q)", 
            9: "Yabancı Oranı", 10: "Yabancı Oranı Değişim",
            11: "Piyasa Değeri", 12: "PD/DD",
            13: "FAVÖK (Q)", 14: "FD / FAVÖK (Yıllıklandırılmış Q)"
        }
        for col_num, text in headers_hist.items(): ws_hist.cell(row=1, column=col_num, value=text)
        
        curr_row = 2
        for res_task in hist_q_results:
            sym = res_task['symbol']
            hist_data_list = res_task['data']
            
            for d in hist_data_list:
                is_b = d.get('is_bank', False)
                ws_hist.cell(row=curr_row, column=1, value=sym)
                ws_hist.cell(row=curr_row, column=2, value=d.get('industry', ''))
                ws_hist.cell(row=curr_row, column=3, value=d['resolved_period'])
                
                c4 = ws_hist.cell(row=curr_row, column=4, value=d['debt_equity'])
                if isinstance(c4.value, (int, float)): c4.number_format = '0.00'
                
                c5 = ws_hist.cell(row=curr_row, column=5, value=d.get('op_income_q', ''))
                if isinstance(c5.value, (int, float)): c5.number_format = '#,##0'
                
                c6 = ws_hist.cell(row=curr_row, column=6, value=d.get('op_cash_flow_q', ''))
                if isinstance(c6.value, (int, float)): c6.number_format = '#,##0'
                
                c7 = ws_hist.cell(row=curr_row, column=7, value=d.get('net_fx', ''))
                if isinstance(c7.value, (int, float)): c7.number_format = '#,##0'
                
                c8 = ws_hist.cell(row=curr_row, column=8, value=d.get('export_ratio_q', ''))
                if isinstance(c8.value, (int, float)): c8.number_format = '0.00%'
                
                historical_price = d.get('historical_price', "")
                item_2oa = d.get('item_2oa', "")
                ozk = d.get('ozk', "")
                net_borc = d.get('net_borc', "")
                favok = d.get('favok_q', "")
                
                piyasa_degeri = ""
                pddd = ""
                if isinstance(historical_price, (int, float)) and isinstance(item_2oa, (int, float)):
                    piyasa_degeri = historical_price * item_2oa
                    if isinstance(ozk, (int, float)) and ozk != 0:
                        pddd = piyasa_degeri / ozk
                
                fd_favok = ""
                if isinstance(piyasa_degeri, (int, float)) and isinstance(net_borc, (int, float)):
                    fd = piyasa_degeri + net_borc
                    if isinstance(favok, (int, float)) and favok != 0:
                        fd_favok = fd / (favok * 4)

                historical_yabanci = d.get('historical_yabanci', "")
                c9 = ws_hist.cell(row=curr_row, column=9, value=historical_yabanci)
                if isinstance(c9.value, (int, float)): c9.number_format = '0.00%'
                
                ws_hist.cell(row=curr_row, column=10, value="")
                
                c11 = ws_hist.cell(row=curr_row, column=11, value=piyasa_degeri)
                if isinstance(c11.value, (int, float)): c11.number_format = '#,##0'
                
                c12 = ws_hist.cell(row=curr_row, column=12, value=pddd)
                if isinstance(c12.value, (int, float)): c12.number_format = '0.00'
                
                c13 = ws_hist.cell(row=curr_row, column=13, value=favok)
                if isinstance(c13.value, (int, float)): c13.number_format = '#,##0'
                
                c14 = ws_hist.cell(row=curr_row, column=14, value=fd_favok)
                if isinstance(c14.value, (int, float)): c14.number_format = '0.00'
                
                c4.fill = green_fill if not is_b and isinstance(c4.value, (int, float)) and c4.value < 1.0 else no_fill
                c5.fill = green_fill if isinstance(c5.value, (int, float)) and c5.value > 0 else no_fill
                c6.fill = green_fill if isinstance(c6.value, (int, float)) and c6.value > 0 else no_fill
                c7.fill = green_fill if not is_b and isinstance(c7.value, (int, float)) and c7.value > 0 else no_fill
                
                curr_row += 1

    wb.save(file_name)
    print(f"✅ Excel dosyası başarıyla güncellendi ve kapatıldı!")

if __name__ == "__main__":
    main_automation()