# Excel sayfa yazımı: Config, Facts (uzun format), Sinyal (Excel 365 formülleri).
# Sinyal için FILTER / LET / XLOOKUP kullanılır — Microsoft 365 gerekir.

from openpyxl.utils import get_column_letter

FACTS_SHEET = "Facts"
OZET_SHEET = "Özet"
CONFIG_PERIOD_START_ROW = 1
CONFIG_PERIOD_COL = 6  # F — dönem etiketleri (Config!F1:…)
CONFIG_PERIOD_MAX_ROWS = 50

# Yeni metrik: (Excel Facts sütunu "Metrik" kodu, hist hücre üçlüsündeki indeks)
METRICS = [
    ("PDDD", 0),
    ("FDFAVOK", 1),
    ("BORC", 2),
]


def ensure_config_sheet(wb):
    if "Config" in wb.sheetnames:
        return
    ws = wb.create_sheet("Config")
    ws["A1"], ws["B1"] = "analiz_donemi", 12
    ws["A2"], ws["B2"] = "ucuz_esik", 25
    ws["A3"], ws["B3"] = "pahali_esik", 75


def _clear_matrix_sheet(ws, max_row, max_col):
    for r in range(1, max_row + 1):
        for c in range(1, max_col + 1):
            cell = ws.cell(row=r, column=c)
            cell.value = None
            cell.number_format = "General"


def _period_range_end_col_letter():
    return get_column_letter(CONFIG_PERIOD_COL)


def update_config_period_labels(wb, period_labels):
    """
    Config!F1:F* içine çeyrek etiketlerini yazar (en güncel üstte).
    analiz_donemi (B1) ve diğer kullanıcı alanlarına dokunmaz.
    """
    if "Config" not in wb.sheetnames:
        ensure_config_sheet(wb)
    ws = wb["Config"]
    col = CONFIG_PERIOD_COL
    last = CONFIG_PERIOD_START_ROW + CONFIG_PERIOD_MAX_ROWS
    for r in range(CONFIG_PERIOD_START_ROW, last + 1):
        ws.cell(row=r, column=col, value=None)
    for i, lbl in enumerate(period_labels[:CONFIG_PERIOD_MAX_ROWS]):
        ws.cell(row=CONFIG_PERIOD_START_ROW + i, column=col, value=lbl)


def write_facts(wb, period_labels, matrix_rows):
    """
    matrix_rows: list of (row_idx, symbol, cells) where cells is length N list of
                 (pddd, fd_favok, borc) each float or None.
    Facts: Ticker | Metrik | Dönem | Değer (uzun format).
    row_idx yalnızca Main_TR_Q satırıyla hizalama için taşınır; Facts satır sırasında kullanılmaz.
    """
    n = len(period_labels)
    if FACTS_SHEET not in wb.sheetnames:
        wb.create_sheet(FACTS_SHEET)
    ws = wb[FACTS_SHEET]

    n_metrics = len(METRICS)
    data_rows = max(1, len(matrix_rows) * max(n, 1) * n_metrics)
    _clear_matrix_sheet(ws, max(data_rows + 5, 500), 10)

    ws.cell(row=1, column=1, value="Ticker")
    ws.cell(row=1, column=2, value="Metrik")
    ws.cell(row=1, column=3, value="Dönem")
    ws.cell(row=1, column=4, value="Değer")

    out_r = 2
    for _row_idx, symbol, cells in matrix_rows:
        for j in range(n):
            period = period_labels[j] if j < len(period_labels) else ""
            trip = cells[j] if j < len(cells) else (None, None, None)
            for metric_code, idx in METRICS:
                v = trip[idx] if idx < len(trip) else None
                ws.cell(row=out_r, column=1, value=symbol)
                ws.cell(row=out_r, column=2, value=metric_code)
                ws.cell(row=out_r, column=3, value=period)
                c = ws.cell(row=out_r, column=4, value=v)
                if isinstance(v, (int, float)):
                    c.number_format = "0.00"
                out_r += 1


def _config_win_spill_ref():
    """Dikey dönem penceresi: F1’den itibaren MIN(B1, dolu F sayısı) satır."""
    col = _period_range_end_col_letter()
    return (
        f"OFFSET(Config!${col}${CONFIG_PERIOD_START_ROW},0,0,"
        f"MIN(Config!$B$1,COUNTA(Config!${col}$"
        f"{CONFIG_PERIOD_START_ROW}:"
        f"${col}${CONFIG_PERIOD_START_ROW + CONFIG_PERIOD_MAX_ROWS - 1})),1)"
    )


def _facts_filter_v(metric_code: str, main_row: int) -> str:
    win = _config_win_spill_ref()
    return (
        f"_xlfn.FILTER(Facts!$D:$D,(Facts!$A:$A=$A{main_row})*(Facts!$B:$B=\"{metric_code}\")"
        f"*(ISNUMBER(MATCH(Facts!$C:$C,{win},0)))*(ISNUMBER(Facts!$D:$D)))"
    )


def _facts_current_lookup(metric_code: str, main_row: int) -> str:
    return (
        f"_xlfn.LET(x,_xlfn.XLOOKUP(1,(Facts!$A:$A=$A{main_row})*(Facts!$B:$B=\"{metric_code}\")"
        f"*(Facts!$C:$C=Config!${_period_range_end_col_letter()}${CONFIG_PERIOD_START_ROW}),"
        f"Facts!$D:$D,\"\"),IF(ISNUMBER(x),x,\"\"))"
    )


def _facts_median_let(metric_code: str, main_row: int) -> str:
    v = _facts_filter_v(metric_code, main_row)
    return f"_xlfn.LET(v,{v},IFERROR(MEDIAN(v),\"\"))"


def _facts_pct_let(metric_code: str, main_row: int) -> str:
    v = _facts_filter_v(metric_code, main_row)
    cur = _facts_current_lookup(metric_code, main_row)
    return (
        f"_xlfn.LET(v,{v},cur,{cur},"
        f"IF(OR(Config!$B$1<=0,NOT(ISNUMBER(cur))),\"\","
        f"IFERROR(SUM(--(v<=cur))/Config!$B$1,\"\")))"
    )


def write_sinyal_sheet(wb, tr_tasks, tr_fin_results_map):
    """
    tr_tasks: [(row_idx, symbol), ...] — Main_TR_Q ile aynı satır numaraları.
    Facts + Config (F sütunu dönem listesi) üzerinden Excel 365 formülleri.
    """
    name = "Sinyal"
    if name not in wb.sheetnames:
        wb.create_sheet(name)
    ws = wb[name]

    # A–Q (17 sütun); P=Skor, Q=Genel Sinyal
    headers = {
        1: "Ticker",
        2: "Sektör",
        3: "Banka mı?",
        4: "Güncel PD/DD",
        5: "Medyan PD/DD",
        6: "PD/DD %ile",
        7: "PD/DD Sinyal",
        8: "Güncel FD/FAVÖK",
        9: "Medyan FD/FAVÖK",
        10: "FD/FAVÖK %ile",
        11: "FD/FAVÖK Sinyal",
        12: "Güncel Borç/Özsermaye",
        13: "Medyan Borç/Özsermaye",
        14: "Borç/Özser. %ile",
        15: "Borç/Özser. Sinyal",
        16: "Skor",
        17: "Genel Sinyal",
    }
    max_r = 2
    for r, _ in tr_tasks:
        max_r = max(max_r, r)
    _clear_matrix_sheet(ws, max(max_r + 2, 50), 20)

    for col, text in headers.items():
        ws.cell(row=1, column=col, value=text)

    for row_idx, sym in tr_tasks:
        r = row_idx
        fin = tr_fin_results_map.get(sym) or {}
        data = fin.get("data") or {}
        is_bank = bool(data.get("is_bank", False))
        c_val = "EVET" if is_bank else "HAYIR"

        ws.cell(row=r, column=1, value=f"=Main_TR_Q!A{r}")
        ws.cell(row=r, column=2, value=f"=Main_TR_Q!B{r}")
        ws.cell(row=r, column=3, value=c_val)

        cur_p = _facts_current_lookup("PDDD", r)
        med_p = _facts_median_let("PDDD", r)
        pct_p = _facts_pct_let("PDDD", r)
        ws.cell(row=r, column=4, value=f"={cur_p}")
        ws.cell(row=r, column=5, value=f"={med_p}")
        ws.cell(row=r, column=6, value=f"={pct_p}")
        ws.cell(
            row=r,
            column=7,
            value=(
                f"=IF(F{r}<Config!$B$2/100,\"UCUZ\","
                f"IF(F{r}>Config!$B$3/100,\"PAHALI\",\"NÖTR\"))"
            ),
        )

        cur_f = _facts_current_lookup("FDFAVOK", r)
        med_f = _facts_median_let("FDFAVOK", r)
        pct_f = _facts_pct_let("FDFAVOK", r)
        ws.cell(row=r, column=8, value=f"=IF(C{r}=\"EVET\",\"\",{cur_f})")
        ws.cell(row=r, column=9, value=f"=IF(C{r}=\"EVET\",\"\",{med_f})")
        ws.cell(row=r, column=10, value=f"=IF(C{r}=\"EVET\",\"\",{pct_f})")
        ws.cell(
            row=r,
            column=11,
            value=(
                f"=IF(C{r}=\"EVET\",\"\","
                f"IF(J{r}<Config!$B$2/100,\"UCUZ\","
                f"IF(J{r}>Config!$B$3/100,\"PAHALI\",\"NÖTR\")))"
            ),
        )

        cur_b = _facts_current_lookup("BORC", r)
        med_b = _facts_median_let("BORC", r)
        pct_b = _facts_pct_let("BORC", r)
        ws.cell(row=r, column=12, value=f"=IF(C{r}=\"EVET\",\"\",{cur_b})")
        ws.cell(row=r, column=13, value=f"=IF(C{r}=\"EVET\",\"\",{med_b})")
        ws.cell(row=r, column=14, value=f"=IF(C{r}=\"EVET\",\"\",{pct_b})")
        ws.cell(
            row=r,
            column=15,
            value=(
                f"=IF(C{r}=\"EVET\",\"\","
                f"IF(N{r}<Config!$B$2/100,\"UCUZ\","
                f"IF(N{r}>Config!$B$3/100,\"PAHALI\",\"NÖTR\")))"
            ),
        )

        ws.cell(
            row=r,
            column=16,
            value=(
                f"=IF(C{r}=\"EVET\",ROUND((1-F{r})*100,0),"
                f"ROUND(((1-F{r})*0.4+(1-J{r})*0.4+(1-N{r})*0.2)*100,0))"
            ),
        )
        ws.cell(
            row=r,
            column=17,
            value=(
                f"=IF(P{r}>Config!$B$3,\"UCUZ\","
                f"IF(P{r}<Config!$B$2,\"PAHALI\",\"NÖTR\"))"
            ),
        )


def write_ozet_sheet(wb, tr_tasks):
    """
    Tek ekran özeti: Main_TR_Q, TR_TTM ve Sinyal ile aynı satır numarası (row_idx).
    """
    if not tr_tasks:
        return
    if "Main_TR_Q" not in wb.sheetnames or "TR_TTM" not in wb.sheetnames:
        return
    if "Sinyal" not in wb.sheetnames:
        return

    name = OZET_SHEET
    if name not in wb.sheetnames:
        wb.create_sheet(name)
    ws = wb[name]

    headers = {
        1: "Ticker",
        2: "Sektör",
        3: "Dönem",
        4: "Banka mı?",
        5: "Genel Sinyal",
        6: "Skor",
        7: "PD/DD Sinyal",
        8: "FD/FAVÖK Sinyal",
        9: "Borç/Özser. Sinyal",
        10: "Faaliyet Kârı (TTM)",
        11: "Nakit Akışı (TTM)",
        12: "Net Döviz",
        13: "İhracat (TTM) %",
        14: "Yabancı Oranı",
        15: "Yabancı Oranı Değişim",
        16: "İşletme özü",
        17: "Yapı özü",
    }
    max_r = 2
    for r, _ in tr_tasks:
        max_r = max(max_r, r)
    _clear_matrix_sheet(ws, max(max_r + 2, 50), 22)

    for col, text in headers.items():
        ws.cell(row=1, column=col, value=text)

    for row_idx, _sym in tr_tasks:
        r = row_idx
        ws.cell(row=r, column=1, value=f"=Main_TR_Q!A{r}")
        ws.cell(row=r, column=2, value=f"=Main_TR_Q!B{r}")
        ws.cell(row=r, column=3, value=f"=TR_TTM!C{r}")
        ws.cell(row=r, column=4, value=f"=Sinyal!C{r}")
        ws.cell(row=r, column=5, value=f"=Sinyal!Q{r}")
        ws.cell(row=r, column=6, value=f"=Sinyal!P{r}")
        ws.cell(row=r, column=7, value=f"=Sinyal!G{r}")
        ws.cell(row=r, column=8, value=f"=Sinyal!K{r}")
        ws.cell(row=r, column=9, value=f"=Sinyal!O{r}")
        ws.cell(row=r, column=10, value=f"=TR_TTM!E{r}")
        ws.cell(row=r, column=11, value=f"=TR_TTM!F{r}")
        ws.cell(row=r, column=12, value=f"=TR_TTM!G{r}")
        ws.cell(row=r, column=13, value=f"=TR_TTM!H{r}")
        ws.cell(row=r, column=14, value=f"=TR_TTM!I{r}")
        ws.cell(row=r, column=15, value=f"=TR_TTM!J{r}")
        ws.cell(
            row=r,
            column=16,
            value=(
                f"=IF(AND(ISNUMBER(TR_TTM!E{r}),TR_TTM!E{r}>0,"
                f"ISNUMBER(TR_TTM!F{r}),TR_TTM!F{r}>0),\"İşletme+\",\"Dikkat\")"
            ),
        )
        ws.cell(
            row=r,
            column=17,
            value=(
                f"=IF(Sinyal!C{r}=\"EVET\",\"Banka\","
                f"IF(AND(ISNUMBER(TR_TTM!D{r}),TR_TTM!D{r}<1,"
                f"ISNUMBER(TR_TTM!G{r}),TR_TTM!G{r}>0),\"Yapı+\",\"Yapı kontrol\"))"
            ),
        )

        for c in (10, 11, 12):
            cell = ws.cell(row=r, column=c)
            cell.number_format = "#,##0"
        ws.cell(row=r, column=13).number_format = "0.00%"
        ws.cell(row=r, column=14).number_format = "0.00%"
        ws.cell(row=r, column=15).number_format = "0.00%"
        ws.cell(row=r, column=6).number_format = "0"
