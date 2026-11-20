#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
build_wide.py - sjednocení ročních reportů aukcí EU ETS (EEX, primární trh) 2012-2025.

Spuštění (od začátku do konce bez ručních zásahů):
    python build_wide.py [DATA_PATH] [OUT_PATH]

Cesty lze zadat i proměnnými prostředí ETS_DATA_PATH / ETS_OUT_PATH.
Zdrojové soubory se jen čtou. Skript je deterministický a idempotentní (výstupy se přepisují).

KLÍČOVÉ ZJIŠTĚNÍ (ověřeno v datech, viz validace):
  Hodnoty "po státech" (sloupce států od 2017 i řetězec 'Auction Details' u 2012-2016) NEJSOU objemy
  v tCO2, ale VÝNOSY aukce v EUR přiřazené jednotlivým státům. Jejich součet = 'Total Revenue'
  (přesně). Objem státu je proto DOPOČÍTÁN jako výnos_státu / aukční_cena (součet přes státy pak
  přesně odpovídá 'Auction Volume'). Surové výnosy zůstávají ve výstupu (revenue_eur).
"""
import glob
import logging
import os
import re
import sys
import warnings

import numpy as np
import pandas as pd

warnings.filterwarnings("ignore")

# --------------------------------------------------------------------------------------
# KONSTANTY / CESTY
# --------------------------------------------------------------------------------------
DATA_PATH = os.environ.get("ETS_DATA_PATH") or (
    r"G:\My Drive\EU_ETS1\emission-spot-primary-market-auction-report-2012-2025-data"
    r"\emission-spot-primary-market-auction-report-2012-2025-data"
)
OUT_PATH = os.environ.get("ETS_OUT_PATH") or ""          # prázdné = DATA_PATH (resp. ../output_wide)
if len(sys.argv) > 1:
    DATA_PATH = sys.argv[1]
if len(sys.argv) > 2:
    OUT_PATH = sys.argv[2]

DATA_NAME_TO2019 = "emission-spot-primary-market-auction-report-{year}-data.xls"
DATA_NAME_FROM2020 = "emission-spot-primary-market-auction-report-{year}-data.xlsx"
YEARS = list(range(2012, 2026))
EXTS = (".xls", ".xlsx")          # hranice .xls/.xlsx se nehardcoduje, hledá se podle roku a obou přípon

OUT_WIDE_XLSX = "eu_ets_auctions_wide.xlsx"
OUT_WIDE_CSV = "eu_ets_auctions_wide.csv"
OUT_DAILY_CSV = "eu_ets_auctions_wide_daily.csv"
OUT_REVENUE_CSV = "eu_ets_auctions_wide_revenue_eur.csv"
OUT_LONG_CSV = "eu_ets_auctions_long.csv"
OUT_CANCELLED_CSV = "eu_ets_auctions_cancelled.csv"
OUT_SURVEY_CSV = "source_survey.csv"
OUT_VALID_CSV = "validation_volume_check.csv"
OUT_YEAR_CSV = "validation_year_counts.csv"
OUT_STATES_CSV = "validation_state_history.csv"
OUT_SPOT_CSV = "validation_spot_checks.csv"
OUT_REPORT = "validation_report.txt"
OUT_LOG = "build_wide.log"

# Mapování kódu -> název (z hlaviček nového formátu a z listu Caption; XA není v Caption nikde uveden)
CODE_NAMES = {
    "AT": "Austria", "BE": "Belgium", "BG": "Bulgaria", "CY": "Cyprus", "CZ": "Czech Republic",
    "DE": "Germany", "DK": "Denmark", "EE": "Estonia", "EL": "Greece", "ES": "Spain", "FI": "Finland",
    "FR": "France", "HR": "Croatia", "HU": "Hungary", "IE": "Ireland", "IF": "Innovation Fund",
    "IS": "Iceland", "IT": "Italy", "IX": "InnoFund RRF", "LI": "Liechtenstein", "LT": "Lithuania",
    "LU": "Luxembourg", "LV": "Latvia", "MF": "Modernisation Fund", "MT": "Malta", "MX": "MS RRF",
    "NL": "Netherlands", "NO": "Norway", "PL": "Poland", "PT": "Portugal", "RO": "Romania",
    "SE": "Sweden", "SF": "Social Climate Fund", "SI": "Slovenia", "SK": "Slovakia",
    "XI": "Northern Ireland", "XF": "Flanders (Caption)", "UK": "United Kingdom (Caption, v datech nenalezeno)",
    "EU": "Europe (společná aukce EEX)",
}
# sloupce bez kódu v závorce, které se mapují jinak (2019: "Innovation\nFund" bez kódu)
NAME_TO_CODE = {"innovation fund": "IF"}
CONTRACT_TO_CERT = {"T2PA": "EUA", "T3PA": "EUA", "EAA2": "EUAA", "EAA3": "EUAA"}

# regexy pro sjednocení názvů sloupců (po normalizaci: bez \n, jedna mezera, lower)
META_PATTERNS = [
    ("date", r"^date$"),
    ("time", r"^(auction )?time$"),
    ("auction_name_raw", r"^auction name$"),
    ("contract", r"^contract$"),
    ("certificate_src", r"^certificate$"),
    ("status_src", r"^status$"),
    ("auction_price_eur", r"^auction price"),
    ("min_bid_eur", r"^min(imum|imal)?( bid| price)"),
    ("max_bid_eur", r"^max(imum|imal)?( bid| price)"),
    ("mean_price_eur", r"^mean"),
    ("median_price_eur", r"^median"),
    ("auction_volume_tco2", r"^auction volume"),
    ("total_bids_tco2", r"^total amount of bids"),
    ("cover_ratio", r"^cover ratio"),
    ("n_bidders", r"^total number of bidders"),
    ("n_successful_bidders", r"^number of succ?ess?ful bidders"),
    ("total_revenue_eur", r"^total revenue"),
    ("country_src", r"^country$"),
    ("details_src", r"^auction details$"),
]

log = logging.getLogger("build_wide")


# --------------------------------------------------------------------------------------
# POMOCNÉ FUNKCE
# --------------------------------------------------------------------------------------
def norm(s):
    """Sjednocení textu hlavičky: \\n, NBSP, dvojité mezery, lower-case."""
    s = str(s).replace("\xa0", " ").replace("\n", " ")
    return re.sub(r"\s+", " ", s).strip().lower()


def find_files(data_path):
    """year -> cesta; hledá podle roku a obou přípon, rekurzivně (kvůli dvojité složce)."""
    found, missing = {}, []
    for y in YEARS:
        hits = []
        for ext in EXTS:
            hits += glob.glob(os.path.join(data_path, "**", f"*{y}-data{ext}"), recursive=True)
        hits = sorted(set(hits))
        if not hits:
            missing.append(y)
            continue
        if len(hits) > 1:
            log.warning("Rok %s: více souborů, beru první: %s", y, hits)
        found[y] = hits[0]
    return found, missing


def detect_header_row(raw, max_scan=20):
    for i in range(min(max_scan, len(raw))):
        if raw.iloc[i].map(norm).eq("date").any():
            return i
    raise ValueError("Hlavička s 'Date' nenalezena")


def read_sheet(path):
    """Vrátí (raw_df bez hlavičky, název listu) pro list s daty aukcí (první list obsahující hlavičku 'Date')."""
    xl = pd.ExcelFile(path)
    for sh in xl.sheet_names:
        raw = xl.parse(sh, header=None)
        try:
            hr = detect_header_row(raw)
        except ValueError:
            continue
        return raw, hr, sh, xl.sheet_names
    raise ValueError(f"{path}: žádný list s hlavičkou 'Date'")


def clean_number(v):
    """Text/číslo -> float. Odstraní mezery/NBSP a tisícové tečky (evropský zápis 2.309.100)."""
    if v is None or (isinstance(v, float) and np.isnan(v)):
        return np.nan
    if isinstance(v, (int, float, np.integer, np.floating)):
        return float(v)
    s = str(v).replace("\xa0", "").replace(" ", "").strip()
    if s == "" or s.lower() in ("nan", "none", "-"):
        return np.nan
    if re.fullmatch(r"-?\d{1,3}(\.\d{3})+", s):       # 1.234.567 / 941.700 -> tisícové oddělovače
        return float(s.replace(".", ""))
    if re.fullmatch(r"-?\d+", s):
        return float(s)
    if re.fullmatch(r"-?\d+,\d+", s):                  # desetinná čárka
        return float(s.replace(",", "."))
    if re.fullmatch(r"-?\d+\.\d+", s):                 # desetinná tečka (ne tisíce)
        return float(s)
    raise ValueError(f"nečíselná hodnota: {v!r}")


def load_year(year, path):
    """Načte jeden roční soubor: tabulka s kanonickými metadaty + surové sloupce + info o formátu."""
    raw, hr, sheet, sheets = read_sheet(path)
    header = list(raw.iloc[hr])
    body = raw.iloc[hr + 1:].copy()
    body["_excel_row"] = body.index + 1                # číslo řádku v Excelu (1-based)
    body.columns = list(range(len(header))) + ["_excel_row"]

    meta_cols, state_cols, ignored = {}, {}, []
    for j, h in enumerate(header):
        n = norm(h)
        if n in ("nan", ""):
            continue
        m = re.search(r"\(([A-Z]{2})\)", str(h))
        if m:
            state_cols[j] = m.group(1)
            continue
        if n in NAME_TO_CODE:
            state_cols[j] = NAME_TO_CODE[n]
            log.info("%s: sloupec %r bez kódu v závorce mapován na %s", year, str(h).replace("\n", "\\n"), NAME_TO_CODE[n])
            continue
        for key, pat in META_PATTERNS:
            if re.search(pat, n) and key not in meta_cols.values():
                meta_cols[j] = key
                break
        else:
            ignored.append(str(h).replace("\n", "\\n"))

    df = pd.DataFrame({"_excel_row": body["_excel_row"]})
    for j, key in meta_cols.items():
        df[key] = body[j]
    fmt = "new" if state_cols else "old"
    if fmt == "old" and "details_src" not in df:
        raise ValueError(f"{year}: ani sloupce států, ani 'Auction Details'")

    # --- odstranění pomocných / souhrnných / opakovaných hlaviček ---
    n0 = len(df)
    is_hdr = df["date"].map(lambda v: isinstance(v, str) and norm(v) == "date")
    df = df[~is_hdr]
    dates = pd.to_datetime(df["date"], errors="coerce", dayfirst=True) if df["date"].map(lambda v: isinstance(v, str)).any() \
        else pd.to_datetime(df["date"], errors="coerce")
    bad = dates.isna()
    if bad.any():
        log.info("%s: odstraněno %d řádků bez platného data (poznámky/součty)", year, int(bad.sum()))
    df = df[~bad].copy()
    df["date"] = dates[~bad].dt.normalize()
    if (df["date"].dt.year != year).any():
        raise ValueError(f"{year}: data mimo rok souboru -> možná záměna den/měsíc: {df.loc[df['date'].dt.year != year, 'date'].unique()[:5]}")
    df["source_year"] = year
    df["source_file"] = os.path.basename(path)
    df["source_format"] = fmt

    for c in ["auction_price_eur", "auction_volume_tco2", "total_revenue_eur", "min_bid_eur", "max_bid_eur",
              "mean_price_eur", "median_price_eur", "total_bids_tco2", "cover_ratio", "n_bidders", "n_successful_bidders"]:
        if c in df:
            df[c] = df[c].map(clean_number)
    for c in ["contract", "auction_name_raw", "certificate_src", "status_src", "country_src"]:
        if c in df:
            df[c] = df[c].astype(str).str.strip().replace({"nan": np.nan})

    info = dict(year=year, file=os.path.basename(path), ext=os.path.splitext(path)[1], sheets="|".join(sheets),
                data_sheet=sheet, raw_rows=len(raw), header_row=hr, n_cols=len(header), fmt=fmt,
                data_rows=len(df), n_state_cols=len(state_cols), ignored_cols="|".join(ignored))
    return df, raw, hr, state_cols, info


def auction_type(name_raw, country_src, contract):
    """Sjednocený typ aukce: EU (společná), DE, PL, XI (Severní Irsko, 'NIR'), nebo kód národní aukce (HU, LT, CZ, NL, XF, XA)."""
    if isinstance(country_src, str) and re.fullmatch(r"[A-Z]{2}", country_src):
        return country_src
    s = str(name_raw)
    if re.fullmatch(r"[A-Z]{2}", s):
        return s
    toks = [t.upper() for t in re.findall(r"[A-Za-z]+", s)]
    if "NIR" in toks:
        return "XI"
    for t in ("EU", "DE", "PL"):
        if t in toks:
            return t
    raise ValueError(f"neznámý typ aukce: {name_raw!r}")


# --------------------------------------------------------------------------------------
# PARSERY (long tabulka)
# --------------------------------------------------------------------------------------
LONG_COLS = ["date", "source_year", "auction_name_raw", "auction_type", "contract", "certificate", "status",
             "auction_price_eur", "auction_volume_tco2", "total_revenue_eur", "country_code",
             "revenue_eur", "volume_tco2", "volume_method", "source_file", "source_row"]


def _finalize_common(df):
    df = df.copy()
    df["auction_type"] = [auction_type(a, c, k) for a, c, k in
                          zip(df["auction_name_raw"], df.get("country_src", pd.Series([np.nan] * len(df), index=df.index)), df["contract"])]
    df["certificate"] = df["contract"].map(CONTRACT_TO_CERT)
    if "certificate_src" in df:                       # kontrola vůči explicitnímu sloupci (2016)
        mism = df["certificate_src"].notna() & (df["certificate_src"] != df["certificate"])
        if mism.any():
            log.warning("Certificate != Contract mapování v %d řádcích (%s)", mism.sum(), df["source_year"].iloc[0])
    if "status_src" in df:
        df["status"] = df["status_src"].str.lower()
    else:                                              # starý formát nemá Status: úspěšná = má cenu
        df["status"] = np.where(df["auction_price_eur"].notna(), "successful", "unknown")
    if "country_src" in df:
        mm = df["country_src"].notna() & (df["country_src"] != df["auction_type"])
        if mm.any():
            log.warning("Country != typ z Auction Name v %d řádcích", mm.sum())
    return df


def parse_old_format(df):
    """Starý formát (2012-2016): rozparsuje 'Auction Details' ('AT: 941.700\\nLU: ...' nebo 'AT:431520,BE:...')."""
    df = _finalize_common(df)
    rows, bad_tokens = [], []
    for r in df.to_dict("records"):
        base = dict(date=r["date"], source_year=r["source_year"], auction_name_raw=r["auction_name_raw"],
                    auction_type=r["auction_type"], contract=r["contract"], certificate=r["certificate"], status=r["status"],
                    auction_price_eur=r["auction_price_eur"], auction_volume_tco2=r["auction_volume_tco2"],
                    total_revenue_eur=r["total_revenue_eur"], source_file=r["source_file"], source_row=r["_excel_row"])
        det = r["details_src"]
        if r["status"] != "successful":
            rows.append(dict(base, country_code=np.nan, revenue_eur=np.nan, volume_tco2=np.nan, volume_method="cancelled"))
            continue
        if isinstance(det, str) and det.strip():
            for tok in re.split(r"[\n,;]+", det.strip()):
                tok = tok.strip()
                if not tok:
                    continue
                if ":" not in tok:
                    bad_tokens.append((r["date"], tok))
                    continue
                code, val = tok.split(":", 1)
                try:
                    rev = clean_number(val)
                except ValueError:
                    bad_tokens.append((r["date"], tok))
                    continue
                rows.append(dict(base, country_code=code.strip(), revenue_eur=rev, volume_tco2=np.nan,
                                 volume_method="derived_revenue_over_price"))
        else:
            # národní aukce bez rozpisu (DE, PL, HU, LT, CZ, NL, XF, XA...): celý objem patří pořádajícímu státu
            rev = r["total_revenue_eur"] if pd.notna(r["total_revenue_eur"]) else r["auction_price_eur"] * r["auction_volume_tco2"]
            rows.append(dict(base, country_code=r["auction_type"], revenue_eur=rev, volume_tco2=r["auction_volume_tco2"],
                             volume_method="national_auction_full_volume"))
    if bad_tokens:
        log.warning("Nenaparsované tokeny v Auction Details (%d): %s", len(bad_tokens), bad_tokens[:5])
    return pd.DataFrame(rows)


def parse_new_format(df, raw, state_cols):
    """Nový formát (2017+): kódy států z hlaviček regexem \\(([A-Z]{2})\\), hodnoty ve sloupcích = výnos v EUR."""
    df = _finalize_common(df)
    out = []
    rawv = raw.copy()
    rawv.index = rawv.index + 1                         # excel row
    for code_idx, code in state_cols.items():
        vals = rawv.loc[df["_excel_row"].values, code_idx].map(clean_number).fillna(0.0).values
        part = pd.DataFrame({
            "date": df["date"].values, "source_year": df["source_year"].values,
            "auction_name_raw": df["auction_name_raw"].values, "auction_type": df["auction_type"].values,
            "contract": df["contract"].values, "certificate": df["certificate"].values, "status": df["status"].values,
            "auction_price_eur": df["auction_price_eur"].values, "auction_volume_tco2": df["auction_volume_tco2"].values,
            "total_revenue_eur": df["total_revenue_eur"].values, "country_code": code, "revenue_eur": vals,
            "volume_tco2": np.nan, "volume_method": "derived_revenue_over_price",
            "source_file": df["source_file"].values, "source_row": df["_excel_row"].values})
        out.append(part)
    long = pd.concat(out, ignore_index=True)
    canc = long["status"] != "successful"
    long.loc[canc, ["revenue_eur", "volume_method"]] = [np.nan, "cancelled"]
    # státy, které v dané aukci nemají nic (0), do long tabulky nepatří (wide je doplní nulou)
    keep = canc | (long["revenue_eur"].fillna(0) != 0)
    # zrušené aukce: ponechat jen jeden řádek na aukci
    first_canc = canc & ~long.loc[canc].duplicated(["date", "auction_type", "contract"]).reindex(long.index, fill_value=False)
    long = long[(keep & ~canc) | first_canc].copy()
    long.loc[long["volume_method"] == "cancelled", "country_code"] = np.nan
    return long


def derive_volume(long):
    """volume = revenue / price (pokud už není známo); zaokrouhlení na celé tCO2, je-li odchylka < 1e-6."""
    m = long["volume_tco2"].isna() & long["revenue_eur"].notna() & long["auction_price_eur"].notna()
    est = long.loc[m, "revenue_eur"] / long.loc[m, "auction_price_eur"]
    dev = (est - est.round()).abs().max() if len(est) else 0
    log.info("Dopočet objemu revenue/price: %d řádků, max odchylka od celého čísla = %.2e", int(m.sum()), dev)
    long.loc[m, "volume_tco2"] = est.round() if dev < 1e-6 else est.round(6)
    return long, dev


# --------------------------------------------------------------------------------------
# WIDE / DAILY
# --------------------------------------------------------------------------------------
META_WIDE = ["date", "year", "auction_type", "certificate", "contract", "status", "auction_name_raw",
             "auction_price_eur", "auction_volume_tco2", "total_revenue_eur"]


def build_wide(long, auctions, value_col, prefix, states):
    key = ["date", "auction_type", "certificate", "contract"]
    ok = long[long["status"] == "successful"]
    dups = ok.duplicated(key + ["country_code"], keep=False)
    if dups.any():
        log.warning("Duplicitní klíč v long tabulce (%d řádků) -> sčítám a vypisuji", int(dups.sum()))
        log.warning("\n%s", ok[dups].head(20).to_string())
    pv = ok.pivot_table(index=key, columns="country_code", values=value_col, aggfunc="sum", fill_value=0.0)
    pv = pv.reindex(columns=states, fill_value=0.0)
    pv.columns = [f"{prefix}{c}" for c in pv.columns]
    wide = auctions.merge(pv.reset_index(), on=key, how="left")
    wide[list(pv.columns)] = wide[list(pv.columns)].fillna(0.0)
    return wide


def main():
    global OUT_PATH
    out = OUT_PATH or DATA_PATH
    try:
        os.makedirs(out, exist_ok=True)
        testf = os.path.join(out, ".write_test")
        open(testf, "w").close()
        os.remove(testf)
    except Exception:
        out = os.path.join(os.path.dirname(os.path.normpath(DATA_PATH)), "output_wide")
        os.makedirs(out, exist_ok=True)
    logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s",
                        handlers=[logging.StreamHandler(sys.stdout),
                                  logging.FileHandler(os.path.join(out, OUT_LOG), mode="w", encoding="utf-8")])
    log.info("DATA_PATH=%s", DATA_PATH)
    log.info("OUT_PATH=%s", out)
    if os.path.isdir(DATA_PATH) is False:
        raise SystemExit(f"DATA_PATH neexistuje: {DATA_PATH}")

    # ---- 0) soubory ----
    files, missing = find_files(DATA_PATH)
    log.info("Nalezeno souborů: %d z %d; chybí: %s", len(files), len(YEARS), missing)
    if missing:
        raise SystemExit(f"Chybí soubory pro roky: {missing}")

    # ---- 1) průzkum + 2) parsery ----
    survey, longs, cancelled_src, src_counts = [], [], [], {}
    for y in YEARS:
        df, raw, hr, state_cols, info = load_year(y, files[y])
        if info["fmt"] == "old":
            lg = parse_old_format(df)
        else:
            lg = parse_new_format(df, raw, state_cols)
        # průzkumné údaje
        d = df
        info.update(
            auction_names="|".join(sorted(d["auction_name_raw"].dropna().unique())),
            contracts="|".join(sorted(d["contract"].dropna().unique())),
            certificate_col="Certificate" if "certificate_src" in d else "(odvozeno z Contract)",
            status_col=("|".join(f"{k}:{v}" for k, v in d["status_src"].value_counts().items())
                        if "status_src" in d else "(chybí; úspěšná = má cenu)"),
            details_rows=int(d["details_src"].notna().sum()) if "details_src" in d else 0,
            unique_days=int(d["date"].nunique()),
            max_auctions_per_day=int(d.groupby("date").size().max()),
            days_with_multiple_auctions=int((d.groupby("date").size() > 1).sum()),
            states_in_header="|".join(sorted(set(state_cols.values()))) if state_cols else "(v Auction Details)")
        survey.append(info)
        src_counts[y] = len(df)
        longs.append(lg)
        log.info("%s: %s formát, hlavička v řádku %s, %d aukcí, %d dnů", y, info["fmt"], hr, len(df), info["unique_days"])
    survey = pd.DataFrame(survey)

    long = pd.concat(longs, ignore_index=True)
    long, dev = derive_volume(long)

    # zrušené/neuskutečněné aukce -> samostatný soubor, mimo long/wide
    canc = long[long["volume_method"] == "cancelled"].copy()
    long = long[long["volume_method"] != "cancelled"].copy()
    long = long[LONG_COLS].sort_values(["date", "auction_type", "contract", "country_code"]).reset_index(drop=True)

    # ---- 4) klíč: seznam aukcí ----
    key = ["date", "auction_type", "certificate", "contract"]
    auctions = (long.groupby(key, as_index=False)
                .agg(year=("source_year", "first"), status=("status", "first"), auction_name_raw=("auction_name_raw", "first"),
                     auction_price_eur=("auction_price_eur", "first"), auction_volume_tco2=("auction_volume_tco2", "first"),
                     total_revenue_eur=("total_revenue_eur", "first")))
    # národní aukce z 2012 nemají Total Revenue -> doplnit cena*objem (označeno)
    miss_rev = auctions["total_revenue_eur"].isna()
    auctions["total_revenue_imputed"] = miss_rev
    auctions.loc[miss_rev, "total_revenue_eur"] = auctions["auction_price_eur"] * auctions["auction_volume_tco2"]
    log.info("Total Revenue doplněno ze cena*objem u %d aukcí (2012 národní aukce)", int(miss_rev.sum()))
    k3 = ["date", "auction_type", "certificate"]
    n_dup3 = int(auctions.duplicated(k3, keep=False).sum())
    log.info("Klíč Date x Auction type x Certificate: duplicity=%d (viz rozhodnutí v README)", n_dup3)
    auctions = auctions.sort_values(key).reset_index(drop=True)

    states = sorted(long["country_code"].dropna().unique())
    unknown = [s for s in states if s not in CODE_NAMES]
    log.info("Kódy (%d): %s; nenamapované názvy: %s", len(states), states, unknown)

    wide = build_wide(long, auctions, "volume_tco2", "volume_tco2_", states)
    wide_rev = build_wide(long, auctions, "revenue_eur", "revenue_eur_", states)
    vcols = [f"volume_tco2_{s}" for s in states]
    wide["volume_tco2_sum_states"] = wide[vcols].sum(axis=1)
    wide["volume_diff_pct"] = (wide["volume_tco2_sum_states"] - wide["auction_volume_tco2"]) / wide["auction_volume_tco2"] * 100
    wide = wide[META_WIDE + ["total_revenue_imputed", "volume_tco2_sum_states", "volume_diff_pct"] + vcols]
    rcols = [f"revenue_eur_{s}" for s in states]
    wide_rev = wide_rev[META_WIDE + ["total_revenue_imputed"] + rcols]

    # denní agregace
    d = wide.copy()
    d["vol_eua"] = np.where(d["certificate"] == "EUA", d["auction_volume_tco2"], 0.0)
    d["vol_euaa"] = np.where(d["certificate"] == "EUAA", d["auction_volume_tco2"], 0.0)
    daily = d.groupby("date", as_index=False).agg(
        year=("year", "first"), n_auctions=("auction_type", "size"),
        auction_types=("auction_type", lambda s: "|".join(sorted(set(s)))),
        auction_volume_tco2=("auction_volume_tco2", "sum"), auction_volume_tco2_EUA=("vol_eua", "sum"),
        auction_volume_tco2_EUAA=("vol_euaa", "sum"), total_revenue_eur=("total_revenue_eur", "sum"),
        **{c: (c, "sum") for c in vcols})
    daily["avg_price_eur_volume_weighted"] = daily["total_revenue_eur"] / daily["auction_volume_tco2"]
    daily["volume_tco2_sum_states"] = daily[vcols].sum(axis=1)
    daily["volume_diff_pct"] = (daily["volume_tco2_sum_states"] - daily["auction_volume_tco2"]) / daily["auction_volume_tco2"] * 100
    daily = daily[["date", "year", "n_auctions", "auction_types", "auction_volume_tco2", "auction_volume_tco2_EUA",
                   "auction_volume_tco2_EUAA", "total_revenue_eur", "avg_price_eur_volume_weighted",
                   "volume_tco2_sum_states", "volume_diff_pct"] + vcols]

    # ---- 6) validace ----
    rep = []
    P = rep.append
    P("VALIDAČNÍ REPORT build_wide.py")
    P("=" * 80)
    P("1) Soubory: nalezeno %d/%d (%s)" % (len(files), len(YEARS), ", ".join(f"{y}{survey.set_index('year').loc[y, 'ext']}" for y in YEARS)))

    # interpretace: součet surových hodnot po státech vs Total Revenue / Auction Volume
    chk = long[long["volume_method"] == "derived_revenue_over_price"].groupby(key, as_index=False).agg(
        sum_raw=("revenue_eur", "sum"), rev=("total_revenue_eur", "first"), vol=("auction_volume_tco2", "first"))
    chk["rel_vs_revenue"] = (chk["sum_raw"] - chk["rev"]) / chk["rev"]
    chk["ratio_vs_volume"] = chk["sum_raw"] / chk["vol"]
    P("\n2) INTERPRETACE HODNOT PO STÁTECH (společné aukce 2012-2025 a nové formáty; n=%d aukcí)" % len(chk))
    P("   max |součet_po_státech - Total Revenue| / Total Revenue = %.3e" % chk["rel_vs_revenue"].abs().max())
    P("   součet_po_státech / Auction Volume: min %.2f, max %.2f (tj. NEjsou to tCO2, ale EUR)" % (chk["ratio_vs_volume"].min(), chk["ratio_vs_volume"].max()))
    P("   => hodnoty po státech = VÝNOS z aukce v EUR přiřazený státu; objem = výnos / aukční cena.")
    P("   max odchylka dopočteného objemu od celého čísla: %.2e" % dev)

    # objem: součet přes státy vs Auction Volume
    v = wide[["date", "auction_type", "certificate", "contract", "auction_volume_tco2", "volume_tco2_sum_states", "volume_diff_pct"]].copy()
    v["abs_diff_tco2"] = v["volume_tco2_sum_states"] - v["auction_volume_tco2"]
    P("\n3) SOUČET OBJEMŮ PŘES STÁTY vs Auction Volume (úspěšné aukce, n=%d)" % len(v))
    P("   max |odchylka| = %.3e %%, počet aukcí s |odchylkou| > 0.001 %%: %d" % (v["volume_diff_pct"].abs().max(), int((v["volume_diff_pct"].abs() > 1e-3).sum())))
    worst = v.reindex(v["volume_diff_pct"].abs().sort_values(ascending=False).index).head(10)
    P("   10 největších odchylek:\n" + worst.to_string(index=False))

    # počty po letech
    yc = []
    cancelled_by_year = canc.groupby("source_year").size() if len(canc) else pd.Series(dtype=int)
    for y in YEARS:
        w = wide[wide["year"] == y]
        yc.append(dict(year=y, source_rows=src_counts[y], cancelled=int(cancelled_by_year.get(y, 0)),
                       wide_rows=len(w), unique_days_wide=int(w["date"].nunique()),
                       unique_days_source=int(survey.set_index("year").loc[y, "unique_days"]),
                       rows_reconcile=(src_counts[y] - int(cancelled_by_year.get(y, 0)) == len(w))))
    yc = pd.DataFrame(yc)
    P("\n4) POČTY ŘÁDKŮ PO LETECH (zdroj vs wide)")
    P(yc.to_string(index=False))
    P("   všech 14 roků přítomno: %s; všechny sedí (zdroj - zrušené = wide): %s" % (sorted(files) == YEARS, bool(yc["rows_reconcile"].all())))
    P("   denní soubor: %d dní; wide: %d řádků; součet objemů daily vs wide: %s" % (
        len(daily), len(wide), np.isclose(daily["auction_volume_tco2"].sum(), wide["auction_volume_tco2"].sum())))
    P("   součet objemů po státech daily vs wide: %s" % np.isclose(daily[vcols].sum().sum(), wide[vcols].sum().sum()))

    P("\n5) KÓDY STÁTŮ (%d): %s" % (len(states), ", ".join(states)))
    P("   bez názvu v mapě: %s" % (unknown or "žádné"))
    P("   kódy mimo EU-27/EEA, které je třeba vysvětlit: XA (není v Caption), XF (Flandry, Caption), IF/MF/MX/IX/SF/XI = fondy / Severní Irsko")

    # historie států
    hist = []
    for s in states:
        nz = long[(long["country_code"] == s) & (long["volume_tco2"] > 0)]
        hist.append(dict(code=s, name=CODE_NAMES.get(s, "?"), first_nonzero=nz["date"].min(), last_nonzero=nz["date"].max(),
                         n_auctions=len(nz), total_volume_tco2=nz["volume_tco2"].sum()))
    hist = pd.DataFrame(hist)
    P("\n6) HISTORIE STÁTŮ (první a poslední nenulový den)")
    P(hist.to_string(index=False))
    if "HR" in states:
        P("   Chorvatsko první nenulový: %s" % hist.set_index("code").loc["HR", "first_nonzero"].date())
    P("   VB/UK/GB: kód %s v datech" % ("NALEZEN" if any(c in states for c in ("UK", "GB")) else "NEnalezen (EEX report neobsahuje VB; neodhaduje se)"))

    # spot checky: 5 řádků z každého období, nezávislé čtení původní buňky (xlrd/openpyxl)
    spot = spot_checks(wide, long, files, survey)
    P("\n7) NAMÁTKOVÁ KONTROLA proti zdroji (5 řádků na období, seed=0)")
    P(spot.to_string(index=False))
    P("   shoda: %d/%d" % (int(spot["match"].sum()), len(spot)))

    # duplicity / zrušené
    P("\n8) ZRUŠENÉ AUKCE (vyřazeny z wide/long; objem nebyl přidělen)")
    P(canc[["date", "auction_type", "contract", "status", "auction_volume_tco2"]].to_string(index=False) if len(canc) else "   žádné")
    P("\n9) Duplicity klíče date x auction_type x certificate: %d řádků" % n_dup3)
    if n_dup3:
        P(auctions[auctions.duplicated(k3, keep=False)].to_string(index=False))

    # ---- zápis ----
    def w(df, name, **kw):
        df.to_csv(os.path.join(out, name), index=False, encoding="utf-8-sig", **kw)

    w(long, OUT_LONG_CSV, date_format="%Y-%m-%d")
    w(wide, OUT_WIDE_CSV, date_format="%Y-%m-%d")
    w(daily, OUT_DAILY_CSV, date_format="%Y-%m-%d")
    w(wide_rev, OUT_REVENUE_CSV, date_format="%Y-%m-%d")
    w(canc[["date", "source_year", "auction_name_raw", "auction_type", "contract", "certificate", "status",
            "auction_volume_tco2", "source_file", "source_row"]], OUT_CANCELLED_CSV, date_format="%Y-%m-%d")
    w(survey, OUT_SURVEY_CSV)
    w(v.sort_values(["date", "auction_type"]), OUT_VALID_CSV, date_format="%Y-%m-%d")
    w(yc, OUT_YEAR_CSV)
    w(hist, OUT_STATES_CSV, date_format="%Y-%m-%d")
    w(spot, OUT_SPOT_CSV)
    with pd.ExcelWriter(os.path.join(out, OUT_WIDE_XLSX), engine="openpyxl", datetime_format="yyyy-mm-dd", date_format="yyyy-mm-dd") as xw:
        wide.to_excel(xw, sheet_name="wide", index=False)
        daily.to_excel(xw, sheet_name="daily", index=False)
        wide_rev.to_excel(xw, sheet_name="wide_revenue_eur", index=False)
        for sh in xw.book.worksheets:
            sh.freeze_panes = "B2" if sh.title == "daily" else "E2"
    with open(os.path.join(out, OUT_REPORT), "w", encoding="utf-8") as f:
        f.write("\n".join(rep))
    log.info("\n" + "\n".join(rep))
    log.info("HOTOVO: wide %s, daily %s, long %s; výstup: %s", wide.shape, daily.shape, long.shape, out)


def spot_checks(wide, long, files, survey, n=5, seed=0):
    """Pro náhodné řádky z každého období znovu přečte surové buňky zdroje (xlrd/openpyxl) a porovná s wide."""
    periods = {"2012-2015 (old, Details s tečkami)": (2012, 2015), "2016 (old, plain)": (2016, 2016),
               "2017-2019 (new)": (2017, 2019), "2020-2025 (new, Status)": (2020, 2025)}
    res = []
    for pname, (a, b) in periods.items():
        sub = wide[(wide["year"] >= a) & (wide["year"] <= b) & (wide["volume_tco2_sum_states"] > 0)]
        samp = sub.sample(n=min(n, len(sub)), random_state=seed)
        for _, r in samp.iterrows():
            y = int(r["year"])
            lrow = long[(long["date"] == r["date"]) & (long["auction_type"] == r["auction_type"]) & (long["contract"] == r["contract"])]
            srow = int(lrow["source_row"].iloc[0])
            cells, hdr = raw_row(files[y], srow)
            ok_all, nchk = True, 0
            nh0 = [norm(h) for h in hdr]
            vcell = cells[next(i for i, x in enumerate(nh0) if x.startswith("auction volume"))]
            ok_all &= abs(float(vcell) - r["auction_volume_tco2"]) < 1e-6
            if any("(" in str(h) for h in hdr):                          # nový formát: výnos ve sloupci / cena
                price = float(cells[[norm(h) for h in hdr].index(next(x for x in [norm(h) for h in hdr] if x.startswith("auction price")))])
                for h, c in zip(hdr, cells):
                    m = re.search(r"\(([A-Z]{2})\)", str(h))
                    if m and c not in (None, ""):
                        exp = float(c) / price
                        got = r.get(f"volume_tco2_{m.group(1)}", 0.0)
                        ok_all &= abs(exp - got) < 1e-3
                        nchk += 1
            else:
                nh = [norm(h) for h in hdr]
                det = cells[nh.index("auction details")]
                price = float(cells[[i for i, x in enumerate(nh) if x.startswith("auction price")][0]])
                if isinstance(det, str) and det.strip():
                    for tok in re.split(r"[\n,;]+", det.strip()):
                        if ":" in tok:
                            code, val = tok.split(":", 1)
                            exp = float(val.strip().replace(".", "")) / price
                            ok_all &= abs(exp - r.get(f"volume_tco2_{code.strip()}", 0.0)) < 1e-3
                            nchk += 1
                else:
                    ok_all &= abs(r.get(f"volume_tco2_{r['auction_type']}", 0.0) - r["auction_volume_tco2"]) < 1e-6
                    nchk = 1
            res.append(dict(period=pname, date=r["date"].date(), auction_type=r["auction_type"], contract=r["contract"],
                            source_file=os.path.basename(files[y]), source_row=srow, n_state_cells_checked=nchk, match=bool(ok_all)))
    return pd.DataFrame(res)


def raw_row(path, excel_row):
    """Nezávislé čtení jednoho řádku + hlavičky přímo přes xlrd / openpyxl."""
    if path.lower().endswith(".xls"):
        import xlrd
        wb = xlrd.open_workbook(path)
        sh = next(s for s in wb.sheets() if any(norm(c.value) == "date" for r in range(min(10, s.nrows)) for c in s.row(r)))
        hr = next(r for r in range(sh.nrows) if any(norm(c.value) == "date" for c in sh.row(r)))
        return [c.value for c in sh.row(excel_row - 1)], [c.value for c in sh.row(hr)]
    import openpyxl
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    for ws in wb.worksheets:
        rows = list(ws.iter_rows(values_only=True))
        for hr, rr in enumerate(rows[:20]):
            if any(norm(c) == "date" for c in rr if c is not None):
                return list(rows[excel_row - 1]), list(rr)
    raise ValueError("hlavička nenalezena")


if __name__ == "__main__":
    main()
