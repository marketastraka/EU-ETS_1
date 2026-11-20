import pandas as pd
data_path = r"G:\My Drive\EU_ETS1\emission-spot-primary-market-auction-report-2012-2025-data\emission-spot-primary-market-auction-report-2012-2025-data"
data_name_to2019 = "emission-spot-primary-market-auction-report-{year}-data.xls"
data_name_from2020 = "emission-spot-primary-market-auction-report-{year}-data.xlsx"


# 1) Country code -> target column label
code_labels = {
    "AT": "Austria\n(AT)", "BE": "Belgium\n(BE)", "BG": "Bulgaria\n(BG)",
    "CY": "Cyprus\n(CY)", "CZ": "Czech Republic\n(CZ)", "DE": "Germany\n(DE)",
    "DK": "Denmark\n(DK)", "EE": "Estonia\n(EE)", "EL": "Greece\n(EL)",
    "ES": "Spain\n(ES)", "FI": "Finland\n(FI)", "FR": "France\n(FR)",
    "HR": "Croatia\n(HR)", "HU": "Hungary\n(HU)", "IE": "Ireland\n(IE)",
    "IF": "Innovation Fund\n(IF)", "IS": "Iceland\n(IS)", "IT": "Italy\n(IT)",
    "IX": "InnoFund RRF\n(IX)", "LI": "Liechtenstein\n(LI)", "LT": "Lithuania\n(LT)",
    "LU": "Luxembourg\n(LU)", "LV": "Latvia\n(LV)", "MF": "Modernisation Fund\n(MF)",
    "MT": "Malta\n(MT)", "MX": "MS RRF\n(MX)", "NL": "Netherlands\n(NL)",
    "NO": "Norway\n(NO)", "PL": "Poland\n(PL)", "PT": "Portugal\n(PT)",
    "RO": "Romania\n(RO)", "SE": "Sweden\n(SE)", "SI": "Slovenia\n(SI)",
    "SK": "Slovakia\n(SK)", "XI": "Northern Ireland\n(XI)",
}

# 2) Parse the details string into a dict {code: value}
def parse_details(s):
    if pd.isna(s) or not str(s).strip():
        return {}
    out = {}
    for part in str(s).split(","):
        if ":" in part:
            k, v = part.split(":", 1)
            out[k.strip()] = pd.to_numeric(v.strip(), errors="coerce")
    return out

details_col = "Auction Details"

def load_auction_data_2012_2015(year):
    frames = []

    df = pd.read_excel(f"{data_path}\\{data_name_to2019.format(year=year)}", skiprows=2)
    country_df = pd.DataFrame(df[details_col].apply(parse_details).tolist(), index=df.index)

    # Keep all possible countries in a fixed order (missing ones become NaN; use .fillna(0) if you prefer 0)
    country_df = country_df.reindex(columns=list(code_labels.keys()))
    country_df = country_df.rename(columns=code_labels)

    # Safety check: any codes in the data that aren't in the mapping?
    unknown = set().union(*df[details_col].apply(parse_details).apply(dict.keys)) - set(code_labels)
    if unknown:
        print("Codes not in mapping:", unknown)

    # 3) Rename the other columns to the target names
    rename_map = {
        "Minimal Price €/tCO2": "Minimum Bid €/tCO2",
        "Maximal Price €/tCO2": "Maximum Bid €/tCO2",
        "Total Revenue €": "Total Revenue €",
    }
    result = df.drop(columns=[details_col]).rename(columns=rename_map)

    # 4) Add the target columns that don't exist in the source (empty for now)
    extra_cols = ["Number of bids submitted"]
    for c in extra_cols:
        result[c] = pd.NA

    # 5) Combine and put columns in the target order
    out = pd.concat([result, country_df], axis=1)

    target_order = [
        "Date", "Time", "Auction Name", "Contract",
        "Auction Price €/tCO2", "Minimum Bid €/tCO2", "Maximum Bid €/tCO2",
        "Mean €/tCO2", "Median €/tCO2", "Auction Volume tCO2",
        "Total Amount of Bids ", "Cover Ratio",
        "Total Number of Bidders", "Number of Successful Bidders",
        "Total Revenue €", "Country",
    ] + list(code_labels.values())

    frames.append(out[target_order])

    return pd.concat(frames, ignore_index=True)


# short format loading data
data_2012 = pd.read_excel(f"{data_path}\\{data_name_to2019.format(year=2012)}", skiprows=2)
data_2012 = load_auction_data_2012_2015(2012)
data_2016 = load_auction_data_2012_2015(2016)

# long format loading data
data_2017 = pd.read_excel(f"{data_path}\\{data_name_to2019.format(year=2017)}", skiprows=5)
data_2024 = pd.read_excel(f"{data_path}\\{data_name_from2020.format(year=2024)}", skiprows=5)
data_2024.columns

