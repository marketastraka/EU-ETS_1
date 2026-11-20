import pandas as pd
data = pd.read_excel("G:\\My Drive\\EU_ETS1\\verified_emissions_2025_en.xlsx", sheet_name="data", skiprows=2)
data.columns

cz = data[data["REGISTRY_CODE"] == "CZ"]
cz
data
data = data.rename(columns={"ALLOCATION2008": "ALLOCATION_2008"})

year_cols = [c for c in data.columns if c.endswith(tuple(str(y) for y in range(2000, 2030)))]
id_cols = [c for c in data.columns if c not in year_cols]

df = data.set_index(id_cols)
df
df.columns = pd.MultiIndex.from_tuples(
    [tuple(c.rsplit("_", 1)) for c in df.columns],
    names = ["metric", "year"]
)

df_long = df.stack("year").reset_index()
df_long
df_long.columns
df_long = df_long.dropna(subset=['ALLOCATION', 'ALLOCATION_RESERVE',
       'ALLOCATION_TRANSITIONAL', 'ALLOCATION_FEETS', 'ALLOCATION_BYICELAND',
       'CH_ALLOCATION', 'VERIFIED_EMISSIONS', 'CH_VERIFIED_EMISSIONS'], how='all')

df_long = df_long[df_long['VERIFIED_EMISSIONS']!='Excluded']

metric_cols = ['ALLOCATION', 'ALLOCATION_RESERVE', 'ALLOCATION_TRANSITIONAL',
               'ALLOCATION_FEETS', 'ALLOCATION_BYICELAND',
               'CH_ALLOCATION', 'VERIFIED_EMISSIONS', 'CH_VERIFIED_EMISSIONS']

# nejdřív převést na čísla (text typu "Excluded" se změní na NaN)
df_long[metric_cols] = df_long[metric_cols].apply(pd.to_numeric, errors="coerce")

df_long = df_long[(df_long[metric_cols] > 1).any(axis=1)]

df_long["pomer"] = (df_long["ALLOCATION"] + df_long["ALLOCATION_RESERVE"] + df_long["ALLOCATION_TRANSITIONAL"] + df_long["ALLOCATION_FEETS"] + df_long["ALLOCATION_BYICELAND"] + df_long["CH_ALLOCATION"]) / (df_long["VERIFIED_EMISSIONS"] + df_long["CH_VERIFIED_EMISSIONS"])

NA = df_long[df_long["VERIFIED_EMISSIONS"].isna()]
NA

cz = df_long[df_long["REGISTRY_CODE"] == "CZ"]
cz.columns
