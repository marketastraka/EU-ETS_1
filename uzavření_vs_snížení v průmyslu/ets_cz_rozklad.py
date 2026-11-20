# -*- coding: utf-8 -*-
"""
Emise českých stacionárních zařízení v EU ETS1: kolik poklesu jde na vrub
zavírání zařízení a kolik snižování emisí v provozech, které dál fungují.

Spuštění:
    python ets_cz_rozklad.py [cesta_k_xlsx] [vystupni_slozka]

Výchozí vstup: G:\\My Drive\\EU_ETS1\\verified_emissions_2025_en.xlsx
Výstup: ets_cz_rozklad.png (300 dpi), .svg, ets_cz_rozklad_podklady.csv,
        ets_cz_citlivost.csv, ets_cz_top10_zavrena.csv, ets_cz_zarizeni_klasifikace.csv

==========================================================================
POZNÁMKY K DATŮM (ověřeno na souboru, extrakce EUTL z 2026-04-01)
==========================================================================
1) Soubor NENÍ v „long“ formátu, ale „wide“: sloupce VERIFIED_EMISSIONS_<rok>,
   ALLOCATION_<rok> atd. Sloupce `year` a `pomer` v něm nejsou (pomer nepoužíváme).
   Tabulku proto přetavujeme (melt) do long formátu: 1 řádek = zařízení x rok.
2) IDENTIFIER_IN_REG NENÍ identifikátor zařízení, ale název držitele (firmy);
   dvě zařízení mohou mít stejného držitele. Stabilní unikátní klíč je
   INSTALLATION_IDENTIFIER (v CZ unikátní, bez NaN). Ten používáme jako klíč.
3) REGISTRY_CODE == "CZ"; ACCOUNT_TYPE: "OPERATOR HOLDING ACCOUNT" (stacionární)
   vs. "AIRCRAFT OPERATOR HOLDING ACCOUNT" (letectví, kód 10). Stacionární = první.
   Kódy 50 (námoří) a 60 (ETS2) v CZ nejsou; pro jistotu se vyřazují také.
4) Read Me listu: prázdná buňka = nebyl přidělen příděl / nebyly ověřeny emise,
   0 = skutečná nula, "Excluded" = zařízení je pro daný rok MIMO ROZSAH ETS.
   "Excluded" je tedy změna rozsahu (opt-out malých zařízení od 2021), ne zavření.
5) Data jsou srovnatelná až od 2013 (rozšíření rozsahu ve fázi III); zařízení
   zaniklá před 2013 v souboru nejsou. Proto je výchozí rok 2013 (2005 není k dispozici).
6) CH_* sloupce se týkají švýcarského propojení, jen letectví; u stacionárních
   zařízení CZ se ověřuje, že jsou prázdné (assert níže).
8) Ledvice 4 (id 424) a „ČEZ, a. s. - Elektrárna Ledvice“ (id 121) se berou jako JEDNO zařízení
   (společnost): emise i příděly se sečtou (viz SLOUCENI). Důvod: v roce 2021 emise Ledvic 4
   zmizí (účet uzavřen 23. 12. 2021) a ve stejném roce vzrostou emise Ledvic o 3,2 Mt; data
   tak silně naznačují administrativní sloučení v registru. Bez sečtení by Ledvice 4 vypadaly
   jako největší zavřené zařízení a Ledvice jako nárůst emisí, což by zkreslilo rozklad.
   Sloučené Ledvice jsou zařazeny jako ROZŠÍŘENÍ KAPACITY existujícího zařízení (blok Ledvice 4
   rozšířil elektrárnu), tj. v kategorii „pokračující“, ne jako nové zařízení ani zavření.
7) ALLOCATION_FEETS / ALLOCATION_BYICELAND existují jen pro rok 2025 a u CZ
   stacionárních zařízení se ověřuje, jestli jsou nenulové (print níže).
   Pro analýzu se používá jen ALLOCATION.

==========================================================================
METODIKA ROZKLADU (pro každé t = 2014..2025; E = ověřené emise)
==========================================================================
E_t: NaN a "Excluded" se pro součet berou jako 0 (chybějící hlášení = žádné
ověřené emise v souboru), ale druh chybějící hodnoty se uchovává zvlášť.
A_t: příděl; NaN = 0 (v Read Me: prázdné = nebyl přidělen příděl).
Změna dE_t = E_t - E_{t-1} se rozdělí zařízení po zařízení na:

 - ZAVŘENÁ:       E_{t-1} > 0 a E_t == 0 (nebo hlášení chybí), zařízení není v t
                  „Excluded“. Počítají se sem i zařízení bez bezplatného přídělu
                  (např. elektrárny) a případy, kdy příděl neklesl
                  o 15 % (dříve skupina „nejasné“). Příspěvek = -E_{t-1}.
                  Původní přísnější kritérium (příděl klesl o víc než práh) zůstává
                  jen v citlivostní analýze (mode="alokace").
 - NOVÁ:          zařízení, které v žádném dřívějším roce nemělo E > 0 a v t má
                  E > 0. Příspěvek = +E_t.
 - MIMO ROZSAH:   E_{t-1} > 0 a v t je "Excluded" (opt-out). Příspěvek = -E_{t-1}.
                  Nejde ani o úsporu, ani o zavření -> vykazuje se zvlášť.
 - NEJASNÉ:       v hlavním rozkladu prázdná skupina (vše zastavené je zavřené);
                  vzniká jen v přísném kritériu přes příděl (citlivostní analýza).
                  Rizika: dočasná odstávka, sloučení, přečíslování – viz top 10.
 - POKRAČUJÍCÍ:   zbytek: E_t - E_{t-1}. Zahrnuje i zařízení, která se po
                  mezeře vrátila (E_{t-1}=0 -> E_t>0, ne poprvé), a zařízení, jejichž
                  emise klesly, ale nedosáhly 0.
Součet všech složek přesně odpovídá změně celkových emisí (assert).

Kumulativní pohled vůči roku 2013: součet ročních příspěvků od 2014 do t.
Výchozí rok 2013 je vhodnější než 2005 (data před 2013 nejsou srovnatelná
ani dostupná, viz bod 5), a zároveň je to začátek fáze III.

Zařízení, která z datasetu úplně zmizí (řádek chybí): v tomto souboru řádky
nechybí (každé zařízení má řádek, jen s prázdnými hodnotami); chybějící
hodnoty se proto řeší jako NaN -> 0 a vstupují do logiky výše (ZAVŘENÁ/NEJASNÉ).

Změny rozsahu: nová zařízení s činností zařazenou později (např. spalování
odpadu od 2024) jsou v NOVÁ; v souboru se vypíšou po rocích s kódem činnosti,
aby šlo posoudit, co je skutečná nová kapacita a co jen změna rozsahu.

Rok 2025: extrakce k 1. 4. 2026, krátce po termínu ověřených hlášení (31. 3.);
hlášení mohou být neúplná/předběžná. Skript vypíše počet zařízení s emisemi
v 2025 vs. 2024 a počet zařízení, která se v 2025 nově objevují jako „nehlášená".
"""
import re
import sys
import numpy as np
import pandas as pd
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from pathlib import Path

XLSX = sys.argv[1] if len(sys.argv) > 1 else r"G:\My Drive\EU_ETS1\verified_emissions_2025_en.xlsx"
OUT = Path(sys.argv[2]) if len(sys.argv) > 2 else Path(".")
OUT.mkdir(parents=True, exist_ok=True)

YEARS = list(range(2013, 2026))
BASE = 2013
THR_MAIN = 0.15
CATS = ["zavrena", "nova", "pokracujici", "nejasne", "mimo_rozsah"]

# Zařízení, která se berou jako jedno (sečtou se emise i příděly): {zdroj: cíl}
# (INSTALLATION_IDENTIFIER). Ledvice 4 (id 424) + "ČEZ, a. s. - Elektrárna Ledvice" (id 121).
SLOUCENI = {424: 121}
SLOUCENY_NAZEV = {121: "ČEZ - Elektrárna Ledvice (včetně Ledvic 4)"}

# číselník EUTL (list "activity codes", sloupec "aligned") -> srozumitelné sektory.
# Staré kódy 1-9 (fáze II) odpovídají novým 20+ (kód 6 sdružuje cement i vápno).
SEKTOR = {
    1: "Energetika a teplo", 20: "Energetika a teplo",
    2: "Rafinerie", 21: "Rafinerie",
    3: "Koks", 22: "Koks",
    4: "Železo a ocel", 5: "Železo a ocel", 24: "Železo a ocel",
    25: "Železo a ocel",
    6: "Cement a vápno", 29: "Cement a vápno", 30: "Cement a vápno",
    7: "Sklo", 31: "Sklo",
    8: "Keramika", 32: "Keramika",
    9: "Papír a celulóza", 35: "Papír a celulóza", 36: "Papír a celulóza",
    27: "Neželezné kovy", 26: "Neželezné kovy", 28: "Neželezné kovy",
    33: "Minerální vlna a sádra", 34: "Minerální vlna a sádra",
    38: "Chemie", 42: "Chemie", 41: "Chemie", 37: "Chemie", 39: "Chemie",
    99: "Jiné (opt-in)",
}


# ---------------------------------------------------------------- 1) načtení
def sluc(cz):
    """Sečte emise a příděly zařízení ze SLOUCENI do cílového zařízení; zdrojový řádek se odstraní."""
    cz = cz.copy()
    cols = [c for c in cz.columns if re.match(r"^(VERIFIED_EMISSIONS|ALLOCATION\w*)_\d{4}$", str(c))]
    for src, dst in SLOUCENI.items():
        assert src in cz.index and dst in cz.index, "Sloučená zařízení nejsou v datech"
        vals = cz.loc[[src, dst], cols]
        assert not vals.astype(str).apply(lambda x: x.str.strip().str.lower().eq("excluded")).any().any()
        cz.loc[dst, cols] = vals.apply(pd.to_numeric, errors="coerce").sum(min_count=1)   # NaN + NaN zůstane NaN
        if dst in SLOUCENY_NAZEV:
            cz.loc[dst, "INSTALLATION_NAME"] = SLOUCENY_NAZEV[dst]
        print(f"Sloučeno: id {src} -> id {dst} ({SLOUCENY_NAZEV.get(dst, '')})")
        cz = cz.drop(index=src)
    return cz


def nacti():
    raw = pd.read_excel(XLSX, sheet_name="data", header=2)
    cz = raw[raw["REGISTRY_CODE"] == "CZ"].copy()
    print("Unikátní ACCOUNT_TYPE v CZ:", cz["ACCOUNT_TYPE"].value_counts().to_dict())
    cz = cz[cz["ACCOUNT_TYPE"].str.upper().str.startswith("OPERATOR")]
    cz = cz[~cz["MAIN_ACTIVITY_TYPE_CODE"].isin([10, 50, 60])]
    assert cz["INSTALLATION_IDENTIFIER"].is_unique and cz["INSTALLATION_IDENTIFIER"].notna().all()
    ch = [c for c in cz.columns if c.startswith("CH_")]
    assert cz[ch].isna().all().all() or print("POZOR: CH sloupce nejsou prázdné") is None
    fe = ["ALLOCATION_FEETS_2025", "ALLOCATION_BYICELAND_2025"]
    print("FEETS/BYICELAND nenulové u CZ stacionárních:",
          {c: int((pd.to_numeric(cz[c], errors="coerce").fillna(0) != 0).sum()) for c in fe})
    cz = cz.set_index("INSTALLATION_IDENTIFIER", drop=False)
    cz = sluc(cz)
    print(f"Zařízení: {len(cz)}, roky {YEARS[0]}-{YEARS[-1]}")
    return cz


def matice(cz):
    """Z wide formátu: E (ověřené emise), A (příděl), EXCL (příznak Excluded)."""
    raw = pd.DataFrame({y: cz[f"VERIFIED_EMISSIONS_{y}"] for y in YEARS})
    EXCL = raw.apply(lambda s: s.astype(str).str.strip().str.lower() == "excluded")
    E = raw.apply(pd.to_numeric, errors="coerce")
    A = pd.DataFrame({y: pd.to_numeric(cz[f"ALLOCATION_{y}"], errors="coerce") for y in YEARS})
    return E, A, EXCL


# ---------------------------------------------------------------- 2) rozklad
def rozklad(E, A, EXCL, cz, thr=0.15, mode="vse"):
    """
    mode: 'vse'      = HLAVNÍ definice: zavřeno = emise na 0/chybí (kromě Excluded)
          'alokace'  = původní přísné kritérium (pokles přídělu > thr), zbytek „nejasné“
          'posledni' = zavřeno, pokud t-1 je poslední rok s E>0 (trvalé ukončení v datech)
          'closure'  = jako 'posledni' a ACCOUNT_CLOSURE != OPEN
    Vrací (agregace po letech, klasifikace po zařízeních a letech).
    """
    E0, A0 = E.fillna(0.0), A.fillna(0.0)
    ever_before = (E0 > 0).cumsum(axis=1).shift(1, axis=1).fillna(0) > 0
    pos = E0 > 0
    last_pos = pos.apply(lambda r: max([y for y in YEARS if r[y]] or [np.nan]), axis=1)
    closure_set = cz["ACCOUNT_CLOSURE"].astype(str).str.upper().ne("OPEN")
    rows = []
    for t in YEARS[1:]:
        p = t - 1
        d = E0[t] - E0[p]
        stopped = (E0[p] > 0) & (E0[t] == 0)
        excl = stopped & EXCL[t]
        if mode == "vse":
            closed = stopped & ~excl
        elif mode == "alokace":
            drop = (A0[p] > 0) & ((A0[p] - A0[t]) / A0[p].replace(0, np.nan) > thr)
            closed = stopped & ~excl & drop
        elif mode == "posledni":
            closed = stopped & ~excl & (last_pos == p)
        else:
            closed = stopped & ~excl & (last_pos == p) & closure_set
        unclear = stopped & ~excl & ~closed
        new = (E0[t] > 0) & ~ever_before[t]
        cont = ~(closed | unclear | excl | new)
        cat = pd.Series("pokracujici", index=E.index)
        cat[new] = "nova"; cat[closed] = "zavrena"; cat[unclear] = "nejasne"; cat[excl] = "mimo_rozsah"
        rows.append(pd.DataFrame({"rok": t, "kat": cat, "zmena": d, "E_prev": E0[p], "E_t": E0[t]}))
    long = pd.concat(rows).reset_index()
    agg = long.pivot_table(index="rok", columns="kat", values="zmena", aggfunc="sum", fill_value=0.0)
    for k in CATS:
        if k not in agg:
            agg[k] = 0.0
    agg = agg[CATS]
    agg["celkem_zmena"] = E0.sum().diff().reindex(agg.index)
    agg["emise_celkem"] = E0.sum().reindex(agg.index)
    agg["n_zavrena"] = long[long.kat == "zavrena"].groupby("rok").size().reindex(agg.index).fillna(0).astype(int)
    agg["n_nejasne"] = long[long.kat == "nejasne"].groupby("rok").size().reindex(agg.index).fillna(0).astype(int)
    # kontrola: složky přesně sedí na celkovou změnu
    soucet = agg[CATS].sum(axis=1)
    assert np.allclose(soucet, agg["celkem_zmena"], atol=1e-6), "Složky nesedí na celkovou změnu!"
    return agg, long


def kumulativne(agg):
    """Kumulativní změna proti 2013 a podíl složek na kumulativním poklesu."""
    k = agg[CATS].cumsum()
    k["celkem"] = k.sum(axis=1)
    assert np.allclose(k["celkem"], agg["emise_celkem"] - agg["emise_celkem"].iloc[0] + agg["celkem_zmena"].iloc[0], atol=1e-6)
    return k


# ------------------------------------------------- 3) citlivostní analýza
def citlivost(E, A, EXCL, cz):
    var = [("HLAVNÍ: zavřeno = všechna zastavená (emise na 0/chybí)", dict(mode="vse")),
           ("jen příděl: práh 10 %", dict(thr=0.10, mode="alokace")),
           ("jen příděl: práh 15 %", dict(thr=0.15, mode="alokace")),
           ("jen příděl: práh 25 %", dict(thr=0.25, mode="alokace")),
           ("zavřeno = poslední rok s emisemi > 0 v datech", dict(mode="posledni")),
           ("zavřeno = poslední rok s emisemi > 0 a ACCOUNT_CLOSURE != OPEN", dict(mode="closure"))]
    out = []
    for nm, kw in var:
        a, _ = rozklad(E, A, EXCL, cz, **kw)
        c = a[CATS].sum()
        pokles = -(c.sum())
        r = {"varianta": nm, "pokles_celkem_Mt": pokles / 1e6}
        for k in c.index:
            r[f"{k}_Mt"] = c[k] / 1e6
        # podíl na hrubém poklesu: složky se zápornou hodnotou / celkový pokles
        for k in ["zavrena", "nejasne", "pokracujici"]:
            r[f"podil_{k}_%"] = 100 * -c[k] / pokles
        r["podil_nova_%"] = 100 * -c["nova"] / pokles   # záporný = kompenzuje pokles
        out.append(r)
    return pd.DataFrame(out).round(2)


# ------------------------------------------- 4) top 10 zavřených zařízení
def top10(long, cz):
    z = long[long.kat == "zavrena"].copy()
    z = z.join(cz[["INSTALLATION_NAME", "IDENTIFIER_IN_REG", "MAIN_ACTIVITY_TYPE_CODE", "ACCOUNT_CLOSURE"]], on="INSTALLATION_IDENTIFIER")
    z["sektor"] = z["MAIN_ACTIVITY_TYPE_CODE"].map(SEKTOR).fillna("neznámý kód " + z["MAIN_ACTIVITY_TYPE_CODE"].astype(str))
    nova = long[long.kat == "nova"].join(cz[["IDENTIFIER_IN_REG"]], on="INSTALLATION_IDENTIFIER")
    poz = []
    for _, r in z.iterrows():
        pozn = []
        s = nova[(nova.IDENTIFIER_IN_REG == r.IDENTIFIER_IN_REG) & (nova.rok.between(r.rok - 1, r.rok + 1))]
        if len(s):
            pozn.append("stejný držitel má v okolních letech nové zařízení – možné přečíslování/převod")
        if str(r.ACCOUNT_CLOSURE).upper() == "OPEN":
            pozn.append("účet je stále OPEN – možná jen odstávka")
        poz.append("; ".join(pozn))
    z["poznamka"] = poz
    z = z.sort_values("E_prev", ascending=False).head(10)
    return pd.DataFrame({
        "název": z.INSTALLATION_NAME.values, "držitel": z.IDENTIFIER_IN_REG.values, "sektor": z.sektor.values,
        "emise v posledním roce provozu [t]": z.E_prev.round(0).values,
        "poslední rok provozu": (z.rok - 1).values, "rok zavření (první rok bez emisí)": z.rok.values,
        "ACCOUNT_CLOSURE": z.ACCOUNT_CLOSURE.values, "poznámka": z.poznamka.values})


# --------------------------------------------------------------- 5) graf
COL = {"zavrena": "#D55E00", "nova": "#009E73", "pokracujici": "#0072B2", "nejasne": "#8C8C8C"}
LAB = {"zavrena": "Zavřená zařízení (emise klesly na 0)", "nova": "Nová zařízení",
       "pokracujici": "Pokračující zařízení (skutečné snižování)", "nejasne": "Mimo rozsah ETS"}
INK, INK2, GRID = "#1f1f1f", "#52514e", "#e5e4e0"
matplotlib.rcParams.update({"font.family": "DejaVu Sans", "svg.fonttype": "none"})


def graf(agg, k, shares, nejasne_top, soubor):
    # sloučení „nejasné“ a „mimo rozsah“ pro přehlednost
    a = agg.copy(); a["nejasne"] = a["nejasne"] + a["mimo_rozsah"]
    kk = k.copy(); kk["nejasne"] = kk["nejasne"] + kk["mimo_rozsah"]
    cats = [c for c in ["pokracujici", "zavrena", "nejasne", "nova"] if c == "nova" or a[c].abs().sum() > 5e4]  # <0,05 Mt = zanedbatelné, nezobrazuje se
    yrs = list(a.index)
    fig = plt.figure(figsize=(11.5, 15.2), dpi=100)
    gs = fig.add_gridspec(3, 1, height_ratios=[1.05, 1.35, 1.25], hspace=0.42,
                          left=0.085, right=0.965, top=0.800, bottom=0.095)
    # titulek
    fig.text(0.085, 0.965, shares["titulek"], fontsize=19, fontweight="bold", color=INK, va="top", wrap=True)
    fig.text(0.085, 0.905, shares["podtitulek"], fontsize=11.5, color=INK2, va="top")

    def styl(ax):
        for s in ("top", "right", "left"):
            ax.spines[s].set_visible(False)
        ax.spines["bottom"].set_color("#bdbcb8")
        ax.tick_params(colors=INK2, length=0, labelsize=10.5)
        ax.grid(axis="y", color=GRID, lw=0.8); ax.set_axisbelow(True)

    # panel 1: celkové emise (jedna osa, samostatný panel)
    ax = fig.add_subplot(gs[0]); styl(ax)
    tot = pd.concat([pd.Series({BASE: a["emise_celkem"].iloc[0] - a["celkem_zmena"].iloc[0]}), a["emise_celkem"]]) / 1e6
    ax.plot(tot.index, tot.values, color=INK, lw=2.4, marker="o", ms=6, mfc="white", mew=2, zorder=3)
    for y in (BASE, 2019, 2021, 2025):
        ax.annotate(f"{tot[y]:.1f}", (y, tot[y]), textcoords="offset points", xytext=(0, 10), ha="center",
                    fontsize=10.5, color=INK, fontweight="bold")
    ax.set_ylim(0, tot.max() * 1.18); ax.set_xlim(BASE - .5, 2025.5)
    ax.set_xticks(range(BASE, 2026)); ax.set_ylabel("Mt CO₂e", color=INK2)
    ax.set_title("Celkové ověřené emise českých stacionárních zařízení v EU ETS1", loc="left", fontsize=12.5, color=INK, pad=10)

    # panel 2: meziroční změna rozložená na složky (diverging stacked bars)
    ax = fig.add_subplot(gs[1]); styl(ax)
    pos_b = np.zeros(len(yrs)); neg_b = np.zeros(len(yrs))
    for c in cats:
        v = a[c].values / 1e6
        base = np.where(v >= 0, pos_b, neg_b)
        ax.bar(yrs, v, bottom=base, width=0.68, color=COL[c], label=LAB[c], edgecolor="white", linewidth=1.2,
               hatch="////" if c == "nejasne" else None, zorder=3)
        pos_b += np.where(v >= 0, v, 0); neg_b += np.where(v < 0, v, 0)
    ax.plot(yrs, a["celkem_zmena"].values / 1e6, ls="none", marker="D", ms=7, color=INK, mfc="white", mew=1.8, zorder=5,
            label="Celková meziroční změna")
    ax.axhline(0, color=INK2, lw=1)
    ax.set_ylim(min(neg_b.min(), a["celkem_zmena"].min() / 1e6) * 1.22, pos_b.max() * 1.12)
    ax.set_xlim(BASE - .5, 2025.5); ax.set_xticks(range(BASE, 2026)); ax.set_ylabel("Meziroční změna, Mt CO₂e", color=INK2)
    ax.set_title("Meziroční změna emisí podle toho, kde vznikla", loc="left", fontsize=12.5, color=INK, pad=10)
    ax.legend(loc="lower left", fontsize=9.6, frameon=False, ncol=1, labelcolor=INK2)
    for y in yrs:
        if abs(a.loc[y, "celkem_zmena"]) > 3e6:
            ax.annotate(f"{a.loc[y,'celkem_zmena']/1e6:+.1f}", (y, a.loc[y, "celkem_zmena"] / 1e6),
                        textcoords="offset points", xytext=(0, -16 if a.loc[y, "celkem_zmena"] < 0 else 8),
                        ha="center", fontsize=9.5, color=INK, bbox=dict(boxstyle="round,pad=0.15", fc="white", ec="none", alpha=0.85))

    # panel 3: kumulativní změna proti 2013
    ax = fig.add_subplot(gs[2]); styl(ax)
    pos_b = np.zeros(len(yrs)); neg_b = np.zeros(len(yrs))
    for c in cats:
        v = kk[c].values / 1e6
        base = np.where(v >= 0, pos_b, neg_b)
        ax.bar(yrs, v, bottom=base, width=0.68, color=COL[c], edgecolor="white", linewidth=1.2,
               hatch="////" if c == "nejasne" else None, zorder=3)
        pos_b += np.where(v >= 0, v, 0); neg_b += np.where(v < 0, v, 0)
    ax.plot(yrs, kk["celkem"].values / 1e6, color=INK, lw=2, marker="D", ms=6, mfc="white", mew=1.8, zorder=5)
    ax.axhline(0, color=INK2, lw=1)
    ax.set_ylim(min(neg_b.min(), kk["celkem"].min() / 1e6) * 1.2, 1.5)
    ax.set_xlim(BASE - .5, 2025.5); ax.set_xticks(range(BASE, 2026)); ax.set_ylabel("Kumulativní změna vs. 2013, Mt CO₂e", color=INK2)
    ax.set_title("Kumulativní pokles od roku 2013: z čeho se skládá", loc="left", fontsize=12.5, color=INK, pad=10)
    ax.annotate(f"{kk['celkem'].iloc[-1]/1e6:.1f} Mt\n({shares['pct_total']:.0f} %)", (2025, kk["celkem"].iloc[-1] / 1e6),
                textcoords="offset points", xytext=(0, -30), ha="center", fontsize=10.5, color=INK, fontweight="bold")
    # legenda s podíly (barevné čtverečky = barvy sloupců)
    from matplotlib.patches import Patch
    h = [Patch(fc=COL["pokracujici"], label=f"pokračující zařízení: {shares['pokracujici']:.0f} %"),
         Patch(fc=COL["zavrena"], label=f"zavřená zařízení: {shares['zavrena']:.0f} %")]
    if "nejasne" in cats:
        h.append(Patch(fc=COL["nejasne"], hatch="////", ec="white", label=f"mimo rozsah: {shares['nejasne']:.0f} %"))
    h.append(Patch(fc=COL["nova"], label=f"nová zařízení: {shares['nova']:+.1f} p. b. (zmírňují pokles)"))
    ax.legend(handles=h, title="Podíl na poklesu 2013–2025", title_fontsize=10, loc="lower left", fontsize=10,
              frameon=True, facecolor="white", edgecolor=GRID, labelcolor=INK2, borderpad=0.8)
    fig.text(0.085, 0.022,
             "Zdroj: EU ETS / EUTL, verified emissions 2025 (extrakce 1. 4. 2026), vlastní výpočet Fakta o klimatu. "
             "Stacionární zařízení v ČR; letectví vyloučeno.\n"
             "Zavřená = zařízení, jehož emise klesly na 0 (nebo přestalo hlásit), včetně zařízení bez bezplatného přídělu. "
             "Ledvice 4 = rozšíření kapacity elektrárny Ledvice (sečteno, v pokračujících).\nRok 2025 může být předběžný.",
             fontsize=8.6, color=INK2, va="bottom", linespacing=1.5)
    for ext, kw in (("png", dict(dpi=300)), ("svg", {})):
        fig.savefig(str(soubor) + "." + ext, facecolor="white", **kw)
    plt.close(fig)


# ------------------------------- 6) největší uzavřené společnosti (seznam)
def uzavrene_spolecnosti(E, EXCL, cz, long):
    """
    Zařízení, která TRVALE přestala emitovat: poslední rok s emisemi > 0 je <= 2024
    a po něm už žádné emise nejsou. Nezávisí na prahu přídělu, takže zahrnuje i
    „nejasné“ případy Seskupeno podle držitele (IDENTIFIER_IN_REG).
    Pozor: držitel je název z registru; firmy ze stejné skupiny mohou mít různé
    názvy (např. „ČEZ - Elektrárna X“ a „ČEZ, a. s. - Elektrárna Y“) a nesloučí se.
    """
    E0 = E.fillna(0.0)
    pos = E0 > 0
    rows = []
    for i in E0.index[pos.any(axis=1)]:
        last = max(y for y in YEARS if pos.loc[i, y])
        if last >= YEARS[-1]:
            continue
        # po posledním roce s emisemi: byl někdy "Excluded"? (mimo rozsah, ne zavření)
        excl_after = bool(EXCL.loc[i, [y for y in YEARS if y > last]].any())
        rows.append(dict(INSTALLATION_IDENTIFIER=i, posledni_rok=last, E_posledni=E0.loc[i, last],
                         E_max=E0.loc[i].max(), rok_max=int(E0.loc[i].idxmax()), excl=excl_after))
    z = pd.DataFrame(rows).join(cz[["INSTALLATION_NAME", "IDENTIFIER_IN_REG", "MAIN_ACTIVITY_TYPE_CODE",
                                    "ACCOUNT_CLOSURE"]], on="INSTALLATION_IDENTIFIER")
    z["sektor"] = z["MAIN_ACTIVITY_TYPE_CODE"].map(SEKTOR).fillna("neznámý")
    z["rok_zavreni"] = z["posledni_rok"] + 1
    z["ucet_open"] = z["ACCOUNT_CLOSURE"].astype(str).str.upper().eq("OPEN")
    # možný převod/přečíslování: stejný držitel má nové zařízení v okolních letech
    nova = long[long.kat == "nova"].join(cz[["IDENTIFIER_IN_REG"]], on="INSTALLATION_IDENTIFIER")
    def pozn(r):
        p = []
        if r.ucet_open: p.append("účet stále OPEN – ověřit, zda nejde o odstávku")
        if r.excl: p.append("později vyřazeno z rozsahu ETS (Excluded)")
        if len(nova[(nova.IDENTIFIER_IN_REG == r.IDENTIFIER_IN_REG) & nova.rok.between(r.rok_zavreni - 1, r.rok_zavreni + 1)]):
            p.append("stejný držitel má nové zařízení – možný převod/přečíslování")
        return "; ".join(p)
    z["poznamka"] = z.apply(pozn, axis=1)
    zar = z.sort_values("E_posledni", ascending=False)
    g = z.groupby("IDENTIFIER_IN_REG").agg(
        pocet_zarizeni=("INSTALLATION_IDENTIFIER", "size"),
        emise_posledni_rok_Mt=("E_posledni", lambda v: v.sum() / 1e6),
        emise_max_Mt=("E_max", lambda v: v.sum() / 1e6),
        roky_zavreni=("rok_zavreni", lambda v: str(min(v)) if min(v) == max(v) else f"{min(v)}–{max(v)}"),
        sektor=("sektor", lambda v: ", ".join(sorted(set(v)))),
        poznamka=("poznamka", lambda v: "; ".join(sorted({x for p in v for x in p.split("; ") if x})))
    ).sort_values("emise_max_Mt", ascending=False).reset_index().rename(columns={"IDENTIFIER_IN_REG": "držitel"})
    return g, zar


def graf_top(g, soubor, n=10):
    t = g.head(n).iloc[::-1].reset_index(drop=True)
    fig, ax = plt.subplots(figsize=(11.5, 7.6), dpi=100)
    fig.subplots_adjust(left=0.34, right=0.94, top=0.80, bottom=0.14)
    y = np.arange(len(t)); h = 0.36
    ax.barh(y + h / 2, t["emise_max_Mt"], height=h, color="#F2B38A", zorder=3, label="Maximum emisí v letech 2013–2025")
    ax.barh(y - h / 2, t["emise_posledni_rok_Mt"], height=h, color=COL["zavrena"], zorder=3, label="Emise v posledním roce provozu")
    for yy, (mx, v, ry, nz) in enumerate(zip(t.emise_max_Mt, t.emise_posledni_rok_Mt, t.roky_zavreni, t.pocet_zarizeni)):
        ax.text(mx + 0.04, yy + h / 2, f"{mx:.2f} Mt", va="center", fontsize=9.5, color=INK2)
        ax.text(max(v, 0) + 0.04, yy - h / 2, f"{v:.2f} Mt · zavřeno {ry}" + (f" ({nz} zař.)" if nz > 1 else ""),
                va="center", fontsize=9.5, color=INK)
    ax.set_yticks(y); ax.set_yticklabels(t["držitel"])
    ax.set_xlim(0, t.emise_max_Mt.max() * 1.42)
    ax.grid(axis="x", color=GRID, lw=0.8); ax.set_axisbelow(True)
    for s_ in ("top", "right", "left"):
        ax.spines[s_].set_visible(False)
    ax.tick_params(colors=INK2, length=0, labelsize=10.5)
    ax.set_xlabel("Ověřené emise, Mt CO₂e", color=INK2)
    ax.legend(loc="lower right", frameon=False, fontsize=10, labelcolor=INK2)
    top3 = ", ".join(g["držitel"].head(3).str.replace(r" a\.s\.$", "", regex=True))
    fig.text(0.02, 0.965, f"Největší uzavřené zařízení v ETS1 (ČR) podle maxima emisí:\n{top3}",
             fontsize=16, fontweight="bold", color=INK, va="top")
    fig.text(0.02, 0.025, "Zdroj: EU ETS / EUTL, verified emissions 2025, vlastní výpočet Fakta o klimatu. Uzavřená = zařízení, které po daném roce už nehlásí\n"
             "žádné emise (2013–2025). Seskupeno podle držitele účtu; firmy ze stejné skupiny s různým názvem se nesčítají. "
             "U účtů ve stavu OPEN (např. Liberty Ostrava) ověřit.",
             fontsize=8.6, color=INK2, va="bottom", linespacing=1.5)
    for ext, kw in (("png", dict(dpi=300)), ("svg", {})):
        fig.savefig(str(soubor) + "." + ext, facecolor="white", **kw)
    plt.close(fig)


# ------------------------------------------- 7) nová zařízení (počet + největší)
def nova_zarizeni(E, cz):
    """Zařízení, která poprvé vykázala emise > 0 v letech 2014–2025 (v 2013 už byla v provozu)."""
    E0 = E.fillna(0.0)
    pos = E0 > 0
    rows = []
    for i in E0.index[pos.any(axis=1)]:
        f = min(y for y in YEARS if pos.loc[i, y])
        if f == YEARS[0]:
            continue
        rows.append(dict(INSTALLATION_IDENTIFIER=i, rok_vzniku=f, E_prvni_rok=E0.loc[i, f], E_max=E0.loc[i].max(),
                         rok_max=int(E0.loc[i].idxmax()), E_2025=E0.loc[i, YEARS[-1]],
                         posledni_rok=max(y for y in YEARS if pos.loc[i, y])))
    n = pd.DataFrame(rows).join(cz[["INSTALLATION_NAME", "IDENTIFIER_IN_REG", "MAIN_ACTIVITY_TYPE_CODE", "ACCOUNT_CLOSURE"]],
                                on="INSTALLATION_IDENTIFIER")
    n["sektor"] = n["MAIN_ACTIVITY_TYPE_CODE"].map(SEKTOR).fillna("neznámý")
    n["stav"] = np.where(n.posledni_rok == YEARS[-1], "v provozu 2025", "ukončeno")
    return n.sort_values("E_max", ascending=False).reset_index(drop=True)


def graf_nova(n, e_total_2025, soubor, top=10):
    G, G_L, G_G = COL["nova"], "#9AD6C2", "#C8C8C4"
    por = n.groupby("rok_vzniku").size().reindex(range(YEARS[1], YEARS[-1] + 1), fill_value=0)
    stav = {st: n[n.stav == st].groupby("rok_vzniku").size().reindex(por.index, fill_value=0)
            for st in ["v provozu 2025", "ukončeno"]}
    barva = {"v provozu 2025": G, "ukončeno": G_G}
    e25 = n.E_2025.sum()
    cz_ = lambda x: str(x).replace(".", ",")
    fig = plt.figure(figsize=(11.5, 12.2), dpi=100)
    fig.text(0.04, 0.965, f"Nových zařízení v českém ETS1 vzniklo od roku 2014 jen {len(n)} a jsou malá:\n"
             f"v roce 2025 emitují dohromady {cz_(f'{e25/1e6:.2f}')} Mt, tedy {cz_(f'{100*e25/e_total_2025:.1f}')} % emisí ČR",
             fontsize=17.5, fontweight="bold", color=INK, va="top")
    fig.text(0.04, 0.895, f"Žádné nové zařízení nepřekročilo {n.E_max.max()/1e3:.0f} kt CO₂e ročně "
             f"(cca {cz_(f'{n.E_max.max()/1e6:.3f}')} Mt).\nPro srovnání: emise ČR v ETS1 klesly od 2013 o 27,7 Mt.",
             fontsize=11.5, color=INK2, va="top", linespacing=1.5)
    # panel 1: počet po letech
    ax = fig.add_axes([0.07, 0.60, 0.88, 0.17])
    spodek = np.zeros(len(por))
    for st, v in stav.items():
        if v.sum() == 0:
            continue
        ax.bar(por.index, v.values, bottom=spodek, color=barva[st], width=0.68, edgecolor="white", linewidth=1.2, label=st, zorder=3,
               hatch="////" if st.startswith("sloučeno") else None)
        spodek += v.values
    for x, v in por.items():
        if v:
            ax.text(x, v + 0.12, str(int(v)), ha="center", fontsize=11, fontweight="bold", color=INK)
    ax.set_xticks(range(YEARS[1], YEARS[-1] + 1)); ax.set_ylim(0, por.max() * 1.5); ax.set_yticks(range(0, int(por.max()) + 1))
    fig.text(0.04, 0.795, "Počet nových zařízení podle roku prvních emisí", fontsize=12.5, color=INK, va="bottom")
    for s_ in ("top", "right", "left"):
        ax.spines[s_].set_visible(False)
    ax.spines["bottom"].set_color("#bdbcb8"); ax.tick_params(colors=INK2, length=0, labelsize=10.5)
    ax.grid(axis="y", color=GRID, lw=0.8); ax.set_axisbelow(True)
    ax.legend(loc="upper center", frameon=False, fontsize=10, labelcolor=INK2, ncol=3)
    ax.set_ylabel("Počet zařízení", color=INK2)
    # panel 2: největší nová zařízení
    ax = fig.add_axes([0.34, 0.145, 0.61, 0.375])
    t = n.head(top).iloc[::-1].reset_index(drop=True)
    y = np.arange(len(t)); h = 0.36
    ax.barh(y + h / 2, t.E_max / 1e3, height=h, color=G_L, zorder=3, label="Maximum emisí v letech 2014–2025")
    ax.barh(y - h / 2, t.E_2025 / 1e3, height=h, color=G, zorder=3, label="Emise v roce 2025")
    lab = [f"{(r.INSTALLATION_NAME if len(r.INSTALLATION_NAME) < 38 else r.INSTALLATION_NAME[:36] + '…')}\n{r.sektor.lower()}, od {r.rok_vzniku}" for r in t.itertuples()]
    ax.set_yticks(y); ax.set_yticklabels(lab, fontsize=9.6)
    mx = t.E_max.max() / 1e3
    for yy, r in enumerate(t.itertuples()):
        ax.text(r.E_max / 1e3 + mx * 0.015, yy + h / 2, cz_(f"{r.E_max/1e3:.1f}") + " kt",
                va="center", fontsize=9.5, color=INK2)
        ax.text(max(r.E_2025 / 1e3, 0) + mx * 0.015, yy - h / 2, cz_(f"{r.E_2025/1e3:.1f}") + " kt" if r.E_2025 > 0 else "ukončeno",
                va="center", fontsize=9.5, color=INK)
    ax.set_xlim(0, mx * 1.22)
    fig.text(0.04, 0.53, f"{top} největších nových zařízení (podle maxima emisí), kt CO₂e", fontsize=12.5, color=INK, va="bottom")
    ax.grid(axis="x", color=GRID, lw=0.8); ax.set_axisbelow(True)
    for s_ in ("top", "right", "left"):
        ax.spines[s_].set_visible(False)
    ax.tick_params(colors=INK2, length=0, labelsize=10)
    ax.legend(loc="lower right", frameon=False, fontsize=10, labelcolor=INK2)
    ax.set_xlabel("kt CO₂e (1 kt = 1 000 t)", color=INK2)
    ax.set_xticks(range(0, int(mx * 1.2) + 1, 5))
    fig.text(0.04, 0.025,
             "Zdroj: EU ETS / EUTL, verified emissions 2025 (extrakce 1. 4. 2026), vlastní výpočet Fakta o klimatu. Nové zařízení = poprvé vykázalo emise > 0\n"
             "v letech 2014–2025. Ledvice 4 jsou rozšíření kapacity elektrárny Ledvice (sečteno), nejsou nové zařízení.\n"
             "Nová zařízení z roku 2025 (např. dvě sádrokartonová) je třeba ověřit: může jít i o změnu povolení či rozsahu.",
             fontsize=8.6, color=INK2, va="bottom", linespacing=1.5)
    for ext, kw in (("png", dict(dpi=300)), ("svg", {})):
        fig.savefig(str(soubor) + "." + ext, facecolor="white", **kw)
    plt.close(fig)


# ----------------------------------------------------------------- main
if __name__ == "__main__":
    cz = nacti()
    E, A, EXCL = matice(cz)
    E0 = E.fillna(0)
    print("\nEmise ČR po letech [Mt CO2e]:")
    print((E0.sum() / 1e6).round(2).to_string())
    print("\nPočet zařízení s emisemi >0:", (E0 > 0).sum().to_dict())
    print("2025 vs 2024: s emisemi", int((E0[2025] > 0).sum()), "vs", int((E0[2024] > 0).sum()),
          "| nehlášeno v 2025 (E_2024>0, E_2025 prázdné):",
          int(((E0[2024] > 0) & E[2025].isna() & ~EXCL[2025]).sum()))

    agg, long = rozklad(E, A, EXCL, cz, THR_MAIN, "vse")
    k = kumulativne(agg)
    pd.set_option("display.width", 220, "display.max_columns", 30, "display.max_colwidth", 60)
    print("\nRozklad [Mt]:\n", (agg[["zavrena", "nova", "pokracujici", "nejasne", "mimo_rozsah", "celkem_zmena", "emise_celkem"]] / 1e6).round(2).to_string())

    # nová zařízení po letech (změna rozsahu, např. spalování odpadu od 2024)
    nv = long[long.kat == "nova"].join(cz[["INSTALLATION_NAME", "MAIN_ACTIVITY_TYPE_CODE"]], on="INSTALLATION_IDENTIFIER")
    print("\nNová zařízení 2024–2025 (kontrola změn rozsahu):\n",
          nv[nv.rok >= 2024][["rok", "INSTALLATION_NAME", "MAIN_ACTIVITY_TYPE_CODE", "E_t"]].round(0).to_string(index=False))

    # sektorový rozpis zavřených (kontrola kódů)
    neznam = sorted(set(cz.MAIN_ACTIVITY_TYPE_CODE) - set(SEKTOR))
    print("Kódy činností bez mapování na sektor:", neznam or "žádné")

    # podíly na kumulativním poklesu
    tot = k["celkem"].iloc[-1]
    sh = {c: 100 * k[c].iloc[-1] / tot for c in ["pokracujici", "zavrena", "nova"]}
    sh["nejasne"] = 100 * (k["nejasne"].iloc[-1] + k["mimo_rozsah"].iloc[-1]) / tot
    sh["pct_total"] = 100 * tot / (agg["emise_celkem"].iloc[0] - agg["celkem_zmena"].iloc[0])
    sh["nova"] = -sh["nova"]  # kladný příspěvek nových zařízení snižuje čistý pokles
    top_closed = long[long.kat == "zavrena"].join(cz[["INSTALLATION_NAME"]], on="INSTALLATION_IDENTIFIER").sort_values("zmena").iloc[0]
    sh["titulek"] = (f"Pokles emisí ČR v ETS1 jde hlavně z provozů, které dál fungují\n"
                     f"({sh['pokracujici']:.0f} % poklesu od 2013), zavírání zařízení tvoří {sh['zavrena']:.0f} %")
    sh["podtitulek"] = (f"Emise českých zařízení klesly z {(agg['emise_celkem'].iloc[0]-agg['celkem_zmena'].iloc[0])/1e6:.1f} Mt (2013) "
                        f"na {agg['emise_celkem'].iloc[-1]/1e6:.1f} Mt (2025), tedy o {abs(sh['pct_total']):.0f} %.\n"
                        f"Největší jednotlivé zavření: {top_closed.INSTALLATION_NAME} ({-top_closed.zmena/1e6:.1f} Mt, {top_closed.rok}).")
    print("\nPodíly na kumulativním poklesu:", {x: round(v, 1) for x, v in sh.items() if isinstance(v, float)})

    cit = citlivost(E, A, EXCL, cz)
    print("\nCitlivostní analýza (součty 2014–2025):\n", cit.to_string(index=False))
    t10 = top10(long, cz)
    print("\n10 největších zavřených zařízení (hlavní definice):\n", t10.to_string(index=False))

    # podklady pro graf a výstupy
    pod = pd.concat([(agg / 1e6).add_suffix("_Mt").drop(columns=["n_zavrena_Mt", "n_nejasne_Mt"]),
                     agg[["n_zavrena", "n_nejasne"]],
                     (k / 1e6).add_prefix("kum_").add_suffix("_Mt")], axis=1).round(4)
    pod.index.name = "rok"
    pod.to_csv(OUT / "ets_cz_rozklad_podklady.csv", encoding="utf-8-sig")
    cit.to_csv(OUT / "ets_cz_citlivost.csv", index=False, encoding="utf-8-sig")
    t10.to_csv(OUT / "ets_cz_top10_zavrena.csv", index=False, encoding="utf-8-sig")
    kl = long.drop_duplicates(["INSTALLATION_IDENTIFIER", "rok"], keep="first").pivot(index="INSTALLATION_IDENTIFIER", columns="rok", values="kat")
    kl.insert(0, "název", cz["INSTALLATION_NAME"]); kl.insert(1, "kód činnosti", cz["MAIN_ACTIVITY_TYPE_CODE"])
    kl.to_csv(OUT / "ets_cz_zarizeni_klasifikace.csv", encoding="utf-8-sig")
    graf(agg, k, sh, "", OUT / "ets_cz_rozklad")
    g, zar = uzavrene_spolecnosti(E, EXCL, cz, long)
    g.round(4).to_csv(OUT / "ets_cz_uzavrene_spolecnosti.csv", index=False, encoding="utf-8-sig")
    zar.drop(columns=["MAIN_ACTIVITY_TYPE_CODE"]).round(0).to_csv(OUT / "ets_cz_uzavrena_zarizeni.csv", index=False, encoding="utf-8-sig")
    print("\nNejvětší uzavřené společnosti (trvale bez emisí):\n", g.head(15).round(3).to_string(index=False))
    graf_top(g, OUT / "ets_cz_top_uzavrene")
    nz = nova_zarizeni(E, cz)
    nz.drop(columns=["MAIN_ACTIVITY_TYPE_CODE"]).round(0).to_csv(OUT / "ets_cz_nova_zarizeni.csv", index=False, encoding="utf-8-sig")
    print("\nNová zařízení: počet po letech", nz.groupby("rok_vzniku").size().to_dict(), "| součet emisí 2025 [Mt]:", round(nz.E_2025.sum() / 1e6, 3))
    print(nz.head(12)[["INSTALLATION_NAME", "sektor", "rok_vzniku", "E_max", "E_2025", "stav"]].round(0).to_string(index=False))
    graf_nova(nz, E0[YEARS[-1]].sum(), OUT / "ets_cz_nova_zarizeni")
    print("\nHotovo ->", OUT.resolve())
