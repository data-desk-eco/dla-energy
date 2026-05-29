"""Build the refinery capacity reference table from cited sources.

Refinery nameplate capacities feed only the *marine-visibility* estimate
(how much of a refinery's crude intake seaborne data can see). They are not
invented here: the bulk are parsed straight from Wikipedia's "List of oil
refineries" at a pinned revision, and the handful that list omits are taken
from named industry sources, one citation per row.

Output: data/sources/refinery_capacity.csv — the single source of truth the
main build reads. Re-run to refresh; every row carries its source URL and the
access date so the figures stay auditable.

    uv run --with requests scripts/build_capacity.py
"""
from __future__ import annotations

import csv
import re
import sys
from pathlib import Path
from urllib.parse import quote
from urllib.request import Request, urlopen

ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / "data" / "sources" / "refinery_capacity.csv"

# Wikipedia "List of oil refineries", pinned to a specific revision so the
# build is reproducible. Bump these two together when refreshing.
WP_LIST = "List of oil refineries"
WP_REVID = 1356770312
WP_ACCESSED = "2026-05-29"
WP_URL = f"https://en.wikipedia.org/w/index.php?title={quote(WP_LIST)}&oldid={WP_REVID}"

# bbl/day -> kt/year for crude (1 bbl ~ 0.1366 t; x365 /1000).
BBL_D_TO_KT_YR = 0.0499

# Crosswalk: Kpler installation name -> capacity source.
#   ("wp", anchor)               capacity parsed from the WP list; `anchor` is
#                                a substring that uniquely identifies the row.
#   ("ext", bbl_d, url, note)    a refinery the WP list omits or mislabels,
#                                taken from a named source.
# Anchors are chosen to defeat the list's name collisions (e.g. two "Lake
# Charles" refineries, two "Texas City" refineries, PTT vs IRPC at Rayong).
CROSSWALK: dict[str, tuple] = {
    # --- DLA-contract refineries ---
    "Jamnagar Refinery": ("wp", "Jamnagar Refinery]]"),
    "Hengyi Refinery": ("ext", 160000,
        "https://www.hydrocarbonprocessing.com/news/2026/01/chinas-hengyi-petrochemical-to-proceed-with-brunei-refinery-expansion/",
        "Hengyi Pulau Muara Besar phase 1; WP list omits it"),
    "Tema Oil Refinery": ("wp", "Tema Oil Refinery"),
    "Petron Bataan Refinery": ("wp", "Bataan Refinery]] ([[Petron"),
    "Petronas Melaka Refinery": ("wp", "MRCSB"),
    "Lytton Refinery": ("wp", "Lytton Oil Refinery"),
    "Dangote Refinery": ("wp", "Lekki"),
    "Sinopec Hainan": ("ext", 184000,
        "https://www.offshore-technology.com/marketdata/hainan-refinery-china/",
        "Sinopec Hainan; WP list omits it"),
    "La Rabida": ("ext", 190000,
        "https://www.power-technology.com/projects/huelva/",
        "Cepsa/Moeve La Rabida (Huelva); WP list omits it"),
    "Bizerte": ("wp", "Bizerte Refinery"),
    "Rayong IRPC Refinery": ("ext", 215000,
        "https://en.wikipedia.org/wiki/IRPC",
        "IRPC Rayong; WP list lists only the separate PTT/Shell Rayong refinery"),
    "MIDOR Refinery": ("wp", "MIDOR (Middle East Oil Refinery"),
    "Mostorod Refinery I": ("ext", 71500,
        "http://abarrelfull.wikidot.com/cairo-oil-refining-company",
        "CORC Mostorod I; immaterial (no seaborne crude in window)"),
    "Mostorod Refinery II": ("wp", "Egypt Refining Company (ERC)"),
    "NATREF Refinery": ("wp", "Natref"),
    "Zarqa Refinery": ("wp", "Jordan Refinery, Zarqa"),
    "MAF Refinery": ("wp", "[[Mina Al Fahal]]"),
    "Pertamina Dumai": ("wp", "[[Dumai]] Refinery"),
    "M'Bao Oil Refinery": ("ext", 30000,
        "https://arda.africa/societe-africaine-de-raffinage-sar/",
        "SAR M'Bao (Dakar), ~1.5 Mt/yr; WP list omits it"),
    # --- Refineries that ship jet into a DLA storage terminal (upstream) ---
    "Ruwais Refinery": ("wp", "Ruwais Refinery]]"),
    "Mina Abdullah": ("wp", "Mina Abdullah Refinery"),
    "Duqm Refinery": ("wp", "Duqm refinery"),
    "MAA Refinery": ("ext", 346000,
        "https://www.knpc.com.kw/en/AboutKNPC/Pages/MAA.aspx",
        "KNPC Mina Al Ahmadi; WP list omits it"),
    "Al Zour": ("wp", "Al Zour Refinery]]"),
    "Samref": ("wp", "SAMREF"),
    "Sitra Refinery": ("ext", 265000,
        "https://www.hydrocarbonprocessing.com/news/2025/09/bapco-expects-to-commission-400-000-bpd-sitra-refinery-in-q4/",
        "Bapco Sitra pre-BMP capacity (expansion to ~400k commissioning Q4 2025)"),
    "S-Oil Onsan": ("wp", "Onsan Refinery ([[S-Oil"),
    "Sinopec HK Hainan": ("ext", 184000,
        "https://www.offshore-technology.com/marketdata/hainan-refinery-china/",
        "Same complex as Sinopec Hainan under a second Kpler label"),
    "CITGO Lake Charles Refinery": ("wp", "Lake Charles Refinery ([[Citgo"),
    "KNOC Daesan": ("wp", "Daesan Refinery ([[Hyundai Oilbank"),
    "ExxonMobil Baton Rouge Refinery": ("wp", "Baton Rouge Refinery]]"),
    "Total Daesan Hanwha Refinery": ("ext", 200000,
        "https://www.offshore-technology.com/marketdata/daesan-ii-refinery-south-korea/",
        "Hanwha TotalEnergies Daesan condensate splitter; approximate, immaterial (clean)"),
    "SK Ulsan": ("wp", "[[Ulsan]] Refinery ([[SK Energy"),
    "Petrochina Jieyang Refinery": ("ext", 400000,
        "https://www.rigzone.com/news/wire/chinese_mega_refinery_lifts_heavy_oil_prices_from_the_doldrums-30-mar-2023-172408-article/",
        "PetroChina Guangdong (Jieyang); WP list omits it"),
    "Mesaieed Refinery": ("wp", "[[Mesaieed]] Refinery"),
    "Marathon Texas City Refinery": ("wp", "Marathon Petroleum]]|| United States, Texas, [[Texas City"),
    "BP Rotterdam": ("wp", "Rotterdam #2 Refinery ([[BP"),
    "Sasol Augusta": ("wp", "[[Augusta, Sicily|Augusta]] Refinery"),
}


def fetch_wp_pairs() -> list[tuple[str, int]]:
    """Parse (row-context, bbl/day) pairs from the pinned WP list revision."""
    api = ("https://en.wikipedia.org/w/api.php?action=parse"
           f"&oldid={WP_REVID}&prop=wikitext&format=json")
    req = Request(api, headers={"User-Agent": "dla-energy-research/1.0"})
    import json
    wt = json.loads(urlopen(req, timeout=60).read())["parse"]["wikitext"]["*"]
    pairs: list[tuple[str, int]] = []
    for m in re.finditer(r"\{\{cvt\|([\d,]+)\|oilbbl/d", wt):  # bulleted lists
        ctx = wt[wt.rfind("\n", 0, m.start()):m.start()]
        pairs.append((ctx, int(m.group(1).replace(",", ""))))
    for m in re.finditer(r"\|\|\s*([\d,]{4,})\s*<ref", wt):    # sortable tables
        ctx = wt[wt.rfind("\n", 0, m.start()):m.start()]
        pairs.append((ctx, int(m.group(1).replace(",", ""))))
    return pairs


def resolve(name: str, spec: tuple, pairs: list[tuple[str, int]]) -> dict:
    kind = spec[0]
    if kind == "ext":
        _, bbl_d, url, note = spec
        return {"kpler_name": name, "capacity_bbl_d": bbl_d,
                "capacity_kt_yr": round(bbl_d * BBL_D_TO_KT_YR),
                "source_url": url, "source_note": note, "accessed": WP_ACCESSED}
    anchor = spec[1]
    hits = [bbl_d for ctx, bbl_d in pairs if anchor in ctx]
    if not hits:
        raise SystemExit(f"ERROR: anchor not found for {name!r}: {anchor!r}")
    bbl_d = hits[0]
    return {"kpler_name": name, "capacity_bbl_d": bbl_d,
            "capacity_kt_yr": round(bbl_d * BBL_D_TO_KT_YR),
            "source_url": WP_URL,
            "source_note": f"Wikipedia 'List of oil refineries' (rev {WP_REVID})",
            "accessed": WP_ACCESSED}


def main() -> None:
    pairs = fetch_wp_pairs()
    print(f"parsed {len(pairs)} capacity datapoints from WP list rev {WP_REVID}")
    rows = [resolve(name, spec, pairs) for name, spec in CROSSWALK.items()]
    OUT.parent.mkdir(parents=True, exist_ok=True)
    cols = ["kpler_name", "capacity_bbl_d", "capacity_kt_yr",
            "source_url", "source_note", "accessed"]
    with OUT.open("w", newline="") as f:
        w = csv.DictWriter(f, fieldnames=cols)
        w.writeheader()
        w.writerows(rows)
    n_wp = sum(1 for r in rows if "wikipedia" in r["source_url"].lower())
    print(f"wrote {OUT.relative_to(ROOT)}: {len(rows)} refineries "
          f"({n_wp} from the WP list, {len(rows) - n_wp} from named gap sources)")


if __name__ == "__main__":
    main()
