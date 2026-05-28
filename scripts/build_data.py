"""Build the AML Global / DLA Energy contract investigation database.

For each "Fuel Source Location" named in the DLA Energy contract spreadsheet,
pull jet-fuel arrivals (since 2025-01-01) and — for refineries — the upstream
crude. For sites that are storage / transshipment terminals the jet inflows
reveal which refineries fed them, and we pull crude into those refineries too.
"""
from __future__ import annotations

import json
import subprocess
import sys
import time
from datetime import datetime
from pathlib import Path

import duckdb
from openpyxl import load_workbook

ROOT = Path(__file__).resolve().parent.parent
KPLER = ROOT / "skills/kpler/kpler"
DATA = ROOT / "data"
RAW = DATA / "raw"
DB = DATA / "data.duckdb"
XLSX = DATA / "DLA Energy - AML Contracts.xlsx"

JET_PRODUCT = 1644
CRUDE_GROUP = 1370
SINCE = datetime(2025, 1, 1)
PAGE = 500
MAX_OFFSET = 9500

# Manually mapped from the DLA spreadsheet's (Source, Fuel Source Location)
# columns to Kpler installation IDs. Landlocked sites (Barauni, Fergana,
# Zambia, Sasolburg is pipeline-fed but tracked here as the destination
# refinery) and shut-down sites (Mohammedia/SAMIR for Vivo Maroc) are noted
# but the latter are omitted from the Kpler pulls.
#
# Some DLA rows map to multiple installations (e.g. CORC = Mostorod I + II);
# some installations serve multiple DLA contracts.
DLA_SITES = [
    # (kpler_id, role, source_label, location_label, country_code, refueler, airports)
    (1676,  "terminal", "ATT Tanjung Bin Terminal", "Tanjung Bin", "MY",
        "PACIFIC ISLAND ENERGY", "Pago Pago (American Samoa)"),
    (1308,  "refinery", "Lytton Refinery (Ampol)", "Brisbane", "AU",
        "AMPOL", "Canberra; Townsville"),
    (1685,  "terminal", "Ocean Point Terminals (St Croix Products)", "St Croix", "VI",
        "AML GLOBAL", "Bridgetown (Barbados)"),
    (2097,  "terminal", "Ocean Point Terminals (St Croix)", "St Croix", "VI",
        "AML GLOBAL", "Bridgetown (Barbados)"),
    (7038,  "refinery", "NATREF (Sasol/TotalEnergies)", "Sasolburg", "ZA",
        "PUMA ENERGY", "Gaborone (Botswana)"),
    (6782,  "refinery", "Hengyi Pulau Muara Besar", "Pulau Muara Besar", "BN",
        "GLAMCO UDARA", "Brunei Apt"),
    (1554,  "terminal", "Vopak Terminal Europoort", "Rotterdam", "NL",
        "ENACOL", "Amilcar Cabral (Cape Verde)"),
    (9617,  "refinery", "Mostorod I (CORC)", "Cairo", "EG",
        "MISER PETROLEUM", "Cairo Intl"),
    (9709,  "refinery", "Mostorod II (CORC/ERC)", "Cairo", "EG",
        "MISER PETROLEUM", "Cairo Intl"),
    (1782,  "refinery", "Tema Oil Refinery", "Tema", "GH",
        "PUMA ENERGY", "Kotoka (Accra)"),
    (2187,  "refinery", "Pertamina Dumai Refinery", "Riau", "ID",
        "PERTAMINA AVIATION", "Soekarno-Hatta (Jakarta)"),
    (9595,  "refinery", "JOPETROL Zarqa Refinery", "Zarqa", "JO",
        "JORDAN PETROLEUM", "Marka; Aqaba"),
    (1672,  "terminal", "Aqaba Terminal (JOPETROL crude port)", "Aqaba", "JO",
        "JORDAN PETROLEUM", "Marka; Aqaba"),
    (1369,  "terminal", "Kipevu Oil Storage Facility", "Mombasa", "KE",
        "OLA ENERGY KENYA", "Mombasa Moi; JKIA Nairobi"),
    (6798,  "refinery", "MIDOR Refinery", "Alexandria (Amerya)", "EG",
        "FUEL & AVIATION TRADING", "Beirut Rafic Hariri"),
    (1344,  "refinery", "Petronas Melaka Refinery (MRC)", "Melaka", "MY",
        "PETRONAS DAGANGAN", "Kuala Lumpur"),
    (1247,  "refinery", "Cepsa La Rabida (Huelva)", "Huelva", "ES",
        "OLA ENERGY", "Rabat Sale"),
    (1266,  "refinery", "Reliance Jamnagar", "Jamnagar", "IN",
        "ASHARAMI SYNERGY", "Abuja Nnamdi Azikiwe"),
    (2385,  "refinery", "OQ Mina Al Fahal Refinery", "Muscat", "OM",
        "AL MAHA PETROLEUM", "Muscat Intl"),
    (1355,  "terminal", "OQ Mina Al Fahal terminal", "Muscat", "OM",
        "AL MAHA PETROLEUM", "Muscat Intl"),
    (4720,  "refinery", "Petron Bataan Refinery", "Limay Bataan", "PH",
        "PETRON DCMJR / DCMJR", "Davao; Zamboanga"),
    (11317, "refinery", "Dangote Petroleum Refinery", "Lagos", "NG",
        "PUMA ENERGY AVIATION", "San Juan (Puerto Rico)"),
    (1246,  "refinery", "SAR M'Bao Refinery", "Dakar", "SN",
        "OLA ENERGY SENEGAL", "Leopold Sedar Senghor (Dakar)"),
    (1374,  "refinery", "IRPC Rayong Refinery", "Rayong", "TH",
        "PTTOR", "Phuket"),
    (4341,  "refinery", "STIR Bizerte Refinery", "Bizerte", "TN",
        "OLA ENERGY TUNISIE", "Tunis Carthage"),
    (1384,  "refinery", "Sinopec Hainan Refinery", "Hainan", "CN",
        "SKYPEC", "Noi Bai (Hanoi)"),
]

# DLA rows that have no realistic Kpler footprint, kept for reference in the
# notebook but not pulled.
UNTRACKED = [
    ("Indeni / Puma Energy storage", "Chongwe", "ZM", "PUMA ENERGY",
        "Kenneth Kaunda (Lusaka)", "Zambia is landlocked, pipeline-fed from Tazama"),
    ("Vivo Energy Maroc", "Marrakech", "MA", "ASE MOROCCO",
        "Marrakech Menara", "SAMIR Mohammedia refinery shut in 2015; only storage remains"),
    ("Puma Energy Tanzania storage", "Dar es Salaam", "TZ", "PUMA ENERGY",
        "Julius Nyerere (Dar)", "No operating refinery in Tanzania"),
    ("Barauni Refinery (Indian Oil)", "Barauni Bihar", "IN", "NEPAL OIL",
        "Tribhuvan (Kathmandu)", "Landlocked refinery; jet pipelined to Nepal"),
    ("Fergana Oil Refinery", "Fergana", "UZ", "SANOAT ENERGTIKA",
        "Tashkent Islam Karimov", "Landlocked; no marine crude trail"),
    ("Puma Energy Napa Napa / Port Moresby", "Port Moresby", "PG",
        "PACIFIC ENERGY AVIATION", "Jacksons / Port Moresby",
        "Napa Napa refinery shut 2024; only LPG terminal on Kpler"),
]


def kpler_trades(*, scope_flag: str, scope_id: int, products: int | None,
                 label: str, cache: Path | None = None) -> list[dict]:
    if cache and cache.exists():
        print(f"  {label}: cache hit ({cache.name})")
        return json.loads(cache.read_text())
    out: list[dict] = []
    offset = 0
    while offset <= MAX_OFFSET:
        cmd = [
            str(KPLER), "trades",
            scope_flag, str(scope_id),
            "--size", str(PAGE),
            "--offset", str(offset),
            "--no-forecasted",
        ]
        if products is not None:
            cmd += ["--products", str(products)]
        proc = subprocess.run(cmd, capture_output=True, text=True, check=False)
        if proc.returncode != 0:
            print(f"  {label} @ offset={offset}: {proc.stderr.strip()[:200]}", file=sys.stderr)
            break
        page = [json.loads(line) for line in proc.stdout.splitlines() if line.strip()]
        if not page:
            break
        out.extend(page)
        oldest = min((p.get("start") or "") for p in page)
        print(f"  {label} offset={offset:>5}  +{len(page):>3}  oldest={oldest[:10]}")
        if oldest and oldest < SINCE.isoformat():
            break
        if len(page) < PAGE:
            break
        offset += PAGE
        time.sleep(0.2)
    if cache:
        cache.write_text(json.dumps(out))
    return out


def flatten(trade: dict) -> dict:
    pco = trade.get("portCallOrigin") or {}
    pcd = trade.get("portCallDestination") or {}
    o_inst = pco.get("installation") or {}
    d_inst = pcd.get("installation") or {}
    o_zone = pco.get("zone") or {}
    d_zone = pcd.get("zone") or {}

    fq_list = trade.get("flowQuantities") or []
    fq = fq_list[0] if fq_list else {}
    confirmed = fq.get("confirmedProduct") or {}
    flow = confirmed.get("flowQuantity") or {}
    grade = (fq.get("closestAncestorGrade") or {})
    commodity = (fq.get("closestAncestorCommodity") or {})
    group = (fq.get("closestAncestorGroup") or {})

    osi = ((trade.get("orgSpecificInfo") or {}).get("default") or {})
    seqs = osi.get("bestTradeLinkSequences") or osi.get("tradeLinkSequences") or []
    seller = buyer = None
    if seqs:
        for link in (seqs[0].get("tradeLinks") or []):
            if not seller and link.get("seller"):
                seller = link["seller"].get("name")
            if not buyer and link.get("buyer"):
                buyer = link["buyer"].get("name")

    vessels = trade.get("vessels") or []
    v0 = vessels[0] if vessels else {}

    return {
        "trade_id": trade.get("id"),
        "status": trade.get("status"),
        "start": trade.get("start"),
        "end": trade.get("end"),
        "origin_installation_id": o_inst.get("id"),
        "origin_installation": o_inst.get("name"),
        "origin_port": o_zone.get("name"),
        "origin_country": (o_zone.get("country") or {}).get("name"),
        "dest_installation_id": d_inst.get("id"),
        "dest_installation": d_inst.get("name"),
        "dest_port": d_zone.get("name"),
        "dest_country": (d_zone.get("country") or {}).get("name"),
        "product": commodity.get("name"),
        "grade": grade.get("name"),
        "product_group": group.get("name"),
        "mass_t": flow.get("mass"),
        "volume_bbl": flow.get("volume"),
        "seller": seller,
        "buyer": buyer,
        "vessel_name": v0.get("name"),
        "vessel_imo": v0.get("imo"),
    }


def write_table(con, name: str, rows: list[dict]) -> None:
    con.execute(f"DROP TABLE IF EXISTS {name}")
    if not rows:
        con.execute(f"CREATE TABLE {name} (trade_id BIGINT)")
        return
    raw_path = RAW / f"{name}.json"
    raw_path.write_text(json.dumps(rows))
    con.execute(f"CREATE TABLE {name} AS SELECT * FROM read_json_auto('{raw_path}')")
    con.execute(f"ALTER TABLE {name} ALTER start TYPE TIMESTAMP USING start::TIMESTAMP")
    con.execute(f"ALTER TABLE {name} ALTER \"end\" TYPE TIMESTAMP USING \"end\"::TIMESTAMP")
    con.execute(f"DELETE FROM {name} WHERE start < TIMESTAMP '2025-01-01'")


def load_contracts() -> list[dict]:
    """Parse the DLA Energy contract spreadsheet into a list of dicts."""
    wb = load_workbook(XLSX, data_only=True)
    ws = wb.active
    rows_iter = ws.iter_rows(values_only=True)
    headers = [str(h).strip() if h else "" for h in next(rows_iter)]
    out = []
    for row in rows_iter:
        if not any(row):
            continue
        rec = {h: (str(v).strip() if v is not None else None) for h, v in zip(headers, row)}
        out.append(rec)
    return out


def main():
    RAW.mkdir(parents=True, exist_ok=True)
    if DB.exists():
        DB.unlink()
    con = duckdb.connect(str(DB))

    contracts = load_contracts()
    print(f"\nLoaded {len(contracts)} DLA contracts from spreadsheet")
    raw = RAW / "contracts.json"
    raw.write_text(json.dumps(contracts))
    con.execute(f"CREATE TABLE contracts AS SELECT * FROM read_json_auto('{raw}')")

    # Reference table of the mapped sites (for the notebook to join against).
    sites_rows = [
        {"installation_id": iid, "role": role, "source_label": src,
         "location_label": loc, "country_code": cc, "refueler": ref,
         "airports": ap}
        for (iid, role, src, loc, cc, ref, ap) in DLA_SITES
    ]
    untracked_rows = [
        {"source_label": s, "location_label": loc, "country_code": cc,
         "refueler": ref, "airports": ap, "note": note}
        for (s, loc, cc, ref, ap, note) in UNTRACKED
    ]
    (RAW / "sites.json").write_text(json.dumps(sites_rows))
    (RAW / "untracked.json").write_text(json.dumps(untracked_rows))
    con.execute(f"CREATE TABLE dla_sites AS SELECT * FROM read_json_auto('{RAW / 'sites.json'}')")
    con.execute(f"CREATE TABLE dla_untracked AS SELECT * FROM read_json_auto('{RAW / 'untracked.json'}')")

    # ---------------- Jet inflows to each DLA installation ----------------
    print(f"\n[1/3] Jet inflows to {len(DLA_SITES)} DLA-listed installations")
    jet_all: list[dict] = []
    seen_jet: set = set()
    for iid, role, src, *_ in DLA_SITES:
        page = kpler_trades(
            scope_flag="--to-installations", scope_id=iid,
            products=JET_PRODUCT, label=f"jet→{src[:40]} ({iid})",
            cache=RAW / f"jet_{iid}.json",
        )
        for t in page:
            tid = t.get("id")
            if tid and tid not in seen_jet:
                seen_jet.add(tid)
                jet_all.append(t)
    jet_flat = [flatten(t) for t in jet_all]
    write_table(con, "jet_trades", jet_flat)
    n_jet = con.execute("SELECT COUNT(*) FROM jet_trades").fetchone()[0]
    print(f"  wrote jet_trades: {n_jet} rows")

    # ---------------- Crude inflows ----------------
    # Direct refineries from DLA list:
    refinery_targets: dict[int, str] = {}
    terminal_ids = set()
    for iid, role, src, *_ in DLA_SITES:
        if role == "refinery":
            refinery_targets[iid] = src
        else:
            terminal_ids.add(iid)

    # Storage terminals: trace one step upstream via jet inflows to identify
    # the refineries that supplied them, and pull crude there too.
    upstream_from_terminals: dict[int, str] = {}
    for row in jet_flat:
        if row.get("dest_installation_id") in terminal_ids:
            up_id = row.get("origin_installation_id")
            up_name = row.get("origin_installation")
            if up_id and up_id not in refinery_targets and up_id not in upstream_from_terminals:
                upstream_from_terminals[up_id] = up_name or f"installation {up_id}"

    crude_targets = {**refinery_targets, **upstream_from_terminals}
    print(f"\n[2/3] Crude inflows to {len(crude_targets)} refineries "
          f"({len(refinery_targets)} from DLA list + "
          f"{len(upstream_from_terminals)} upstream of DLA terminals)")

    crude_all: list[dict] = []
    seen_crude: set = set()
    for iid, name in crude_targets.items():
        page = kpler_trades(
            scope_flag="--to-installations", scope_id=iid,
            products=CRUDE_GROUP, label=f"crude→{name[:40]} ({iid})",
            cache=RAW / f"crude_{iid}.json",
        )
        for t in page:
            tid = t.get("id")
            if tid and tid not in seen_crude:
                seen_crude.add(tid)
                crude_all.append(t)
    crude_flat = [flatten(t) for t in crude_all]
    write_table(con, "crude_trades", crude_flat)
    n_crude = con.execute("SELECT COUNT(*) FROM crude_trades").fetchone()[0]
    print(f"  wrote crude_trades: {n_crude} rows")

    # ---------------- Display-name macro (similar to HKIA notebook) ----------------
    print("\n[3/3] Refinery display-name macro")
    con.execute("""
        CREATE OR REPLACE MACRO refinery_name(name) AS
          CASE name
            WHEN 'Petronas Melaka Refinery' THEN 'Petronas Melaka (MRC)'
            WHEN 'La Rabida'                THEN 'Cepsa La Rabida (Huelva)'
            WHEN 'MAF Refinery'             THEN 'OQ Mina Al Fahal Refinery'
            WHEN 'Mina Al Fahal'            THEN 'OQ Mina Al Fahal terminal'
            WHEN 'NATREF Refinery'          THEN 'NATREF (Sasolburg)'
            WHEN 'M''Bao Oil Refinery'      THEN 'SAR M''Bao (Dakar)'
            WHEN 'Saint Croix'              THEN 'Ocean Point (St Croix)'
            WHEN 'Saint Croix Products'     THEN 'Ocean Point Products (St Croix)'
            WHEN 'ATB'                      THEN 'ATT Tanjung Bin'
            WHEN 'Lytton Refinery'          THEN 'Lytton (Ampol)'
            WHEN 'Tema Oil Refinery'        THEN 'Tema Oil Refinery'
            WHEN 'Pertamina Dumai'          THEN 'Pertamina Dumai'
            WHEN 'Hengyi Refinery'          THEN 'Hengyi (Brunei)'
            WHEN 'MIDOR Refinery'           THEN 'MIDOR (Alexandria)'
            WHEN 'Mostorod Refinery I'      THEN 'CORC Mostorod I'
            WHEN 'Mostorod Refinery II'     THEN 'ERC Mostorod II'
            WHEN 'Zarqa Refinery'           THEN 'JOPETROL Zarqa'
            WHEN 'Aqaba Terminal'           THEN 'Aqaba Terminal'
            WHEN 'Kipevu'                   THEN 'Kipevu (Mombasa)'
            WHEN 'Vopak Europoort'          THEN 'Vopak Europoort'
            WHEN 'Jamnagar Refinery'        THEN 'Reliance Jamnagar'
            WHEN 'Petron Bataan Refinery'   THEN 'Petron Bataan'
            WHEN 'Dangote Refinery'         THEN 'Dangote (Lagos)'
            WHEN 'Rayong IRPC Refinery'     THEN 'IRPC Rayong'
            WHEN 'Bizerte'                  THEN 'STIR Bizerte'
            WHEN 'Sinopec Hainan'           THEN 'Sinopec Hainan'
            ELSE name
          END
    """)

    print("\nDone.")
    con.close()


if __name__ == "__main__":
    main()
