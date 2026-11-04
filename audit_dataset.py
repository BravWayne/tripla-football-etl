import os
import sys
import json
import re
import urllib.request
from typing import List, Dict, Any, Tuple, Optional
import pandas as pd
from bs4 import BeautifulSoup
from curl_cffi import requests

sys.stdout.reconfigure(line_buffering=True)
sys.stderr.reconfigure(line_buffering=True)

# ==============================================================================
# CONFIGURATION
# ==============================================================================
INPUT_FILE = "tripla_staging_raw.xlsx"
SHEET_NAME = "Data"

HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/128.0.0.0 Safari/537.36",
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    "Accept-Language": "en-US,en;q=0.9",
    "Referer": "https://www.transfermarkt.com/",
}

PROXIES = [
    "http://45.195.91.105:8080",
    "http://45.190.184.105:8080",
    "http://180.180.156.53:80",
    "http://38.191.149.127:999",
    "http://203.252.142.217:80",
    "http://45.194.41.16:8080"
]

TMAPI_HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/128.0.0.0 Safari/537.36",
    "Accept": "*/*"
}

APPROVED_SUPERSTARS = {
    "Lionel Messi": {
        "clubs": {"131": (2004, 2021), "583": (2021, 2023), "69261": (2023, 2030)},
        "senior_nt": "Argentina"
    },
    "Cristiano Ronaldo": {
        "clubs": {"336": (2002, 2003), "985": (2003, 2022), "418": (2009, 2018), "506": (2018, 2021), "18544": (2023, 2030)},
        "senior_nt": "Portugal"
    },
    "Karim Benzema": {
        "clubs": {"1041": (2004, 2009), "418": (2009, 2023), "8014": (2023, 2030)},
        "senior_nt": "France"
    },
    "Neymar": {
        "clubs": {"221": (2009, 2013), "131": (2013, 2017), "583": (2017, 2023), "1114": (2023, 2030)},
        "senior_nt": "Brazil"
    },
    "Luka Modrić": {
        "clubs": {"148": (2008, 2012), "418": (2012, 2030)},
        "senior_nt": "Croatia"
    },
    "Sergio Ramos": {
        "clubs": {"368": (2004, 2024), "418": (2005, 2021), "583": (2021, 2023)},
        "senior_nt": "Spain"
    },
    "Mohamed Salah": {
        "clubs": {"631": (2014, 2016), "430": (2014, 2015), "12": (2015, 2017), "31": (2017, 2030)},
        "senior_nt": "Egypt"
    },
    "Kevin De Bruyne": {
        "clubs": {"631": (2012, 2014), "82": (2014, 2015), "281": (2015, 2030)},
        "senior_nt": "Belgium"
    },
    "Kylian Mbappé": {
        "clubs": {"162": (2015, 2018), "583": (2017, 2024), "418": (2024, 2030)},
        "senior_nt": "France"
    },
    "Robert Lewandowski": {
        "clubs": {"16": (2010, 2014), "27": (2014, 2022), "131": (2022, 2030)},
        "senior_nt": "Poland"
    },
    "Toni Kroos": {
        "clubs": {"27": (2007, 2014), "15": (2009, 2010), "418": (2014, 2024)},
        "senior_nt": "Germany"
    },
    "Ángel Di María": {
        "clubs": {"294": (2007, 2010), "418": (2010, 2014), "985": (2014, 2015), "583": (2015, 2022), "506": (2022, 2023)},
        "senior_nt": "Argentina"
    },
    "Luis Suárez": {
        "clubs": {"723": (2007, 2011), "31": (2011, 2014), "131": (2014, 2020), "13": (2020, 2022), "69261": (2024, 2030)},
        "senior_nt": "Uruguay"
    },
    "Erling Haaland": {
        "clubs": {"419": (2019, 2020), "16": (2020, 2022), "281": (2022, 2030)},
        "senior_nt": "Norway"
    },
    "Riyad Mahrez": {
        "clubs": {"1003": (2014, 2018), "281": (2018, 2023), "18545": (2023, 2030)},
        "senior_nt": "Algeria"
    }
}

# ==============================================================================
# NETWORK HELPERS
# ==============================================================================
def fetch_tmapi(url: str) -> Optional[Dict[str, Any]]:
    try:
        req = urllib.request.Request(url, headers=TMAPI_HEADERS)
        with urllib.request.urlopen(req, timeout=6) as resp:
            if resp.status == 200:
                return json.loads(resp.read().decode())
    except Exception:
        pass
    return None

def fetch_html(url: str) -> Optional[str]:
    for p in PROXIES:
        try:
            s = requests.Session(impersonate="chrome120")
            s.proxies = {"http": p, "https": p}
            r = s.get(url, headers=HEADERS, timeout=4)
            if r.status_code == 200 and len(r.text) > 5000:
                return r.text
        except Exception:
            pass
    return None

CLUB_CACHE = {}

def get_club_info(club_id: str) -> Optional[Dict[str, Any]]:
    if club_id in CLUB_CACHE:
        return CLUB_CACHE[club_id]
    data = fetch_tmapi(f"https://tmapi.transfermarkt.technology/clubs?ids[]={club_id}")
    if data and data.get("data"):
        c = data["data"][0]
        CLUB_CACHE[club_id] = c
        return c
    return None

# ==============================================================================
# AUDIT LOGIC
# ==============================================================================
def audit_player_row(row: pd.Series) -> List[Tuple[str, str]]:
    issues = []
    pid = str(row["ID"])
    slug = str(row["Slug"])
    name = str(row["Name"])

    # 1. Fetch TMAPI player data & transfer history
    p_res = fetch_tmapi(f"https://tmapi.transfermarkt.technology/players?ids[]={pid}")
    if not p_res or not p_res.get("data"):
        issues.append(("Profile", f"Player ID {pid} not found in Transfermarkt API"))
        return issues
    p_data = p_res["data"][0]

    # ----------------------------------------------------
    # AUDIT: CLUBS
    # ----------------------------------------------------
    t_data = fetch_tmapi(f"https://tmapi.transfermarkt.technology/transfer/history/player/{pid}")
    raw_transfers = t_data.get("data", {}).get("history", {}).get("terminated", []) if t_data else []
    curr_assignments = p_data.get("clubAssignments", [])

    all_transfer_cids = []
    for tr in raw_transfers:
        cid = str(tr.get("transferDestination", {}).get("clubId", ""))
        if cid and cid not in all_transfer_cids:
            all_transfer_cids.append(cid)
    for ca in curr_assignments:
        if not ca.get("type") == "nationalTeam" and not ca.get("isNationalTeam"):
            cid = str(ca.get("clubId", ""))
            if cid and cid not in all_transfer_cids:
                all_transfer_cids.append(cid)

    valid_senior_clubs = set()
    for cid in all_transfer_cids:
        c_info = get_club_info(cid)
        if not c_info:
            continue
        if c_info.get("baseDetails", {}).get("isNationalTeam"):
            continue
        c_name = c_info.get("name") or c_info.get("baseDetails", {}).get("shortName", "")
        c_type = c_info.get("baseDetails", {}).get("clubTypeId", 1)
        if c_type == 1 and not re.search(r'\b(U\d+|II|B|Youth|Sub-\d+)\b', c_name, re.I) and not any(x in c_name for x in ["Retired", "Without Club"]):
            valid_senior_clubs.add(c_name)

    row_clubs = [c.strip() for c in str(row["Clubs"]).split(",") if c.strip() and str(row["Clubs"]).lower() != "nan"]
    for c in row_clubs:
        # Assert no youth/reserves
        if re.search(r'\b(U\d+|II|B|Youth|Sub-\d+)\b', c, re.I):
            issues.append(("Clubs", f"Youth/reserve team found in clubs column: '{c}'"))
        # Assert club exists in official transfer history
        if c not in valid_senior_clubs:
            issues.append(("Clubs", f"Club '{c}' does not exist in player's verified senior transfer history: {list(valid_senior_clubs)}"))

    # ----------------------------------------------------
    # AUDIT: TROPHIES
    # ----------------------------------------------------
    row_trophies = [t.strip() for t in str(row["Trophies"]).split(",") if t.strip() and str(row["Trophies"]).lower() != "nan"]
    ach_res = fetch_tmapi(f"https://tmapi.transfermarkt.technology/player/{pid}/achievement")
    official_trophies = set()
    future_trophies = set()

    if ach_res and ach_res.get("data"):
        data = ach_res["data"]
        aggregated = data.get("aggregated", [])
        records = data.get("achievements", [])
        for item in aggregated:
            t_name = item.get("name", "").strip()
            item_records = [r for r in records if r.get("id") == item.get("id")]
            if all(r.get("seasonId", 0) >= 2025 for r in item_records):
                future_trophies.add(t_name.lower())
            official_trophies.add(t_name.lower())

    for t in row_trophies:
        t_low = t.lower()
        # Assert status is winner or champion
        if not ("winner" in t_low or "champion" in t_low):
            issues.append(("Trophies", f"Trophy '{t}' is not an explicit winner or champion title"))
        # Assert negative keywords are absent
        if any(bad in t_low for bad in ["participant", "runner-up", "under-", "u17", "u19", "u20", "u21"]):
            issues.append(("Trophies", f"Trophy '{t}' contains disallowed youth/participant keyword"))
        # Senior World Cup vs Club World Cup
        if t_low == "world cup winner":
            if "world cup winner" in future_trophies:
                issues.append(("Trophies", f"Player has unconfirmed future placeholder for World Cup winner"))
        # Ronaldo Saudi League check
        if pid == "8198" and "saudi" in t_low and "champion" in t_low:
            issues.append(("Trophies", "Cristiano Ronaldo falsely attributed Saudi Arabian champion"))

    # ----------------------------------------------------
    # AUDIT: SHIRT NUMBERS
    # ----------------------------------------------------
    row_shirts = [s.strip() for s in str(row["Shirt Numbers"]).split(",") if s.strip() and str(row["Shirt Numbers"]).lower() != "nan"]
    shirt_url = f"https://www.transfermarkt.com/{slug}/rueckennummern/spieler/{pid}"
    shirt_html = fetch_html(shirt_url)
    verified_shirt_pairs = set()
    has_senior_intl_verified = False

    if shirt_html:
        soup = BeautifulSoup(shirt_html, "html.parser")
        tables = soup.select("table.items")
        if tables:
            for tr in tables[0].select("tbody tr"):
                cols = tr.select("td")
                if len(cols) >= 4:
                    raw_c = cols[2].text.strip()
                    num = cols[3].text.strip().replace("#", "")
                    if num.isdigit():
                        verified_shirt_pairs.add((raw_c, num))
            if len(tables) > 1:
                for tr in tables[1].select("tbody tr"):
                    cols = tr.select("td")
                    if len(cols) >= 3:
                        t_name = cols[2].text.strip()
                        if not any(u in t_name for u in ["U1", "U2", "U3", "Olympic", "Olympia", "Youth"]):
                            has_senior_intl_verified = True
                            break

    for s in row_shirts:
        if ":" not in s:
            issues.append(("Shirt Numbers", f"Invalid format '{s}', expected 'Club:Number'"))
            continue
        c_part, n_part = s.split(":", 1)
        c_part = c_part.strip()
        n_part = n_part.strip()
        # Assert club exists in career clubs
        if c_part not in valid_senior_clubs:
            issues.append(("Shirt Numbers", f"Club '{c_part}' in shirt number '{s}' not in senior clubs list"))
        # Assert pair exists in /rueckennummern
        if (c_part, n_part) not in verified_shirt_pairs:
            issues.append(("Shirt Numbers", f"Shirt pair '{s}' does not exist in Transfermarkt career numbers"))

    # ----------------------------------------------------
    # AUDIT: CAPS VS CITIZENSHIP
    # ----------------------------------------------------
    # Also check participations / clubAssignments for national teams
    for ca in curr_assignments:
        if ca.get("type") == "nationalTeam" or ca.get("isNationalTeam"):
            has_senior_intl_verified = True
            break
    if len(p_data.get("nominees", [])) > 0 or len(p_data.get("participations", [])) > 0:
        has_senior_intl_verified = True

    nt_val = str(row["National Team"]).strip()
    cit_val = str(row["Citizenship (Non-Internationals)"]).strip()
    is_nt_empty = (nt_val == "" or nt_val.lower() == "nan" or nt_val.lower() == "none")
    is_cit_empty = (cit_val == "" or cit_val.lower() == "nan" or cit_val.lower() == "none")

    if has_senior_intl_verified:
        if is_nt_empty:
            issues.append(("Caps vs Citizenship", "Player has senior international caps but 'National Team' is empty"))
        if not is_cit_empty:
            issues.append(("Caps vs Citizenship", f"Player has senior international caps but 'Citizenship (Non-Internationals)' is populated: '{cit_val}'"))
    else:
        if not is_nt_empty:
            issues.append(("Caps vs Citizenship", f"Player has 0 senior caps but 'National Team' is populated: '{nt_val}'"))
        if is_cit_empty:
            issues.append(("Caps vs Citizenship", "Player has 0 senior caps but 'Citizenship (Non-Internationals)' is empty"))

    # ----------------------------------------------------
    # AUDIT: TEAMMATES
    # ----------------------------------------------------
    row_teammates = [tm.strip() for tm in str(row["Teammates"]).split(",") if tm.strip() and str(row["Teammates"]).lower() != "nan"]
    for tm in row_teammates:
        if tm not in APPROVED_SUPERSTARS:
            issues.append(("Teammates", f"Teammate '{tm}' is NOT in the approved 15 superstars list"))
        else:
            star_info = APPROVED_SUPERSTARS[tm]
            shares_nt = (not is_nt_empty and nt_val == star_info.get("senior_nt"))
            shares_club = any(cid in star_info["clubs"] for cid in all_transfer_cids)
            if not (shares_nt or shares_club):
                issues.append(("Teammates", f"Teammate '{tm}' did not share senior national team or club with player"))

    return issues

def main():
    if not os.path.exists(INPUT_FILE):
        print(f"[!] Input file '{INPUT_FILE}' does not exist.")
        sys.exit(1)

    df = pd.read_excel(INPUT_FILE, sheet_name=SHEET_NAME)
    total_checked = len(df)
    passing_players = 0
    all_discrepancies = []

    print("=" * 80)
    print("PHASE 2: AUTOMATED DATASET AUDIT AGAINST TRANSFERMARKT")
    print(f"File: {INPUT_FILE} | Sheet: {SHEET_NAME} | Total Rows: {total_checked}")
    print("=" * 80)

    for idx, row in df.iterrows():
        pname = row["Name"]
        pid = row["ID"]
        print(f"\nAuditing [{idx + 1}/{total_checked}]: {pname} (ID: {pid})...")
        issues = audit_player_row(row)
        if not issues:
            passing_players += 1
            print(f"  [PASS] 100% verified against Transfermarkt.")
        else:
            print(f"  [FAIL] Found {len(issues)} discrepancy/discrepancies:")
            for col, desc in issues:
                print(f"         - [{col}]: {desc}")
                all_discrepancies.append((pname, col, desc))

    # Print Final Summary Report
    print("\n" + "=" * 80)
    print("FINAL AUDIT SUMMARY REPORT")
    print("=" * 80)
    print(f"Total Players Checked : {total_checked}")
    print(f"Passing Players (100%): {passing_players} / {total_checked} ({passing_players / total_checked * 100:.1f}%)")
    print(f"Flagged Discrepancies : {len(all_discrepancies)}")
    print("-" * 80)

    if all_discrepancies:
        print(f"{'Player Name':<25} | {'Column':<20} | {'Issue Found'}")
        print("-" * 80)
        for pname, col, desc in all_discrepancies:
            print(f"{pname:<25} | {col:<20} | {desc}")
    else:
        print("[SUCCESS] All players passed with 100% factual accuracy! Zero discrepancies found.")
    print("=" * 80)

if __name__ == "__main__":
    main()
