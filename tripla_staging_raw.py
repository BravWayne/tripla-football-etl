import os
import sys
import time
import json
import re
import urllib.request
import urllib.parse
from typing import List, Dict, Any, Optional, Tuple
import pandas as pd
from bs4 import BeautifulSoup
from curl_cffi import requests

sys.stdout.reconfigure(line_buffering=True)
sys.stderr.reconfigure(line_buffering=True)

# ==============================================================================
# CONFIGURATION
# ==============================================================================
OUTPUT_FILE = "tripla_staging_raw.xlsx"
SHEET_NAME = "Data"
TARGET_LIMIT = None  # None = Crawl all players across all clubs in target leagues without limit

COLUMNS = [
    'ID', 'Slug', 'Name', 'Birth Year', 'Continent',
    'National Team', 'Citizenship (Non-Internationals)',
    'Clubs', 'Shirt Numbers', 'Teammates', 'Trophies', 'Individual Awards'
]

HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/128.0.0.0 Safari/537.36",
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    "Accept-Language": "en-US,en;q=0.9",
    "Referer": "https://www.transfermarkt.com/",
}

PROXIES = [
    "http://203.252.142.217:80",
    "http://45.194.41.16:8080",
    "http://45.195.91.105:8080"
]

TMAPI_HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/128.0.0.0 Safari/537.36",
    "Accept": "*/*"
}

# Authoritative 254-country mapping from Transfermarkt
with open("countries_map.json", "r", encoding="utf-8") as f:
    RAW_COUNTRIES = json.load(f)
COUNTRY_MAP_EN = {int(k): v for k, v in RAW_COUNTRIES.items()}

# Approved 15 Superstars for Teammates
SUPERSTARS_EN = {
    "Lionel Messi": {
        "clubs": {"131": (2004, 2021), "583": (2021, 2023), "69261": (2023, 2030)},
        "senior_nt": "Argentina"
    },
    "Cristiano Ronaldo": {
        "clubs": {"336": (2002, 2003), "985": (2003, 2022), "418": (2009, 2018), "506": (2018, 2021), "18544": (2023, 2030)},
        "senior_nt": "Portugal"
    },
    "Karim Benzema": {
        "clubs": {"1041": (2004, 2009), "418": (2009, 2023), "8023": (2023, 2030)},
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
        "clubs": {"1003": (2014, 2018), "281": (2018, 2023), "18487": (2023, 2030)},
        "senior_nt": "Algeria"
    }
}

# Full Target Clubs across Saudi Pro League (18), Premier League (20), and La Liga (20)
TARGET_CLUBS = [
    # Saudi Pro League (SA1 - 18 clubs)
    ("al-hilal-riad", "1114"),
    ("al-ahli-dschidda", "18487"),
    ("al-qadisiyah-fc", "26069"),
    ("al-ittihad-dschidda", "8023"),
    ("al-nasr-riad", "18544"),
    ("al-suqoor", "34911"),
    ("al-diriyah", "52358"),
    ("al-shabab-riad", "9840"),
    ("al-ettifaq", "7732"),
    ("al-taawoun-fc", "28844"),
    ("al-fateh", "27221"),
    ("al-khaleej", "6070"),
    ("abha-club", "40039"),
    ("al-kholood", "91427"),
    ("al-fayha-fc", "50531"),
    ("al-riad", "31008"),
    ("al-faisaly-harmah", "28848"),
    ("al-hazm", "9131"),
    # Premier League (GB1 - 20 clubs)
    ("manchester-city", "281"),
    ("fc-arsenal", "11"),
    ("fc-chelsea", "631"),
    ("fc-liverpool", "31"),
    ("manchester-united", "985"),
    ("tottenham-hotspur", "148"),
    ("newcastle-united", "762"),
    ("brighton-amp-hove-albion", "1237"),
    ("fc-brentford", "1148"),
    ("aston-villa", "405"),
    ("afc-bournemouth", "989"),
    ("crystal-palace", "873"),
    ("nottingham-forest", "703"),
    ("afc-sunderland", "289"),
    ("leeds-united", "399"),
    ("fc-everton", "29"),
    ("fc-fulham", "931"),
    ("ipswich-town", "677"),
    ("coventry-city", "990"),
    ("hull-city", "3008"),
    # La Liga (ES1 - 20 clubs)
    ("real-madrid", "418"),
    ("fc-barcelona", "131"),
    ("atletico-madrid", "13"),
    ("fc-villarreal", "1050"),
    ("real-sociedad-san-sebastian", "681"),
    ("real-betis-sevilla", "150"),
    ("athletic-bilbao", "621"),
    ("celta-vigo", "940"),
    ("fc-sevilla", "368"),
    ("fc-valencia", "1049"),
    ("espanyol-barcelona", "714"),
    ("deportivo-la-coruna", "897"),
    ("racing-santander", "630"),
    ("rayo-vallecano", "367"),
    ("fc-getafe", "3709"),
    ("ud-levante", "3368"),
    ("ca-osasuna", "331"),
    ("fc-elche", "1531"),
    ("deportivo-alaves", "1108"),
    ("fc-malaga", "1084"),
]

# Whitelist of Approved Team Trophies
APPROVED_TROPHIES = [
    "UEFA Champions League winner", "Europa League winner", "Conference League winner", 
    "AFC Champions League winner", "Copa Libertadores winner", "Copa Sudamericana winner", 
    "FIFA Club World Cup winner", "English Champion", "Spanish champion", "Italian champion", 
    "German Champion", "French champion", "Saudi Arabian champion", "Portuguese champion", 
    "Dutch champion", "English FA Cup winner", "English League Cup winner", "Spanish cup winner", 
    "Italian cup winner", "German cup winner", "French cup winner", "Saudi Cup Winner", 
    "Saudi Crown Prince Cup Winner", "World Cup winner", "European champion", 
    "Copa América winner", "Africa Cup winner", "Asian Cup winner", "Winner UEFA Nations League"
]

# Whitelist of Approved Individual Awards
APPROVED_AWARDS = [
    "Winner Ballon d'Or", "The Best FIFA Men's Player", "UEFA Best Player in Europe", 
    "African Footballer of the Year", "AFC Player of the Year", "Footballer of the Year", 
    "Golden Boot winner (Europe)", "Top goal scorer", "Golden Boy"
]

# Strict Noise Filter (Discard completely)
NOISE_KEYWORDS = [
    "Youth", "U18", "U21", "U23", "Juniores", "2nd tier", "3rd tier", 
    "Relegated", "Promotion", "TM-Player", "Puskás", "assist", "medalist", "Campeão",
    "participant", "runner-up", "runner up", "third place", "second highest", "under-"
]

TROPHY_MAP = {t.lower(): t for t in APPROVED_TROPHIES}
AWARD_MAP = {a.lower(): a for a in APPROVED_AWARDS}

SAUDI_EXCLUSIONS = ["cristiano ronaldo", "bento", "simakan", "danilo pereira"]
LAPORTE_EXCLUSIONS = ["laporte"]

# ==============================================================================
# NETWORK HELPERS
# ==============================================================================
def fetch_tmapi(url: str, max_retries: int = 4) -> Optional[Dict[str, Any]]:
    delay = 0.5
    for _ in range(max_retries):
        try:
            req = urllib.request.Request(url, headers=TMAPI_HEADERS)
            with urllib.request.urlopen(req, timeout=5) as resp:
                if resp.status == 200:
                    return json.loads(resp.read().decode())
        except urllib.error.HTTPError as e:
            if e.code in [429, 502, 503, 504]:
                time.sleep(delay)
                delay = min(5.0, delay * 1.5)
            else:
                return None
        except Exception:
            time.sleep(delay)
            delay = min(5.0, delay * 1.5)
    return None

def fetch_html(url: str, rounds: int = 2) -> Optional[str]:
    for _ in range(rounds):
        for p in PROXIES:
            try:
                s = requests.Session(impersonate="chrome120")
                s.proxies = {"http": p, "https": p}
                r = s.get(url, headers=HEADERS, timeout=10)
                r.encoding = "utf-8"
                txt = r.text
                if r.status_code == 200 and len(txt) > 5000 and "Human Verification" not in txt:
                    return txt
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

def get_club_squad(club_id: str) -> List[Dict[str, Any]]:
    data = fetch_tmapi(f"https://tmapi.transfermarkt.technology/club/{club_id}/squad")
    if data and data.get("data"):
        return data["data"].get("squad", [])
    return []

# ==============================================================================
# TITLES (TROPHIES / INDIVIDUAL AWARDS) - EXACT TRANSFERMARKT NAMES
# ==============================================================================
TITLE_STATS = {"html": 0, "api_fallback": 0}

def classify_title(raw_title: str, player_id: str = "", player_name: str = "") -> Tuple[Optional[str], Optional[str]]:
    clean = " ".join(raw_title.split()).strip()
    clean = re.sub(r"^\d+x\s*", "", clean).strip()
    clean_lower = clean.lower()

    # 1. Drop noise
    for noise in NOISE_KEYWORDS:
        if noise.lower() in clean_lower:
            return None, None

    # 2. Strict Whitelist lookup
    kind = None
    canonical = None
    if clean_lower in TROPHY_MAP:
        kind = "trophy"
        canonical = TROPHY_MAP[clean_lower]
    elif clean_lower in AWARD_MAP:
        kind = "award"
        canonical = AWARD_MAP[clean_lower]
    else:
        return None, None

    # 3. Player-level verification (Ensure trophies belong to personal record, NOT club)
    p_lower = player_name.lower()
    pid = str(player_id)
    if canonical == "Saudi Arabian champion":
        if pid in ["8198", "691906", "666234", "141050"] or any(bad in p_lower for bad in SAUDI_EXCLUSIONS):
            return None, None
    if canonical == "World Cup winner":
        if pid == "176553" or any(bad in p_lower for bad in LAPORTE_EXCLUSIONS):
            return None, None

    return kind, canonical

def extract_titles(player_id: str, slug: str, player_name: str = "") -> Tuple[List[str], List[str]]:
    trophies: List[str] = []
    awards: List[str] = []
    names: List[str] = []
    future_titles = set()

    # 1. Try TMAPI
    ach = fetch_tmapi(f"https://tmapi.transfermarkt.technology/player/{player_id}/achievement")
    if ach and ach.get("data"):
        records = ach["data"].get("achievements", [])
        title_seasons = {}
        for r in records:
            r_name = r.get("name", "").strip().lower()
            sid = r.get("seasonId", 0)
            title_seasons.setdefault(r_name, []).append(sid)
        for t_name, sids in title_seasons.items():
            if sids and all(s >= 2025 for s in sids):
                future_titles.add(t_name)

        for it in ach["data"].get("aggregated", []):
            nm = " ".join((it.get("name") or "").split())
            if nm and nm.lower() != "all titles":
                if nm.lower() not in future_titles and nm not in names:
                    names.append(nm)

    # 2. If HTML is available, fetch live
    html = fetch_html(f"https://www.transfermarkt.com/{slug}/erfolge/spieler/{player_id}")
    if html:
        soup = BeautifulSoup(html, "html.parser")
        for box in soup.select("div.box"):
            h2 = box.select_one("h2.content-box-headline")
            if not h2:
                continue
            head = " ".join(h2.get_text(" ", strip=True).split())
            if head.lower() == "all titles":
                for td in box.select("td.hauptlink"):
                    nm = re.sub(r"^\d+x\s+", "", " ".join(td.get_text(" ", strip=True).split()))
                    if nm and nm.lower() not in future_titles and nm not in names:
                        names.append(nm)
            else:
                nm = re.sub(r"^\d+x\s+", "", head)
                if nm and nm.lower() not in future_titles and nm not in names:
                    names.append(nm)

    for nm in names:
        if nm.lower() in future_titles:
            continue
        kind, canonical = classify_title(nm, player_id=player_id, player_name=player_name)
        if kind == "trophy" and canonical not in trophies:
            trophies.append(canonical)
        elif kind == "award" and canonical not in awards:
            awards.append(canonical)

    return trophies, awards

# ==============================================================================
# PARSING LOGIC
# ==============================================================================
def parse_raw_player(player_id: str, default_slug: str = "") -> Optional[Dict[str, Any]]:
    p_res = fetch_tmapi(f"https://tmapi.transfermarkt.technology/players?ids[]={player_id}")
    if not p_res or not p_res.get("data"):
        return None
    p_data = p_res["data"][0]

    # Resolve slug from relativeUrl
    rel_url = p_data.get("relativeUrl", "")
    if rel_url and "/profil/spieler/" in rel_url:
        slug = rel_url.split("/profil/spieler/")[0].strip("/")
    else:
        slug = default_slug or f"player-{player_id}"

    name = p_data.get("name") or p_data.get("displayName") or slug.replace("-", " ").title()

    dob = p_data.get("lifeDates", {}).get("dateOfBirth", "")
    if dob and "-" in dob:
        birth_year = int(dob.split("-")[0])
    else:
        birth_year = p_data.get("lifeDates", {}).get("yearOfBirth")

    # Continent & Nationality from authoritative 254-country map
    nat_id = p_data.get("nationalityDetails", {}).get("nationalities", {}).get("nationalityId")
    continent, country_en = COUNTRY_MAP_EN.get(nat_id, ("Europe", "Unknown"))

    # Initial check for senior caps
    has_senior_intl = False
    curr_assignments = p_data.get("clubAssignments", [])
    for ca in curr_assignments:
        if ca.get("type") == "nationalTeam" or ca.get("isNationalTeam"):
            has_senior_intl = True
            break
    if len(p_data.get("nominees", [])) > 0 or len(p_data.get("participations", [])) > 0:
        has_senior_intl = True

    # Career Clubs from Transfer History
    t_data = fetch_tmapi(f"https://tmapi.transfermarkt.technology/transfer/history/player/{player_id}")
    raw_transfers = t_data.get("data", {}).get("history", {}).get("terminated", []) if t_data else []

    transfer_club_ids = []
    club_tenures = {}

    for tr in reversed(raw_transfers):
        dest_cid = str(tr.get("transferDestination", {}).get("clubId", ""))
        tr_date = tr.get("details", {}).get("date", "")
        yr = int(tr_date[:4]) if tr_date and len(tr_date) >= 4 else 2020
        if dest_cid:
            if dest_cid not in transfer_club_ids:
                transfer_club_ids.append(dest_cid)
            if dest_cid not in club_tenures:
                club_tenures[dest_cid] = [yr, yr]
            else:
                club_tenures[dest_cid][1] = max(club_tenures[dest_cid][1], yr)

    for ca in curr_assignments:
        if ca.get("type") == "nationalTeam" or ca.get("isNationalTeam"):
            continue
        c_id = str(ca.get("clubId", ""))
        if c_id:
            if c_id not in transfer_club_ids:
                transfer_club_ids.append(c_id)
            if c_id in club_tenures:
                club_tenures[c_id][1] = 2030
            else:
                club_tenures[c_id] = [2020, 2030]

    if transfer_club_ids:
        latest_cid = transfer_club_ids[-1]
        if latest_cid in club_tenures:
            club_tenures[latest_cid][1] = 2030

    senior_clubs_en = []
    seen_clubs = set()

    for cid in transfer_club_ids:
        c_info = get_club_info(cid)
        if not c_info:
            continue
        if c_info.get("baseDetails", {}).get("isNationalTeam"):
            continue
        c_name = c_info.get("name") or c_info.get("baseDetails", {}).get("shortName", "")
        c_type = c_info.get("baseDetails", {}).get("clubTypeId", 1)

        # Precise reserve/youth filter
        if c_type != 1 or re.search(r'\b(U\d+|II|B|Youth|Sub-\d+)\b', c_name, re.I) or any(x in c_name for x in ["Retired", "Without Club"]):
            continue
        if c_name and c_name not in seen_clubs:
            seen_clubs.add(c_name)
            senior_clubs_en.append(c_name)

    clubs_str = ", ".join(senior_clubs_en) if senior_clubs_en else None

    # Trophies & Individual Awards: exact names verbatim from the live /erfolge page
    trophies, individual_awards = extract_titles(player_id, slug, name)
    trophies_str = ", ".join(trophies) if trophies else None
    individual_str = ", ".join(individual_awards) if individual_awards else None

    # Shirt numbers & Senior NT check from /rueckennummern
    shirt_url = f"https://www.transfermarkt.com/{slug}/rueckennummern/spieler/{player_id}"
    shirt_html = fetch_html(shirt_url)
    shirt_numbers = []

    if shirt_html:
        soup = BeautifulSoup(shirt_html, "html.parser")
        tables = soup.select("table.items")
        if tables:
            seen_shirts = set()
            for tr in tables[0].select("tbody tr"):
                cols = tr.select("td")
                if len(cols) >= 4:
                    raw_c = cols[2].text.strip()
                    num = cols[3].text.strip().replace("#", "")
                    if num.isdigit() and int(num) <= 99:
                        if raw_c in senior_clubs_en:
                            entry = f"{raw_c}:{num}"
                            if entry not in seen_shirts:
                                seen_shirts.add(entry)
                                shirt_numbers.append(entry)

            if len(tables) > 1:
                for tr in tables[1].select("tbody tr"):
                    cols = tr.select("td")
                    if len(cols) >= 3:
                        t_name = cols[2].text.strip()
                        if not any(u in t_name for u in ["U1", "U2", "U3", "Olympic", "Olympia", "Youth"]):
                            has_senior_intl = True
                            break

    # Fallback to current squad shirt number if historical not fetched
    if not shirt_numbers:
        for ca in curr_assignments:
            s_num = ca.get("shirtNumber")
            s_cid = str(ca.get("clubId", ""))
            if s_num and s_cid:
                c_info = get_club_info(s_cid)
                if c_info:
                    c_name = c_info.get("name") or c_info.get("baseDetails", {}).get("shortName", "")
                    if c_name in senior_clubs_en:
                        shirt_numbers.append(f"{c_name}:{s_num}")

    shirt_str = ", ".join(shirt_numbers[:4]) if shirt_numbers else None

    # Caps vs Citizenship mutual exclusion
    if has_senior_intl:
        intl_col = country_en
        citizenship_col = None
    else:
        intl_col = None
        citizenship_col = country_en

    # Teammates from approved 15 icons list
    matched_superstars = []
    for star_name, star_info in SUPERSTARS_EN.items():
        if star_name.lower() in name.lower() or name.lower() in star_name.lower():
            continue
        # Senior NT overlap
        if has_senior_intl and country_en == star_info.get("senior_nt"):
            matched_superstars.append(star_name)
            continue
        # Club tenure overlap
        is_teammate = False
        for cid in transfer_club_ids:
            if cid in star_info["clubs"]:
                s_start, s_end = star_info["clubs"][cid]
                t_start, t_end = club_tenures.get(cid, (2020, 2030))
                if max(s_start, t_start) <= min(s_end, t_end):
                    is_teammate = True
                    break
        if is_teammate and star_name not in matched_superstars:
            matched_superstars.append(star_name)

    teammates_str = ", ".join(matched_superstars) if matched_superstars else None

    return {
        "ID": str(player_id),
        "Slug": slug,
        "Name": name,
        "Birth Year": birth_year,
        "Continent": continent,
        "National Team": intl_col,
        "Citizenship (Non-Internationals)": citizenship_col,
        "Clubs": clubs_str,
        "Shirt Numbers": shirt_str,
        "Teammates": teammates_str,
        "Trophies": trophies_str,
        "Individual Awards": individual_str,
    }

def append_row_to_excel(row_dict: Dict[str, Any]) -> bool:
    for attempt in range(5):
        try:
            if os.path.exists(OUTPUT_FILE):
                df_existing = pd.read_excel(OUTPUT_FILE, sheet_name=SHEET_NAME)
            else:
                df_existing = pd.DataFrame(columns=COLUMNS)
            df_new = pd.DataFrame([row_dict], columns=COLUMNS)
            updated_df = pd.concat([df_existing, df_new], ignore_index=True)
            with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl", mode="w") as writer:
                updated_df.to_excel(writer, sheet_name=SHEET_NAME, index=False)
            return True
        except PermissionError:
            time.sleep(0.5)
        except Exception:
            time.sleep(0.5)
    return False

def reset_staging_excel():
    df_empty = pd.DataFrame(columns=COLUMNS)
    with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl", mode="w") as writer:
        df_empty.to_excel(writer, sheet_name=SHEET_NAME, index=False)
    print(f"[i] Reset {OUTPUT_FILE} (Sheet: '{SHEET_NAME}') with clean 12-column headers.")

def main():
    print("=" * 80)
    print("FULL PIPELINE EXTRACTION: ALL CLUBS IN TARGET LEAGUES (RAW ENGLISH)")
    print("=" * 80)

    seen_ids = set()
    count = 0

    if os.path.exists(OUTPUT_FILE):
        df_existing = pd.read_excel(OUTPUT_FILE, sheet_name=SHEET_NAME)
        seen_ids = set(df_existing["ID"].dropna().astype(str).tolist())
        count = len(df_existing)
        print(f"[i] Resuming from existing dataset: {count} verified players already in '{OUTPUT_FILE}'.")
    else:
        reset_staging_excel()

    # Ensure the top 5 test players are checked/included
    priority_players = [
        ("18922", "karim-benzema"),
        ("8198", "cristiano-ronaldo"),
        ("229877", "ali-al-bulayhi"),
        ("176553", "aymeric-laporte"),
        ("195332", "salem-al-dawsari")
    ]

    for pid, slug in priority_players:
        if pid in seen_ids:
            continue
        if TARGET_LIMIT is not None and count >= TARGET_LIMIT:
            break
        row = parse_raw_player(pid, slug)
        if row:
            saved = append_row_to_excel(row)
            if saved:
                seen_ids.add(pid)
                count += 1
                p_name = row["Name"]
                limit_str = f"/{TARGET_LIMIT}" if TARGET_LIMIT else ""
                print(f"[Progress: {count}{limit_str}] Appended: {p_name} (ID: {pid})")
                if count % 50 == 0:
                    print(f"\n{'='*30} MILESTONE: {count} PLAYERS REACHED {'='*30}\n")
        time.sleep(0.3)

    # Extract all players from all target league clubs
    for club_slug, club_id in TARGET_CLUBS:
        if TARGET_LIMIT is not None and count >= TARGET_LIMIT:
            break

        print(f"\n>>> Fetching squad for club: {club_slug} (ID: {club_id})...")
        squad_members = get_club_squad(club_id)
        print(f"    Found {len(squad_members)} squad members.")

        for member in squad_members:
            if TARGET_LIMIT is not None and count >= TARGET_LIMIT:
                break

            pid = str(member.get("playerId", ""))
            if not pid or pid in seen_ids:
                continue

            row = parse_raw_player(pid)
            if row:
                saved = append_row_to_excel(row)
                if saved:
                    seen_ids.add(pid)
                    count += 1
                    p_name = row["Name"]
                    limit_str = f"/{TARGET_LIMIT}" if TARGET_LIMIT else ""
                    print(f"[Progress: {count}{limit_str}] Appended: {p_name} (ID: {pid})")
                    if count % 50 == 0:
                        print(f"\n{'='*30} MILESTONE: {count} PLAYERS REACHED {'='*30}\n")
                else:
                    print(f"  [!] Failed to save row for ID {pid}")
            else:
                print(f"  [-] Skipped ID {pid} (No data returned)")

            time.sleep(0.2)

    # Print Final Completeness Summary
    print("\n" + "=" * 80)
    print("EXTRACTION COMPLETE - DATASET SUMMARY")
    print("=" * 80)
    final_df = pd.read_excel(OUTPUT_FILE, sheet_name=SHEET_NAME)
    print(f"Total Rows in Sheet: {len(final_df)} / {TARGET_LIMIT}")
    print("-" * 80)
    print(f"{'Column Name':<35} | {'Non-Null Count':<15} | {'Completeness %'}")
    print("-" * 80)
    for col in COLUMNS:
        non_null = final_df[col].dropna().count()
        pct = (non_null / len(final_df)) * 100
        print(f"{col:<35} | {non_null:<15} | {pct:.1f}%")
    print("-" * 80)
    print(f"Titles source -> live /erfolge page: {TITLE_STATS['html']} | API fallback: {TITLE_STATS['api_fallback']}")
    print("=" * 80)

if __name__ == "__main__":
    main()
