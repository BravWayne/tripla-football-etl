import openpyxl
import json
import os
import re
from test_translation import CONTINENTS_MAP, COUNTRIES_MAP, SUPERSTARS_MAP, TROPHIES_MAP, AWARDS_MAP

STAGING_FILE = "tripla_staging_raw.xlsx"
OUTPUT_FILE = "tripla-football-data.xlsx"
SHEET_NAME = "البيانات"
CLUBS_CACHE_FILE = "clubs_cache.json"
NAMES_CACHE_FILE = "names_cache.json"

TARGET_COLUMNS = [
    'الاسم بالعربي', 'الاسم بالإنجليزي', 'سنة الميلاد', 'القارة', 
    'المنتخبات', 'الجنسية (لغير الدوليين)', 'الأندية', 'الأرقام', 
    'الزملاء', 'الألقاب', 'الجوائز الفردية'
]

def load_cache(filepath):
    if os.path.exists(filepath):
        with open(filepath, "r", encoding="utf-8") as f:
            return json.load(f)
    return {}

def translate_dataset():
    clubs_cache = load_cache(CLUBS_CACHE_FILE)
    names_cache = load_cache(NAMES_CACHE_FILE)

    print(f"Loaded {len(clubs_cache)} cached clubs and {len(names_cache)} cached names.")

    wb_in = openpyxl.load_workbook(STAGING_FILE, read_only=True)
    ws_in = wb_in["Data"]
    headers = [str(cell) for cell in next(ws_in.iter_rows(values_only=True))]
    
    col_idx = {h: idx for idx, h in enumerate(headers)}

    wb_out = openpyxl.Workbook()
    ws_out = wb_out.active
    ws_out.title = SHEET_NAME
    ws_out.append(TARGET_COLUMNS)

    count = 0
    missing_names = 0
    missing_clubs = 0

    for row in ws_in.iter_rows(min_row=2, values_only=True):
        eng_name = str(row[col_idx["Name"]]).strip() if row[col_idx["Name"]] else None
        birth_year = int(row[col_idx["Birth Year"]]) if row[col_idx["Birth Year"]] else None
        continent_raw = str(row[col_idx["Continent"]]).strip() if row[col_idx["Continent"]] else None
        nat_team_raw = str(row[col_idx["National Team"]]).strip() if row[col_idx["National Team"]] else None
        cit_raw = str(row[col_idx["Citizenship (Non-Internationals)"]]).strip() if row[col_idx["Citizenship (Non-Internationals)"]] else None
        clubs_raw = str(row[col_idx["Clubs"]]).strip() if row[col_idx["Clubs"]] else None
        shirts_raw = str(row[col_idx["Shirt Numbers"]]).strip() if row[col_idx["Shirt Numbers"]] else None
        teammates_raw = str(row[col_idx["Teammates"]]).strip() if row[col_idx["Teammates"]] else None
        trophies_raw = str(row[col_idx["Trophies"]]).strip() if row[col_idx["Trophies"]] else None
        awards_raw = str(row[col_idx["Individual Awards"]]).strip() if row[col_idx["Individual Awards"]] else None

        # 1. Arabic Name
        ar_name = names_cache.get(eng_name, eng_name)
        if ar_name == eng_name:
            missing_names += 1

        # 2. Continent
        ar_continent = CONTINENTS_MAP.get(continent_raw, continent_raw)

        # 3. National Team vs Citizenship
        ar_nat_team = COUNTRIES_MAP.get(nat_team_raw, nat_team_raw) if nat_team_raw else None
        ar_cit = COUNTRIES_MAP.get(cit_raw, cit_raw) if cit_raw else None

        # 4. Clubs & Club Mapping Dictionary for Shirt Numbers
        club_to_ar = {}
        ar_clubs_list = []
        if clubs_raw:
            for c in clubs_raw.split(", "):
                c_clean = c.strip()
                if not c_clean:
                    continue
                c_ar = clubs_cache.get(c_clean, c_clean)
                if c_ar == c_clean:
                    missing_clubs += 1
                club_to_ar[c_clean] = c_ar
                if c_ar not in ar_clubs_list:
                    ar_clubs_list.append(c_ar)
        ar_clubs_str = "، ".join(ar_clubs_list) if ar_clubs_list else None

        # 5. Shirt Numbers (Consistent with Clubs)
        ar_shirts_list = []
        if shirts_raw:
            for item in shirts_raw.split(", "):
                item = item.strip()
                if ":" in item:
                    c_part, num_part = item.rsplit(":", 1)
                    c_part = c_part.strip()
                    num_part = num_part.strip()
                    # Use identical Arabic club translation
                    ar_c_shirt = club_to_ar.get(c_part) or clubs_cache.get(c_part, c_part)
                    ar_shirts_list.append(f"{ar_c_shirt}:{num_part}")
        ar_shirts_str = "، ".join(ar_shirts_list) if ar_shirts_list else None

        # 6. Teammates
        ar_teammates_list = []
        if teammates_raw:
            for t in teammates_raw.split(", "):
                t_clean = t.strip()
                if t_clean in SUPERSTARS_MAP:
                    ar_teammates_list.append(SUPERSTARS_MAP[t_clean])
                else:
                    ar_teammates_list.append(names_cache.get(t_clean, t_clean))
        ar_teammates_str = "، ".join(ar_teammates_list) if ar_teammates_list else None

        # 7. Trophies
        ar_trophies_list = []
        if trophies_raw:
            for tr in trophies_raw.split(", "):
                tr_clean = tr.strip()
                if tr_clean in TROPHIES_MAP:
                    ar_trophies_list.append(TROPHIES_MAP[tr_clean])
                else:
                    ar_trophies_list.append(tr_clean)
        ar_trophies_str = "، ".join(ar_trophies_list) if ar_trophies_list else None

        # 8. Individual Awards
        ar_awards_list = []
        if awards_raw:
            for aw in awards_raw.split(", "):
                aw_clean = aw.strip()
                if aw_clean in AWARDS_MAP:
                    ar_awards_list.append(AWARDS_MAP[aw_clean])
                else:
                    ar_awards_list.append(aw_clean)
        ar_awards_str = "، ".join(ar_awards_list) if ar_awards_list else None

        row_out = [
            ar_name,
            eng_name,
            birth_year,
            ar_continent,
            ar_nat_team,
            ar_cit,
            ar_clubs_str,
            ar_shirts_str,
            ar_teammates_str,
            ar_trophies_str,
            ar_awards_str
        ]

        ws_out.append(row_out)
        count += 1

    wb_out.save(OUTPUT_FILE)
    print(f"\n[SUCCESS] Successfully translated and exported {count} player rows into {OUTPUT_FILE} (Sheet: '{SHEET_NAME}').")
    print(f"  - Missing name translations: {missing_names}")
    print(f"  - Missing club translations: {missing_clubs}")

if __name__ == "__main__":
    translate_dataset()
