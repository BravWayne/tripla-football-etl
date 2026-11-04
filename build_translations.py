import openpyxl
import json
import os
import re
import ssl
import time
import urllib.request
import urllib.parse
from concurrent.futures import ThreadPoolExecutor, as_completed

SSL_CTX = ssl.create_default_context()
SSL_CTX.check_hostname = False
SSL_CTX.verify_mode = ssl.CERT_NONE
WIKI_HEADERS = {"User-Agent": "TriplaFootballCrawler/2.0 (https://tripla.app; contact@tripla.app) python-urllib"}
GOOGLE_HEADERS = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"}

CLUBS_CACHE_FILE = "clubs_cache.json"
NAMES_CACHE_FILE = "names_cache.json"

# Load existing caches if available
if os.path.exists(CLUBS_CACHE_FILE):
    with open(CLUBS_CACHE_FILE, "r", encoding="utf-8") as f:
        clubs_cache = json.load(f)
else:
    clubs_cache = {}

if os.path.exists(NAMES_CACHE_FILE):
    with open(NAMES_CACHE_FILE, "r", encoding="utf-8") as f:
        names_cache = json.load(f)
else:
    names_cache = {}

# Seed with known translations
from test_translation import KNOWN_CLUBS

for k, v in KNOWN_CLUBS.items():
    if k not in clubs_cache:
        clubs_cache[k] = v

KNOWN_NAMES = {
    "Cristiano Ronaldo": "كريستيانو رونالدو",
    "Lionel Messi": "ليونيل ميسي",
    "Karim Benzema": "كريم بنزيما",
    "Neymar": "نيمار",
    "Luka Modrić": "لوكا مودريتش",
    "Sergio Ramos": "سيرجيو راموس",
    "Mohamed Salah": "محمد صلاح",
    "Kevin De Bruyne": "كيفين دي بروين",
    "Kylian Mbappé": "كيليان مبابي",
    "Robert Lewandowski": "روبرت ليفاندوفسكي",
    "Toni Kroos": "توني كروس",
    "Ángel Di María": "أنخل دي ماريا",
    "Luis Suárez": "لويس سواريز",
    "Erling Haaland": "إيرلينغ هالاند",
    "Riyad Mahrez": "رياض محرز",
    "Aymeric Laporte": "إيميرك لابورت",
    "Salem Al-Dawsari": "سالم الدوسري",
    "Ali Al-Bulayhi": "علي البليهي",
    "Yassine Bounou": "ياسين بونو",
    "Sadio Mané": "ساديو ماني",
    "N'Golo Kanté": "نغولو كانتي",
    "Fabinho": "فابينيو",
    "Roberto Firmino": "روبرتو فيرمينو",
    "Rúben Neves": "روبن نيفيز",
    "Aleksandar Mitrović": "ألكساندر ميتروفيتش",
    "Kalidou Koulibaly": "خاليدو كوليبالي",
    "Sergej Milinković-Savić": "سيرغي ميلينكوفيتش-سافيتش",
    "Marcelo Brozović": "مارسيلو بروزوفيتش",
    "Otávio": "أوتافيو",
    "Sultan Al-Ghannam": "سلطان الغنام",
    "Abderrazak Hamdallah": "عبد الرزاق حمد الله",
    "Saud Abdulhamid": "سعود عبد الحميد",
    "Nasser Al-Dawsari": "ناصر الدوسري",
    "Abdullah Al-Khaibari": "عبد الله الخيبري",
    "Ayman Yahya": "أيمن يحيى",
    "Franck Kessié": "فرانك كيسي",
    "Riyad Mahrez": "رياض محرز",
    "Ivan Toney": "إيفان توني",
    "Moussa Diaby": "موسى ديابي",
    "Houssem Aouar": "حسام عوار",
    "Steven Bergwijn": "ستيفن بيرجوين",
    "Danilo Pereira": "دانيلو بيريرا",
    "Mohamed Simakan": "محمد سيماكان",
    "Bento": "بينتو",
}

for k, v in KNOWN_NAMES.items():
    if k not in names_cache:
        names_cache[k] = v

def translate_via_google(text: str) -> str:
    url = f"https://translate.googleapis.com/translate_a/single?client=gtx&sl=en&tl=ar&dt=t&q={urllib.parse.quote(text)}"
    req = urllib.request.Request(url, headers=GOOGLE_HEADERS)
    with urllib.request.urlopen(req, timeout=5) as resp:
        res = json.loads(resp.read().decode())
        return "".join([part[0] for part in res[0] if part[0]]).strip()

def clean_club_arabic(text: str) -> str:
    t = re.sub(r"^(نادي|فريق)\s+", "", text)
    t = re.sub(r"\s+(إف سي|اف سي|سي إف|إف\.سي|لكرة القدم|الرياضي)$", "", t)
    return t.strip()

def get_wikidata_arabic(query: str, is_club: bool = False) -> str:
    try:
        url = f"https://www.wikidata.org/w/api.php?action=wbsearchentities&search={urllib.parse.quote(query)}&language=en&format=json"
        req = urllib.request.Request(url, headers=WIKI_HEADERS)
        with urllib.request.urlopen(req, timeout=4, context=SSL_CTX) as resp:
            res = json.loads(resp.read().decode())
        if res.get("search"):
            for item in res["search"][:2]:
                eid = item["id"]
                url2 = f"https://www.wikidata.org/w/api.php?action=wbgetentities&ids={eid}&languages=ar&format=json"
                req2 = urllib.request.Request(url2, headers=WIKI_HEADERS)
                with urllib.request.urlopen(req2, timeout=4, context=SSL_CTX) as resp2:
                    entity_res = json.loads(resp2.read().decode())
                ar_label = entity_res.get("entities", {}).get(eid, {}).get("labels", {}).get("ar", {}).get("value")
                if ar_label:
                    clean = re.sub(r"[^\u0600-\u06FF\s\-]", "", ar_label).strip()
                    if clean and "لاعب" not in clean:
                        if is_club:
                            return clean_club_arabic(clean)
                        return clean
    except Exception:
        pass
    return ""

def resolve_club(club: str) -> str:
    if club in clubs_cache and clubs_cache[club]:
        return clubs_cache[club]
    
    # Try Wikidata
    ar = get_wikidata_arabic(club, is_club=True)
    if not ar:
        # Fallback to Google Translate
        try:
            ar_g = translate_via_google(club)
            ar = clean_club_arabic(ar_g)
        except Exception:
            ar = club
    
    clean_final = re.sub(r"[^\u0600-\u06FF\s\-]", "", ar).strip()
    return clean_final if clean_final else club

def resolve_name(name: str) -> str:
    if name in names_cache and names_cache[name]:
        return names_cache[name]
    
    # Try Wikidata
    ar = get_wikidata_arabic(name, is_club=False)
    if not ar:
        # Fallback to Google Translate
        try:
            ar = translate_via_google(name)
        except Exception:
            ar = name
            
    clean_final = re.sub(r"[^\u0600-\u06FF\s\-]", "", ar).strip()
    return clean_final if clean_final else name

def main():
    wb = openpyxl.load_workbook("tripla_staging_raw.xlsx", read_only=True)
    ws = wb["Data"]
    headers = next(ws.iter_rows(values_only=True))
    name_idx = headers.index("Name")
    club_idx = headers.index("Clubs")

    unique_names = set()
    unique_clubs = set()

    for r in ws.iter_rows(min_row=2, values_only=True):
        if r[name_idx]:
            unique_names.add(str(r[name_idx]).strip())
        if r[club_idx]:
            for c in str(r[club_idx]).split(", "):
                unique_clubs.add(c.strip())

    print(f"Total Unique Names: {len(unique_names)} (Cached: {len([n for n in unique_names if n in names_cache])})")
    print(f"Total Unique Clubs: {len(unique_clubs)} (Cached: {len([c for c in unique_clubs if c in clubs_cache])})")

    # 1. Resolve Clubs
    unresolved_clubs = [c for c in unique_clubs if c not in clubs_cache]
    print(f"\nResolving {len(unresolved_clubs)} clubs...")
    
    def process_c(c):
        ar = resolve_club(c)
        return c, ar

    with ThreadPoolExecutor(max_workers=5) as executor:
        futures = {executor.submit(process_c, c): c for c in unresolved_clubs}
        done_count = 0
        for f in as_completed(futures):
            c, ar = f.result()
            clubs_cache[c] = ar
            done_count += 1
            if done_count % 100 == 0 or done_count == len(unresolved_clubs):
                print(f"  [Clubs: {done_count}/{len(unresolved_clubs)}] Sample: {c} -> {ar}")
                with open(CLUBS_CACHE_FILE, "w", encoding="utf-8") as out:
                    json.dump(clubs_cache, out, ensure_ascii=False, indent=2)

    with open(CLUBS_CACHE_FILE, "w", encoding="utf-8") as out:
        json.dump(clubs_cache, out, ensure_ascii=False, indent=2)
    print("✓ All clubs resolved and saved.")

    # 2. Resolve Names
    unresolved_names = [n for n in unique_names if n not in names_cache]
    print(f"\nResolving {len(unresolved_names)} names...")

    def process_n(n):
        ar = resolve_name(n)
        return n, ar

    with ThreadPoolExecutor(max_workers=5) as executor:
        futures = {executor.submit(process_n, n): n for n in unresolved_names}
        done_count = 0
        for f in as_completed(futures):
            n, ar = f.result()
            names_cache[n] = ar
            done_count += 1
            if done_count % 100 == 0 or done_count == len(unresolved_names):
                print(f"  [Names: {done_count}/{len(unresolved_names)}] Sample: {n} -> {ar}")
                with open(NAMES_CACHE_FILE, "w", encoding="utf-8") as out:
                    json.dump(names_cache, out, ensure_ascii=False, indent=2)

    with open(NAMES_CACHE_FILE, "w", encoding="utf-8") as out:
        json.dump(names_cache, out, ensure_ascii=False, indent=2)
    print("✓ All names resolved and saved.")

if __name__ == "__main__":
    main()
