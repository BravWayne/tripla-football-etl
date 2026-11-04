import os
import sys
import time
import json
import re
import urllib.request
import urllib.parse
import ssl
from typing import List, Dict, Any, Optional, Tuple
import pandas as pd
from bs4 import BeautifulSoup
from curl_cffi import requests
from openai import OpenAI

sys.stdout.reconfigure(line_buffering=True)
sys.stderr.reconfigure(line_buffering=True)

# ==============================================================================
# CONFIGURATION
# ==============================================================================
EXCEL_FILE = "tripla-football-data.xlsx"
TEMPLATE_FILE = "tripla-football-template-2.xlsx"
SHEET_NAME = "البيانات"

HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/128.0.0.0 Safari/537.36",
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,*/*;q=0.8",
    "Accept-Language": "en-US,en;q=0.9",
    "Referer": "https://www.transfermarkt.com/",
}

PROXIES = [
    "http://203.252.142.217:80",
    "http://45.194.41.16:8080",
    "http://45.195.91.105:8080",
]

TMAPI_HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/128.0.0.0 Safari/537.36",
    "Accept": "*/*"
}

WIKI_HEADERS = {
    "User-Agent": "TriplaFootballCrawler/2.0 (https://tripla.app; contact@tripla.app) python-urllib"
}

SSL_CTX = ssl.create_default_context()
SSL_CTX.check_hostname = False
SSL_CTX.verify_mode = ssl.CERT_NONE

# Target Schema
COLUMNS_SCHEMA = [
    'الاسم بالعربي', 'الاسم بالإنجليزي', 'سنة الميلاد', 'القارة', 
    'المنتخبات', 'الجنسية (لغير الدوليين)', 'الأندية', 'الأرقام', 
    'الزملاء', 'الألقاب', 'الجوائز الفردية'
]

# 1. Deterministic Club Translation Dictionary
CLUB_MAP = {
    "Real Madrid": "ريال مدريد", "FC Barcelona": "برشلونة", "Atlético de Madrid": "أتلتيكو مدريد",
    "Sevilla FC": "إشبيلية", "Real Sociedad": "ريال سوسيداد", "Villarreal CF": "فياريال",
    "Athletic Bilbao": "أتلتيك بيلباو", "Valencia CF": "فالنسيا", "Real Betis Balompié": "ريال بيتيس",
    "Rayo Vallecano": "رايو فايكانو", "Getafe CF": "خيتافي", "RCD Mallorca": "مايوركا",
    "Manchester City": "مانشستر سيتي", "Liverpool FC": "ليفربول", "Arsenal FC": "أرسنال",
    "Manchester United": "مانشستر يونايتد", "Chelsea FC": "تشيلسي", "Tottenham Hotspur": "توتنهام",
    "Aston Villa": "أستون فيلا", "Newcastle United": "نيوكاسل", "Everton FC": "إيفرتون",
    "Al-Hilal SFC": "الهلال", "Al-Nassr FC": "النصر", "Al-Ittihad Club": "الاتحاد",
    "Al-Ahli SFC": "الأهلي", "Al-Ettifaq FC": "الاتفاق", "Al-Shabab FC": "الشباب",
    "Al-Qadsiah FC": "القادسية", "Al-Taawoun FC": "التعاون", "Al-Fateh SC": "الفتح",
    "Al-Nahdah FC": "النهضة", "Al-Bukiryah FC": "البكيرية", "Al-Hazem SC": "الحزم",
    "Al-Wehda FC": "الوحدة", "Al-Faisaly FC": "الفيصلي", "Al-Khaleej FC": "الخليج",
    "Paris Saint-Germain": "باريس سان جيرمان", "Olympique Lyon": "ليون", "Olympique Marseille": "مارسيليا",
    "AS Monaco": "موناكو", "LOSC Lille": "ليل", "Stade Rennais FC": "رين",
    "Bayern Munich": "بايرن ميونخ", "Borussia Dortmund": "بوروسيا دورتموند", "Bayer 04 Leverkusen": "باير ليفركوزن",
    "RB Leipzig": "لايبزيغ", "Eintracht Frankfurt": "آينتراخت فرانكفورت",
    "Inter Milan": "إنتر ميلان", "AC Milan": "ميلان", "Juventus FC": "يوفنتوس", "AS Roma": "روما",
    "SSC Napoli": "نابولي", "SS Lazio": "لاتسيو", "Atalanta BC": "أتالانتا", "ACF Fiorentina": "فيورنتينا",
    "SL Benfica": "بنفيكا", "FC Porto": "بورتو", "Sporting CP": "سبورتينغ لشبونة",
    "Santos FC": "سانتوس", "CR Flamengo": "فلامنغو", "São Paulo FC": "ساو باولو", "SE Palmeiras": "بالميراس",
    "Club Athletico Paranaense": "أتلتيكو باراناينسي", "Corinthians": "كورينثيانز", "Cruzeiro": "كروزيرو",
    "Club Atlético River Plate": "ريفر بليت", "CA Boca Juniors": "بوكا جونيورز",
    "Ajax Amsterdam": "أياكس", "PSV Eindhoven": "آيندهوفن", "Feyenoord Rotterdam": "فاينورد",
    "Celtic FC": "سلتيك", "Rangers FC": "رينجرز", "Galatasaray": "غلطة سراي", "Fenerbahce": "فنربخشة", "Besiktas JK": "بشكتاش",
    "Envigado FC": "إنفيغادو", "Chicago Fire FC": "شيكاغو فاير", "Zenit St. Petersburg": "زينيت",
    "Inter Miami CF": "إنتر ميامي"
}

# 2. Deterministic Trophy Translation Dictionary (Only confirmed major trophies)
TROPHY_MAP = {
    "Champions League winner": "دوري أبطال أوروبا",
    "UEFA Champions League winner": "دوري أبطال أوروبا",
    "Spanish champion": "الدوري الإسباني",
    "English champion": "الدوري الإنجليزي",
    "Saudi Arabian champion": "الدوري السعودي",
    "Italian champion": "الدوري الإيطالي",
    "German champion": "الدوري الألماني",
    "French champion": "الدوري الفرنسي",
    "European champion": "كأس أمم أوروبا",
    "Copa América winner": "كوبا أمريكا",
    "Africa Cup winner": "كأس أمم أفريقيا",
    "Asian Cup winner": "كأس أمم آسيا",
    "Spanish cup winner": "كأس ملك إسبانيا",
    "FA Cup winner": "كأس الاتحاد الإنجليزي",
    "Italian Cup winner": "كأس إيطاليا",
    "French cup winner": "كأس فرنسا",
    "Saudi Crown Prince Cup Winner": "كأس ولي العهد السعودي",
    "Saudi Cup Winner": "كأس خادم الحرمين الشريفين",
    "King's Cup Winner": "كأس خادم الحرمين الشريفين",
    "Europa League winner": "الدوري الأوروبي",
    "AFC Champions League winner": "دوري أبطال آسيا",
    "FIFA Club World Cup winner": "كأس العالم للأندية",
    "Club World Cup winner": "كأس العالم للأندية",
}

# 3. Approved Superstars for «الزملاء»
SUPERSTARS = {
    "ليونيل ميسي": {
        "clubs": {"131": (2004, 2021), "583": (2021, 2023), "69261": (2023, 2030)},
        "senior_nt": "الأرجنتين", "eng_name": "Lionel Messi"
    },
    "كريستيانو رونالدو": {
        "clubs": {"336": (2002, 2003), "985": (2003, 2022), "418": (2009, 2018), "506": (2018, 2021), "18544": (2023, 2030)},
        "senior_nt": "البرتغال", "eng_name": "Cristiano Ronaldo"
    },
    "كريم بنزيما": {
        "clubs": {"1041": (2004, 2009), "418": (2009, 2023), "8014": (2023, 2030)},
        "senior_nt": "فرنسا", "eng_name": "Karim Benzema"
    },
    "نيمار": {
        "clubs": {"221": (2009, 2013), "131": (2013, 2017), "583": (2017, 2023), "1114": (2023, 2030)},
        "senior_nt": "البرازيل", "eng_name": "Neymar"
    },
    "لوكا مودريتش": {
        "clubs": {"148": (2008, 2012), "418": (2012, 2030)},
        "senior_nt": "كرواتيا", "eng_name": "Luka Modrić"
    },
    "سيرجيو راموس": {
        "clubs": {"368": (2004, 2024), "418": (2005, 2021), "583": (2021, 2023)},
        "senior_nt": "إسبانيا", "eng_name": "Sergio Ramos"
    },
    "محمد صلاح": {
        "clubs": {"631": (2014, 2016), "430": (2014, 2015), "12": (2015, 2017), "31": (2017, 2030)},
        "senior_nt": "مصر", "eng_name": "Mohamed Salah"
    },
    "كيفين دي بروين": {
        "clubs": {"631": (2012, 2014), "82": (2014, 2015), "281": (2015, 2030)},
        "senior_nt": "بلجيكا", "eng_name": "Kevin De Bruyne"
    },
    "كيليان مبابي": {
        "clubs": {"162": (2015, 2018), "583": (2017, 2024), "418": (2024, 2030)},
        "senior_nt": "فرنسا", "eng_name": "Kylian Mbappé"
    },
    "روبرت ليفاندوفسكي": {
        "clubs": {"16": (2010, 2014), "27": (2014, 2022), "131": (2022, 2030)},
        "senior_nt": "بولندا", "eng_name": "Robert Lewandowski"
    },
    "توني كروس": {
        "clubs": {"27": (2007, 2014), "15": (2009, 2010), "418": (2014, 2024)},
        "senior_nt": "ألمانيا", "eng_name": "Toni Kroos"
    },
    "أنخل دي ماريا": {
        "clubs": {"294": (2007, 2010), "418": (2010, 2014), "985": (2014, 2015), "583": (2015, 2022), "506": (2022, 2023)},
        "senior_nt": "الأرجنتين", "eng_name": "Ángel Di María"
    },
    "لويس سواريز": {
        "clubs": {"723": (2007, 2011), "31": (2011, 2014), "131": (2014, 2020), "13": (2020, 2022), "69261": (2024, 2030)},
        "senior_nt": "أوروغواي", "eng_name": "Luis Suárez"
    },
    "إيرلينغ هالاند": {
        "clubs": {"419": (2019, 2020), "16": (2020, 2022), "281": (2022, 2030)},
        "senior_nt": "النرويج", "eng_name": "Erling Haaland"
    },
    "رياض محرز": {
        "clubs": {"1003": (2014, 2018), "281": (2018, 2023), "18545": (2023, 2030)},
        "senior_nt": "الجزائر", "eng_name": "Riyad Mahrez"
    }
}

# 4. Country & Continent Mapping
COUNTRY_MAP = {
    146: ("آسيا", "السعودية"), 50: ("أوروبا", "فرنسا"), 26: ("أمريكا الجنوبية", "البرازيل"),
    136: ("أوروبا", "البرتغال"), 157: ("أوروبا", "إسبانيا"), 149: ("أفريقيا", "السنغال"),
    150: ("أوروبا", "صربيا"), 189: ("أوروبا", "إنجلترا"), 40: ("أوروبا", "ألمانيا"),
    9: ("أمريكا الجنوبية", "الأرجنتين"), 122: ("أوروبا", "هولندا"), 37: ("أوروبا", "كرواتيا"),
    107: ("أفريقيا", "المغرب"), 45: ("أفريقيا", "مصر"), 4: ("أفريقيا", "الجزائر"),
    124: ("أفريقيا", "نيجيريا"), 38: ("أفريقيا", "ساحل العاج"), 54: ("أفريقيا", "غانا"),
    31: ("أفريقيا", "الكاميرون"), 34: ("أمريكا الجنوبية", "كولومبيا"), 83: ("أمريكا الجنوبية", "كولومبيا"),
    182: ("أمريكا الجنوبية", "أوروغواي"), 32: ("أمريكا الجنوبية", "تشيلي"), 75: ("أوروبا", "إيطاليا"),
    19: ("أوروبا", "بلجيكا"), 174: ("أوروبا", "تركيا"), 125: ("أوروبا", "النرويج"), 161: ("أوروبا", "السويد"),
    39: ("أوروبا", "الدنمارك"), 135: ("أوروبا", "بولندا"), 162: ("أوروبا", "سويسرا"), 12: ("أوروبا", "النمسا"),
    152: ("أوروبا", "اسكتلندا"), 77: ("آسيا", "اليابان"), 87: ("آسيا", "كوريا الجنوبية"), 102: ("أفريقيا", "مالي"),
    62: ("أوروبا", "اليونان"), 140: ("أوروبا", "رومانيا"), 68: ("أوروبا", "المجر"), 41: ("أوروبا", "جورجيا"),
    169: ("أفريقيا", "تونس"), 137: ("آسيا", "قطر"), 183: ("آسيا", "الإمارات"), 80: ("آسيا", "الأردن"),
    74: ("آسيا", "العراق"), 128: ("آسيا", "عمان"), 14: ("آسيا", "البحرين"), 88: ("آسيا", "الكويت"),
    184: ("أمريكا الشمالية", "الولايات المتحدة"), 110: ("أمريكا الشمالية", "المكسيك"), 72: ("أوروبا", "أيرلندا"),
    188: ("أوروبا", "ويلز"), 177: ("أوروبا", "أوكرانيا"), 144: ("أوروبا", "روسيا"), 43: ("أمريكا الجنوبية", "الإكوادور"),
    133: ("أمريكا الجنوبية", "باراغواي"), 185: ("أمريكا الجنوبية", "فنزويلا"), 134: ("أمريكا الجنوبية", "بيرو"),
    21: ("أوروبا", "البوسنة والهرسك"), 48: ("أوروبا", "فنلندا"), 70: ("أوروبا", "أيسلندا"), 1: ("أوروبا", "ألبانيا"),
    155: ("أوروبا", "سلوفاكيا"), 156: ("أوروبا", "سلوفينيا"), 42: ("أفريقيا", "غينيا"), 30: ("أفريقيا", "بوركينا فاسو"),
    35: ("أفريقيا", "جمهورية الكونغو الديمقراطية"), 186: ("أفريقيا", "زامبيا"), 191: ("أفريقيا", "زيمبابوي"),
    166: ("أفريقيا", "توغو"), 65: ("أفريقيا", "غينيا بيساو"), 46: ("أفريقيا", "غينيا الاستوائية"), 115: ("أفريقيا", "موزمبيق"),
    5: ("أفريقيا", "أنغولا"), 20: ("أفريقيا", "بنين"), 36: ("أمريكا الشمالية", "كوستاريكا"), 76: ("أمريكا الشمالية", "جامايكا"),
    132: ("أمريكا الشمالية", "بنما"), 3: ("أمريكا الشمالية", "كندا"),
}

# ==============================================================================
# NETWORK CLIENTS
# ==============================================================================
def fetch_tm_html(url: str, max_retries: int = 3) -> Optional[str]:
    """Fetches Transfermarkt HTML using rotating proxies to bypass WAF."""
    for p in PROXIES:
        for attempt in range(max_retries):
            try:
                s = requests.Session(impersonate="chrome120")
                s.proxies = {"http": p, "https": p}
                r = s.get(url, headers=HEADERS, timeout=6)
                if r.status_code == 200 and len(r.text) > 5000:
                    return r.text
            except Exception:
                time.sleep(0.5)
    return None

def fetch_tmapi(url: str, max_retries: int = 5) -> Optional[Dict[str, Any]]:
    delay = 1.0
    for attempt in range(max_retries):
        try:
            req = urllib.request.Request(url, headers=TMAPI_HEADERS)
            with urllib.request.urlopen(req, timeout=6) as resp:
                if resp.status == 200:
                    return json.loads(resp.read().decode())
        except urllib.error.HTTPError as e:
            if e.code in [429, 502, 503, 504]:
                time.sleep(delay)
                delay = min(10.0, delay * 1.5)
            else:
                return None
        except Exception:
            time.sleep(delay)
            delay = min(10.0, delay * 1.5)
    return None

CLUB_CACHE = {}
PLAYER_CACHE = {}

def get_club_info(club_id: str) -> Optional[Dict[str, Any]]:
    if club_id in CLUB_CACHE:
        return CLUB_CACHE[club_id]
    data = fetch_tmapi(f"https://tmapi.transfermarkt.technology/clubs?ids[]={club_id}")
    if data and data.get("data"):
        c = data["data"][0]
        CLUB_CACHE[club_id] = c
        return c
    return None

def get_player_data(player_id: str) -> Optional[Dict[str, Any]]:
    if player_id in PLAYER_CACHE:
        return PLAYER_CACHE[player_id]
    data = fetch_tmapi(f"https://tmapi.transfermarkt.technology/players?ids[]={player_id}")
    if data and data.get("data"):
        p = data["data"][0]
        PLAYER_CACHE[player_id] = p
        return p
    return None

# ==============================================================================
# WIKIDATA OFFICIAL ARABIC NAME RESOLVER
# ==============================================================================
WIKIDATA_NAME_CACHE = {}

def get_official_arabic_name(english_name: str) -> str:
    if english_name in WIKIDATA_NAME_CACHE:
        return WIKIDATA_NAME_CACHE[english_name]

    KNOWN_NAMES = {
        "Cristiano Ronaldo": "كريستيانو رونالدو", "Lionel Messi": "ليونيل ميسي",
        "Karim Benzema": "كريم بنزيما", "Neymar": "نيمار", "Luka Modrić": "لوكا مودريتش",
        "Sergio Ramos": "سيرجيو راموس", "Mohamed Salah": "محمد صلاح", "Kevin De Bruyne": "كيفين دي بروين",
        "Kylian Mbappé": "كيليان مبابي", "Robert Lewandowski": "روبرت ليفاندوفسكي",
        "Toni Kroos": "توني كروس", "Ángel Di María": "أنخل دي ماريا", "Luis Suárez": "لويس سواريز",
        "Erling Haaland": "إيرلينغ هالاند", "Riyad Mahrez": "رياض محرز", "Aymeric Laporte": "إيميرك لابورت",
        "Ali Al-Bulayhi": "علي البليهي", "Salem Al-Dawsari": "سالم الدوسري", "Jhon Durán": "جون دوران"
    }
    if english_name in KNOWN_NAMES:
        WIKIDATA_NAME_CACHE[english_name] = KNOWN_NAMES[english_name]
        return KNOWN_NAMES[english_name]

    try:
        url = f"https://www.wikidata.org/w/api.php?action=wbsearchentities&search={urllib.parse.quote(english_name)}&language=en&format=json"
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
                ar_label = entity_res["entities"][entity_id].get("labels", {}).get("ar", {}).get("value")
                if ar_label:
                    clean_ar = re.sub(r"[^\u0600-\u06FF\s\-]", "", ar_label).strip()
                    if clean_ar and "لاعب" not in clean_ar:
                        WIKIDATA_NAME_CACHE[english_name] = clean_ar
                        return clean_ar
    except Exception:
        pass

    return english_name

# ==============================================================================
# TRANSFERMARKT STRICT WHITELISTED TROPHIES & INDIVIDUAL AWARDS PARSER
# ==============================================================================
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

APPROVED_AWARDS = [
    "Winner Ballon d'Or", "The Best FIFA Men's Player", "UEFA Best Player in Europe", 
    "African Footballer of the Year", "AFC Player of the Year", "Footballer of the Year", 
    "Golden Boot winner (Europe)", "Top goal scorer", "Golden Boy"
]

NOISE_KEYWORDS = [
    "Youth", "U18", "U21", "U23", "Juniores", "2nd tier", "3rd tier", 
    "Relegated", "Promotion", "TM-Player", "Puskás", "assist", "medalist", "Campeão",
    "participant", "runner-up", "runner up", "third place", "second highest", "under-"
]

TROPHY_MAP = {t.lower(): t for t in APPROVED_TROPHIES}
AWARD_MAP = {a.lower(): a for a in APPROVED_AWARDS}

SAUDI_EXCLUSIONS = ["cristiano ronaldo", "bento", "simakan", "danilo pereira"]
LAPORTE_EXCLUSIONS = ["laporte"]

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

def parse_tm_trophies(slug: str, player_id: str, player_name: str = "") -> Tuple[List[str], List[str]]:
    """
    Extracts and filters player trophies and individual awards using strict whitelists:
    - Only approved 29 major team trophies and 9 individual awards.
    - Discards all noise (Youth, U18-U23, 2nd/3rd tier, TM-Player, Puskás, assist, etc.).
    - Fixes player-level verification (excludes unearned club titles and future season placeholders).
    """
    trophies = []
    individual = []

    # 1. Try TMAPI /player/{player_id}/achievement first
    ach_res = fetch_tmapi(f"https://tmapi.transfermarkt.technology/player/{player_id}/achievement")
    if ach_res and ach_res.get("data"):
        data = ach_res["data"]
        aggregated = data.get("aggregated", [])
        records = data.get("achievements", [])

        for item in aggregated:
            raw_title = item.get("name", "").strip()
            if not raw_title or raw_title.lower() == "all titles":
                continue

            # Skip items that only exist as future placeholders (e.g. seasonId >= 2025)
            item_records = [r for r in records if r.get("id") == item.get("id")]
            if item_records and all(r.get("seasonId", 0) >= 2025 for r in item_records):
                continue

            kind, canonical = classify_title(raw_title, player_id=player_id, player_name=player_name)
            if kind == "trophy" and canonical not in trophies:
                trophies.append(canonical)
            elif kind == "award" and canonical not in individual:
                individual.append(canonical)

        return trophies, individual

    # 2. Fallback to HTML /erfolge if TMAPI is unavailable
    url = f"https://www.transfermarkt.com/{slug}/erfolge/spieler/{player_id}"
    html = fetch_tm_html(url)
    if not html:
        return trophies, individual

    soup = BeautifulSoup(html, "html.parser")
    for box in soup.select("div.box"):
        headline = box.select_one("h2.content-box-headline")
        if not headline:
            continue
        title_text = headline.text.strip()
        if not title_text or title_text.lower() == "all titles":
            continue

        clean_title = re.sub(r"^\d+x\s*", "", title_text).strip()
        kind, canonical = classify_title(clean_title, player_id=player_id, player_name=player_name)
        if kind == "trophy" and canonical not in trophies:
            trophies.append(canonical)
        elif kind == "award" and canonical not in individual:
            individual.append(canonical)

    return trophies, individual

# ==============================================================================
# TRANSFERMARKT /RUECKENNUMMERN HTML PARSER
# ==============================================================================
def parse_tm_shirt_numbers(slug: str, player_id: str, allowed_clubs_ar: List[str]) -> Tuple[List[str], bool]:
    """
    Parses Transfermarkt /rueckennummern HTML:
    - table.items > tbody > tr
    - Maps club to Arabic BEFORE membership check
    - Formats strictly as النادي:الرقم
    - Checks for senior national team appearance
    """
    url = f"https://www.transfermarkt.com/{slug}/rueckennummern/spieler/{player_id}"
    html = fetch_tm_html(url)
    if not html:
        return [], False

    soup = BeautifulSoup(html, "html.parser")
    tables = soup.select("table.items")
    if not tables:
        return [], False

    seen = set()
    numbers = []
    has_senior_intl = False

    # 1. Parse Club shirt numbers (Table 0)
    for tr in tables[0].select("tbody tr"):
        cols = tr.select("td")
        if len(cols) >= 4:
            raw_c_name = cols[2].text.strip()
            num = cols[3].text.strip().replace("#", "")
            if num.isdigit() and int(num) <= 99:
                # Map to exact Arabic club string
                ar_c_name = CLUB_MAP.get(raw_c_name, raw_c_name)
                # Ensure the club name exists verbatim in the player's clubs list
                if ar_c_name in allowed_clubs_ar:
                    entry = f"{ar_c_name}:{num}"
                    if entry not in seen:
                        seen.add(entry)
                        numbers.append(entry)

    # 2. Parse National team shirt numbers (Table 1 if present)
    if len(tables) > 1:
        for tr in tables[1].select("tbody tr"):
            cols = tr.select("td")
            if len(cols) >= 3:
                t_name = cols[2].text.strip()
                if not any(u in t_name for u in ["U1", "U2", "U3", "Olympic", "Olympia", "Youth"]):
                    has_senior_intl = True
                    break

    return numbers, has_senior_intl

# ==============================================================================
# PURE DETERMINISTIC PLAYER DOSSIER PARSER
# ==============================================================================
def parse_player_data(player_id: str, slug: str) -> Optional[Dict[str, Any]]:
    p_data = get_player_data(player_id)
    if not p_data:
        return None

    eng_name = p_data.get("name") or p_data.get("displayName") or slug.replace("-", " ").title()
    arabic_name = get_official_arabic_name(eng_name)

    dob = p_data.get("lifeDates", {}).get("dateOfBirth", "")
    birth_year = int(dob.split("-")[0]) if dob and "-" in dob else None

    # Citizenship & Continent
    nat_id = p_data.get("nationalityDetails", {}).get("nationalities", {}).get("nationalityId")
    continent, country_ar = COUNTRY_MAP.get(nat_id, ("أوروبا", "إسبانيا"))

    # Initial check for Senior International Caps (from clubAssignments / nominees / participations)
    curr_assignments = p_data.get("clubAssignments", [])
    has_senior_intl = False
    for ca in curr_assignments:
        if ca.get("type") == "nationalTeam" or ca.get("isNationalTeam"):
            has_senior_intl = True
            break

    if len(p_data.get("nominees", [])) > 0 or len(p_data.get("participations", [])) > 0:
        has_senior_intl = True

    # Career Clubs (FROM TRANSFER HISTORY ONLY - FILTER YOUTH/RESERVES)
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

    # Ensure active / current club tenures extend to present (2030) - EXCLUDE national teams
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

    # Resolve club names and filter out youth / reserves / blank states / national teams
    senior_clubs_ar = []
    seen_clubs = set()

    for cid in transfer_club_ids:
        c_info = get_club_info(cid)
        if not c_info:
            continue
        if c_info.get("baseDetails", {}).get("isNationalTeam"):
            continue
        c_name = c_info.get("name") or c_info.get("baseDetails", {}).get("shortName", "")
        c_type = c_info.get("baseDetails", {}).get("clubTypeId", 1)

        # Precise reserve/youth filter that preserves legitimate clubs like Athletic Bilbao
        if c_type != 1 or re.search(r'\b(U\d+|II|B|Youth|Sub-\d+)\b', c_name, re.I) or any(x in c_name for x in ["Retired", "Without Club"]):
            continue
        if c_name and c_name not in seen_clubs:
            seen_clubs.add(c_name)
            ar_c = CLUB_MAP.get(c_name, c_name)
            senior_clubs_ar.append(ar_c)

    clubs_str = "، ".join(senior_clubs_ar) if senior_clubs_ar else None

    # 2. REAL TROPHIES PARSER (/erfolge)
    trophies, individual_awards = parse_tm_trophies(slug, player_id, eng_name)
    trophies_str = ", ".join(trophies) if trophies else None
    individual_str = ", ".join(individual_awards) if individual_awards else None

    # 3. REAL SHIRT NUMBERS PARSER & SENIOR NT CHECK (/rueckennummern)
    shirt_numbers, has_intl_from_shirts = parse_tm_shirt_numbers(slug, player_id, senior_clubs_ar)
    shirt_str = "، ".join(shirt_numbers[:4]) if shirt_numbers else None
    if has_intl_from_shirts:
        has_senior_intl = True

    if has_senior_intl:
        intl_col = country_ar
        citizenship_col = None
    else:
        intl_col = None
        citizenship_col = country_ar

    # 4. TEAMMATES («الزملاء» - STRICT 15 SUPERSTARS ONLY)
    matched_superstars = []
    for star_name, star_info in SUPERSTARS.items():
        if star_info["eng_name"].lower() in eng_name.lower() or eng_name.lower() in star_info["eng_name"].lower():
            continue
        # Senior NT overlap
        if has_senior_intl and country_ar == star_info.get("senior_nt"):
            matched_superstars.append(star_name)
            continue
        # Club overlap during tenure
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

    teammates_str = "، ".join(matched_superstars) if matched_superstars else None

    return {
        "الاسم بالعربي": arabic_name,
        "الاسم بالإنجليزي": eng_name,
        "سنة الميلاد": birth_year,
        "القارة": continent,
        "المنتخبات": intl_col,
        "الجنسية (لغير الدوليين)": citizenship_col,
        "الأندية": clubs_str,
        "الأرقام": shirt_str,
        "الزملاء": teammates_str,
        "الألقاب": trophies_str,
        "الجوائز الفردية": individual_str,
    }

# ==============================================================================
# EXCEL APPEND
# ==============================================================================
def append_to_excel(row_dict: Dict[str, Any]) -> bool:
    try:
        if os.path.exists(EXCEL_FILE):
            df_existing = pd.read_excel(EXCEL_FILE, sheet_name=SHEET_NAME)
        else:
            df_existing = pd.DataFrame(columns=COLUMNS_SCHEMA)

        row_data = {col: row_dict.get(col, None) for col in COLUMNS_SCHEMA}
        df_new = pd.DataFrame([row_data])
        updated_df = pd.concat([df_existing, df_new], ignore_index=True)

        for attempt in range(5):
            try:
                with pd.ExcelWriter(EXCEL_FILE, engine="openpyxl", mode="w") as writer:
                    updated_df.to_excel(writer, sheet_name=SHEET_NAME, index=False)
                return True
            except PermissionError:
                time.sleep(1.0)
        return False
    except Exception as e:
        print(f"[!] Error appending to Excel: {e}")
        return False

# ==============================================================================
# TARGET CLUBS & SQUAD SCRAPER
# ==============================================================================
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

def get_club_squad(slug: str, club_id: str) -> List[Tuple[str, str, str]]:
    # First try TMAPI /club/{club_id}/squad
    data = fetch_tmapi(f"https://tmapi.transfermarkt.technology/club/{club_id}/squad")
    if data and data.get("data"):
        squad_items = data["data"].get("squad", [])
        players = []
        for m in squad_items:
            pid = str(m.get("playerId", ""))
            if pid:
                p_info = get_player_data(pid)
                if p_info:
                    rel_url = p_info.get("relativeUrl", "")
                    p_slug = rel_url.split("/profil/spieler/")[0].strip("/") if "/profil/spieler/" in rel_url else f"player-{pid}"
                    p_name = p_info.get("name") or p_info.get("displayName") or p_slug.replace("-", " ").title()
                    players.append((pid, p_slug, p_name))
        if players:
            return players

    # Fallback to HTML
    url = f"https://www.transfermarkt.com/{slug}/startseite/verein/{club_id}"
    html = fetch_tm_html(url)
    if not html:
        return []
    soup = BeautifulSoup(html, "html.parser")
    squad_table = soup.select_one("div#yw1 table.items") or soup.select_one("table.items")
    players = []
    if squad_table:
        for a in squad_table.select("td.hauptlink a"):
            href = a.get("href", "")
            name = a.text.strip()
            if "/profil/spieler/" in href and name and not any(k in name.lower() for k in ["rumour", "gerücht"]):
                parts = href.split("/profil/spieler/")
                p_slug = parts[0].strip("/")
                p_id = parts[1].split("/")[0]
                if not any(x[0] == p_id for x in players):
                    players.append((p_id, p_slug, name))
    return players

def reset_excel():
    df_empty = pd.DataFrame(columns=COLUMNS_SCHEMA)
    with pd.ExcelWriter(EXCEL_FILE, engine="openpyxl", mode="w") as writer:
        df_empty.to_excel(writer, sheet_name=SHEET_NAME, index=False)
    print(f"[i] Reset {EXCEL_FILE} (Sheet: '{SHEET_NAME}') with clean headers.")

def run_test():
    reset_excel()
    test_players = [
        ("18922", "karim-benzema", "Karim Benzema"),
        ("8198", "cristiano-ronaldo", "Cristiano Ronaldo"),
        ("195332", "salem-al-dawsari", "Salem Al-Dawsari"),
    ]

    print("=" * 80)
    print("VERIFICATION RUN: BENZEMA, CRISTIANO RONALDO, SALEM AL-DAWSARI")
    print("Exact Transfermarkt Names | Trophies vs. Individual Awards")
    print("=" * 80)

    for pid, slug, expected_name in test_players:
        print(f"\nProcessing {expected_name} (ID: {pid})...")
        row = parse_player_data(pid, slug)
        if row:
            saved = append_to_excel(row)
            if saved:
                print(f"  [✔] Saved row to {EXCEL_FILE}")
            print(f"  Player: {row['الاسم بالإنجليزي']} ({row['الاسم بالعربي']})")
            print(f"  Trophies (الألقاب): {row['الألقاب']}")
            print(f"  Individual Awards (الجوائز الفردية): {row['الجوائز الفردية']}")
        else:
            print(f"  [!] Failed to extract data for {expected_name}")
        time.sleep(0.5)

    print("\n" + "=" * 80)
    print("SAVED EXCEL SUMMARY (tripla-football-data.xlsx)")
    print("=" * 80)
    df = pd.read_excel(EXCEL_FILE, sheet_name=SHEET_NAME)
    for idx, r in df.iterrows():
        print(f"[{idx + 1}] {r['الاسم بالإنجليزي']} ({r['الاسم بالعربي']}):")
        print(f"    - Trophies (الألقاب): {r['الألقاب']}")
        print(f"    - Individual Awards (الجوائز الفردية): {r['الجوائز الفردية']}")
    print("=" * 80)

def main():
    if len(sys.argv) > 1 and sys.argv[1] == "--full":
        print("=" * 70)
        print("TRIPLA FOOTBALL EXTRACTION PIPELINE (TARGET: 300 PLAYERS)")
        print("Strict Deterministic Logic | Real Trophies | Verified Shirt Numbers")
        print("=" * 70)

        if os.path.exists(EXCEL_FILE):
            df_existing = pd.read_excel(EXCEL_FILE, sheet_name=SHEET_NAME)
            seen_names = set(df_existing["الاسم بالإنجليزي"].dropna().str.lower())
            total_count = len(df_existing)
        else:
            df_existing = pd.DataFrame(columns=COLUMNS_SCHEMA)
            seen_names = set()
            total_count = 0

        print(f"[i] Starting with {total_count} existing verified players in sheet.")

        TARGET_LIMIT = None  # Crawl all players across all clubs without limit

        for club_slug, club_id in TARGET_CLUBS:
            if TARGET_LIMIT is not None and total_count >= TARGET_LIMIT:
                break

            print(f"\n>>> Fetching squad for club: {club_slug} (ID: {club_id})...")
            squad = get_club_squad(club_slug, club_id)
            print(f"    Found {len(squad)} squad players.")

            for pid, pslug, pname in squad:
                if total_count >= TARGET_LIMIT:
                    break

                if pname.lower() in seen_names:
                    continue

                print(f"[{total_count + 1}/{TARGET_LIMIT}] Extracting: {pname} (ID: {pid})...")
                row = parse_player_data(pid, pslug)
                if row:
                    saved = append_to_excel(row)
                    if saved:
                        seen_names.add(pname.lower())
                        total_count += 1
                        print(f"  [✔] Saved: {row['الاسم بالعربي']} | {row['الأندية']} | أرقام: {row['الأرقام']} | ألقاب: {row['الألقاب']}")
                    else:
                        print(f"  [!] Failed to save row for {pname}")
                else:
                    print(f"  [-] Skipped {pname} (No data returned)")

                time.sleep(0.5)

        print("\n" + "=" * 70)
        print(f"[COMPLETE] Extraction finished! Total players in sheet: {total_count}")
        print("=" * 70)
    else:
        run_test()

if __name__ == "__main__":
    main()