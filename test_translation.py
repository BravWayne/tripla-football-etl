import openpyxl
import json
import re
import urllib.request
import urllib.parse
import ssl
import time

SSL_CTX = ssl.create_default_context()
SSL_CTX.check_hostname = False
SSL_CTX.verify_mode = ssl.CERT_NONE
WIKI_HEADERS = {"User-Agent": "TriplaFootballCrawler/2.0 (https://tripla.app; contact@tripla.app) python-urllib"}

# 1. Continents
CONTINENTS_MAP = {
    "Europe": "أوروبا",
    "South America": "أمريكا الجنوبية",
    "Africa": "أفريقيا",
    "Asia": "آسيا",
    "North America": "أمريكا الشمالية",
    "Oceania": "أوقيانوسيا",
}

# 2. Countries
COUNTRIES_MAP = {
    'Albania': 'ألبانيا', 'Algeria': 'الجزائر', 'Angola': 'أنغولا', 'Argentina': 'الأرجنتين',
    'Armenia': 'أرمينيا', 'Australia': 'أستراليا', 'Austria': 'النمسا', 'Belgium': 'بلجيكا',
    'Bosnia-Herzegovina': 'البوسنة والهرسك', 'Brazil': 'البرازيل', 'Bulgaria': 'بلغاريا',
    'Burkina Faso': 'بوركينا فاسو', 'Cameroon': 'الكاميرون', 'Canada': 'كندا', 'Cape Verde': 'الرأس الأخضر',
    'Chile': 'تشيلي', 'Colombia': 'كولومبيا', 'Comoros': 'جزر القمر', "Cote d'Ivoire": 'ساحل العاج',
    'Croatia': 'كرواتيا', 'Curacao': 'كوراساو', 'Czech Republic': 'التشيك',
    'DR Congo': 'جمهورية الكونغو الديمقراطية', 'Denmark': 'الدنمارك', 'Dominican Republic': 'جمهورية الدومينيكان',
    'Ecuador': 'الإكوادور', 'Egypt': 'مصر', 'England': 'إنجلترا', 'Equatorial Guinea': 'غينيا الاستوائية',
    'Finland': 'فنلندا', 'France': 'فرنسا', 'French Guiana': 'غيانا الفرنسية', 'Gabon': 'الغابون',
    'Georgia': 'جورجيا', 'Germany': 'ألمانيا', 'Ghana': 'غانا', 'Greece': 'اليونان',
    'Guadeloupe': 'غوادلوب', 'Guinea': 'غينيا', 'Guinea-Bissau': 'غينيا بيساو', 'Haiti': 'هايتي',
    'Hungary': 'المجر', 'Iceland': 'أيسلندا', 'Indonesia': 'إندونيسيا', 'Iraq': 'العراق',
    'Ireland': 'أيرلندا', 'Israel': 'إسرائيل', 'Italy': 'إيطاليا', 'Jamaica': 'جامايكا',
    'Japan': 'اليابان', 'Jordan': 'الأردن', 'Kenya': 'كينيا', 'Korea, South': 'كوريا الجنوبية',
    'Kosovo': 'كوسوفو', 'Liberia': 'ليبيريا', 'Lithuania': 'ليتوانيا', 'Luxembourg': 'لوكسمبورغ',
    'Malaysia': 'ماليزيا', 'Mali': 'مالي', 'Mauritania': 'موريتانيا', 'Mexico': 'المكسيك',
    'Morocco': 'المغرب', 'Mozambique': 'موزمبيق', 'Netherlands': 'هولندا', 'New Caledonia': 'كاليدونيا الجديدة',
    'New Zealand': 'نيوزيلندا', 'Nigeria': 'نيجيريا', 'North Macedonia': 'مقدونيا الشمالية',
    'Northern Ireland': 'أيرلندا الشمالية', 'Norway': 'النرويج', 'Panama': 'بنما', 'Paraguay': 'باراغواي',
    'Poland': 'بولندا', 'Portugal': 'البرتغال', 'Romania': 'رومانيا', 'Russia': 'روسيا',
    'Rwanda': 'رواندا', 'Saudi Arabia': 'السعودية', 'Scotland': 'اسكتلندا', 'Senegal': 'السنغال',
    'Serbia': 'صربيا', 'Slovakia': 'سلوفاكيا', 'Slovenia': 'سلوفينيا', 'Spain': 'إسبانيا',
    'Suriname': 'سورينام', 'Sweden': 'السويد', 'Switzerland': 'سويسرا', 'The Gambia': 'غامبيا',
    'Togo': 'توغو', 'Tunisia': 'تونس', 'Türkiye': 'تركيا', 'Ukraine': 'أوكرانيا',
    'United States': 'الولايات المتحدة', 'Uruguay': 'أوروغواي', 'Uzbekistan': 'أوزبكستان',
    'Venezuela': 'فنزويلا', 'Wales': 'ويلز', 'Zambia': 'زامبيا'
}

# 3. Superstars
SUPERSTARS_MAP = {
    "Lionel Messi": "ليونيل ميسي", "Cristiano Ronaldo": "كريستيانو رونالدو",
    "Karim Benzema": "كريم بنزيما", "Neymar": "نيمار", "Luka Modrić": "لوكا مودريتش",
    "Sergio Ramos": "سيرجيو راموس", "Mohamed Salah": "محمد صلاح",
    "Kevin De Bruyne": "كيفين دي بروين", "Kylian Mbappé": "كيليان مبابي",
    "Robert Lewandowski": "روبرت ليفاندوفسكي", "Toni Kroos": "توني كروس",
    "Ángel Di María": "أنخل دي ماريا", "Luis Suárez": "لويس سواريز",
    "Erling Haaland": "إيرلينغ هالاند", "Riyad Mahrez": "رياض محرز",
}

# 4. Trophies
TROPHIES_MAP = {
    "UEFA Champions League winner": "دوري أبطال أوروبا",
    "Europa League winner": "الدوري الأوروبي",
    "Conference League winner": "دوري المؤتمر الأوروبي",
    "AFC Champions League winner": "دوري أبطال آسيا",
    "Copa Libertadores winner": "كوبا ليبرتادوريس",
    "Copa Sudamericana winner": "كوبا سود أمريكانا",
    "FIFA Club World Cup winner": "كأس العالم للأندية",
    "English Champion": "الدوري الإنجليزي",
    "Spanish champion": "الدوري الإسباني",
    "Italian champion": "الدوري الإيطالي",
    "German Champion": "الدوري الألماني",
    "French champion": "الدوري الفرنسي",
    "Saudi Arabian champion": "الدوري السعودي",
    "Portuguese champion": "الدوري البرتغالي",
    "Dutch champion": "الدوري الهولندي",
    "English FA Cup winner": "كأس الاتحاد الإنجليزي",
    "English League Cup winner": "كأس الرابطة الإنجليزية",
    "Spanish cup winner": "كأس ملك إسبانيا",
    "Italian cup winner": "كأس إيطاليا",
    "German cup winner": "كأس ألمانيا",
    "French cup winner": "كأس فرنسا",
    "Saudi Cup Winner": "كأس خادم الحرمين الشريفين",
    "Saudi Crown Prince Cup Winner": "كأس ولي العهد السعودي",
    "World Cup winner": "كأس العالم",
    "European champion": "كأس أمم أوروبا",
    "Copa América winner": "كوبا أمريكا",
    "Africa Cup winner": "كأس أمم أفريقيا",
    "Asian Cup winner": "كأس أمم آسيا",
    "Winner UEFA Nations League": "دوري الأمم الأوروبية",
}

# 5. Individual Awards
AWARDS_MAP = {
    "Winner Ballon d'Or": "الكرة الذهبية",
    "The Best FIFA Men's Player": "جائزة الأفضل من فيفا",
    "UEFA Best Player in Europe": "أفضل لاعب في أوروبا",
    "African Footballer of the Year": "أفضل لاعب في أفريقيا",
    "AFC Player of the Year": "أفضل لاعب في آسيا",
    "Footballer of the Year": "لاعب العام",
    "Golden Boot winner (Europe)": "الحذاء الذهبي الأوروبي",
    "Top goal scorer": "هداف البطولة",
    "Golden Boy": "الفتى الذهبي (غولدن بوي)",
}

# Load known club maps
from tripla_crawler import CLUB_MAP

KNOWN_CLUBS = dict(CLUB_MAP)
KNOWN_CLUBS.update({
    "Brentford FC": "برينتفورد",
    "AFC Bournemouth": "بورنموث",
    "Crystal Palace": "كريستال بالاس",
    "Nottingham Forest": "نوتينغهام فورست",
    "Sunderland AFC": "سندرلاند",
    "Leeds United": "ليدز يونايتد",
    "Fulham FC": "فولهام",
    "Ipswich Town": "إيبسويتش تاون",
    "Coventry City": "كوفنتري سيتي",
    "Hull City": "هال سيتي",
    "Brighton & Hove Albion": "برايتون",
    "Wolverhampton Wanderers": "وولفرهامبتون",
    "West Ham United": "وست هام يونايتد",
    "Southampton FC": "ساوثهامبتون",
    "Leicester City": "ليستر سيتي",
    "CA Osasuna": "أوساسونا",
    "Deportivo Alavés": "ديبورتيفو ألافيس",
    "Celta de Vigo": "سيلتا فيغو",
    "RCD Espanyol Barcelona": "إسبانيول",
    "Deportivo de La Coruña": "ديبورتيفو لاكورونيا",
    "Racing Santander": "راسينغ سانتاندير",
    "UD Levante": "ليفانتي",
    "Elche CF": "إلتشي",
    "Málaga CF": "مالقا",
    "NEOM SC": "نيوم",
    "Al-Diriyah FC": "الدرعية",
    "Al-Kholood": "الخلود",
    "Al-Fayha FC": "الفيحاء",
    "Al-Riyadh SC": "الرياض",
    "Al-Hazem SC": "الحزم",
    "Al-Khaleej FC": "الخليج",
    "Al-Faisaly FC": "الفيصلي",
    "Abha Club": "أبها",
    "Damac FC": "ضمك",
    "Al-Okhdood Club": "الأخدود",
    "Al-Wehda FC": "الوحدة",
    "Al-Raed SFC": "الرائد",
    "Al-Tai FC": "الطائي",
    "Al-Adalah FC": "العدالة",
    "Al-Batin FC": "الباطن",
})

print(f"Loaded {len(KNOWN_CLUBS)} known clubs.")
