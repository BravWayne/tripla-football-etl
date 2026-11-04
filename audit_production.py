import openpyxl
import re

DATA_FILE = "tripla-football-data.xlsx"
SHEET_NAME = "البيانات"

EXPECTED_COLUMNS = [
    'الاسم بالعربي', 'الاسم بالإنجليزي', 'سنة الميلاد', 'القارة', 
    'المنتخبات', 'الجنسية (لغير الدوليين)', 'الأندية', 'الأرقام', 
    'الزملاء', 'الألقاب', 'الجوائز الفردية'
]

ARABIC_CONTINENTS = {'أوروبا', 'أمريكا الجنوبية', 'أفريقيا', 'آسيا', 'أمريكا الشمالية', 'أوقيانوسيا'}

APPROVED_TROPHIES_AR = {
    "دوري أبطال أوروبا", "الدوري الأوروبي", "دوري المؤتمر الأوروبي", "دوري أبطال آسيا",
    "كوبا ليبرتادوريس", "كوبا سود أمريكانا", "كأس العالم للأندية", "الدوري الإنجليزي",
    "الدوري الإسباني", "الدوري الإيطالي", "الدوري الألماني", "الدوري الفرنسي",
    "الدوري السعودي", "الدوري البرتغالي", "الدوري الهولندي", "كأس الاتحاد الإنجليزي",
    "كأس الرابطة الإنجليزية", "كأس ملك إسبانيا", "كأس إيطاليا", "كأس ألمانيا",
    "كأس فرنسا", "كأس خادم الحرمين الشريفين", "كأس ولي العهد السعودي", "كأس العالم",
    "كأس أمم أوروبا", "كوبا أمريكا", "كأس أمم أفريقيا", "كأس أمم آسيا", "دوري الأمم الأوروبية"
}

APPROVED_AWARDS_AR = {
    "الكرة الذهبية", "جائزة الأفضل من فيفا", "أفضل لاعب في أوروبا",
    "أفضل لاعب في أفريقيا", "أفضل لاعب في آسيا", "لاعب العام",
    "الحذاء الذهبي الأوروبي", "هداف البطولة", "الفتى الذهبي (غولدن بوي)"
}

APPROVED_SUPERSTARS_AR = {
    "ليونيل ميسي", "كريستيانو رونالدو", "كريم بنزيما", "نيمار", "لوكا مودريتش",
    "سيرجيو راموس", "محمد صلاح", "كيفين دي بروين", "كيليان مبابي",
    "روبرت ليفاندوفسكي", "توني كروس", "أنخل دي ماريا", "لويس سواريز",
    "إيرلينغ هالاند", "رياض محرز"
}

def audit():
    wb = openpyxl.load_workbook(DATA_FILE, read_only=True)
    assert SHEET_NAME in wb.sheetnames, f"Missing sheet {SHEET_NAME}"
    ws = wb[SHEET_NAME]

    rows = list(ws.iter_rows(values_only=True))
    headers = list(rows[0])
    assert headers == EXPECTED_COLUMNS, f"Headers mismatch: {headers}"
    data_rows = rows[1:]

    print(f"Total Rows Audited: {len(data_rows)}")
    assert len(data_rows) == 1575, f"Expected 1575 rows, found {len(data_rows)}"

    ar_char_re = re.compile(r'[\u0600-\u06FF]')

    for idx, r in enumerate(data_rows, start=2):
        ar_name, eng_name, birth_year, continent, nat_team, cit, clubs, shirts, teammates, trophies, awards = r

        # 1. Names
        assert ar_name and ar_char_re.search(str(ar_name)), f"Row {idx}: Invalid Arabic name '{ar_name}'"
        assert eng_name and len(str(eng_name)) > 1, f"Row {idx}: Invalid English name '{eng_name}'"

        # 2. Birth Year
        assert birth_year and 1970 <= int(birth_year) <= 2012, f"Row {idx}: Invalid birth year {birth_year}"

        # 3. Continent
        assert continent in ARABIC_CONTINENTS, f"Row {idx}: Invalid continent '{continent}'"

        # 4. Mutual Exclusivity: National Team vs Citizenship
        has_nt = bool(nat_team and str(nat_team).strip())
        has_cit = bool(cit and str(cit).strip())
        assert has_nt != has_cit, f"Row {idx}: NT ({nat_team}) and Citizenship ({cit}) must be mutually exclusive!"

        # 5. Clubs
        assert clubs and str(clubs).strip(), f"Row {idx}: Clubs empty"
        clubs_list = [c.strip() for c in str(clubs).split("، ")]

        # 6. Shirt Numbers (Check club name alignment with Clubs column)
        if shirts:
            for s in str(shirts).split("، "):
                assert ":" in s, f"Row {idx}: Invalid shirt format '{s}'"
                c_part, num_part = s.rsplit(":", 1)
                assert c_part.strip() in clubs_list, f"Row {idx}: Shirt club '{c_part}' not found in clubs '{clubs}'"

        # 7. Teammates
        if teammates:
            for t in str(teammates).split("، "):
                assert t.strip() in APPROVED_SUPERSTARS_AR, f"Row {idx}: Teammate '{t}' not in approved superstars"

        # 8. Trophies
        if trophies:
            for tr in str(trophies).split("، "):
                assert tr.strip() in APPROVED_TROPHIES_AR, f"Row {idx}: Trophy '{tr}' not in approved trophies"

        # 9. Individual Awards
        if awards:
            for aw in str(awards).split("، "):
                assert aw.strip() in APPROVED_AWARDS_AR, f"Row {idx}: Award '{aw}' not in approved awards"

        # 10. Player-Specific Integrity Assertions
        eng_lower = str(eng_name).lower()
        if any(p in eng_lower for p in ["cristiano ronaldo", "bento", "simakan", "danilo pereira"]):
            if trophies:
                assert "الدوري السعودي" not in str(trophies), f"Row {idx}: {eng_name} falsely has 'الدوري السعودي'!"

        if "laporte" in eng_lower:
            if trophies:
                assert "كأس العالم" not in [t.strip() for t in str(trophies).split("، ") if t.strip() == "كأس العالم"], f"Row {idx}: Laporte falsely has 'كأس العالم'!"

        if "benzema" in eng_lower:
            if trophies:
                tr_items = [t.strip() for t in str(trophies).split("، ")]
                assert "كأس العالم للأندية" in tr_items, f"Row {idx}: Benzema missing 'كأس العالم للأندية'!"
                assert "كأس العالم" not in [t for t in tr_items if t == "كأس العالم"], f"Row {idx}: Benzema falsely has 'كأس العالم'!"

    print("[SUCCESS] All 1,575 rows passed 100% of audit checks perfectly!")

if __name__ == "__main__":
    audit()
