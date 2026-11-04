import re
import pandas as pd

OUTPUT_FILE = "tripla_staging_raw.xlsx"
SHEET_NAME = "Data"

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

def clean_title(item: str):
    item = " ".join(item.split()).strip()
    item = re.sub(r"^\d+x\s*", "", item).strip()
    for noise in NOISE_KEYWORDS:
        if noise.lower() in item.lower():
            return None, None
    low = item.lower()
    if low in TROPHY_MAP:
        return "trophy", TROPHY_MAP[low]
    if low in AWARD_MAP:
        return "award", AWARD_MAP[low]
    return None, None

def filter_staging_dataset():
    df = pd.read_excel(OUTPUT_FILE, sheet_name=SHEET_NAME)
    print(f"Read {len(df)} rows from {OUTPUT_FILE}")

    for idx, row in df.iterrows():
        p_name = str(row["Name"]).lower()
        
        # 1. Filter Trophies
        raw_t = str(row["Trophies"]) if pd.notna(row["Trophies"]) else ""
        valid_t = []
        if raw_t:
            for it in raw_t.split(","):
                kind, canonical = clean_title(it)
                if kind == "trophy":
                    # Player-level exclusions
                    if canonical == "Saudi Arabian champion" and any(bad in p_name for bad in SAUDI_EXCLUSIONS):
                        continue
                    if canonical == "World Cup winner" and any(bad in p_name for bad in LAPORTE_EXCLUSIONS):
                        continue
                    if canonical not in valid_t:
                        valid_t.append(canonical)
        
        # 2. Filter Individual Awards
        raw_a = str(row["Individual Awards"]) if pd.notna(row["Individual Awards"]) else ""
        valid_a = []
        if raw_a:
            for it in raw_a.split(","):
                kind, canonical = clean_title(it)
                if kind == "award":
                    if canonical not in valid_a:
                        valid_a.append(canonical)

        df.at[idx, "Trophies"] = ", ".join(valid_t) if valid_t else None
        df.at[idx, "Individual Awards"] = ", ".join(valid_a) if valid_a else None

    with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl", mode="w") as writer:
        df.to_excel(writer, sheet_name=SHEET_NAME, index=False)
    print(f"[✓] Successfully re-filtered {OUTPUT_FILE} (Sheet: '{SHEET_NAME}').")

    # Verification Print
    print("\n" + "=" * 80)
    print("VERIFICATION OF TARGET PLAYERS")
    print("=" * 80)
    target_names = ["Cristiano Ronaldo", "Aymeric Laporte", "Bento", "Mohamed Simakan", "Danilo Pereira", "Karim Benzema", "Salem Al-Dawsari"]
    for t_name in target_names:
        sub = df[df["Name"].str.contains(t_name, case=False, na=False)]
        if not sub.empty:
            r = sub.iloc[0]
            print(f"Player: {r['Name']} (ID: {r['ID']})")
            print(f"  Trophies: {r['Trophies']}")
            print(f"  Individual Awards: {r['Individual Awards']}\n")

    # Overall Summary
    print("=" * 80)
    print("FINAL DATASET SUMMARY")
    print("=" * 80)
    print(f"Total Rows: {len(df)}")
    print(f"Trophies populated: {df['Trophies'].dropna().count()} / {len(df)} ({df['Trophies'].dropna().count()/len(df)*100:.1f}%)")
    print(f"Individual Awards populated: {df['Individual Awards'].dropna().count()} / {len(df)} ({df['Individual Awards'].dropna().count()/len(df)*100:.1f}%)")
    
    # Assert no non-whitelisted items exist
    all_t = set()
    for t in df["Trophies"].dropna():
        for x in t.split(","):
            all_t.add(x.strip())
    assert all_t.issubset(set(APPROVED_TROPHIES)), f"Non-whitelisted trophies found: {all_t - set(APPROVED_TROPHIES)}"

    all_a = set()
    for a in df["Individual Awards"].dropna():
        for x in a.split(","):
            all_a.add(x.strip())
    assert all_a.issubset(set(APPROVED_AWARDS)), f"Non-whitelisted awards found: {all_a - set(APPROVED_AWARDS)}"

    print("[✓] All assertions passed! 100% of trophies and awards match the approved whitelists.")

if __name__ == "__main__":
    filter_staging_dataset()
