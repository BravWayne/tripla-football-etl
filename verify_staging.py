import pandas as pd

df = pd.read_excel("tripla_staging_raw.xlsx", sheet_name="Data")

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

all_t = set()
for t in df["Trophies"].dropna():
    for x in t.split(","):
        all_t.add(x.strip())

all_a = set()
for a in df["Individual Awards"].dropna():
    for x in a.split(","):
        all_a.add(x.strip())

print(f"Distinct Trophies in Sheet ({len(all_t)}):")
for x in sorted(all_t):
    assert x in APPROVED_TROPHIES, f"Illegal trophy: {x}"
    print("  ✓", x)

print(f"\nDistinct Awards in Sheet ({len(all_a)}):")
for x in sorted(all_a):
    assert x in APPROVED_AWARDS, f"Illegal award: {x}"
    print("  ✓", x)

print("\nPlayer-Level Verification:")
for target, forbidden_t in [
    ("Cristiano Ronaldo", "Saudi Arabian champion"),
    ("Bento", "Saudi Arabian champion"),
    ("Mohamed Simakan", "Saudi Arabian champion"),
    ("Danilo Pereira", "Saudi Arabian champion"),
    ("Aymeric Laporte", "World Cup winner")
]:
    sub = df[df["Name"] == target]
    t_val = str(sub["Trophies"].values[0]) if not sub.empty else ""
    items = [x.strip() for x in t_val.split(",")]
    assert forbidden_t not in items, f"{target} STILL HAS {forbidden_t}"
    print(f"  ✓ Verified: {target} does NOT have '{forbidden_t}'")

print("\n[SUCCESS] 100% of trophies and individual awards are verified against all rules!")
