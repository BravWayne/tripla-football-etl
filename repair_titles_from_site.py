"""Re-fetch Trophies / Individual Awards from the live Transfermarkt /erfolge page
for every row in tripla_staging_raw.xlsx, so all titles are verbatim from the website.
Rows whose page still cannot be loaded are listed at the end (left unchanged)."""
import pandas as pd
import tripla_staging_raw as t

df = pd.read_excel(t.OUTPUT_FILE, sheet_name=t.SHEET_NAME, dtype={"ID": str})
failed = []
changed = 0
for i, row in df.iterrows():
    before = t.TITLE_STATS["html"]
    tr, aw = t.extract_titles(str(row["ID"]), row["Slug"])
    if t.TITLE_STATS["html"] == before:
        # one more attempt with extra rounds
        html = t.fetch_html(f"https://www.transfermarkt.com/{row['Slug']}/erfolge/spieler/{row['ID']}", rounds=4)
        if html:
            tr, aw = t.extract_titles(str(row["ID"]), row["Slug"])
        if t.TITLE_STATS["html"] == before:
            failed.append((row["ID"], row["Name"]))
            continue
    new_t = ", ".join(tr) if tr else None
    new_a = ", ".join(aw) if aw else None
    if (new_t, new_a) != (row["Trophies"] if pd.notna(row["Trophies"]) else None,
                          row["Individual Awards"] if pd.notna(row["Individual Awards"]) else None):
        changed += 1
    df.at[i, "Trophies"] = new_t
    df.at[i, "Individual Awards"] = new_a
    print(f"[{i + 1}/{len(df)}] {row['Name']}: OK")

with pd.ExcelWriter(t.OUTPUT_FILE, engine="openpyxl", mode="w") as w:
    df.to_excel(w, sheet_name=t.SHEET_NAME, index=False)

print(f"\nRows verified from live site: {len(df) - len(failed)}/{len(df)} | rows changed: {changed}")
print("Still not loadable:", failed if failed else "none")
