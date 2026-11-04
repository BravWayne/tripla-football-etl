import pandas as pd
import re
from openai import OpenAI
import json

client = OpenAI(base_url="http://127.0.0.1:11434/v1", api_key="ollama")
df = pd.read_excel("tripla-football-data.xlsx", sheet_name="البيانات")
non_ar = re.compile(r"[^\u0600-\u06FF\s\-\.\'\’]")
remaining = [(idx, row["الاسم بالإنجليزي"]) for idx, row in df.iterrows() if non_ar.search(str(row["الاسم بالعربي"]))]

print(f"Remaining to fix: {len(remaining)}")
for idx, eng in remaining:
    prompt = f"Translate player name '{eng}' into Arabic script only. Output JSON like: {{\"arabic\": \"...\"}}"
    try:
        res = client.chat.completions.create(
            model="qwen2.5:7b",
            response_format={"type": "json_object"},
            messages=[
                {"role": "system", "content": "You translate names into Arabic letters only. JSON only."},
                {"role": "user", "content": prompt}
            ],
            temperature=0.1
        )
        data = json.loads(res.choices[0].message.content)
        val = list(data.values())[0] if data else ""
        clean_val = re.sub(r"[^\u0600-\u06FF\s\-]", "", str(val)).strip()
        if clean_val:
            df.at[idx, "الاسم بالعربي"] = clean_val
            print(f"Fixed [{idx}]: {eng} -> {clean_val}")
    except Exception as e:
        print(f"Error on {eng}: {e}")

with pd.ExcelWriter("tripla-football-data.xlsx", engine="openpyxl", mode="w") as writer:
    df.to_excel(writer, sheet_name="البيانات", index=False)
print("Complete cleanup finished successfully!")
