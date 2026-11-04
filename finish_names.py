import json
import re
import time
import urllib.request
import urllib.parse
from concurrent.futures import ThreadPoolExecutor, as_completed

NAMES_CACHE_FILE = "names_cache.json"

def translate_mymemory(text: str) -> str:
    try:
        url = f"https://api.mymemory.translated.net/get?q={urllib.parse.quote(text)}&langpair=en|ar"
        req = urllib.request.Request(url, headers={"User-Agent": "TriplaFootball/2.0 (contact@tripla.app)"})
        with urllib.request.urlopen(req, timeout=4) as resp:
            data = json.loads(resp.read().decode())
            tr = data.get("responseData", {}).get("translatedText", "")
            clean = re.sub(r"[^\u0600-\u06FF\s\-]", "", tr).strip()
            if clean and "MYMEMORY" not in tr.upper():
                return clean
    except Exception:
        pass
    return ""

def main():
    with open(NAMES_CACHE_FILE, "r", encoding="utf-8") as f:
        names_cache = json.load(f)

    ar_pattern = re.compile(r'[\u0600-\u06FF]')
    remaining = [k for k, v in names_cache.items() if not ar_pattern.search(str(v))]
    print(f"Total remaining names to translate: {len(remaining)}", flush=True)

    def process(name):
        ar = translate_mymemory(name)
        return name, ar

    done = 0
    with ThreadPoolExecutor(max_workers=5) as executor:
        futures = {executor.submit(process, n): n for n in remaining}
        for f in as_completed(futures):
            name, ar = f.result()
            if ar:
                names_cache[name] = ar
            done += 1
            if done % 50 == 0 or done == len(remaining):
                print(f"  [{done}/{len(remaining)}] Sample: {name} -> {ar}", flush=True)
                with open(NAMES_CACHE_FILE, "w", encoding="utf-8") as out:
                    json.dump(names_cache, out, ensure_ascii=False, indent=2)

    with open(NAMES_CACHE_FILE, "w", encoding="utf-8") as out:
        json.dump(names_cache, out, ensure_ascii=False, indent=2)

    total_ar = sum(1 for v in names_cache.values() if ar_pattern.search(str(v)))
    print(f"\nFinished! Total Arabic names in cache: {total_ar} / {len(names_cache)}", flush=True)

if __name__ == "__main__":
    main()
