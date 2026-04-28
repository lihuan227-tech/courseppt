#!/usr/bin/env python3
"""
Fetch artwork images from Wikipedia for embedding in the lesson decks.
Uses MediaWiki API — no external dependencies (pure stdlib).

Run:  python3 download_artworks.py
Then re-run any of:
      python3 create_day2_masters.py
      python3 create_day3_chinese_ink.py
      python3 create_day4_contemporary.py

If an image is in ./images/ with the matching key, the deck embeds it
automatically; otherwise a placeholder + URL is shown.
"""
import json, os, sys, time, urllib.parse, urllib.request

OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "images")
os.makedirs(OUT, exist_ok=True)

# (key, wikipedia title, lang)
ARTWORKS = [
    # D1 references
    ("mona_lisa",       "Mona Lisa", "en"),
    # D2 — Picasso
    ("dora_maar",       "Portrait of Dora Maar", "en"),
    ("demoiselles",     "Les Demoiselles d'Avignon", "en"),
    ("weeping_woman",   "The Weeping Woman", "en"),
    ("guernica",        "Guernica (Picasso)", "en"),
    # D2 — Matisse
    ("polynesia_sea",   "Polynésie, la mer", "en"),
    ("snail",           "The Snail", "en"),
    ("icarus",          "Icarus (Matisse)", "en"),
    ("matisse_red_room","The Dessert: Harmony in Red", "en"),
    # D2 — Van Gogh
    ("starry_rhone",    "Starry Night Over the Rhône", "en"),
    ("sunflowers",      "Sunflowers (Van Gogh series)", "en"),
    ("starry_night",    "The Starry Night", "en"),
    # D3 — Chinese ink (these may not have pageimages on en.wp; try zh.wp too)
    ("xishan_xinglu",   "溪山行旅圖", "zh"),
    ("qi_baishi",       "齊白石", "zh"),
    ("bada_shanren",    "八大山人", "zh"),
    ("zhang_daqian",    "张大千", "zh"),
    ("wu_zuoren",       "吴作人", "zh"),
    ("zheng_banqiao",   "郑燮", "zh"),
    ("han_meilin",      "韩美林", "zh"),
    # D4 — contemporary
    ("dali_clocks",     "The Persistence of Memory", "en"),
    ("magritte_pipe",   "The Treachery of Images", "en"),
    ("bicycle_wheel",   "Bicycle Wheel", "en"),
    ("kusama",          "Yayoi Kusama", "en"),
    ("riley",           "Bridget Riley", "en"),
    ("warhol_soup",     "Campbell's Soup Cans", "en"),
]

def fetch_main_image(title, lang):
    api = f"https://{lang}.wikipedia.org/w/api.php"
    params = {
        "action": "query",
        "format": "json",
        "titles": title,
        "prop": "pageimages",
        "pithumbsize": "2000",
        "redirects": "1",
    }
    url = api + "?" + urllib.parse.urlencode(params)
    req = urllib.request.Request(url, headers={"User-Agent": "GR-EDU-ArtClass/1.0 (educational)"})
    with urllib.request.urlopen(req, timeout=15) as r:
        data = json.load(r)
    pages = data.get("query", {}).get("pages", {})
    for p in pages.values():
        thumb = p.get("thumbnail", {}).get("source")
        if thumb: return thumb
    return None

def save_image(url, key):
    ext = url.rsplit(".", 1)[-1].lower().split("?")[0]
    if ext not in ("jpg", "jpeg", "png", "webp"): ext = "jpg"
    out = os.path.join(OUT, f"{key}.{ext}")
    req = urllib.request.Request(url, headers={"User-Agent": "GR-EDU-ArtClass/1.0 (educational)"})
    with urllib.request.urlopen(req, timeout=30) as r:
        with open(out, "wb") as f:
            f.write(r.read())
    return out

def main():
    print(f"💾 Saving images to: {OUT}")
    ok, fail = [], []
    for i,(key, title, lang) in enumerate(ARTWORKS):
        if i: time.sleep(1.5)  # throttle to avoid 429
        if any(os.path.exists(os.path.join(OUT, f"{key}.{ext}")) for ext in ("jpg","jpeg","png","webp")):
            print(f"⏭️  {key}: already downloaded")
            ok.append(key); continue
        try:
            src = fetch_main_image(title, lang)
            if not src:
                print(f"❌ {key}: no free image on {lang}.wikipedia for '{title}' (likely in copyright)")
                fail.append(key); continue
            path = save_image(src, key)
            print(f"✅ {key}: {path}")
            ok.append(key)
        except Exception as e:
            print(f"❌ {key}: {e}")
            fail.append(key)
    print(f"\nDone: {len(ok)} ok, {len(fail)} failed")
    if fail:
        print("Failed:", ", ".join(fail))
        print("Tip: these may be in copyright (no free image on Wikipedia) —")
        print("the slides will fall back to placeholder + URL text.")

if __name__ == "__main__":
    main()
