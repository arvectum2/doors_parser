"""Scrape Hausdoors 4 missing lines"""
import requests, json
from bs4 import BeautifulSoup

HEADERS = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"}

lines = {
    "\u0424\u0430\u043a\u0442\u0443\u0440\u043d\u044b\u0435 \u043b\u0438\u043d\u0438\u0438": "https://hausdoors.ru/dveri-facturnye-linii",
    "\u041d\u043e\u0432\u0430\u044f \u043a\u043b\u0430\u0441\u0441\u0438\u043a\u0430": "https://hausdoors.ru/dveri-new-classic",
    "\u041c\u0438\u043d\u0438\u043c\u0430\u043b\u0438\u0437\u043c": "https://hausdoors.ru/dveri-minimalism",
    "\u0421\u043e\u0432\u0440\u0435\u043c\u0435\u043d\u043d\u0430\u044f \u043a\u043b\u0430\u0441\u0441\u0438\u043a\u0430": "https://hausdoors.ru/dveri-soft-classic",
}

results = {}
for line, url in lines.items():
    print(f"\n--- {line} ---")
    try:
        resp = requests.get(url, timeout=15, headers=HEADERS)
        print(f"  Status: {resp.status_code}, Final: {resp.url}")
        if resp.status_code == 200:
            soup = BeautifulSoup(resp.text, "html.parser")
            h1 = soup.find("h1")
            print(f"  H1: {h1.get_text(strip=True)[:50] if h1 else 'N/A'}")
            
            # Find all product-like elements
            cards = soup.select(".product-card, .catalog-item, .card, [class*=product], .item, .door")
            print(f"  Product cards: {len(cards)}")
            
            # Find all links with text
            links = soup.select("a[href]")
            model_links = []
            for a in links:
                href = a.get("href", "")
                text = a.get_text(strip=True)[:60]
                if text and len(text) > 3 and "/dveri" in href:
                    model_links.append({"href": href, "text": text})
            print(f"  Door links: {len(model_links)}")
            for ml in model_links[:10]:
                print(f"    {ml['href'][:50]} | {ml['text'][:40]}")
            
            # Find images
            imgs = soup.select("img")
            door_imgs = []
            for img in imgs:
                src = img.get("src") or img.get("data-src") or ""
                if src and ("dver" in src.lower() or "product" in src.lower() or "catalog" in src.lower()):
                    door_imgs.append(src)
            print(f"  Door images: {len(door_imgs)}")
            
            results[line] = {"models": model_links, "cards": len(cards)}
    except Exception as e:
        print(f"  Error: {e}")

print(f"\n{'='*60}")
print("Summary:")
for line, data in results.items():
    print(f"  {line}: {len(data['models'])} models, {data['cards']} cards")
