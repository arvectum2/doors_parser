import csv, os, json, urllib.request, ssl, re, time

ctx = ssl.create_default_context()
ctx.check_hostname = False
ctx.verify_mode = ssl.CERT_NONE
opener = urllib.request.build_opener(urllib.request.HTTPSHandler(context=ctx))

BASE = r"C:\Users\GN\Desktop\doors_parser_ready_fix"
CATALOG = os.path.join(BASE, "output", "delivery_v5_2", "doors_catalog.csv")
V4 = os.path.join(BASE, "output", "final_rework_ready_v4", "doors_catalog.csv")

with open(CATALOG, encoding='utf-8-sig') as f:
    rows = list(csv.DictReader(f, delimiter=';'))
fieldnames = list(rows[0].keys())
with open(V4, encoding='utf-8-sig') as f:
    v4_rows = list(csv.DictReader(f, delimiter=';'))
v4_asd = {r['model_name']: r for r in v4_rows if r['producer'] == 'АСД'}

missing = [(i,r) for i,r in enumerate(rows) if r['producer']=='АСД' and not r.get('image_url_inside','').strip()]
print(f"Missing inside: {len(missing)}")

def wb_available(url):
    try:
        req = urllib.request.Request(f"https://archive.org/wayback/available?url={url}", headers={'User-Agent':'Mozilla/5.0'})
        resp = opener.open(req, timeout=15)
        d = json.loads(resp.read())
        s = d.get('archived_snapshots',{}).get('closest')
        return s['url'] if s else None
    except: return None

def wb_images(wb_url):
    try:
        req = urllib.request.Request(wb_url, headers={'User-Agent':'Mozilla/5.0'})
        resp = opener.open(req, timeout=15)
        html = resp.read().decode('utf-8', errors='ignore')
        imgs = re.findall(r'(?:src|data-src)="([^"]+)"', html, re.I)
        result = []
        for img in imgs:
            if not any(x in img.lower() for x in ['jpg','png','webp']): continue
            if any(x in img.lower() for x in ['logo','icon','star','arrow','banner','badge','loading','toolbar','novinka_round']): continue
            m = re.search(r'/web/\d+im_/(https?://.+)', img)
            if m: img = m.group(1)
            result.append(img)
        return result
    except: return []

def is_img(url):
    for u in [url.replace('https://','http://'), url]:
        try:
            req = urllib.request.Request(u, headers={'User-Agent':'Mozilla/5.0'})
            resp = opener.open(req, timeout=10)
            data = resp.read(4)
            if data[:3]==b'\xff\xd8\xff' or data[:4]==b'\x89PNG' or data[:4]==b'RIFF': return True
        except: pass
    return False

found = 0
for idx, r in missing:
    name = r['model_name']
    v4r = v4_asd.get(name, {})
    source = v4r.get('source_url', '')
    front = r.get('image_url_front','').strip()
    print(f"\n--- {name[:55]} ---")
    print(f"  source: {source[:70]}")

    # 1) Wayback via source_url
    if source:
        wb = wb_available(source)
        if wb:
            print(f"  wayback: {wb[:60]}")
            imgs = wb_images(wb)
            prod = [i for i in imgs if 'novinka_round2' not in i]
            if len(prod) >= 2:
                cand = prod[1]
                if is_img(cand):
                    rows[idx]['image_url_inside'] = cand.replace('https://','http://')
                    found += 1
                    print(f"  -> INSIDE from wayback img2: {cand[:60]}")
                    continue
            print(f"  wayback imgs: {len(prod)} product imgs, no inside candidate")
        else:
            print(f"  no wayback snapshot")
        time.sleep(0.3)

    # 2) Try source_url with .html variants
    if source:
        for variant in [source.replace('.html',''), source.rstrip('/')]:
            wb = wb_available(variant)
            if wb:
                imgs = wb_images(wb)
                prod = [i for i in imgs if 'novinka_round2' not in i]
                if len(prod) >= 2 and is_img(prod[1]):
                    rows[idx]['image_url_inside'] = prod[1].replace('https://','http://')
                    found += 1
                    print(f"  -> INSIDE from variant wayback: {prod[1][:60]}")
                    break
            time.sleep(0.3)
        else:
            continue
        continue

    # 3) Try from front URL directory
    if front:
        m = re.search(r'(/(?:img|images)/\d{4}/[^/]+/)', front)
        if m:
            d = m.group(1)
            base = front.split('/')[-1]
            bm = re.match(r'(.+?)(?:-\d+)?\.(.+)$', base)
            if bm:
                bn, ext = bm.group(1), bm.group(2)
                for n in range(1,30):
                    t = f"http://as-doors.ru{d}{bn}-{n}.{ext}"
                    if t != front and is_img(t):
                        rows[idx]['image_url_inside'] = t
                        found += 1
                        print(f"  -> INSIDE from pattern: {t[:60]}")
                        break

print(f"\nTotal found: {found}")

with open(CATALOG, 'w', encoding='utf-8-sig', newline='') as f:
    csv.DictWriter(f, fieldnames=fieldnames, delimiter=';').writerows(rows)

asd = [r for r in rows if r['producer']=='АСД']
f=sum(1 for r in asd if r.get('image_url_front','').strip())
i=sum(1 for r in asd if r.get('image_url_inside','').strip())
b=sum(1 for r in asd if r.get('image_url_front','').strip() and r.get('image_url_inside','').strip())
print(f"Final: front={f}/199 inside={i}/199 both={b}/199")
miss = [r['model_name'][:50] for r in asd if not r.get('image_url_inside','').strip()]
print(f"Still missing ({len(miss)}):")
for n in miss: print(f"  {n}")
