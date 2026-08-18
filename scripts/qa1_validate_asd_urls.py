"""QA-1: HTTP-validate all ASD image URLs (front + inside) using requests"""
import csv, json, os, time
from concurrent.futures import ThreadPoolExecutor, as_completed
import requests

BASE = r"C:\Users\GN\Desktop\doors_parser_ready_fix"
CAT = os.path.join(BASE, "output", "delivery_v5_2", "doors_catalog.csv")
OUT_DIR = os.path.join(BASE, "output", "delivery_v5_2")

ASD_KEY = "\u0410\u0421\u0414"

with open(CAT, encoding='utf-8-sig') as f:
    rows = list(csv.DictReader(f, delimiter=';'))
asd = [r for r in rows if r['producer'] == ASD_KEY]
print(f"ASD rows: {len(asd)}")

# Collect unique URLs
url_tasks = []
for r in asd:
    name = r['model_name']
    for side in ('front', 'inside'):
        url = r.get(f'image_url_{side}', '').strip()
        if url:
            url_tasks.append((name, side, url))
print(f"URLs to validate: {len(url_tasks)}")

# HTTP check with requests
session = requests.Session()
session.headers.update({'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'})
session.verify = False
import urllib3
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

def check_url(name, side, url, timeout=15):
    try:
        resp = session.head(url, timeout=timeout, allow_redirects=True)
        ct = resp.headers.get('Content-Type', '')
        is_image = ct.startswith('image/')
        if not is_image and resp.status_code == 200:
            # HEAD might not return image content-type for some servers, try GET with range
            try:
                resp2 = session.get(url, timeout=timeout, stream=True, headers={'Range': 'bytes=0-1023'})
                body = resp2.raw.read(16)
                ct2 = resp2.headers.get('Content-Type', ct)
                is_image = ct2.startswith('image/') or body[:4] in (b'\xff\xd8\xff\xe0', b'\xff\xd8\xff\xe1', b'\x89PNG', b'GIF89a', b'GIF87a', b'RIFF')
                if is_image:
                    return name, side, url, resp2.status_code, ct2, None
            except:
                pass
        return name, side, url, resp.status_code, ct, None
    except requests.exceptions.SSLError as e:
        # Try without SSL
        try:
            url_http = url.replace('https://', 'http://')
            resp = session.head(url_http, timeout=timeout, allow_redirects=True, verify=False)
            ct = resp.headers.get('Content-Type', '')
            return name, side, url, resp.status_code, ct, None
        except Exception as e2:
            return name, side, url, 0, '', str(e2)
    except Exception as e:
        return name, side, url, 0, '', str(e)

# Run validation
results = {}
errors = 0
t0 = time.time()

with ThreadPoolExecutor(max_workers=20) as ex:
    fut_map = {ex.submit(check_url, *t): t for t in url_tasks}
    for i, fut in enumerate(as_completed(fut_map), 1):
        name, side, url, status, ct, err = fut.result()
        results[(name, side)] = {'url': url, 'http_status': status, 'content_type': ct, 'error': err}
        is_ok = status == 200 and ct.startswith('image/')
        if not is_ok:
            errors += 1
        if i % 20 == 0 or i == len(url_tasks):
            elapsed = time.time() - t0
            print(f"  [{i}/{len(url_tasks)}] {elapsed:.0f}s, failures={errors}")

elapsed = time.time() - t0
print(f"\nValidation: {len(url_tasks)} URLs in {elapsed:.0f}s, failures={errors}")

# Save
result_path = os.path.join(OUT_DIR, "asd_url_validation_results.json")
serializable = {}
for (name, side), v in results.items():
    serializable[f"{name}||{side}"] = v
with open(result_path, 'w', encoding='utf-8') as f:
    json.dump(serializable, f, ensure_ascii=False, indent=2)
print(f"Saved: {result_path}")
