import sys, re, urllib.request
sys.stdout.reconfigure(encoding='utf-8')

pages = {
    'Париж 1': 'https://olymp-doors.ru/paris-1',
    'Париж 2': 'https://olymp-doors.ru/paris-2',
    'Париж 3': 'https://olymp-doors.ru/paris-3',
    'Париж 4': 'https://olymp-doors.ru/paris-4',
    'Париж 5': 'https://olymp-doors.ru/paris-5',
    'Париж 6': 'https://olymp-doors.ru/paris-6',
    'Париж 7': 'https://olymp-doors.ru/paris-7',
}

for name, url in pages.items():
    req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
    resp = urllib.request.urlopen(req, timeout=15)
    html = resp.read().decode('utf-8', errors='replace')
    
    # og:image
    m = re.search(r'og:image"\s+content="([^"]+)"', html)
    img = m.group(1) if m else 'NOT FOUND'
    
    # title
    m2 = re.search(r'<title>([^<]+)</title>', html)
    title = m2.group(1).strip() if m2 else 'N/A'
    
    print(f"{name}")
    print(f"  URL: {url}")
    print(f"  Title: {title}")
    print(f"  Image: {img}")
    print()
