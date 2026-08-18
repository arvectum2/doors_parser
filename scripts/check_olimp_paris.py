import urllib.request
import ssl

ctx = ssl.create_default_context()
ctx.check_hostname = False
ctx.verify_mode = ssl.CERT_NONE

for model in ['paris-1', 'paris-2', 'paris-3', 'paris-4', 'paris-5', 'paris-6', 'paris-7']:
    url = f"https://olymp-doors.ru.tilda.ws/{model}.png"
    try:
        req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'})
        resp = urllib.request.urlopen(req, timeout=15, context=ctx)
        data = resp.read(16)
        ct = resp.headers.get('Content-Type', '')
        actual = 'unknown'
        if data[:3] == b'\xff\xd8\xff': actual = 'JPEG'
        elif data[:4] == b'\x89PNG': actual = 'PNG'
        elif data[:4] == b'RIFF': actual = 'WEBP'
        print(f"  {model}: OK ({actual}, {ct})")
    except urllib.error.HTTPError as e:
        print(f"  {model}: HTTP {e.code}")
    except Exception as e:
        print(f"  {model}: {str(e)[:60]}")
