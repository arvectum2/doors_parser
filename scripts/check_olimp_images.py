import urllib.request
import ssl

ctx = ssl.create_default_context()
ctx.check_hostname = False
ctx.verify_mode = ssl.CERT_NONE

# Check various image extensions for Paris-3 and Paris-7
for model in ['paris-3', 'paris-7']:
    for ext in ['.png', '.jpg', '.jpeg', '.webp', '.gif']:
        url = f"https://olymp-doors.ru.tilda.ws/{model}{ext}"
        try:
            req = urllib.request.Request(url, method='HEAD', headers={'User-Agent': 'Mozilla/5.0'})
            resp = urllib.request.urlopen(req, timeout=10, context=ctx)
            ct = resp.headers.get('Content-Type', '')
            print(f"  OK: {url} ({ct})")
        except urllib.error.HTTPError as e:
            print(f"  {e.code}: {url}")
        except Exception as e:
            print(f"  ERR: {url}: {str(e)[:60]}")

# Check if the page itself has any image references
for model in ['paris-1', 'paris-3', 'paris-7']:
    url = f"https://olymp-doors.ru.tilda.ws/{model}"
    try:
        req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
        resp = urllib.request.urlopen(req, timeout=10, context=ctx)
        html = resp.read().decode('utf-8', errors='replace')
        # Find image references
        import re
        imgs = re.findall(r'(?:src|href)=["\']([^"\']*\.(?:png|jpg|jpeg|webp))', html)
        print(f"\n{model} page images: {imgs[:5]}")
        # Also check for og:image
        og = re.findall(r'og:image.*?content=["\']([^"\']+)', html)
        print(f"  og:image: {og}")
    except Exception as e:
        print(f"  {model} page error: {str(e)[:60]}")
