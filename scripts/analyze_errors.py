import json
from collections import Counter
from urllib.parse import urlparse

with open('output/delivery/image_url_validation_report.json', encoding='utf-8') as f:
    report = json.load(f)

errors = report.get('errors_sampled', [])
print(f"Total sampled errors: {len(errors)}")

domain_errors = Counter()
domain_error_types = {}
for e in errors:
    url = e['url']
    err = e['error']
    domain = urlparse(url).netloc
    domain_errors[domain] += 1
    if domain not in domain_error_types:
        domain_error_types[domain] = Counter()
    if '503' in err: domain_error_types[domain]['503'] += 1
    elif 'timed out' in err.lower(): domain_error_types[domain]['timeout'] += 1
    elif '302' in err: domain_error_types[domain]['302'] += 1
    elif 'ascii' in err.lower(): domain_error_types[domain]['ascii'] += 1
    else: domain_error_types[domain][err[:30]] += 1

print("\nErrors by domain:")
for d, c in domain_errors.most_common():
    print(f"  {d}: {c}")
    for t, tc in domain_error_types[d].most_common():
        print(f"    {t}: {tc}")

ascii_errors = [e for e in errors if 'ascii' in e['error'].lower()]
print(f"\nASCII encoding errors: {len(ascii_errors)}")
for e in ascii_errors[:5]:
    print(f"  {e['url'][:100]}")
    print(f"    Error: {e['error'][:80]}")
