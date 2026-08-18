import json

# The corrupted JSON
with open('output/delivery_v5_2/wayback_inside_all_results.json') as f:
    wb = json.load(f)

# Try to fix double encoding
def fix_mojibake(s):
    """Try to fix UTF-8 -> CP1251 -> UTF-8 double encoding"""
    try:
        return s.encode('cp1251').decode('utf-8')
    except (UnicodeEncodeError, UnicodeDecodeError):
        pass
    try:
        return s.encode('latin-1').decode('utf-8')
    except:
        return s

# Test on first key
k0 = list(wb.keys())[0]
fixed = fix_mojibake(k0)
print('Original:', repr(k0[:60]))
print('Fixed:   ', repr(fixed[:60]))
print()

# Check if fixed keys match catalog
import csv
ASD = '\u0410\u0421\u0414'
with open('output/final_rework_ready_v5/doors_catalog.csv', encoding='utf-8-sig') as f:
    cat = list(csv.DictReader(f, delimiter=';'))
asd = [r for r in cat if r['producer'] == ASD]
cat_names = [r['model_name'] for r in asd]

# Fix all wb keys
wb_fixed = {}
for orig_k, v in wb.items():
    fixed_k = fix_mojibake(orig_k)
    wb_fixed[fixed_k] = v

matched = sum(1 for cn in cat_names if cn in wb_fixed)
print(f'After fixing: {matched}/199 matched')

# Check which ones match
matched_names = [cn for cn in cat_names if cn in wb_fixed]
unmatched = [cn for cn in cat_names if cn not in wb_fixed]
print(f'Matched: {len(matched_names)}, Unmatched: {len(unmatched)}')

# Show sample
if matched_names:
    print('Sample matched:', repr(matched_names[0][:60]))
    print('Has inside:', bool(wb_fixed[matched_names[0]].get('inside_url')))
