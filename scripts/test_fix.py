import json

# Read corrupted JSON
with open('output/delivery_v5_2/wayback_inside_all_results.json') as f:
    wb = json.load(f)

k0 = list(wb.keys())[0]
print('Original key:', repr(k0[:80]))
print('Original hex:', k0.encode('utf-8').hex(':'))

# Try fix
try:
    encoded = k0.encode('cp1251')
    print(f'CP1251 encoded: {encoded.hex(":")}')
    decoded = encoded.decode('utf-8')
    print(f'UTF-8 decoded: {repr(decoded[:80])}')
    print(f'Fixed hex: {decoded.encode("utf-8").hex(":")}')
except Exception as e:
    print(f'CP1251 failed: {e}')

# Try latin-1
try:
    encoded = k0.encode('latin-1')
    print(f'Latin-1 encoded: {encoded.hex(":")}')
    decoded = encoded.decode('utf-8')
    print(f'Fixed via latin-1: {repr(decoded[:80])}')
except Exception as e:
    print(f'Latin-1 failed: {e}')
