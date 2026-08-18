import json

with open('output/delivery_v5_2/wayback_inside_all_results.json') as f:
    content = f.read()
# Check raw content for first key
import re
m = re.match(r'\{\s*\n\s*"(.+?)":', content, re.DOTALL)
if m:
    raw_key = m.group(1)
    print('Raw JSON key:', raw_key)
    print('Raw hex:', raw_key.encode('utf-8').hex(':'))

# Now parse and check
f.seek(0)
wb = json.loads(content)
parsed_key = list(wb.keys())[0]
print()
print('Parsed key: ', parsed_key)
print('Parsed hex:', parsed_key.encode('utf-8').hex(':'))

# Compare with an expected string
expected = '\u0022\u041f\u0440\u043e\u0442\u0438\u0432\u043e\u043f\u043e\u0436\u0430\u0440\u043d\u0430\u044f\u0022 \u0434\u0432\u0435\u0440\u044c. \u0417\u0430\u043a\u0430\u0437\u043d\u0430\u044f \u043c\u043e\u0434\u0435\u043b\u044c'
print()
print('Expected:   ', expected)
print('Expected hex:', expected.encode('utf-8').hex(':'))
print()
print('raw_key == expected:', raw_key == expected)
print('parsed_key == expected:', parsed_key == expected)
print('raw_key == parsed_key:', raw_key == parsed_key)
