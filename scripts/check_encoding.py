import sys, locale, json

print('Default encoding:', sys.getdefaultencoding())
print('Filesystem encoding:', sys.getfilesystemencoding())
print('Locale preferred encoding:', locale.getpreferredencoding())
print('Python version:', sys.version)

for enc in ['cp1251', 'utf-8', None]:
    try:
        if enc:
            with open('output/delivery_v5_2/fixed_test.json', encoding=enc) as f:
                data = json.load(f)
        else:
            with open('output/delivery_v5_2/fixed_test.json') as f:
                data = json.load(f)
        k0 = list(data.keys())[0]
        h = k0.encode('utf-8').hex(':')
        is_correct = 'd0:9f' in h[:30]
        print(f'enc={enc or "default"}: hex={h[:30]}... correct={is_correct}')
    except Exception as e:
        print(f'enc={enc}: ERROR {e}')
