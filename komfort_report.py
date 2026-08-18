import sys
sys.stdout.reconfigure(encoding='utf-8')

print('=' * 80)
print('KOMFORT DOORS - COMPLETE SCRAPING RESULTS')
print('=' * 80)
print()
print('SOURCE: https://komfortdoors.ru/')
print()
print('STEP 1: CURRENT CATALOG (v4 zip)')
print('-' * 40)
print('Total Komfort rows: 58')
print()
print('Collections in catalog:')
print('  Art          - 1 model (placeholder)')
print('  Gift         - 1 model (placeholder)')
print('  Malta        - 1 model (placeholder)')
print('  Molding      - 1 model (placeholder)')
print('  Neo          - 1 model (placeholder)')
print('  New line     - 1 model (placeholder)')
print('  Premier      - 1 model (placeholder)')
print('  Slope        - 1 model (placeholder)')
print('  True         - 1 model (placeholder)')
print('  Turin        - 1 model (placeholder)')
print('  Ekoshpon     - 1 model (placeholder)')
print('  Emal         - 47 models (with actual model names)')
print()
print('NOTE: Most collection rows have generic names (just the')
print('collection name) with no individual model details.')
print()
print('=' * 80)
print('STEP 2-3: WEBSITE STRUCTURE & ALL MODELS FOUND')
print('=' * 80)
print()
print('Website domain: komfortdoors.ru')
print('Main category: Emal (emal/) - parent for all 10 sub-collections')
print()
print('SITE CATEGORIES:')
print('  1. Emal (emal/) - 10 sub-collections, 119 models total')
print('  2. Tsargovye dveri s pokrytiem PVC - 76+ models (2 pages)')
print('  3. Karkasno-shchitovye dveri s ABS-kromkoy PVC - 18 models')
print('  4. Invisible (skrytye dveri) - 4 models')
print('  5. Furnitura - accessories (not doors)')
print()
print('=' * 80)
print('DETAILED COLLECTIONS UNDER EMAL')
print('=' * 80)

collections = {
    'NEW LINE': {
        'url': 'https://komfortdoors.ru/new-line/',
        'models': [
            ('NL-1 DG', 'new-line/nl-1-dg', '/assets/cache_image/products/639/nline1_0x0_9d7.webp'),
            ('NL-2 DG', 'new-line/nl-2-dg', '/assets/cache_image/products/640/nline2_0x0_9d7.webp'),
            ('NL-3 DG', 'new-line/nl-3-dg', '/assets/cache_image/products/641/nline3_0x0_9d7.webp'),
            ('NL-4 DG', 'new-line/nl-4-dg', '/assets/cache_image/products/642/nline4_0x0_9d7.webp'),
            ('NL-5 DG', 'new-line/nl-5-dg', '/assets/cache_image/products/643/nline5_0x0_9d7.webp'),
            ('NL-6 DG', 'new-line/nl-6-dg', '/assets/cache_image/products/644/nline6_0x0_9d7.webp'),
            ('NL-7 DG', 'new-line/nl-7-dg', '/assets/cache_image/products/645/nline7_0x0_9d7.webp'),
            ('NL-8 DG', 'new-line/nl-8-dg', '/assets/cache_image/products/646/nline8_0x0_9d7.webp'),
            ('NL-9 DG', 'new-line/nl-9-dg', '/assets/cache_image/products/647/nline9_0x0_9d7.webp'),
            ('NL-10 DG', 'new-line/nl-10-dg', '/assets/cache_image/products/648/nline10_0x0_9d7.webp'),
            ('NL-11 DG', 'new-line/nl-11-dg', '/assets/cache_image/products/649/nline11_0x0_9d7.webp'),
            ('NL-12 DG', 'new-line/nl-12-dg', '/assets/cache_image/products/650/nline12_0x0_9d7.webp'),
        ]
    },
    'SLOPE': {
        'url': 'https://komfortdoors.ru/slope/',
        'models': [
            ('SL 1 DG', 'slope/sl-1-dg', '/assets/cache_image/products/651/slope-1_0x0_9d7.webp'),
            ('SL 2 DG', 'slope/sl-2-dg', '/assets/cache_image/products/652/slope-2_0x0_9d7.webp'),
            ('SL 3 DG', 'slope/sl-3-dg', '/assets/cache_image/products/653/slope-3_0x0_9d7.webp'),
            ('SL 4 DG', 'slope/sl-4-dg', '/assets/cache_image/products/654/slope-4_0x0_9d7.webp'),
            ('SL 5 DG', 'slope/sl-5-dg', '/assets/cache_image/products/655/slope-5_0x0_9d7.webp'),
            ('SL 6 DG', 'slope/sl-6-dg', '/assets/cache_image/products/656/slope-6_0x0_9d7.webp'),
            ('SL 7 DG', 'slope/sl-7-dg', '/assets/cache_image/products/657/slope-7_0x0_9d7.webp'),
            ('SL 8 DG', 'slope/sl-8-dg', '/assets/cache_image/products/658/slope-8_0x0_9d7.webp'),
            ('SL 9 DG', 'slope/sl-9-dg', '/assets/cache_image/products/659/slope-9_0x0_9d7.webp'),
            ('SL 10 DG', 'slope/sl-10-dg', '/assets/cache_image/products/660/slope-10_0x0_9d7.webp'),
            ('SL 11 DG', 'slope/sl-11-dg', '/assets/cache_image/products/661/slope-11_0x0_9d7.webp'),
        ]
    },
    'TRUE': {
        'url': 'https://komfortdoors.ru/true/',
        'models': [
            ('True 1 DG', 'true/true-1-dg', '/assets/cache_image/products/662/true-1-dg_0x0_9d7.webp'),
            ('True 2 DG', 'true/true-2-dg', '/assets/cache_image/products/663/true-2-dg_0x0_9d7.webp'),
            ('True 3 DG', 'true/true-3-dg', '/assets/cache_image/products/664/true-3-dg_0x0_9d7.webp'),
            ('True 3 DO', 'true/true-3-do', '/assets/cache_image/products/665/true-3-do_0x0_9d7.webp'),
            ('True 4 DG', 'true/true-4-dg', '/assets/cache_image/products/666/true-4-dg_0x0_9d7.webp'),
            ('True 4 DO', 'true/true-4-do', '/assets/cache_image/products/667/true-4-do_0x0_9d7.webp'),
            ('True 5 DG', 'true/true-5-dg', '/assets/cache_image/products/668/true-5-dg_0x0_9d7.webp'),
            ('True 5 DO', 'true/true-5-do', '/assets/cache_image/products/669/true-5-do_0x0_9d7.webp'),
            ('True 6 DG', 'true/true-6-dg', '/assets/cache_image/products/670/true-6-dg_0x0_9d7.webp'),
            ('True 6 DO', 'true/true-6-do', '/assets/cache_image/products/671/true-6-do_0x0_9d7.webp'),
            ('True 7 DG', 'true/true-7-dg', '/assets/cache_image/products/672/true-7-dg_0x0_9d7.webp'),
            ('True 7 DO', 'true/true-7-do', '/assets/cache_image/products/673/true-7-do_0x0_9d7.webp'),
        ]
    },
    'GIFT': {
        'url': 'https://komfortdoors.ru/gift/',
        'models': [
            ('GIFT-1 DG', 'gift/gift-1-dg', '/assets/cache_image/products/674/gift1_0x0_9d7.webp'),
            ('GIFT-2 DG', 'gift/gift-2-dg', '/assets/cache_image/products/675/gift2_0x0_9d7.webp'),
            ('GIFT-3 DG', 'gift/gift-3-dg', '/assets/cache_image/products/676/gift3_0x0_9d7.webp'),
            ('GIFT-4 DG', 'gift/gift-4-dg', '/assets/cache_image/products/677/gift4_0x0_9d7.webp'),
            ('GIFT-5 DG', 'gift/gift-5-dg', '/assets/cache_image/products/678/gift5_0x0_9d7.webp'),
            ('GIFT-6 DG', 'gift/gift-6-dg', '/assets/cache_image/products/679/gift6_0x0_9d7.webp'),
        ]
    },
    'ART': {
        'url': 'https://komfortdoors.ru/art/',
        'models': [
            ('ART 3 DG', 'art/art-3-dg', '/assets/cache_image/products/688/art3_0x0_9d7.webp'),
            ('ART 11 DG', 'art/art-11-dg', '/assets/cache_image/products/680/art11_0x0_9d7.webp'),
            ('ART 12 DG', 'art/art-12-dg', '/assets/cache_image/products/681/art12_0x0_9d7.webp'),
            ('ART 13 DG', 'art/art-13-dg', '/assets/cache_image/products/682/art13_0x0_9d7.webp'),
            ('ART 14 DG', 'art/art-14-dg', '/assets/cache_image/products/683/art14_0x0_9d7.webp'),
            ('ART 15 DG', 'art/art-15-dg', '/assets/cache_image/products/684/art15_0x0_9d7.webp'),
            ('ART 16 DG', 'art/art-16-dg', '/assets/cache_image/products/685/art16_0x0_9d7.webp'),
            ('ART 18 DG', 'art/art-18-dg', '/assets/cache_image/products/686/art18_0x0_9d7.webp'),
            ('ART 20 DG', 'art/art-20-dg', '/assets/cache_image/products/687/art20_0x0_9d7.webp'),
        ]
    },
    'PREMIER': {
        'url': 'https://komfortdoors.ru/premier/',
        'models': [
            ('Premier 1 DG', 'premier/premier-1-dg', '/assets/cache_image/products/706/premer-1-nalichnik-bagetnyij-kopiya_0x0_9d7.webp'),
            ('Premier 1 DO', 'premier/premier-1-do', '/assets/cache_image/products/707/premer-1-nalichnik-bagetnyij-do_0x0_9d7.webp'),
            ('Premier 2 DG', 'premier/premier-2-dg', '/assets/cache_image/products/708/premer-2-nalichnik-bagetnyij-kopiya_0x0_9d7.webp'),
            ('Premier 2 DO', 'premier/premier-2-do', '/assets/cache_image/products/709/premer-2-nalichnik-bagetnyij-do_0x0_9d7.webp'),
            ('Premier 3 DG', 'premier/premier-3-dg', '/assets/cache_image/products/710/premer-3-nalichnik-bagetnyij-kopiya_0x0_9d7.webp'),
            ('Premier 3 DO', 'premier/premier-3-do', '/assets/cache_image/products/711/premer-3-nalichnik-bagetnyij-do_0x0_9d7.webp'),
            ('Premier 4/1 DG', 'premier/premier-4/1-dg', '/assets/cache_image/products/712/premer-4-.1-kopiya_0x0_9d7.webp'),
            ('Premier 4/1 DO', 'premier/premier-4/1-do', '/assets/cache_image/products/713/premer-4-.1-do_0x0_9d7.webp'),
            ('Premier 15 DG', 'premier/premier-15-dg', '/assets/cache_image/products/714/premer-15-kopiya_0x0_9d7.webp'),
            ('Premier 16 DG', 'premier/premier-16-dg', '/assets/cache_image/products/715/premer16-kopiya_0x0_9d7.webp'),
            ('Premier 17 DG', 'premier/premier-17-dg', '/assets/cache_image/products/716/premer-17-kopiya_0x0_9d7.webp'),
            ('Premier 18 DG', 'premier/premier-18-dg', '/assets/cache_image/products/717/premer-18-kopiya_0x0_9d7.webp'),
        ]
    },
    'MALTA': {
        'url': 'https://komfortdoors.ru/malta/',
        'models': [
            ('L 6 DG', 'malta/l-6-dg', '/assets/cache_image/products/235/a1-400x800-400x800_0x0_9d7.webp'),
            ('L-10 DO', 'malta/l-10-do', '/assets/cache_image/products/700/l10_0x0_9d7.webp'),
            ('L-12 DG', 'malta/l-12-dg', '/assets/cache_image/products/237/dg-tripleks-400-na-800-s-vodnyim-400x800_0x0_9d7.webp'),
            ('L-12 DO', 'malta/l-12-do', '/assets/cache_image/products/238/tripleks-400-na-800-s-vodnyim-400x800_0x0_9d7.webp'),
            ('L-13 DG', 'malta/l-13-dg', '/assets/cache_image/products/48/malta-l-13-dgbelyij_0x0_9d7.webp'),
            ('L-14 DG', 'malta/l-14-dg', '/assets/cache_image/products/533/malta-l-14-emal-belaya_0x0_9d7.webp'),
            ('L-19 DG', 'malta/l-19-dg', '/assets/cache_image/products/701/l19_0x0_9d7.webp'),
            ('L-20 DG', 'malta/l-20-dg', '/assets/cache_image/products/702/l20_0x0_9d7.webp'),
            ('L-23 DG', 'malta/l-23-dg', '/assets/cache_image/products/229/1-400x800_0x0_9d7.webp'),
            ('L-23 DO', 'malta/l-23-do', '/assets/cache_image/products/230/do-400x800_0x0_9d7.webp'),
            ('L-26 DG', 'malta/l-26-dg', '/assets/cache_image/products/231/22-dg-400x800_0x0_9d7.webp'),
            ('L-26 DO', 'malta/l-26-do', '/assets/cache_image/products/232/2-do-400x800_0x0_9d7.webp'),
            ('L-28 DG', 'malta/l-28-dg', '/assets/cache_image/products/703/l28_0x0_9d7.webp'),
            ('L-29 DG', 'malta/l-29-dg', '/assets/cache_image/products/704/l29_0x0_9d7.webp'),
            ('L-30 DG', 'malta/l-30-dg', '/assets/cache_image/products/705/l-30_0x0_9d7.webp'),
            ('MALTA DG', 'malta/malta-dg', '/assets/cache_image/products/236/d-mono-902kh417-bel-glukh-400x800_0x0_9d7.webp'),
        ]
    },
    'NEO': {
        'url': 'https://komfortdoors.ru/neo/',
        'models': [
            ('Neo 8 DG', 'neo/neo-8-dg', '/assets/cache_image/products/224/400na-800-s-vodnyim-400x800_0x0_9d7.webp'),
            ('Neo 8 DO', 'neo/neo-8-emal-belaya-do', '/assets/cache_image/products/225/do-400-na-800-s-vodnyim-400x800_0x0_9d7.webp'),
            ('Neo 9 DG', 'neo/neo-9-emal-belaya-dg', '/assets/cache_image/products/247/9-400-na-800-s-vodnyim-400x800_0x0_9d7.webp'),
            ('Neo 9 DO', 'neo/neo-9-emal-belaya-do', '/assets/cache_image/products/248/neo9do_0x0_9d7.webp'),
            ('Neo 10 DG', 'neo/neo-10-dg', '/assets/cache_image/products/633/neo-10-dg_0x0_9d7.webp'),
            ('Neo 10 DO', 'neo/neo-8-do', '/assets/cache_image/products/634/neo-10-do_0x0_9d7.webp'),
            ('Neo 11 DG', 'neo/neo-11-dg', '/assets/cache_image/products/635/neo-11-dg_0x0_9d7.webp'),
            ('Neo 11 DO', 'neo/neo-11-do', '/assets/cache_image/products/636/neo-11_0x0_9d7.webp'),
            ('Neo 12 DG', 'neo/neo-12-dg', '/assets/cache_image/products/637/neo-12-dg_0x0_9d7.webp'),
            ('Neo 12 DO', 'neo/neo-12-do', '/assets/cache_image/products/638/neo-12-do_0x0_9d7.webp'),
        ]
    },
    'MOLDING': {
        'url': 'https://komfortdoors.ru/molding/',
        'models': [
            ('M1 DG', 'molding/m1-dg', '/assets/cache_image/products/689/m-1-kopiya_0x0_9d7.webp'),
            ('M2 DG', 'molding/m2-dg', '/assets/cache_image/products/690/m-2-kopiya_0x0_9d7.webp'),
            ('M3 DG', 'molding/m3-dg', '/assets/cache_image/products/691/m-3-kopiya_0x0_9d7.webp'),
            ('M4 DG', 'molding/m4-dg', '/assets/cache_image/products/692/m-4-kopiya_0x0_9d7.webp'),
            ('M5 DG', 'molding/m5-dg', '/assets/cache_image/products/693/m-5-kopiya_0x0_9d7.webp'),
            ('M6 DG', 'molding/m6-dg', '/assets/cache_image/products/694/v6_0x0_9d7.webp'),
            ('M7 DG', 'molding/m7-dg', '/assets/cache_image/products/695/m7_0x0_9d7.webp'),
            ('M8 DG', 'molding/m8-dg', '/assets/cache_image/products/696/m-8_0x0_9d7.webp'),
        ]
    },
    'TURIN': {
        'url': 'https://komfortdoors.ru/turin/',
        'models': [
            ('TURIN 1 DG', 'turin/turin-1-dg', '/assets/cache_image/products/590/turin-1-dg_0x0_9d7.webp'),
            ('TURIN 1 DO', 'turin/turin-1-do', '/assets/cache_image/products/591/turin-1do_0x0_9d7.webp'),
            ('TURIN 2 DG', 'turin/turin-2-dg', '/assets/cache_image/products/592/turin-2-dg-(1)_0x0_9d7.webp'),
            ('TURIN 2 DO', 'turin/turin-2-do', '/assets/cache_image/products/593/turin-2-do-(1)_0x0_9d7.webp'),
            ('TURIN 4 DG', 'turin/turin-4-dg', '/assets/cache_image/products/239/4-dg-kopiya-400x800_0x0_9d7.webp'),
            ('TURIN 4 DO', 'turin/turin-4-do', '/assets/cache_image/products/240/4-do-kopiya-400x800_0x0_9d7.webp'),
            ('TURIN 10 DG', 'turin/turin-10-dg', '/assets/cache_image/products/243/10-dg-kopiya-400x800_0x0_9d7.webp'),
            ('TURIN 10 DO', 'turin/turin-10-do', '/assets/cache_image/products/244/10-do-kopiya-400x800_0x0_9d7.webp'),
            ('TURIN 13 DG', 'turin/turin-13-dg', '/assets/cache_image/products/241/13-dg-kopiya-400x800_0x0_9d7.webp'),
            ('TURIN 13 DO', 'turin/turin-13-do', '/assets/cache_image/products/242/turin-13-do_0x0_9d7.webp'),
            ('TURIN 18 DG', 'turin/turin-18-dg', '/assets/cache_image/products/245/18-emal-belaya-dg-kopiya-400x800-(1)_0x0_9d7.webp'),
            ('TURIN 18 DO', 'turin/turin-18-do', '/assets/cache_image/products/246/18-emal-belaya-do-kopiya-400x800_0x0_9d7.webp'),
            ('TURIN 22 DG', 'turin/turin-22-dg', '/assets/cache_image/products/279/55-dg-400-na-800-s-vodnyim-400x800_0x0_9d7.webp'),
            ('TURIN 22 DO', 'turin/turin-22-do', '/assets/cache_image/products/280/55-do-400-na-800-s-vodnyim-400x800_0x0_9d7.webp'),
            ('TURIN 23 DG', 'turin/turin-23-dg', '/assets/cache_image/products/281/23-dg-400-na-800-s-vodnyim-400x800_0x0_9d7.webp'),
            ('TURIN 23 DO', 'turin/turin-23-do', '/assets/cache_image/products/282/23-do-400-na-800-s-vodnyim-400x800_0x0_9d7.webp'),
            ('TURIN 24 DG', 'turin/turin-24-dg', '/assets/cache_image/products/283/24-dg-bolshaya-s-vodnyim-400x800_0x0_9d7.webp'),
            ('TURIN 24 DO', 'turin/turin-24-emal-belaya-do', '/assets/cache_image/products/284/24-do-400-na-800-s-vodnyim-400x800_0x0_9d7.webp'),
            ('TURIN 25 DG', 'turin/turin-25-dg', '/assets/cache_image/products/285/25-dg-400-na-800-s-vodnyim-400x800_0x0_9d7.webp'),
            ('TURIN 25 DO', 'turin/turin-25-do', '/assets/cache_image/products/286/25-do-400-na-800-s-vodnyim-400x800_0x0_9d7.webp'),
            ('TURIN 32 DG', 'turin/turin-32-dg', '/assets/cache_image/products/697/turin-32_0x0_9d7.webp'),
            ('TURIN 33 DG', 'turin/turin-33-dg', '/assets/cache_image/products/698/turin-33_0x0_9d7.webp'),
            ('TURIN 39 DG', 'turin/turin-39-dg', '/assets/cache_image/products/699/turin-39_0x0_9d7.webp'),
        ]
    },
}

total = 0
for name, data in collections.items():
    count = len(data['models'])
    total += count
    print()
    print('--- {} ({} models) ---'.format(name, count))
    print('Page URL: {}'.format(data['url']))
    for model_name, model_path, img_path in data['models']:
        model_url = 'https://komfortdoors.ru/{}'.format(model_path)
        img_url = 'https://komfortdoors.ru{}'.format(img_path)
        print('  {}'.format(model_name))
        print('    URL:   {}'.format(model_url))
        print('    Image: {}'.format(img_url))

print()
print()
print('Total Emal sub-collection models: {}'.format(total))
print()
print('=' * 80)
print('STEP 4: SUMMARY')
print('=' * 80)
print()
print('Current Komfort rows in v4 catalog: 58')
print('Total models found on site (Emal only): {}'.format(total))
print()
print('Models per collection (on site):')
for name, data in collections.items():
    print('  {:12s} - {:3d} models'.format(name, len(data['models'])))
print()
print('EXPECTED COLLECTIONS STATUS:')
print('  Art          - FOUND on site:  9 models | In catalog: 1 (placeholder)')
print('  Gift         - FOUND on site:  6 models | In catalog: 1 (placeholder)')
print('  Malta        - FOUND on site: 16 models | In catalog: 1 (placeholder)')
print('  Molding      - FOUND on site:  8 models | In catalog: 1 (placeholder)')
print('  Neo          - FOUND on site: 10 models | In catalog: 1 (placeholder)')
print('  New line (NL)- FOUND on site: 12 models | In catalog: 1 (placeholder)')
print('  Premier      - FOUND on site: 12 models | In catalog: 1 (placeholder)')
print('  Slope        - FOUND on site: 11 models | In catalog: 1 (placeholder)')
print('  True         - FOUND on site: 12 models | In catalog: 1 (placeholder)')
print('  Turin        - FOUND on site: 23 models | In catalog: 1 (placeholder)')
print('  Ekoshpon     - NOT FOUND on site (404)  | In catalog: 1 (placeholder)')
print('  Emal         - PARENT CATEGORY          | In catalog: 47 models')
print()
print('MISSING FROM CATALOG:')
print('  - 62 individual models from Emal sub-collections (119 on site - 47 in catalog)')
print('  - Ekoshpon collection is NOT on the website (possibly discontinued)')
print()
print('ALL 119 MODELS ON SITE HAVE:')
print('  - model name (as displayed on card)')
print('  - model URL (detail page)')
print('  - image URL (thumbnail from listing)')
print()
print('NOTE: Image URLs are THUMBNAILS from listing pages.')
print('Full-size images may use different paths on detail pages.')
print()
print('OTHER CATEGORIES NOT IN SCOPE (but exist on site):')
print('  - Tsargovye dveri PVC: ~76 models (Klassik, Alpha, Sigma series)')
print('  - Karkasno-shchitovye PVC: 18 models (Kosta, Tekhno series)')
print('  - Invisible: 4 models (AVERS/REVERS variants)')
