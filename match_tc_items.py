import json
import zipfile
import xml.etree.ElementTree as ET

def load_tc_items():
    with open('tc_items_extracted.json', 'r', encoding='utf-8') as f:
        return json.load(f)

def load_fsbc_resources():
    zip_path = 'ФСНБ-2022 (приказ Минстроя России от 12.08.2026 года № 527пр).zip'
    xml_filename = 'ФСБЦ_Мат&Оборуд.xml'

    with zipfile.ZipFile(zip_path, 'r') as z:
        with z.open(xml_filename) as f:
            tree = ET.parse(f)
            root = tree.getroot()

    resources = []
    for res in root.findall('.//Resource'):
        code = res.attrib.get('Code', '')
        name = res.attrib.get('Name', '')
        unit = res.attrib.get('MeasureUnit', '')
        price_elem = res.find('.//Price')
        price = float(price_elem.attrib.get('Cost', 0)) if price_elem is not None and price_elem.attrib.get('Cost') else 0.0
        resources.append({
            'code': code,
            'name': name,
            'unit': unit,
            'price': price
        })
    return resources

def main():
    tc_items = load_tc_items()
    fsbc = load_fsbc_resources()
    print(f"Loaded {len(tc_items)} TC items and {len(fsbc)} FSBC resources.")

    results = []

    for item in tc_items:
        tc_code = item['code']
        tc_name = item['name']
        tc_unit = item['unit']
        tc_price = item['price']

        # Extract group code from TC code, e.g., '01.6.03.01' from 'ТЦ_01.6.03.01_34_...'
        parts = tc_code.split('_')
        group_prefix = parts[1] if len(parts) > 1 else ''
        parent_prefix = '.'.join(group_prefix.split('.')[:3]) if '.' in group_prefix else group_prefix

        # Search candidates
        # 1. Exact group prefix match
        group_matches = [r for r in fsbc if r['code'].startswith(group_prefix)]
        # 2. Parent group prefix match
        parent_matches = [r for r in fsbc if r['code'].startswith(parent_prefix)]

        results.append({
            'tc_item': item,
            'group_prefix': group_prefix,
            'parent_prefix': parent_prefix,
            'group_count': len(group_matches),
            'parent_count': len(parent_matches),
            'group_matches': group_matches,
            'parent_matches': parent_matches
        })

    print("Search complete. Diagnostic summary:")
    for r in results:
        it = r['tc_item']
        print(f"Row {it['row']:4d} | TC: {it['code'][:20]}... | Group {r['group_prefix']}: {r['group_count']} matches | Parent {r['parent_prefix']}: {r['parent_count']} matches | Name: {it['name'][:40]}")

if __name__ == '__main__':
    main()
