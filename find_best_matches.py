import json
import zipfile
import xml.etree.ElementTree as ET

def load_fsbc():
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

def search_fsbc(fsbc, group_prefix=None, keywords=None, exclude_words=None):
    res = []
    for r in fsbc:
        if group_prefix and not r['code'].startswith(group_prefix):
            continue
        n_lower = r['name'].lower()
        if keywords:
            if not all(kw.lower() in n_lower for kw in keywords):
                continue
        if exclude_words:
            if any(ex.lower() in n_lower for ex in exclude_words):
                continue
        res.append(r)
    return res

def main():
    with open('tc_items_extracted.json', 'r', encoding='utf-8') as f:
        tc_items = json.load(f)
    fsbc = load_fsbc()

    for idx, item in enumerate(tc_items, 1):
        print("="*80)
        print(f"Item #{idx} | Row {item['row']} | Pos: {item['pos']}")
        print(f"TC Code:  {item['code']}")
        print(f"TC Name:  {item['name']}")
        print(f"TC Unit:  {item['unit']} | Qty: {item['qty']} | Price: {item['price']}")

        parts = item['code'].split('_')
        group = parts[1] if len(parts) > 1 else ''

        # Print matches depending on item
        matches = search_fsbc(fsbc, group_prefix=group)
        if not matches and '.' in group:
            parent = '.'.join(group.split('.')[:-1])
            matches = search_fsbc(fsbc, group_prefix=parent)
            print(f"(Searched parent group {parent}: {len(matches)} matches)")
        else:
            print(f"(Searched group {group}: {len(matches)} matches)")

        for m in matches[:10]:
            print(f"   [{m['code']}] {m['name']} | Unit: {m['unit']} | Base Price: {m['price']}")

if __name__ == '__main__':
    main()
