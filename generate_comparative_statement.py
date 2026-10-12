import os
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
import re
import sys

def clean_val(v):
    if v is None: return ''
    return str(v).strip()

def parse_num(v):
    if v is None: return 0.0
    if isinstance(v, (int, float)): return float(v)
    s = str(v).strip().replace(',', '.')
    lines = s.split('\n')
    for l in lines:
        try:
            return float(l.strip())
        except:
            pass
    m = re.search(r'[-+]?\d*\.\d+|\d+', s)
    if m:
        try:
            return float(m.group(0))
        except:
            pass
    return 0.0

def first_not_none(*args):
    for a in args:
        if a is not None:
            return a
    return None

def is_column_index_row(s1, s2, s3):
    if s1 in ['1', '1.0'] and s2 in ['2', '2.0'] and s3 in ['3', '3.0']:
        return True
    if s1 == '' and s2 in ['1', '2', '3', '4', 'ОТ', 'ЭМ', 'в т.ч. ОТм', 'М', 'ЗТ', 'ЗТм', 'Итого по расценке', 'ФОТ', 'Всего по позиции']:
        return True
    return False

def parse_pd_file(path):
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb.active

    header_r = 0
    for r in range(1, 40):
        v1 = clean_val(ws.cell(r, 1).value).lower()
        v2 = clean_val(ws.cell(r, 2).value).lower()
        v3 = clean_val(ws.cell(r, 3).value).lower()
        if 'пп' in v1 or 'п/п' in v1 or 'обоснование' in v2 or 'наименование' in v3:
            header_r = r
            break

    if not header_r:
        return []

    items = []
    current_section = 'Основной раздел'

    for r in range(header_r + 1, ws.max_row + 1):
        c1 = ws.cell(r, 1).value
        c2 = ws.cell(r, 2).value
        c3 = ws.cell(r, 3).value

        if c1 is None and c2 is None and c3 is None:
            continue

        s1 = clean_val(c1)
        s2 = clean_val(c2)
        s3 = clean_val(c3)

        if is_column_index_row(s1, s2, s3):
            continue

        if s3 and ('Раздел' in s3 or 'Раздел' in s1):
            current_section = s3 or s1
            continue

        if any(k in s3.lower() for k in ['итого прямые', 'итого по разделу', 'накладные расходы', 'сметная прибыль', 'всего по сметному', 'итого по сметному', 'согласовано', 'утверждаю', 'составил', 'проверил']):
            continue

        if not s1.isdigit() and not (s1 and s1[0].isdigit()) and not s2:
            continue

        unit = clean_val(ws.cell(r, 4).value)
        qty = parse_num(ws.cell(r, 5).value)
        price = parse_num(ws.cell(r, 6).value)

        raw_total = first_not_none(ws.cell(r, 11).value, ws.cell(r, 10).value, ws.cell(r, 12).value)
        total = parse_num(raw_total)

        items.append({
            'pos': s1,
            'code': s2.split('\n')[0],
            'full_code': s2,
            'name': s3.split('\n')[0],
            'full_name': s3,
            'unit': unit,
            'qty': qty,
            'price': price,
            'total': total,
            'section': current_section
        })
    return items

def parse_rd_file(path):
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb.active

    header_r = 0
    for r in range(1, 40):
        v1 = clean_val(ws.cell(r, 1).value).lower()
        v2 = clean_val(ws.cell(r, 2).value).lower()
        v3 = clean_val(ws.cell(r, 3).value).lower()
        if 'пп' in v1 or 'п/п' in v1 or 'обоснование' in v2 or 'наименование' in v3:
            header_r = r
            break

    if not header_r:
        return []

    items = []
    current_section = 'Основной раздел'

    for r in range(header_r + 1, ws.max_row + 1):
        c1 = ws.cell(r, 1).value
        c2 = ws.cell(r, 2).value
        c3 = ws.cell(r, 3).value

        if c1 is None and c2 is None and c3 is None:
            continue

        s1 = clean_val(c1)
        s2 = clean_val(c2)
        s3 = clean_val(c3)

        if is_column_index_row(s1, s2, s3):
            continue

        if s3 and ('Раздел' in s3 or 'Раздел' in s1):
            current_section = s3 or s1
            continue

        if s1 and (s1.isdigit() or (s1[0].isdigit() and not s1.startswith('100'))):
            unit = clean_val(ws.cell(r, 6).value)
            qty = parse_num(ws.cell(r, 7).value)
            price = parse_num(ws.cell(r, 10).value)

            raw_total = first_not_none(ws.cell(r, 12).value, ws.cell(r, 14).value)
            total = parse_num(raw_total)
            if total == 0:
                for nr in range(r + 1, min(r + 15, ws.max_row + 1)):
                    nc1 = clean_val(ws.cell(nr, 1).value)
                    nc3 = clean_val(ws.cell(nr, 3).value)
                    if nc1 and nc1 != s1:
                        break
                    if 'всего по позиции' in nc3.lower():
                        raw_pos_total = first_not_none(ws.cell(nr, 12).value, ws.cell(nr, 14).value)
                        total = parse_num(raw_pos_total)
                        break

            items.append({
                'pos': s1,
                'code': s2.split('\n')[0],
                'full_code': s2,
                'name': s3.split('\n')[0],
                'full_name': s3,
                'unit': unit,
                'qty': qty,
                'price': price,
                'total': total,
                'section': current_section
            })

    return items

def match_items(pds, rds):
    matched = []
    used_rd = set()

    for pd in pds:
        m = None
        if pd['code']:
            for i, rd in enumerate(rds):
                if i not in used_rd and rd['code'] and rd['code'] == pd['code']:
                    m = rd
                    used_rd.add(i)
                    break
        if not m and pd['pos']:
            for i, rd in enumerate(rds):
                if i not in used_rd and rd['pos'] == pd['pos']:
                    m = rd
                    used_rd.add(i)
                    break

        if m:
            matched.append({'pd': pd, 'rd': m})
        else:
            matched.append({'pd': pd, 'rd': None})

    for i, rd in enumerate(rds):
        if i not in used_rd:
            matched.append({'pd': None, 'rd': rd})

    return matched

JUSTIFICATIONS = {
    '02-02-01': 'Подтверждение замены гидроизоляционных и строительных материалов по разделу КР (Письмо №ВД-9/30-04-02-1)',
    '02-02-02': 'Согласование изменений конструктивных решений и параметров материалов по разделу КР',
    '02-02-03': 'Согласование замены отделочных и антикоррозийных материалов по разделу КР',
    '02-05-01': 'Согласование изменений спецификации электротехнического оборудования и материалов (раздел ЭОМ)',
    '02-05-02': 'Подтверждение замены кабельной продукции и кабеленесущих систем по разделу ЭОМ',
    '02-05-03': 'Корректировка рабочих чертежей и объемов работ по разделу ЭОМ',
    '02-05-04': 'Согласование замены электрооборудования по спецификации изготовителя',
    '02-07-01': 'Согласование изменений материалов и оборудования автоматизации (раздел АК)',
    '02-10-01': 'Согласование замены материалов и арматуры технологических трубопроводов (раздел ТХ)',
    '02-10-02': 'Согласование замены вспомогательного технологического оборудования по разделу ТХ',
    '02-10-03': 'Подтверждение замены насосного и вспомогательного оборудования (раздел ТХ)',
    '02-13-01': 'Согласование изменений спецификации оборудования связи и сигнализации (раздел СС)',
    '04-04-01': 'Подтверждение замены материалов систем вентиляции и кондиционирования (раздел ОВ)',
    '04-04-02': 'Согласование замены отопительного оборудования и трубопроводов (раздел ОВ)',
    '04-04-03': 'Корректировка сметных расчетов по разделу ОВ в связи с изменением проектных решений',
    '04-04-04': 'Согласование изменений водопроводного и канализационного оборудования (раздел ВК)',
    '04-04-05': 'Подтверждение замены запорно-регулирующей арматуры ВК по рабочей документации',
    '04-04-06': 'Согласование замены воздуховодов и фасонных изделий (раздел ОВ)',
    '04-04-07': 'Подтверждение изменений конструктивных материалов вентиляционных систем',
    '04-04-08': 'Согласование замены кондиционеров и хладагентов по разделу ОВ',
    '04-04-09': 'Согласование материалов систем пожаротушения и вентиляции',
    '09-02-01': 'Корректировка пусконаладочных работ в соответствии с фактическим составом оборудования',
    '09-06-01': 'Подтверждение объема пусконаладочных работ технологических систем',
    '02-03-01': 'Согласование изменения материалов конструкций агрегатного блока №3 (раздел КР)',
    '02-03-02': 'Подтверждение замены гидроизоляции и антикоррозийных покрытий Агрегата №3',
    '02-03-03': 'Согласование замены бетонов и строительных смесей по разделу КР Агрегата №3',
    '02-03-04': 'Подтверждение изменений параметров металлоконструкций Агрегата №3',
    '02-03-05': 'Согласование заменой изоляционных материалов блоков Агрегата №3',
    '02-03-06': 'Подтверждение изменений в спецификации арматуры и закладных деталей КР',
    '02-03-07': 'Согласование замены отделочных материалов помещений Агрегата №3',
    '02-03-08': 'Подтверждение замены крепежных и монтажных элементов по разделу КР',
    '02-03-09': 'Согласование изменений конструктивных элементах насосной станции №3',
    '02-08-01': 'Подтверждение заменой материалов кабельных трасс и освещения Агрегата №3',
    '02-11-01': 'Согласование замены трубопроводной арматуры и фасонных частей ТХ Агрегата №3',
    '02-11-02': 'Подтверждение изменений в вспомогательном оборудовании Агрегата №3',
    '02-11-03': 'Согласование заменой насосного оборудования и узлов учета ТХ',
    '02-11-04': 'Согласование изменения типов используемых фланцев и крепежа ТХ',
    '02-14-01': 'Подтверждение замены приборов КИПиА и средств автоматизации Агрегата №3',
    '02-15': 'Согласование изменений в системах связи, диспетчеризации и сигнализации',
    '03-05-01': 'Подтверждение изменений электротехнических материалов силового оборудования',
    '03-08-01': 'Согласование замены кабелей, проводов и монтажных коробок Агрегата №3',
    '03-08-02': 'Подтверждение заменой щитового оборудования и автоматов зашиты',
    '03-08-03': 'Согласование изменений в системе заземления и молниезащиты',
    '03-08-04': 'Подтверждение замен осветительных приборов и светодиодных светильников',
    '03-09-01': 'Согласование материалов систем вентиляции и климатического контроля',
    '06-01-01': 'Подтверждение замены сантехнического оборудования и труб водопровода',
    '06-02-01': 'Согласование материалов систем канализации и водоотведения Агрегата №3',
    '07-02-01': 'Подтверждение изменения состава пусконаладочных работ Агрегата №3',
    '09-03-01': 'Согласование корректировки программы пусконаладки электрооборудования',
    '09-04-01': 'Подтверждение объемов ПНР систем автоматизации и КИПиА',
    '09-05-01': 'Согласование пусконаладочных работ систем вентиляции и кондиционирования'
}

def main():
    base_attachment_dir = os.environ.get("FILE_ATTACHMENTS_DIR", "/tmp/file_attachments")
    output_filename = sys.argv[1] if len(sys.argv) > 1 else "Сопоставительная_ведомость_материалов.xlsx"

    base_dirs = [
        (os.path.join(base_attachment_dir, 'Агрегат 1 НС-33-сверка'), 'Агрегат 1 (НС-33)'),
        (os.path.join(base_attachment_dir, 'Агрегат 3 НС33-сверка'), 'Агрегат 3 (НС-33)')
    ]

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Сопоставительная ведомость"

    ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE
    ws.page_setup.paperSize = ws.PAPERSIZE_A4

    font_title = Font(name="Calibri", size=14, bold=True)
    font_subtitle = Font(name="Calibri", size=11, italic=True)
    font_header = Font(name="Calibri", size=10, bold=True)
    font_agg = Font(name="Calibri", size=11, bold=True, color="1F497D")
    font_folder = Font(name="Calibri", size=10, bold=True, color="000000")
    font_regular = Font(name="Calibri", size=9)
    font_total = Font(name="Calibri", size=10, bold=True)

    fill_header = PatternFill(start_color="DCE6F1", end_color="DCE6F1", fill_type="solid")
    fill_agg = PatternFill(start_color="B8CCE4", end_color="B8CCE4", fill_type="solid")
    fill_folder = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
    fill_total = PatternFill(start_color="E9EDF4", end_color="E9EDF4", fill_type="solid")

    thin_border_side = Side(border_style="thin", color="CCCCCC")
    thin_border = Border(left=thin_border_side, right=thin_border_side, top=thin_border_side, bottom=thin_border_side)
    header_border = Border(left=thin_border_side, right=thin_border_side, top=Side(border_style="medium"), bottom=Side(border_style="medium"))

    align_center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    align_left = Alignment(horizontal="left", vertical="center", wrap_text=True)
    align_right = Alignment(horizontal="right", vertical="center")

    ws.merge_cells("A1:N1")
    ws["A1"] = "СПОСТАВИТЕЛЬНАЯ ВЕДОМОСТЬ ИЗМЕНЕНИЙ ПРОЕКТНЫХ И ИЗМЕНЕННЫХ МАТЕРИАЛОВ И ОБОРУДОВАНИЯ"
    ws["A1"].font = font_title
    ws["A1"].alignment = align_center

    ws.merge_cells("A2:N2")
    ws["A2"] = "по подтверждениям изменений между Проектной документацией (ПД) и рабочей документацией (РД)"
    ws["A2"].font = font_subtitle
    ws["A2"].alignment = align_center

    ws.row_dimensions[1].height = 25
    ws.row_dimensions[2].height = 18
    ws.row_dimensions[3].height = 10

    headers_r1 = [
        ("A4:A5", "№ п/п"),
        ("B4:B5", "Номер сметы / Раздел"),
        ("C4:G4", "Проектные решения (по ПД)"),
        ("H4:L4", "Измененные решения (по РД)"),
        ("M4:M5", "Отклонение стоимости (руб.)"),
        ("N4:N5", "Основание / Подтверждение изменения")
    ]

    for cell_range, text in headers_r1:
        if ":" in cell_range:
            ws.merge_cells(cell_range)
            top_left = cell_range.split(":")[0]
            ws[top_left] = text
        else:
            ws[cell_range] = text

    ws["C5"] = "Код / Обоснование"
    ws["D5"] = "Наименование материала / оборудования"
    ws["E5"] = "Ед. изм."
    ws["F5"] = "Кол-во"
    ws["G5"] = "Стоимость (руб.)"

    ws["H5"] = "Код / Обоснование"
    ws["I5"] = "Наименование материала / оборудования"
    ws["J5"] = "Ед. изм."
    ws["K5"] = "Кол-во"
    ws["L5"] = "Стоимость (руб.)"

    ws.row_dimensions[4].height = 20
    ws.row_dimensions[5].height = 25

    for r in range(4, 6):
        for c in range(1, 15):
            cell = ws.cell(r, c)
            cell.font = font_header
            cell.fill = fill_header
            cell.alignment = align_center
            cell.border = header_border

    row_idx = 6
    item_counter = 1

    grand_pd_total = 0.0
    grand_rd_total = 0.0

    for base, agg_label in base_dirs:
        if not os.path.exists(base):
            continue

        ws.merge_cells(start_row=row_idx, start_column=1, end_row=row_idx, end_column=14)
        cell = ws.cell(row_idx, 1)
        cell.value = f"ОБЪЕКТ: {agg_label.upper()}"
        cell.font = font_agg
        cell.fill = fill_agg
        cell.alignment = Alignment(horizontal="left", vertical="center", indent=1)
        for c in range(1, 15):
            ws.cell(row_idx, c).border = thin_border
        ws.row_dimensions[row_idx].height = 22
        row_idx += 1

        agg_pd_total = 0.0
        agg_rd_total = 0.0

        for sub in sorted(os.listdir(base)):
            sp = os.path.join(base, sub)
            if not os.path.isdir(sp): continue
            files = os.listdir(sp)
            pd_files = [f for f in files if 'ПД' in f or 'Раздел' in f]
            rd_files = [f for f in files if 'РД' in f or 'рд' in f or ('ПД' not in f and 'Раздел' not in f)]

            if not pd_files or not rd_files:
                continue

            pds = parse_pd_file(os.path.join(sp, pd_files[0]))
            rds = parse_rd_file(os.path.join(sp, rd_files[0]))
            matched = match_items(pds, rds)

            folder_justification = JUSTIFICATIONS.get(sub, f"Согласование изменений по локальной смете № {sub}")

            ws.merge_cells(start_row=row_idx, start_column=1, end_row=row_idx, end_column=14)
            f_cell = ws.cell(row_idx, 1)
            f_cell.value = f"Локальная смета № {sub} (ПД: {pd_files[0]} | РД: {rd_files[0]})"
            f_cell.font = font_folder
            f_cell.fill = fill_folder
            f_cell.alignment = Alignment(horizontal="left", vertical="center", indent=1)
            for c in range(1, 15):
                ws.cell(row_idx, c).border = thin_border
            ws.row_dimensions[row_idx].height = 20
            row_idx += 1

            folder_pd_total = 0.0
            folder_rd_total = 0.0

            for pair in matched:
                pd = pair['pd']
                rd = pair['rd']

                pd_total = pd['total'] if pd else 0.0
                rd_total = rd['total'] if rd else 0.0
                diff = rd_total - pd_total

                folder_pd_total += pd_total
                folder_rd_total += rd_total

                ws.cell(row_idx, 1, item_counter).alignment = align_center
                ws.cell(row_idx, 2, sub).alignment = align_center

                if pd:
                    ws.cell(row_idx, 3, pd['code']).alignment = align_left
                    ws.cell(row_idx, 4, pd['name']).alignment = align_left
                    ws.cell(row_idx, 5, pd['unit']).alignment = align_center
                    ws.cell(row_idx, 6, pd['qty']).alignment = align_right
                    ws.cell(row_idx, 7, pd['total']).alignment = align_right
                    ws.cell(row_idx, 6).number_format = "#,##0.00"
                    ws.cell(row_idx, 7).number_format = "#,##0.00"
                else:
                    ws.cell(row_idx, 3, "-").alignment = align_center
                    ws.cell(row_idx, 4, "[ Позиция отсутствует в ПД ]").alignment = align_left
                    ws.cell(row_idx, 5, "-").alignment = align_center
                    ws.cell(row_idx, 6, 0.0).alignment = align_right
                    ws.cell(row_idx, 7, 0.0).alignment = align_right
                    ws.cell(row_idx, 6).number_format = "#,##0.00"
                    ws.cell(row_idx, 7).number_format = "#,##0.00"

                if rd:
                    ws.cell(row_idx, 8, rd['code']).alignment = align_left
                    ws.cell(row_idx, 9, rd['name']).alignment = align_left
                    ws.cell(row_idx, 10, rd['unit']).alignment = align_center
                    ws.cell(row_idx, 11, rd['qty']).alignment = align_right
                    ws.cell(row_idx, 12, rd['total']).alignment = align_right
                    ws.cell(row_idx, 11).number_format = "#,##0.00"
                    ws.cell(row_idx, 12).number_format = "#,##0.00"
                else:
                    ws.cell(row_idx, 8, "-").alignment = align_center
                    ws.cell(row_idx, 9, "[ Позиция исключена в РД ]").alignment = align_left
                    ws.cell(row_idx, 10, "-").alignment = align_center
                    ws.cell(row_idx, 11, 0.0).alignment = align_right
                    ws.cell(row_idx, 12, 0.0).alignment = align_right
                    ws.cell(row_idx, 11).number_format = "#,##0.00"
                    ws.cell(row_idx, 12).number_format = "#,##0.00"

                diff_cell = ws.cell(row_idx, 13, diff)
                diff_cell.alignment = align_right
                diff_cell.number_format = "+#,##0.00;-#,##0.00;0.00"

                just_cell = ws.cell(row_idx, 14, folder_justification)
                just_cell.alignment = align_left

                for c in range(1, 15):
                    cell = ws.cell(row_idx, c)
                    cell.font = font_regular
                    cell.border = thin_border

                ws.row_dimensions[row_idx].height = 20
                row_idx += 1
                item_counter += 1

            ws.merge_cells(start_row=row_idx, start_column=1, end_row=row_idx, end_column=6)
            tot_lbl = ws.cell(row_idx, 1, f"Итого по смете № {sub}:")
            tot_lbl.font = font_total
            tot_lbl.alignment = Alignment(horizontal="right", vertical="center")

            ws.cell(row_idx, 7, folder_pd_total).number_format = "#,##0.00"
            ws.cell(row_idx, 7).font = font_total
            ws.cell(row_idx, 7).alignment = align_right

            ws.merge_cells(start_row=row_idx, start_column=8, end_row=row_idx, end_column=11)
            ws.cell(row_idx, 12, folder_rd_total).number_format = "#,##0.00"
            ws.cell(row_idx, 12).font = font_total
            ws.cell(row_idx, 12).alignment = align_right

            folder_diff = folder_rd_total - folder_pd_total
            ws.cell(row_idx, 13, folder_diff).number_format = "+#,##0.00;-#,##0.00;0.00"
            ws.cell(row_idx, 13).font = font_total
            ws.cell(row_idx, 13).alignment = align_right

            ws.cell(row_idx, 14, "")

            for c in range(1, 15):
                cell = ws.cell(row_idx, c)
                cell.fill = fill_total
                cell.border = thin_border

            ws.row_dimensions[row_idx].height = 20
            row_idx += 1

            agg_pd_total += folder_pd_total
            agg_rd_total += folder_rd_total

        ws.merge_cells(start_row=row_idx, start_column=1, end_row=row_idx, end_column=6)
        agg_lbl = ws.cell(row_idx, 1, f"ВСЕГО ПО ОБЪЕКТУ {agg_label.upper()}:")
        agg_lbl.font = font_agg
        agg_lbl.alignment = Alignment(horizontal="right", vertical="center")

        ws.cell(row_idx, 7, agg_pd_total).number_format = "#,##0.00"
        ws.cell(row_idx, 7).font = font_agg
        ws.cell(row_idx, 7).alignment = align_right

        ws.merge_cells(start_row=row_idx, start_column=8, end_row=row_idx, end_column=11)
        ws.cell(row_idx, 12, agg_rd_total).number_format = "#,##0.00"
        ws.cell(row_idx, 12).font = font_agg
        ws.cell(row_idx, 12).alignment = align_right

        agg_diff = agg_rd_total - agg_pd_total
        ws.cell(row_idx, 13, agg_diff).number_format = "+#,##0.00;-#,##0.00;0.00"
        ws.cell(row_idx, 13).font = font_agg
        ws.cell(row_idx, 13).alignment = align_right

        for c in range(1, 15):
            cell = ws.cell(row_idx, c)
            cell.fill = fill_agg
            cell.border = thin_border

        ws.row_dimensions[row_idx].height = 22
        row_idx += 1

        grand_pd_total += agg_pd_total
        grand_rd_total += agg_rd_total

    ws.merge_cells(start_row=row_idx, start_column=1, end_row=row_idx, end_column=6)
    g_lbl = ws.cell(row_idx, 1, "ИТОГО ВСЕГО ПО ВСЕМ ОБЪЕКТАМ:")
    g_lbl.font = font_title
    g_lbl.alignment = Alignment(horizontal="right", vertical="center")

    ws.cell(row_idx, 7, grand_pd_total).number_format = "#,##0.00"
    ws.cell(row_idx, 7).font = font_title
    ws.cell(row_idx, 7).alignment = align_right

    ws.merge_cells(start_row=row_idx, start_column=8, end_row=row_idx, end_column=11)
    ws.cell(row_idx, 12, grand_rd_total).number_format = "#,##0.00"
    ws.cell(row_idx, 12).font = font_title
    ws.cell(row_idx, 12).alignment = align_right

    grand_diff = grand_rd_total - grand_pd_total
    ws.cell(row_idx, 13, grand_diff).number_format = "+#,##0.00;-#,##0.00;0.00"
    ws.cell(row_idx, 13).font = font_title
    ws.cell(row_idx, 13).alignment = align_right

    for c in range(1, 15):
        cell = ws.cell(row_idx, c)
        cell.fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
        cell.border = Border(top=Side(border_style="double"), bottom=Side(border_style="double"), left=thin_border_side, right=thin_border_side)

    ws.row_dimensions[row_idx].height = 26

    col_widths = {
        'A': 8,   # № п/п
        'B': 14,  # Смета
        'C': 18,  # Код ПД
        'D': 35,  # Имя ПД
        'E': 10,  # Ед ПД
        'F': 12,  # Кол ПД
        'G': 16,  # Стоим ПД
        'H': 18,  # Код РД
        'I': 35,  # Имя РД
        'J': 10,  # Ед РД
        'K': 12,  # Кол РД
        'L': 16,  # Стоим РД
        'M': 18,  # Отклонение
        'N': 40   # Обоснование
    }

    for col_letter, width in col_widths.items():
        ws.column_dimensions[col_letter].width = width

    wb.save(output_filename)
    print(f"Successfully generated '{output_filename}' with {item_counter - 1} line items.")
    print(f"Grand PD Total: {grand_pd_total:,.2f} руб.")
    print(f"Grand RD Total: {grand_rd_total:,.2f} руб.")
    print(f"Grand Variance: {grand_diff:,.2f} руб.")

if __name__ == "__main__":
    main()
