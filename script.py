import os
import xml.etree.ElementTree as ET
from openpyxl import Workbook


script_dir = os.path.dirname(os.path.abspath(__file__))
xml_file = input("Inserisci il nome del file XML (deve essere nella stessa cartella): ")
xml_path = os.path.join(script_dir, xml_file)


excel_output = os.path.splitext(xml_file)[0] + ".xlsx"
excel_path = os.path.join(script_dir, excel_output)


tree = ET.parse(xml_path)
root = tree.getroot()


wb = Workbook()
wb.remove(wb.active)


for second_level in root:
    sheet_name = second_level.tag[:31]

    counter = 1
    base_name = sheet_name
    while sheet_name in wb.sheetnames:
        sheet_name = f"{base_name}_{counter}"[:31]
        counter += 1

    ws = wb.create_sheet(title=sheet_name)

    column_names = set()
    for record in second_level.findall("record"):
        for column in record.findall("column"):
            if "name" in column.attrib:
                column_names.add(column.attrib["name"])

    column_names = sorted(list(column_names))

    for col_idx, col_name in enumerate(column_names, 1):
        ws.cell(row=1, column=col_idx, value=col_name)

    row_idx = 2
    for record in second_level.findall("record"):
        record_data = {}
        for column in record.findall("column"):
            if "name" in column.attrib:
                col_name = column.attrib["name"]
                col_value = column.text if column.text else ""
                record_data[col_name] = col_value

        for col_idx, col_name in enumerate(column_names, 1):
            if col_name in record_data:
                ws.cell(row=row_idx, column=col_idx, value=record_data[col_name])

        row_idx += 1


wb.save(excel_path)
print(f"File Excel creato con successo: {excel_output}")
