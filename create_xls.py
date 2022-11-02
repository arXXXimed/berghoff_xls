from io import BytesIO
from PIL import Image,ImageDraw, ImageFont
import xlsxwriter
from xlsxwriter.utility import xl_rowcol_to_cell
import base64
import json


def create_excel(data_1C):
    workbook = xlsxwriter.Workbook('template.xlsx')
    format_cell_itemGroup = workbook.add_format(
        {"align": "left", "bold": True,  'bg_color': "#BDD7EE",  "bottom": 1,   "top": 1,   "left": 1, 'indent': 1})
    format_cell_itemGroup_1 = workbook.add_format(
        {"align": "left", "bold": True, 'bg_color': "#BDD7EE", "bottom": 1,   "top": 1})
    format_cell_1 = workbook.add_format(
        {"align": "left",  "valign": "vcenter", "bold": False, "border": 1, "text_wrap": True, "font_size": 10})
    format_cell_2 = workbook.add_format(
        {"align": "center",  "valign": "vcenter", "bold": False, "border": 1, "text_wrap": True, "font_size": 10, 'bg_color': "#C2E59B"})
    format_cell_3 = workbook.add_format(
        {"align": "center",  "valign": "vcenter", "bold": False, "border": 1, "text_wrap": True, "font_size": 10, "num_format": "# ##0.00[$р.-ru-RU]"})

    worksheet = create_templ(workbook)
    worksheet.set_column('Q:T', 20, format_cell_2, {'hidden': 1})

    data_images = data_1C['arrayImages']

    list_data_rows = data_1C['arrayItems']
    for data_rows in list_data_rows:
        itemGroup = data_rows['itemGroup']
        last_row = worksheet.dim_rowmax

        # if last_row > 100:
        #     break     

        worksheet.write(last_row+1, 1, itemGroup, format_cell_itemGroup)
        for i in range(2, 10):
            worksheet.write(last_row+1, i, None, format_cell_itemGroup_1)
        worksheet.write(last_row+1, 11, None, format_cell_itemGroup_1)
        worksheet.write(last_row+1, 13, None, format_cell_itemGroup_1)
        worksheet.write(last_row+1, 15, None, format_cell_itemGroup_1)

        items = data_rows['items']

        
        for item in items:

            current_row = worksheet.dim_rowmax + 1

            '''
            =G15*F15 - L price
            =S15*F15 - q mass
            =T15*F15 - r volueme
            '''
            G15 = xl_rowcol_to_cell(current_row, 6)
            F15 = xl_rowcol_to_cell(current_row, 5)
            S15 = xl_rowcol_to_cell(current_row, 18)
            T15 = xl_rowcol_to_cell(current_row, 19)

            worksheet.set_row(current_row, 47)
            worksheet.write_string(
                current_row, 1, item['artikle'], format_cell_1)
            worksheet.write(current_row, 2, item['artikle'], format_cell_1)
            worksheet.write(current_row, 3, item['name'], format_cell_1)
            worksheet.write(
                current_row, 4, item['amountInPackage'], format_cell_1)
            worksheet.write(current_row, 5, None, format_cell_2)
            worksheet.write(current_row, 6, item['priceGross'], format_cell_3)
            worksheet.write(current_row, 7, round(
                item['priceGross']*0.95, 2), format_cell_3)
            worksheet.write(current_row, 8, round(
                item['priceGross']*0.90, 2), format_cell_3)
            worksheet.write(current_row, 9, round(
                item['priceGross']*0.85, 2), format_cell_3)
            worksheet.write(current_row, 11, f'={G15}*{F15}', format_cell_3)
            worksheet.write(current_row, 13,
                            item['priceReatal'], format_cell_3)
            worksheet.write(current_row, 15, item['remainder'], format_cell_2)
            worksheet.write(current_row, 16, f'={S15}*{F15}', format_cell_3)
            worksheet.write(current_row, 17, f'={T15}*{F15}', format_cell_3)
            worksheet.write(current_row, 18, item['weight'], format_cell_3)
            worksheet.write(current_row, 19, item['volume'], format_cell_3)



            image_item = data_images.get(item['artikle'])
            base64_of_image = image_item['base64']
            dataEncoded = base64.b64decode(base64_of_image)
            image_data = BytesIO(dataEncoded)
            im = Image.open(image_data)
            width, height = im.size
            newsize = (225, 200)
            im1 = im.resize(newsize)
            font = ImageFont.truetype("arial.ttf", size=7)
            idraw = ImageDraw.Draw(im1)
            idraw.text((10, 10), "BergHoff", font=font)
            image_content = BytesIO()
            im1.save(image_content, format='PNG')
            options_image = {
                "x_offset": 1,
                "y_offset": 1,
                "x_scale": 0.30,
                "y_scale": 0.30,
                "object_position": 2,
                "url": None,
                "description": None,
                "decorative": True,
                'image_data':image_content
            }

            worksheet.insert_image(
                current_row, 2, item['artikle'], options_image)

    workbook.close()


def create_templ(workbook):
    worksheet = workbook.add_worksheet()
    worksheet.hide_zero()
    worksheet.freeze_panes(11, 16)
    with open("Datatemplate.json", encoding='utf-8') as f:
        template = json.load(f)

    formats_templ = template['formats']
    formats_excel = {}
    for key, format_cell in formats_templ.items():
        currant_format = workbook.add_format(format_cell)
        formats_excel[key] = currant_format

    def get_format(format):
        if isinstance(format, str):
            cell_format = formats_excel.get(format, None)
        elif isinstance(format, dict):
            cell_format = workbook.add_format(format)
        elif format is None:
            cell_format = None
        return cell_format

    temp_merged = template['merged']
    for merg_cell in temp_merged:

        format_cell = formats_excel.get(
            merg_cell.get("format_cell", "default"), None)
        worksheet.merge_range(merg_cell["adress"], None, format_cell)

    temp_column = template['columns']
    for col, lenght_col in enumerate(temp_column):
        worksheet.set_column(col, col, lenght_col)

    temp_rows = template['rows']
    for row, height in enumerate(temp_rows):
        worksheet.set_row(row, height)

    rows_data = template['rows_data']
    for row_data in rows_data:
        format_cell = get_format(row_data.get("format_cell"))
        worksheet.write_row(row_data["adress"], row_data["data"], format_cell)

    columns_data = template['columns_data']
    for column_data in columns_data:
        format_cell = get_format(column_data.get("format_cell"))
        worksheet.write_column(
            column_data["adress"], column_data["data"], format_cell)

    ranges = template['ranges']
    for range_data in ranges:
        pass

    dataimage = template["data_images"]
    images = template["images"]
    for image in images:
        data_image = dataimage[image['data']]
        dataEncoded = base64.b64decode(data_image)
        image_data = BytesIO(dataEncoded)
        options_image = image['options']
        options_image['image_data'] = image_data
        worksheet.insert_image(image['anchor'], image['data'], options_image)

    cells = template["cells"]
    for cell in cells:
        adress = cell["adress"]
        data = cell["data"]
        format_cell = get_format(cell.get("format_cell"))

        if isinstance(adress, str):
            if ':' in adress:
                worksheet.merge_range(adress, data, format_cell)
            else:
                worksheet.write(adress, data, format_cell)

    worksheet.autofilter(10, 1, 10, 16)

    worksheet.hide_gridlines(2)
    return worksheet


with open('Data_1C\json.txt', encoding='utf-8') as f:
    data_1C = json.load(f)
create_excel(data_1C)
