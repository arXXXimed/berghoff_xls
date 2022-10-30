from io import BytesIO
import xlsxwriter
import base64
import json

with open("Datatemplate.json") as f:
    template = json.load(f)

workbook = xlsxwriter.Workbook('template.xlsx')
worksheet = workbook.add_worksheet()

temp_column = template['columns']
for col, lenght_col in enumerate(temp_column):
    worksheet.set_column(col,col,lenght_col)

temp_rows = template['rows']
for row, height in enumerate(temp_rows):
    worksheet.set_row(row, height)

dataimage = template["data_images"]    
images = template["images"]  
for image in images:
    data_image = dataimage[image['data']]
    dataEncoded = base64.b64decode(data_image)
    image_data = BytesIO(dataEncoded)
    options_image = image['options']
    options_image['image_data'] = image_data
    worksheet.insert_image(image['anchor'], image['data'], options_image)

workbook.close()