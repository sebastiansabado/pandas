# import xlsxwriter

# workbook = xlsxwriter.Workbook(r'C:\Users\sebas\Desktop\projects\pandas\Example2.xlsx')
# worksheet = workbook.add_worksheet()

# worksheet.write('A1', 'Hello')

import xlsxwriter
from datetime import datetime

date = datetime.now().strftime("%Y_%m_%d")

# Create an new Excel file and add a worksheet.
workbook = xlsxwriter.Workbook('returns_'+ date + '.xlsx')
worksheet = workbook.add_worksheet()

# Widen the first column to make the text clearer.
worksheet.set_column('A:A', 20)

# Add a bold format to use to highlight cells.
bold = workbook.add_format({'bold': True})

# Write some simple text.
worksheet.write('A1', 'Hello')

# Text with formatting.
worksheet.write('A2', 'World', bold)

# Write some numbers, with row/column notation.
worksheet.write(2, 0, 123)
worksheet.write(3, 0, 123.456)

# Insert an image.
# worksheet.insert_image('B5', 'logo.png')

workbook.close()