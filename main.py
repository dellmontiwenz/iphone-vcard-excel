# Reading an excel file using Python
import xlrd
from datetime import date
from random import randint

name = input("Enter a filename: ")

# Give the location of the file
loc = ("./excel/" + name + ".xlsx")
loc1 = ("./excel/" + name + ".xls")

# To acquire today's date
today = date.today()
# To create the vcf file.
file1 = open(("./excel/" + name + ".vcf"), "w")

# To generate random whole numbers
def random_with_N_digits(n):
    range_start = 10**(n-1)
    range_end = (10**n)-1
    return randint(range_start, range_end)
# Generate vcf file
def generateVcfFile(_x_name_target, _y_name_target, _y_tel_target):
    # save all location coordinates to a buffer
    x_row_target = _x_name_target + 1
    y_name_loc = _y_name_target
    y_tel_loc = _y_tel_target
    for x_new_target in range (sheet.nrows):
        if (x_new_target >= x_row_target):
            name = str(sheet.cell(x_new_target, y_name_loc)).rstrip("'")[6:]
            m_num = str(sheet.cell(x_new_target, y_tel_loc)).rstrip('0').rstrip('.')[7:]
            print(name + ": " + m_num)
            file1.write("BEGIN:VCARD \n")
            file1.write("VERSION:3.0 \n")
            file1.write("PRODID:-//Apple Inc.//iOS 12.3.2//EN \n")
            file1.write("N:" + name + ";;; \n")
            file1.write("FN:" + name + "\n")
            file1.write("TEL;type=HOME;type=VOICE;type=pref:" + m_num + "\n")
            file1.write("REV:" + str(today) + "T15:05:" + str(random_with_N_digits(3)) + "\n")
            file1.write("END:VCARD \n")

# To open Workbook
try:
    wb = xlrd.open_workbook(loc)
except IOError:
    wb = xlrd.open_workbook(loc1)
sheet = wb.sheet_by_index(0)

# The content of the vcf file
# Coordinates of the cell target
x_name_target = 0
y_name_target = 0
y_tel_target = 0
x = 0
y = 0
while (y < sheet.ncols):
    for x in range (sheet.nrows):
        # Find Name value in cells
        if (sheet.cell_value(x, y) == "Name"):
            # save name location coordinates to a buffer
            x_name_target = x
            y_name_target = y

        if ((sheet.cell_value(x, y) == "Phone") or (sheet.cell_value(x, y) == "phone")):
            # save tel y coordinate to a buffer
            y_tel_target = y
    # Continue searching if not found
    y = y + 1
    x = 0
# Call convert to vcf function
generateVcfFile(x_name_target, y_name_target, y_tel_target)
file1.close()
# Conversion process done.
print("\n")
print('Conversion Done!!!!')
print('created by: Wenz Montifalcon')


