import datetime
import argparse
from reverseCombiner import reverseCombiner
from openpyxl import load_workbook
from openpyxl.styles import Font

# Set up files
parser = argparse.ArgumentParser()
parser.add_argument("file", help="Treba zadať názov súboru na spracovanie", type=str)
args = parser.parse_args()

input_file = args.file
output_file = "new_" + input_file

# Open workbook
print(str(datetime.datetime.now()) + " " + f"Načítavam zošit: {input_file}")
workbook = load_workbook(filename=input_file)

# Get worksheets object, worksheet count for iterator
all_sheets = workbook.worksheets
sheet_count = len(all_sheets)

# Get reference sheet and length for delete algorithm
reference_sheet = workbook.worksheets[3]
reference_sheet_length = len(reference_sheet["A"])

# Get list of rows to delete
print(str(datetime.datetime.now()) + " " + "Vyhľadávam riadky na mazanie.")
rows_list = []

for row in range(reference_sheet_length, 3, -1):
    cell = "A" + str(row)

    if reference_sheet[cell].font.color.rgb != "FFFF0000":
        rows_list.append(row)

rows_tuple = reverseCombiner(rows_list)
print(
    str(datetime.datetime.now())
    + " "
    + f"Nasledovné rozsahy riadkov budú vymazané: {rows_tuple}"
)

# Loop and delete
for ws in range(3, sheet_count):
    for tuple in rows_tuple:
        workbook.worksheets[ws].delete_rows(tuple[0], tuple[1])

# Save modified
print(str(datetime.datetime.now()) + " " + "Ukladám upravený zošit.")
workbook.save(output_file)
