import sys
import xlrd

file_name = str(sys.argv[1])
BUFFER_SIZE = 2048
files_dictionary = {}
print file_name

if file_name:
    loc = file_name
    # To open Workbook
    wb = xlrd.open_workbook(loc)
    sheet = wb.sheet_by_index(0)

    for i in range(sheet.nrows):
        tag_name = str(sheet.cell_value(i, 0))
        bundle_name = str(sheet.cell_value(i, 1)).replace(".0", "")
        bundle_list = files_dictionary.get(tag_name)

        if tag_name in files_dictionary:
            bundle_list.append(bundle_name)
        else:
            files_dictionary[tag_name] = []

for tag in files_dictionary:
    tag_name_to_file = str(tag).replace("/", "")
    txt_file = open(tag_name_to_file.replace("/", "") + ".txt", 'w')
    bundle_list = files_dictionary.get(tag)
    for bundle in bundle_list:
        txt_file.write(bundle + "\n")
    txt_file.close()



