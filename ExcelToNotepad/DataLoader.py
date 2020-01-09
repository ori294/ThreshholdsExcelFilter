
BUFFER_SIZE = 2048


def main():
    files = os.listdir("./")

    for file_name in files:
        if "csv" in file_name:
            csv_name_read = str(file_name)
            print csv_name_read
            csv_name_write = "new" + str(file_name).replace(".csv", ".xlsx")
            print csv_name_write
            read_file = pd.read_csv(csv_name_read)
            read_file.to_excel(csv_name_write, index=None, header=True, engine='xlsxwriter')
            os.remove(csv_name_read)

    files_dictionary = {}
    files = os.listdir("./")
    for file_name in files:
        if "xlsx" in file_name:
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


if __name__ == "__main__":
    import xlrd
    import sys
    import os
    import pandas as pd
