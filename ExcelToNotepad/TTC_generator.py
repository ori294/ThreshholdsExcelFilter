# thresholds to filer, can be changed from here instead being changed from the code
BUFFER_SIZE = 2048
REQTHRESH = 1000000
IMPSTHRESH = 500
REVTHRESH = 5


def main():
    files = os.listdir("./")  # get the files in the dir
    for csv in files:
        if "csv" in csv and "ToClose" not in csv:
            csv_name_read = str(csv)
            csv_name_write = "new" + str(csv).replace(".csv", ".xlsx")
            # convert the csv file, if exists into an excel
            read_file = pd.read_csv(csv_name_read)
            read_file.to_excel(csv_name_write, index=None, header=True, engine='xlsxwriter')

    files = os.listdir("./")  # update the dir after the csv to excel conversion

    for xlsx in files:
        # for each xlsx file in the working directory, do:
        if "xlsx" in xlsx and "ToClose" not in xlsx:
            file_directory = {}
            one_row_list = []
            loc = xlsx
            # To open Workbook
            wb = xlrd.open_workbook(loc)
            sheet = wb.sheet_by_index(0)

            reqs_col_num = 0
            imps_col_num = 0
            rev_col_num = 0

            # find col numbers for each threshold parameter
            for col in range(sheet.ncols):
                if "request" in str(sheet.cell_value(0, col)).lower() and\
                        "fill" not in str(sheet.cell_value(0, col)).lower():
                    reqs_col_num = col
                elif "impression" in str(sheet.cell_value(0, col)).lower():
                    imps_col_num = col
                elif "revenue" in str(sheet.cell_value(0, col)).lower() and\
                        "channel" not in str(sheet.cell_value(0, col)).lower():
                    rev_col_num = col

            # save the first row
            for j in range(sheet.ncols):
                one_row_list.append(str(sheet.cell_value(0, j)))
            file_directory["title"] = one_row_list
            one_row_list = []

            # save the rows that match the thresholds
            for i in range(sheet.nrows):
                if sheet.cell_value(i, reqs_col_num) >= REQTHRESH:
                    if sheet.cell_value(i, imps_col_num) <= IMPSTHRESH:
                        if float(str(sheet.cell_value(i, rev_col_num)).replace("$", "")) <= REVTHRESH:
                            for j in range(sheet.ncols):
                                one_row_list.append(str(sheet.cell_value(i, j)))
                            file_directory[sheet.cell_value(i, 0)] = one_row_list
                            one_row_list = []

            # open a new excel workbook
            nwb = Workbook()
            ws1 = nwb.create_sheet("ToClose")

            # add the first row
            row_list = file_directory.get("title")
            file_directory.pop("title")
            ws1.append(row_list)

            # add the records that were saved beforehand
            for tag in file_directory:
                row_list = file_directory.get(tag)
                ws1.append(row_list)
            nwb.save('ToClose' + str(xlsx))

    # update the dir after the new output files were created.
    files = os.listdir("./")
    moveto = "output/"
    # move the output file into a specific directory
    for outputFile in files:
        if "ToClose" in outputFile:
            shutil.move(outputFile, moveto + outputFile)

        # delete the old csv files
        elif "ToClose" not in outputFile and "csv" in outputFile:
            os.remove(outputFile)
    print "Done"


if __name__ == "__main__":
    from openpyxl import Workbook
    import shutil
    import pandas as pd
    import xlrd
    import os