import xlsxwriter
from settings import app


def push_list_to_xls(my_list, xls_file):
    #
    # Get settings for file locations and names
    #
    home = app['HOME']
    working_dir = app['WORKING_DIR']
    path_to_files = home + '\\' + working_dir + '\\'
    wb_file = path_to_files + xls_file + app['AS_OF_DATE'] + '.xlsx'

    #
    # Write the Excel File
    #
    workbook = xlsxwriter.Workbook(wb_file)
    worksheet = workbook.add_worksheet()

    for this_row, my_val in enumerate(my_list):
        worksheet.write_row(this_row, 0, my_val)
    workbook.close()

    return
