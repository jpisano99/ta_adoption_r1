import xlsxwriter
import datetime
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

    xls_money = workbook.add_format({'num_format': '$#,##0'})
    xls_date = workbook.add_format({'num_format': 'mm / dd/ yyyy'})

    for row_num, my_row in enumerate(my_list):
        for col_num, cell_val in enumerate(my_row):
            if type(cell_val) is float:
                worksheet.write(row_num, col_num, cell_val, xls_money)
            elif isinstance(cell_val, datetime.datetime):
                worksheet.write(row_num, col_num, cell_val, xls_date)
            else:
                worksheet.write(row_num, col_num, cell_val)

    workbook.close()



    # for this_row, my_val in enumerate(my_list):
    #     worksheet.write_row(this_row, 0, my_val)
    # workbook.close()

    return
