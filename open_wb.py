import xlrd
from settings import app


def open_wb(excel_file):
    #
    # Get settings for file locations and names
    #
    home = app['HOME']
    working_dir = app['WORKING_DIR']
    path_to_files = home + '\\' + working_dir + '\\'
    path_to_file = path_to_files + excel_file

    #
    # Open up excel workbook
    #
    my_wb = xlrd.open_workbook(path_to_file)
    my_sheet = my_wb.sheet_by_index(0)

    return my_wb, my_sheet


if __name__ == "__main__":
    my_excel = open_wb(app['BOOKINGS'])
    print('We have: ', my_excel[0], my_excel[1])
