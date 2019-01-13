import time
import datetime
import xlrd
from open_wb import open_wb
from settings import app


def build_renewals_dict(wb, sheet):
    # Return a dict (my_dict) with bookings file info
    my_list = ['End Customer', 'Renewal Date', 'Monthly Charge']
    my_dict = {}

    # Loop across all column headings in the renewals file and
    # Find the column number that matches the col_name in my_list
    for renewals_col_num in range(sheet.ncols):
        # print(sheet.cell_value(0, renewals_col_num))

        # Loop across my_list to find a match from the bookings file
        for idx, val in enumerate(my_list):
            col_name = val
            if col_name == sheet.cell_value(0, renewals_col_num):
                # print('match',col_name, renewals_col_num)
                my_list[idx] = (col_name,renewals_col_num)

    # [('End Customer', 0), ('Renewal Date', 4), ('Monthly Charge', 9)]

    # Now Loop through all the rows and build my_dict
    for renewals_row_num in range(sheet.nrows):

        if renewals_row_num == 0:
            continue

        customer = sheet.cell_value(renewals_row_num, my_list[0][1])
        renewal_date = sheet.cell_value(renewals_row_num, my_list[1][1])

        renewal_date = datetime.datetime(*xlrd.xldate_as_tuple(renewal_date, wb.datemode))
        renewal_date = renewal_date.strftime('%m-%d-%Y')

        my_dict[customer] = [renewal_date, sheet.cell_value(renewals_row_num, my_list[2][1])]

    return my_dict


if __name__ == "__main__":
    wb_renewals, sheet_renewals = open_wb(app['XLS_RENEWALS'])
    renewals_dict = build_renewals_dict(wb_renewals, sheet_renewals)
    print(renewals_dict)
