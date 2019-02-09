import datetime
import xlrd
from settings import app
from sheet_map import sheet_map
from build_sheet_map import build_sheet_map
import copy
import time


def build_renewals_dict(wb, sheet):
    # Return a dict (my_dict) with bookings file info
    my_map = copy.deepcopy(sheet_map) # We need to create a UNIQUE copy of sheet_map
    my_dict = {}
    my_map = build_sheet_map(app['XLS_RENEWALS'], my_map, 'XLS_RENEWALS')

    # Strip out all other un-needed tags from the sheet_map
    # tmp_list = []
    # for idx, x in enumerate(my_map):
    #     if x[1] == 'XLS_RENEWALS':
    #         tmp_list.append(my_map[idx])
    # # my_map = tmp_list
    # print (tmp_list)

    # List comprehension replacement for above
    my_map = [x for x in my_map if x[1] == 'XLS_RENEWALS']

    # Loop over all of the renewal records
    # Build a dict of {customer:[next renewal date, next renewal revenue, upcoming renewals]}


    for row_num in range(1, sheet.nrows):
        customer = sheet.cell_value(row_num, 0)
        if customer in my_dict:
            tmp_record = []
            tmp_records = my_dict[customer]
        else:
            tmp_record = []
            tmp_records = []

        # Loop over the sheet map
        for col_map in my_map:
            my_cell = sheet.cell_value(row_num, col_map[2])

            # Is this cell a Date type (3) ?
            # If so format as a M/D/Y
            if sheet.cell_type(row_num, col_map[2]) == 3:
                my_cell = datetime.datetime(*xlrd.xldate_as_tuple(my_cell, wb.datemode))
                my_cell = my_cell.strftime('%m-%d-%Y')

            tmp_record.append(my_cell)

        tmp_records.append(tmp_record)
        my_dict[customer] = tmp_records

    # for customer, renewals in my_dict.items():
    #     print (customer,renewals)
    #     time.sleep(2)
    # #
    # Here we compress and summarize the dict
    #
    for customer, renewals in my_dict.items():
        renewal_dates = {}

        # Sort this customer renewal dates
        renewals.sort(key=lambda x: x[0])
        # if customer == 'FIRST NATIONAL BANK OF SOUTHERN AFRICA LTD':
        #     print(customer)
        #     print()
        #     print(len(renewals))
        #     print(renewals)
        #     # renewals.sort(key=lambda x: x[0])
        #     print(len(renewals))
        #     print(renewals)
        #     # exit()

        for renewal in renewals:
            renewal_date = renewal[0]
            renewal_revenue = renewal[1]

            if renewal_date in renewal_dates:
                total_revenue = renewal_dates[renewal_date] + renewal_revenue
                renewal_dates[renewal_date] = total_revenue
            else:
                renewal_dates[renewal_date] = renewal_revenue

        #
        # Convert renewals_dates from a dict to a list and
        # Update the my_dict for this customer
        tmp_list = []
        for tmp_date, tmp_revenue in renewal_dates.items():
            tmp_list.append([tmp_date, tmp_revenue])
        my_dict[customer] = tmp_list

        # if customer == 'FIRST NATIONAL BANK OF SOUTHERN AFRICA LTD':
        #     print(my_dict[customer])
        #     exit()
    return my_dict


if __name__ == "__main__":
    from open_wb import open_wb
    wb_renewals, sheet_renewals = open_wb(app['XLS_RENEWALS'])
    renewals_dict = build_renewals_dict(wb_renewals, sheet_renewals)
    print(renewals_dict)
