import datetime
import xlrd
from open_wb import open_wb
from build_sheet_map import build_sheet_map
from sheet_map import sheet_map
from settings import app


def build_renewals_dict(wb, sheet):
    # Return a dict (my_dict) with bookings file info
    my_dict = {}
    my_map = build_sheet_map(app['XLS_RENEWALS'], sheet_map, 'XLS_RENEWALS')

    # Strip out all other un-needed tags from the sheet_map
    tmp_list = []
    for idx, x in enumerate(my_map):
        if x[1] == 'XLS_RENEWALS':
            tmp_list.append(my_map[idx])
    my_map = tmp_list

    # Loop over all of the renewal records
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

    #
    # Here we compress and summarize the dict
    #
    for customer, renewals in my_dict.items():
        renewal_dates = {}

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

    return my_dict


if __name__ == "__main__":
    wb_renewals, sheet_renewals = open_wb(app['XLS_RENEWALS'])
    renewals_dict = build_renewals_dict(wb_renewals, sheet_renewals)
    print(renewals_dict)
