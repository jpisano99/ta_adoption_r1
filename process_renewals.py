import copy
import time
from datetime import datetime
import datetime
import xlrd
from settings import app
from sheet_map import sheet_map
from open_wb import open_wb
from build_sheet_map import build_sheet_map


def process_renewals():
    #
    # Open up the renewals excel workbooks
    wb, sheet = open_wb(app['XLS_RENEWALS'])

    # Get the renewal columns we are looking for
    my_map = build_sheet_map(app['XLS_RENEWALS'], sheet_map, 'XLS_RENEWALS')

    # List comprehension replacement for above
    # Strip out the columns from the sheet map that we don't need
    my_map = [x for x in my_map if x[1] == 'XLS_RENEWALS']

    # Create a simple column name dict
    col_nums = {sheet.cell_value(0, col_num): col_num for col_num in range(0, sheet.ncols)}

    # Loop over all of the renewal records
    # Build a dict of {customer:[next renewal date, next renewal revenue, upcoming renewals]}
    my_dict = {}
    for row_num in range(1, sheet.nrows):
        customer = sheet.cell_value(row_num, col_nums['End Customer'])
        if customer in my_dict:
            tmp_record = []
            tmp_records = my_dict[customer]
        else:
            tmp_record = []
            tmp_records = []

        # Loop over the my map gather the columns we need
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
    print(len(my_dict))

    # Sort each customers renewal dates
    sorted_dict = {}
    summarized_dict = {}
    summarized_rec = []

    for customer, renewals in my_dict.items():
        # Sort this customers renewal records by date order
        renewals.sort(key=lambda x: x[0])
        sorted_dict[customer] = renewals

        next_renewal_date = renewals[0][0]
        tmp_revenue = 0
        for renewal_rec in renewals:
            if renewal_rec[0] == next_renewal_date:
                # Tally this renewal record and get the next
                tmp_revenue = renewal_rec[1] + tmp_revenue
            elif renewal_rec[0] != next_renewal_date:
                # Record these summarized values
                summarized_rec.append([next_renewal_date, tmp_revenue])
                # Reset these values and get the next renewal date for this customer
                next_renewal_date = renewal_rec[0]
                tmp_revenue = renewal_rec[1]

            # Check to see if this is the last renewal record
            # If so exit the loop
            if renewals.index(renewal_rec) == len(renewals)-1:
                break

        summarized_rec.append([next_renewal_date, tmp_revenue])
        summarized_dict[customer] = summarized_rec
        summarized_rec = []

        # if customer == 'FIRST NATIONAL BANK OF SOUTHERN AFRICA LTD':
        #     print(customer, renewals)
        #     renewals.sort(key=lambda x: x[0])
        #     print(customer, renewals)
        # else:
        #     renewals.sort(key=lambda x: x[0])

    print(sorted_dict['FIRST NATIONAL BANK OF SOUTHERN AFRICA LTD'])
    print(summarized_dict['FIRST NATIONAL BANK OF SOUTHERN AFRICA LTD'])
    return


def build_renewals_dict(wb, sheet):
    # Return a dict (my_dict) with bookings file info
    my_map = copy.deepcopy(sheet_map) # We need to create a UNIQUE copy of sheet_map
    my_dict = {}
    my_map = build_sheet_map(app['XLS_RENEWALS'], my_map, 'XLS_RENEWALS')

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

    for customer, renewals in my_dict.items():
        print (customer,renewals)
        time.sleep(2)
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
    # cust_dict = {}
    # cust_list = ['FIRST NATIONAL BANK OF SOUTHERN AFRICA LTD', 'FIRST NATIONAL BANK OF SOUTHERN AFRICA']
    # cust_dict['FIRST NATIONAL BANK OF SOUTHERN AFRICA LTD'] = ['FIRST NATIONAL BANK OF SOUTHERN AFRICA LTD', 'FIRSTRAND LTD', 'Hinson, Richard', 'Pomelli, Luciano', '', '', '', '', '', '', '', '', '', '', '', 0.0, '', 'Product', 'C1-TETRATION', '39RU - Appliance', 0.0, '', 'EMEAR-REGION', 'EMEAR_MEA', 'SUB_SAHARAN_AFRICA', 'SA_ENT_OP', 'SA_ENT', 'SA_ENT_FRG'], ['FIRST NATIONAL BANK OF SOUTHERN AFRICA LTD', 'FIRSTRAND LTD', 'Hinson, Richard', 'Pomelli, Luciano', '', '', '', '', '', '', '', '', '', '', '', 67938.0, '', 'Service', 'ASF-DCV1-TA-QS-M', 'AS Fixed - Medium Service', 0.0, '', 'EMEAR-REGION', 'EMEAR_MEA', 'SUB_SAHARAN_AFRICA', 'SA_ENT_OP', 'SA_ENT', 'SA_ENT_FRG'], ['FIRST NATIONAL BANK OF SOUTHERN AFRICA LTD', 'FIRSTRAND LTD', 'Hinson, Richard', 'Pomelli, Luciano', '', '', '', '', '', '', '', '', '', '', '', 143253.0, '', 'Software', 'C1-TA-BASE-1K-K9', 'TA - Software Subscription 1000 Licenses', 1000.0, '', 'EMEAR-REGION', 'EMEAR_MEA', 'SUB_SAHARAN_AFRICA', 'SA_ENT_OP', 'SA_ENT', 'SA_ENT_FRG'], ['FIRST NATIONAL BANK OF SOUTHERN AFRICA LTD', 'FIRSTRAND LTD', 'Hinson, Richard', 'Pomelli, Luciano', '', '', '', '', '', '', '', '', '', '', '', 648004.0, '', 'Product', 'TA-CL-G1-39-K9', '39RU - Appliance', 0.0, '', 'EMEAR-REGION', 'EMEAR_MEA', 'SUB_SAHARAN_AFRICA', 'SA_ENT_OP', 'SA_ENT', 'SA_ENT_FRG'], ['FIRST NATIONAL BANK OF SOUTHERN AFRICA LTD', 'FIRSTRAND LTD', 'Hinson, Richard', 'Pomelli, Luciano', '', '', '', '', '', '', '', '', '', '', '', 77.0, '', 'Product', 'TA-CL-G1-39-K9', '39RU - Appliance', 0.0, '', 'EMEAR-REGION', 'EMEAR_MEA', 'SUB_SAHARAN_AFRICA', 'SA_ENT_OP', 'SA_ENT', 'SA_ENT_FRG'], ['FIRST NATIONAL BANK OF SOUTHERN AFRICA LTD', 'FIRSTRAND LTD', 'Hinson, Richard', 'Pomelli, Luciano', '', '', '', '', '', '', '', '', '', '', '', 266400.0, '', 'Software', 'C1-TA-BASE-1K-K9', 'TA - Software Subscription 1000 Licenses', 1000.0, '', 'EMEAR-REGION', 'EMEAR_MEA', 'SUB_SAHARAN_AFRICA', 'SA_ENT_OP', 'SA_ENT', 'SA_ENT_FRG'], ['FIRST NATIONAL BANK OF SOUTHERN AFRICA LTD', 'FIRSTRAND LTD', 'Hinson, Richard', 'Pomelli, Luciano', '', '', '', '', '', '', '', '', '', '', '', -66600.0, '', 'Software', 'C1-TA-BASE-1K-K9', 'TA - Software Subscription 1000 Licenses', 1000.0, '', 'EMEAR-REGION', 'EMEAR_MEA', 'SUB_SAHARAN_AFRICA', 'SA_ENT_OP', 'SA_ENT', 'SA_ENT_FRG'], ['FIRST NATIONAL BANK OF SOUTHERN AFRICA LTD', 'FIRSTRAND LTD', 'Hinson, Richard', 'Pomelli, Luciano', '', '', '', '', '', '', '', '', '', '', '', 0.0, '', 'Software', 'C1-TA-BASE-1K-K9', 'TA - Software Subscription 1000 Licenses', 1000.0, '', 'EMEAR-REGION', 'EMEAR_MEA', 'SUB_SAHARAN_AFRICA', 'SA_ENT_OP', 'SA_ENT', 'SA_ENT_FRG'], ['FIRST NATIONAL BANK OF SOUTHERN AFRICA LTD', 'FIRSTRAND LTD', 'Hinson, Richard', 'Pomelli, Luciano', '', '', '', '', '', '', '', '', '', '', '', 143262.0, '', 'Software', 'C1-TA-BASE-1K-K9', 'TA - Software Subscription 1000 Licenses', 1000.0, '', 'EMEAR-REGION', 'EMEAR_MEA', 'SUB_SAHARAN_AFRICA', 'SA_ENT_OP', 'SA_ENT', 'SA_ENT_FRG'], ['FIRST NATIONAL BANK OF SOUTHERN AFRICA LTD', 'FIRSTRAND LTD', 'Hinson, Richard', 'Pomelli, Luciano', '', '', '', '', '', '', '', '', '', '', '', 0.0, '', 'Product', 'C1-TETRATION', '39RU - Appliance', 0.0, '', 'EMEAR-REGION', 'EMEAR_MEA', 'SUB_SAHARAN_AFRICA', 'SA_ENT_OP', 'SA_ENT', 'SA_ENT_FRG'], ['FIRST NATIONAL BANK OF SOUTHERN AFRICA LTD', 'FIRSTRAND LTD', 'Hinson, Richard', 'Pomelli, Luciano', '', '', '', '', '', '', '', '', '', '', '', 0.0, '', 'Software', 'C1-TA-ENF100-K9', 'TA - Software Subscription 100 Enforcement Licenses', 100.0, '', 'EMEAR-REGION', 'EMEAR_MEA', 'SUB_SAHARAN_AFRICA', 'SA_ENT_OP', 'SA_ENT', 'SA_ENT_FRG'], ['FIRST NATIONAL BANK OF SOUTHERN AFRICA LTD', 'FIRSTRAND LTD', 'Hinson, Richard', 'Pomelli, Luciano', '', '', '', '', '', '', '', '', '', '', '', 103356.0, '', 'Software', 'C1-TA-BASE-1K-K9', 'TA - Software Subscription 1000 Licenses', 1000.0, '', 'EMEAR-REGION', 'EMEAR_MEA', 'SUB_SAHARAN_AFRICA', 'SA_ENT_OP', 'SA_ENT', 'SA_ENT_FRG'], ['FIRST NATIONAL BANK OF SOUTHERN AFRICA LTD', 'FIRSTRAND LTD', 'Hinson, Richard', 'Pomelli, Luciano', '', '', '', '', '', '', '', '', '', '', '', -103356.0, '', 'Software', 'C1-TA-BASE-1K-K9', 'TA - Software Subscription 1000 Licenses', 1000.0, '', 'EMEAR-REGION', 'EMEAR_MEA', 'SUB_SAHARAN_AFRICA', 'SA_ENT_OP', 'SA_ENT', 'SA_ENT_FRG'], ['FIRST NATIONAL BANK OF SOUTHERN AFRICA LTD', 'FIRSTRAND LTD', 'Hinson, Richard', 'Pomelli, Luciano', '', '', '', '', '', '', '', '', '', '', '', 0.0, '', 'Service', 'ASF-DCV1-TA-QS-M', 'AS Fixed - Medium Service', 0.0, '', 'EMEAR-REGION', 'EMEAR_MEA', 'SUB_SAHARAN_AFRICA', 'SA_ENT_OP', 'SA_ENT', 'SA_ENT_FRG'], ['FIRST NATIONAL BANK OF SOUTHERN AFRICA LTD', 'FIRSTRAND LTD', 'Hinson, Richard', 'Pomelli, Luciano', '', '', '', '', '', '', '', '', '', '', '', 0.0, '', 'Software', 'C1-TA-ENF100-K9', 'TA - Software Subscription 100 Enforcement Licenses', 100.0, '', 'EMEAR-REGION', 'EMEAR_MEA', 'SUB_SAHARAN_AFRICA', 'SA_ENT_OP', 'SA_ENT', 'SA_ENT_FRG'], ['FIRST NATIONAL BANK OF SOUTHERN AFRICA LTD', 'FIRSTRAND LTD', 'Hinson, Richard', 'Pomelli, Luciano', '', '', '', '', '', '', '', '', '', '', '', 0.0, '', 'Software', 'C1-TA-ENF100-K9', 'TA - Software Subscription 100 Enforcement Licenses', 100.0, '', 'EMEAR-REGION', 'EMEAR_MEA', 'SUB_SAHARAN_AFRICA', 'SA_ENT_OP', 'SA_ENT', 'SA_ENT_FRG'], ['FIRST NATIONAL BANK OF SOUTHERN AFRICA LTD', 'FIRSTRAND LTD', 'Hinson, Richard', 'Pomelli, Luciano', '', '', '', '', '', '', '', '', '', '', '', 0.0, '', 'Software', 'C1-TA-ENF100-K9', 'TA - Software Subscription 100 Enforcement Licenses', 100.0, '', 'EMEAR-REGION', 'EMEAR_MEA', 'SUB_SAHARAN_AFRICA', 'SA_ENT_OP', 'SA_ENT', 'SA_ENT_FRG'], ['FIRST NATIONAL BANK OF SOUTHERN AFRICA LTD', 'FIRSTRAND LTD', 'Hinson, Richard', 'Pomelli, Luciano', '', '', '', '', '', '', '', '', '', '', '', 0.0, '', 'Software', 'C1-TA-BASE-1K-K9', 'TA - Software Subscription 1000 Licenses', 1000.0, '', 'EMEAR-REGION', 'EMEAR_UNALLOCATED', 'MEA_UNALLOCATED_L3', 'UNALLOCATED-STH-AFRICA', 'UNALLOCATED-STH-AFRICA-MISCL5', 'UNALLOCATED-STH-AFRICA-MISCL6'], ['FIRST NATIONAL BANK OF SOUTHERN AFRICA LTD', 'FIRSTRAND LTD', 'Hinson, Richard', 'Pomelli, Luciano', '', '', '', '', '', '', '', '', '', '', '', 0.0, '', 'Software', 'C1-TA-BASE-1K-K9', 'TA - Software Subscription 1000 Licenses', 1000.0, '', 'EMEAR-REGION', 'EMEAR_UNALLOCATED', 'MEA_UNALLOCATED_L3', 'UNALLOCATED-STH-AFRICA', 'UNALLOCATED-STH-AFRICA-MISCL5', 'UNALLOCATED-STH-AFRICA-MISCL6'], ['FIRST NATIONAL BANK OF SOUTHERN AFRICA LTD', 'FIRSTRAND LTD', 'Hinson, Richard', 'Pomelli, Luciano', '', '', '', '', '', '', '', '', '', '', '', 15159.0, '', 'Software', 'C1-TA-BASE-1K-K9', 'TA - Software Subscription 1000 Licenses', 1000.0, '', 'EMEAR-REGION', 'EMEAR_MEA', 'SUB_SAHARAN_AFRICA', 'SA_ENT_OP', 'SA_ENT', 'SA_ENT_FRG'], ['FIRST NATIONAL BANK OF SOUTHERN AFRICA LTD', 'FIRSTRAND LTD', 'Hinson, Richard', 'Pomelli, Luciano', '', '', '', '', '', '', '', '', '', '', '', 29303.0, '', 'Software', 'C1-TA-BASE-1K-K9', 'TA - Software Subscription 1000 Licenses', 1000.0, '', 'EMEAR-REGION', 'EMEAR_MEA', 'SUB_SAHARAN_AFRICA', 'SA_ENT_OP', 'SA_ENT', 'SA_ENT_FRG'], ['FIRST NATIONAL BANK OF SOUTHERN AFRICA LTD', 'FIRSTRAND LTD', 'Hinson, Richard', 'Pomelli, Luciano', '', '', '', '', '', '', '', '', '', '', '', 88197.0, '', 'Software', 'C1-TA-BASE-1K-K9', 'TA - Software Subscription 1000 Licenses', 1000.0, '', 'EMEAR-REGION', 'EMEAR_MEA', 'SUB_SAHARAN_AFRICA', 'SA_ENT_OP', 'SA_ENT', 'SA_ENT_FRG'], ['FIRST NATIONAL BANK OF SOUTHERN AFRICA LTD', 'FIRSTRAND LTD', 'Hinson, Richard', 'Pomelli, Luciano', '', '', '', '', '', '', '', '', '', '', '', 170491.0, '', 'Software', 'C1-TA-BASE-1K-K9', 'TA - Software Subscription 1000 Licenses', 1000.0, '', 'EMEAR-REGION', 'EMEAR_MEA', 'SUB_SAHARAN_AFRICA', 'SA_ENT_OP', 'SA_ENT', 'SA_ENT_FRG']
    # process_renewals(cust_list, cust_dict)
    process_renewals()
    exit()
