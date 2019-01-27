from settings import app
from open_wb import open_wb
from push_list_to_xls import push_list_to_xls
from create_customer_order_dict import create_customer_order_dict
from get_linked_sheet_update import get_linked_sheet_update
from build_sheet_map import build_sheet_map
from sheet_map import sheet_map, sheet_keys
# from push_xls_to_ss import push_xls_to_ss


if __name__ == "__main__":
    #
    # Open the order summary
    #
    wb_orders, sheet_orders = open_wb('tmp_TA Scrubbed Orders_as_of_01_27_2019.xlsx')

    # Loop over the orders XLS worksheet
    # Create a simple list of orders with NO headers
    order_list = []
    for i in range(1, sheet_orders.nrows):  # Skip the header row start at 1
        order_list.append(sheet_orders.row_values(i))

    # Create a dict of customer orders
    customer_order_dict = create_customer_order_dict(order_list)
    print()
    print('We have: ', len(customer_order_dict), ' customers')
    print('with ', len(order_list), ' skus')

    # Build Sheet Maps
    sheet_map = build_sheet_map(app['SS_CX'], sheet_map, 'SS_CX')
    sheet_map = build_sheet_map(app['SS_AS'], sheet_map, 'SS_AS')
    sheet_map = build_sheet_map(app['SS_SAAS'], sheet_map, 'SS_SAAS')

    #
    # Get dict updates from linked sheets CX/AS/SAAS
    #
    cx_dict = get_linked_sheet_update(sheet_map, 'SS_CX', sheet_keys)
    as_dict = get_linked_sheet_update(sheet_map, 'SS_AS', sheet_keys)
    saas_dict = get_linked_sheet_update(sheet_map, 'SS_SAAS', sheet_keys)

    print()
    print('We have CX Updates: ', len(cx_dict))
    print('We have AS Updates: ', len(as_dict))
    print('We have SAAS Updates: ', len(saas_dict))
    print()

    # Create Platform dict for lookup
    platform_dict = {'TA-CL-G1-39-K9': '39RU', 'TA-CL-G1-SFF8-K9': '8RU',
                     'C1-TA-V-SW-K9': 'Sftw Only', 'C1-TAAS-WP-FND-K9': 'SAAS',
                     'E2C1-TAAS-WPFND': 'SAAS'}

    #
    # Init Main Loop Variables
    #
    new_rows = []
    new_row = []
    bookings_col_num = -1
    sensor_col_num = -1
    svc_bookings_col_num = -1
    platform_type_col_num = -1
    sku_col_num = -1
    my_col_idx = {}

    # Create top row for the dashboard
    # also make a dict (my_col_idx) of {column names : column number}
    for col_idx, col in enumerate(sheet_map):
        new_row.append(col[0])
        my_col_idx[col[0]] = col_idx
    new_rows.append(new_row)

    #
    # Main loop
    #
    for customer, orders in customer_order_dict.items():
        new_row = []
        orders_found = len(orders)

        # Default Values
        bookings_total = 0
        sensor_count = 0
        service_bookings = 0
        platform_type = 'Not Identified'

        saas_status = 'No Status'
        cx_contact = 'None assigned'
        cx_status = 'No Update'
        as_pm = ''
        as_cse = ''
        as_complete = ''
        as_comments = ''

        #
        # Get update from linked sheets (if any)
        #
        if customer in saas_dict:
            saas_status = saas_dict[customer][0]
        else:
            saas_status = 'No Status'

        if customer in cx_dict:
            cx_contact = cx_dict[customer][0]
            cx_status = cx_dict[customer][1]
        else:
            cx_contact = 'None assigned'
            cx_status = 'No Update'

        if customer in as_dict:
            if as_dict[customer][0] == '':
                as_pm = 'None Assigned'
            else:
                as_pm = as_dict[customer][0]

            if as_dict[customer][1] == '':
                as_cse = 'None Assigned'
            else:
                as_cse = as_dict[customer][1]

            if as_dict[customer][2] == '':
                as_complete = 'No Update'
            else:
                as_complete = as_dict[customer][2]

            if as_dict[customer][3] == '':
                as_comments = 'No Comments'
            else:
                as_comments = as_dict[customer][3]

        #
        # Loop over this customers orders
        # Create one summary row for this customer
        # Total things
        # Build a list of things that may change order to order (ie Renewal Dates, Customer Names)
        #
        for order_idx, order in enumerate(orders):
            # calculate totals in this loop (ie total_books, sensor count etc)
            bookings_total = bookings_total + order[my_col_idx['Total Bookings']]
            sensor_count = sensor_count + order[my_col_idx['Sensor Count']]

            if order[my_col_idx['Product Type']] == 'Service':
                service_bookings = service_bookings + order[my_col_idx['Total Bookings']]

            if order[my_col_idx['Bundle Product ID']] in platform_dict:
                platform_type = platform_dict[order[my_col_idx['Bundle Product ID']]]
            else:
                platform_type = 'Unknown'

        #
        # Modify/Update this record as needed and then add to the new_rows
        #
        order[my_col_idx['Total Bookings']] = bookings_total
        order[my_col_idx['Sensor Count']] = sensor_count
        order[my_col_idx['Service Bookings']] = service_bookings

        order[my_col_idx['CuSM Name']] = cx_contact
        order[my_col_idx['Next Action']] = cx_status

        order[my_col_idx['AS PM']] = as_pm
        order[my_col_idx['AS CSE']] = as_cse
        order[my_col_idx['Project Status/PM Completion']] = as_complete
        order[my_col_idx['Delivery Comments']] = as_comments

        order[my_col_idx['Provisioning completed']] = saas_status

        order[my_col_idx['Product Description']] = platform_type

        order[my_col_idx['Orders Found']] = orders_found

        new_rows.append(order)
    #
    # End of main loop
    #

    # Do some clean up and ready for output
    #
    # Rename the columns as per the sheet map
    cols_to_delete = []
    for idx, map_info in enumerate(sheet_map):
        if map_info[3] != '':
            if map_info[3] == '*DELETE*':
                # Put the columns to delete in a list
                cols_to_delete.append(idx)
            else:
                # Rename to the new column name
                new_rows[0][idx] = map_info[3]

    # Loop over the new_rows and
    # delete columns we don't need as per the sheet_map
    for col_idx in sorted(cols_to_delete, reverse=True):
        for row_idx, my_row in enumerate(new_rows):
            del new_rows[row_idx][col_idx]

    #
    # Write the Dashboard to an Excel File
    #
    push_list_to_xls(new_rows, app['XLS_DASHBOARD'])
    exit()
