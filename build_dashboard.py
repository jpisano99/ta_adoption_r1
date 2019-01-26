from settings import app
from open_wb import open_wb
from push_list_to_xls import push_list_to_xls
from get_linked_sheet_update import get_linked_sheet_update
from build_sheet_map import build_sheet_map
from sheet_map import sheet_map, sheet_keys
import time
import xlsxwriter

# from push_xls_to_ss import push_xls_to_ss


def create_customer_dict(order_rows):
    # Now we build a Customer Summary/Detail
    # Let's organize as this
    # order_dict: {cust_name:[[order1],[order2],[orderN]]}
    order_dict = {}
    orders = []
    order = []
    x = 0
    for order_row in order_rows:
        customer = order_row[0]

        # Is this in the order dict ?
        if customer in order_dict:
            orders = []
            for order in order_dict[customer]:
                orders.append(order)

            orders.append(order_row)
            order_dict[customer] = orders
        else:
            orders = []
            orders.append(order_row)
            order_dict[customer] = orders
    return order_dict


if __name__ == "__main__":
    # Open the order summary
    wb_orders, sheet_orders = open_wb('TA Order Summary_as_of_01_26_2019.xlsx')
    print(wb_orders, sheet_orders)

    #
    # Build a dict of Customer Orders
    # order_dict: {cust_name:[[order1],[order2],[orderN]]}
    #
    new_rows = []

    # Loop over the orders
    # Create a simple list from orders workbook
    for i in range(sheet_orders.nrows):
        if i == 0:
            continue
        my_row = sheet_orders.row(i)
        new_row = []
        for cell in my_row:
            new_row.append(cell.value)

        new_rows.append(new_row)

    # Create a dict of customer orders
    customer_order_dict = create_customer_dict(new_rows)
    print('We have: ', len(customer_order_dict), ' customers')
    print('with ', len(new_rows), ' skus')

    sheet_map = build_sheet_map(app['SS_CX'], sheet_map, 'SS_CX')
    sheet_map = build_sheet_map(app['SS_AS'], sheet_map, 'SS_AS')
    sheet_map = build_sheet_map(app['SS_SAAS'], sheet_map, 'SS_SAAS')

    print('got sheet maps')
    #
    # Get dict updates from linked sheets CX/AS/SAAS
    #
    cx_dict = get_linked_sheet_update(sheet_map, 'SS_CX', sheet_keys)
    as_dict = get_linked_sheet_update(sheet_map, 'SS_AS', sheet_keys)
    saas_dict = get_linked_sheet_update(sheet_map, 'SS_SAAS', sheet_keys)

    print()
    print('CX Dict ', cx_dict)
    print('AS Dict ', as_dict)
    print('SAAS Dict ', saas_dict)
    print()

    # Create Platform dict for lookup
    platform_dict = {'TA-CL-G1-39-K9': '39RU', 'TA-CL-G1-SFF8-K9': '8RU',
                     'C1-TA-V-SW-K9': 'Sftw Only', 'C1-TAAS-WP-FND-K9': 'SAAS',
                     'E2C1-TAAS-WPFND': 'SAAS'}

    #
    # Create top row for the dashboard
    #
    new_rows = []
    new_row = []
    bookings_col_num = -1
    sensor_col_num = -1
    svc_bookings_col_num = -1
    platform_type_col_num = -1
    sku_col_num = -1
    my_col_idx = {}

    for col_idx, col in enumerate(sheet_map):
        new_row.append(col[0])
        my_col_idx[col[0]] = col_idx # Make a dict of {column names : column number}

    print(my_col_idx)
    new_rows.append(new_row)
    print(new_rows)
    print(my_col_idx)
    print()

    #
    # Main loop
    #
    for customer, orders in customer_order_dict.items():
        print('Customer:', customer, 'has ', len(orders), ' orders')

        new_row = []

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

            if order[sku_col_num] in platform_dict:
                platform_type = platform_dict[order[sku_col_num]]

        #
        # Modify this record as need and add to the new_rows
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

        new_rows.append(order)

    #
    # Write the Dashboard to an Excel File
    #
    push_list_to_xls(new_rows, 'test_dash')
    exit()
