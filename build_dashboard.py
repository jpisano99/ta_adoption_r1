from settings import app
import xlrd
import time
from get_cx_update import get_cx_update
import xlsxwriter
from dashboard_xls import dashboard_xls

def test_it(order_rows):
    # Now we build a Customer Summary/Detail
    # Let's organize as this
    # order_dict: {cust_name:[[order1],[order2],[orderN]]}
    order_dict = {}
    orders = []
    order = []
    x = 0
    for order_row in order_rows:
        customer = order_row[0]
        if x==0:
            x += 1
            continue

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
            order_dict[customer]=orders
    return order_dict

if __name__ == "__main__":
    #
    # Get settings for file locations and names
    #
    home = app['HOME']
    working_dir = app['WORKING_DIR']
    # orders_file = app['XLS_ORDER_SUMMARY'] # Master bookings has 9958 rows as of 12-15-18
    orders_file = 'TA Order Summary_as_of_01_09_2019.xlsx'
    dashboard_file = app['SS_DASHBOARD']

    path_to_files = home +'\\' + working_dir  + '\\'
    path_to_orders = path_to_files + orders_file
    path_to_dashboard = path_to_files + dashboard_file

    wb_orders = xlrd.open_workbook(path_to_orders)
    sheet_orders = wb_orders.sheet_by_index(0)


    #exit()

    #
    # Build a dict of Customer Orders
    # order_dict: {cust_name:[[order1],[order2],[orderN]]}
    #

    new_rows = []
    for i in range(sheet_orders.nrows):
        if i == 0:
            continue
        my_row = sheet_orders.row(i)
        new_row = []
        for cell in my_row:
            #print(cell.value)
            new_row.append(cell.value)

        new_rows.append(new_row)

    customer_order_dict = test_it(new_rows)
    print('We have: ', len(customer_order_dict),' customers')
    print('with ', len(new_rows),' skus')
    print()

    #
    # Get CX update
    #
    cx_dict = get_cx_update()
    # print ('CX Dict ', cx_dict)

    #
    # Create top row for the dashboard
    #
    new_rows = []
    new_row = []
    for x,col in enumerate(dashboard_xls):
        col[1] = x
        #print(x, col)
        new_row.append(col[0])

    new_rows.append(new_row)
    print(new_row)

    #
    # Main loop
    #

    for customer, orders in customer_order_dict.items():
        # print (customer,'\t\t', 'has ', len(orders),' orders')
        sensor_count = 0
        service_bookings = 0
        bookings_total = 0
        new_row = []
        cx_contact = ''
        cx_status = ''

        # Look up the CX update
        if customer in cx_dict:
            cx_contact = cx_dict[customer][0]
            cx_status = cx_dict[customer][1]
            # cx_update = 'FOUND CX Update: ' + str(cx_dict[customer])
        else:
            cx_contact = 'None assigned'
            cx_status = ''
            # cx_update = 'NO CX Update FOUND'

        # Loop over this customers orders and
        # Create one summary row for this customer
        for i, order in enumerate(orders):
            bookings_total = bookings_total + order[11]

            if order[13] == 'Software':  # Sensor Count column
                sensor_count = order[16] + sensor_count
            elif order[13] == 'Service':  # Service Count column
                service_bookings = order[11] + service_bookings

            # print (i+1, '  ', order)
            # time.sleep (.25)

        # Build the new row for this customer
        for x, col in enumerate(dashboard_xls):
            if col[0] == 'Sensor Count':
                new_row.append(sensor_count)
                continue
            elif col[0] == 'Total Bookings':
                new_row.append(bookings_total)
                continue
            elif col[0] == 'CX Contact':
                new_row.append(cx_contact)
                continue
            elif col[0] == 'CX Status':
                new_row.append(cx_status)
                continue
            elif col[0] == 'Service Bookings':
                new_row.append(service_bookings)
                continue

            new_row.append(order[x])


        # print('\t CX Status', cx_update)
        # print('\t Sensors', sensor_count)
        # print('\t Services', service_count)
        # print('\t Total Bookings', bookings_total)
        new_rows.append(new_row)

        #print (new_rows)

        #print('-----------------------------------------')

        #
        # Write the Dashboard to an Excel File

        #
        wb_file = path_to_files + 'jim' + '.xlsx'
        workbook = xlsxwriter.Workbook(wb_file)
        worksheet = workbook.add_worksheet()
        for this_row, my_val in enumerate(new_rows):
            worksheet.write_row(this_row, 0, my_val)
        workbook.close()
