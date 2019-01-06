from settings import app
import xlrd
import time
from get_cx_update import get_cx_update

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
    orders_file = 'TA Order Summary_as_of_01_06_2019.xlsx'
    dashboard_file = app['SS_DASHBOARD']

    path_to_files = home +'\\' + working_dir  + '\\'
    path_to_orders = path_to_files + orders_file
    path_to_dashboard = path_to_files + dashboard_file

    wb_orders = xlrd.open_workbook(path_to_orders)
    sheet_orders = wb_orders.sheet_by_index(0)

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
    #
    #

    dashboard = [['ERP End Customer Name', -1],
                    ['End Customer Global Ultimate Name', -1],
                    ['PSS', -1],
                    ['TSA', -1],
                    ['AS PM', -1],
                    ['AS CSE', -1],
                    ['AS Status', -1],
                    ['SAAS Status', -1],
                    ['CX Contact', -1],
                    ['CX Status', -1],
                    ['Renewal Date(s)', -1],
                    ['Total Bookings', -1],
                    ['Service Orders', -1],
                    ['Sensor Count', -1],
                    ['Active Sensors', -1],
                    ['Sales Level 1', -1],
                    ['Sales Level 2', -1],
                    ['Sales Level 3', -1],
                    ['Sales Level 4', -1],
                    ['Sales Level 5', -1],
                    ['Sales Level 6', -1]]

    new_rows = []
    for customer, orders in customer_order_dict.items():
        print (customer,'\t\t', 'has ', len(orders),' orders')
        sensor_count = 0
        service_count = 0
        bookings_total = 0
        new_row = []
        cx_update = ''

        if customer in cx_dict:
            cx_update = 'FOUND CX Update: ' + str(cx_dict[customer])
        else:
            cx_update = 'NO CX Update FOUND'

        for i, order in enumerate(orders):
            bookings_total =  bookings_total + order[8]
            if order[9] == 'Software': # Sensor Count column
                sensor_count = order[12] + sensor_count
            elif order[9] == 'Service':  # Service Count column
                service_count = order[12] + service_count
            print (i+1, '  ', order)
            time.sleep (1)

        new_row.append(order[0])
        new_row.append(order[1])
        new_row.append(order[2])
        new_row.append(order[3])
        new_row.append(cx_update)
        new_row.append(sensor_count)
        new_row.append(bookings_total)



        print('\t CX Status', cx_update)
        print('\t Sensors', sensor_count)
        print('\t Services', service_count)
        print('\t Total Bookings', bookings_total)
        new_rows.append(new_row)

        print (new_rows)

        print('-----------------------------------------')