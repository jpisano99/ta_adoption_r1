from settings import app
import xlrd
import time


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
    orders_file = 'TA Order Summary_as_of_01_05_2019.xlsx'
    dashboard_file = app['SS_DASHBOARD']

    path_to_files = home +'\\' + working_dir  + '\\'
    path_to_orders = path_to_files + orders_file
    path_to_dashboard = path_to_files + dashboard_file

    wb_orders = xlrd.open_workbook(path_to_orders)
    sheet_orders = wb_orders.sheet_by_index(0)

    #
    # Main loop of bookings excel data
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


    print(len(new_rows))
    customer_order_dict = test_it(new_rows)
    print(len(customer_order_dict))

    for customer, orders in customer_order_dict.items():
        print (customer)
        # print('\t', orders)
        print ('\t\t\t', 'has ', len(orders),' orders')

        for i, order in enumerate(orders):
            print (i+1, '  ', order)
            time.sleep (2)
        print('-----------------------------------------')