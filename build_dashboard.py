from settings import app
import xlrd


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
for i in range(sheet_orders.nrows):
    customer_name = sheet_orders.cell_value(i, 0)
    customer_name_1 = sheet_orders.cell_value(i, 1)
    pss = sheet_orders.cell_value(i, 2)
    print (customer_name,customer_name_1,pss)
