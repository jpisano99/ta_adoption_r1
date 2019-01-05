from settings import *
import xlsxwriter
import smartsheet
from build_coverage_dict import *
from build_sku_dict import *
from build_bookings_dict import *
from build_renewals_dict import *
from cleanup_orders import *
from find_team import *
from push_xls_to_ss import *


#
# Get settings for file locations and names
#
home = app['HOME']
working_dir = app['WORKING_DIR']
bookings_file = app['BOOKINGS'] # Master bookings has 9958 rows as of 12-15-18
renewals_file = app['RENEWALS']
path_to_files = home +'\\' + working_dir  + '\\'
path_to_renewals = path_to_files + renewals_file
path_to_bookings = path_to_files + bookings_file


# Go to Smartsheets and build these two dicts to use reference lookups
# team_dict: {'sales_levels 1-6':[('PSS','TSA')]}
# sku_dict: {sku : [sku_type, sku_description]}
team_dict = build_coverage_dict()
sku_dict = build_sku_dict()

#
# Open up the renewals and bookings excel workbooks
#
wb_renewals = xlrd.open_workbook(path_to_renewals)
sheet_renewals = wb_renewals.sheet_by_index(0)

wb_bookings = xlrd.open_workbook(path_to_bookings)
sheet_bookings = wb_bookings.sheet_by_index(0)

# From the renewals file get renewal dates for lookup
# {erp_customer_name:[renewal_date,monthly_charge]}
renewals_dict = build_renewals_dict(wb_renewals,sheet_renewals)

# From the current up to date bookings file build a simple dict
# that describes the format of the output file we are creating
# and the columns we need to add (ie PSS, TSA, Renewal Dates)
bookings_dict = build_bookings_dict(path_to_bookings,sheet_bookings)

#
# init a bunch a variable we need for the main loop
#
customer_list = []
order_top_row = []
order_rows = []
order_row = []
sku_col_num = -1
col_pss_num = -1
col_tsa_num = -1

# Build the column titles top row
# Also grab
# 1. sku column number
# 2. PSS and TSA column numbers
for idx, val in enumerate(bookings_dict['col_info']):
    col_name = val[0]
    col_num = val[1]
    order_top_row.append(val[0])
    if col_name == 'Bundle Product ID':
        sku_col_num = col_num
    if col_name == 'PSS':
        col_pss_num = idx
    if col_name == 'TSA':
        col_tsa_num = idx

order_row = order_top_row
order_rows.append(order_row)

#
# Main loop of bookings excel data
#
for i in range(sheet_bookings.nrows):

    # Is this SKU of interest ?
    sku = sheet_bookings.cell_value(i,sku_col_num)

    if sku in sku_dict :
        # Let's make a row for this order
        # Since it has an "interesting" sku
        order_row = []
        sales_level = ''
        sales_level_cntr = 0
        sku_desc = sku_dict[sku][1]

        # Walk across the bookings_dict columns
        # to build this output row cell by cell
        for val in bookings_dict['col_info']:
            col_name = val[0]
            col_idx = val[1]

            # Capture both of the Customer names
            if col_name == 'ERP End Customer Name':
                customer_name_erp = sheet_bookings.cell_value(i, col_idx)
            if col_name == 'End Customer Global Ultimate Name':
                customer_name_end = sheet_bookings.cell_value(i, col_idx)

            # If this is a 'Sales Level X' column then
            # Capture it's value for lookup
            if col_name[:-2] == 'Sales Level':
                sales_level = sales_level + sheet_bookings.cell_value(i, col_idx) +','
                sales_level_cntr += 1
                if sales_level_cntr == 6:
                    # We have collected all 6 sales levels
                    # Now go to find_team to do the lookup
                    sales_level = sales_level[:-1]
                    sales_team = find_team(team_dict,sales_level)
                    pss = sales_team[0]
                    tsa = sales_team[1]
                    order_row[col_pss_num] = pss
                    order_row[col_tsa_num] = tsa

            if col_idx != -1:
                # OK we have a cell we need so grab it
                order_row.append(sheet_bookings.cell_value(i, col_idx))
            elif col_name == 'Product Description':
                # Add in the Product Description
                order_row.append(sku_desc)
            elif col_name == 'Renewal Date(s)':
                # Add in the Renewal Date if there is one
                # Else just add a blank string
                if customer_name_erp in renewals_dict:
                    renewal_date = renewals_dict[customer_name_erp]
                    order_row.append(renewal_date[0])
                else:
                    order_row.append('')
            else:
                # this cell is assigned a -1 in the bookings_dict
                # so assign a blank as a placeholder for now
                order_row.append('')

        # Done with all the columns in this row
        # Log this row for BOTH customer names and orders
        # Go to next row of the raw bookings data
        customer_list.append((customer_name_erp, customer_name_end))
        order_rows.append(order_row)

#
# End
#

# OK we now have a full list (order_rows) of just the SKUs we are interested in
# As determined by the sku_dict

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


# Create a simple customer_list list of tuples
# Contains a full set of unique sorted customer names
# customer_list = [(erp_customer_name,end_customer_ultimate), (CustA,CustA)]
customer_list = set(customer_list)

# Convert the SET to a LIST so we can sort it
customer_list = list(customer_list)

# Sort the LIST
customer_list.sort(key=lambda tup: tup[0])
print('We have: ', len(customer_list), ' customers')


# Clean up order_dict to remove:
# 1.  +/- zero sum orders
# 2. zero revenue orders
order_dict = cleanup_orders(customer_list,order_dict,order_top_row)

#
# Create a summary order file out of the order_dict
#
summary_order_rows=[]
summary_order_rows.append(order_top_row)
for key,val in order_dict.items():
    for my_row in val:
        summary_order_rows.append(my_row)

#
# Write the Summary order Excel File
# Includes only those "Interesting" SKU's that are non-zero sum
#
wb_file = path_to_files + app['XLS_ORDER_SUMMARY'] + app['AS_OF_DATE'] + '.xlsx'
workbook = xlsxwriter.Workbook(wb_file)
worksheet = workbook.add_worksheet()
for this_row, my_val in enumerate(summary_order_rows):
    worksheet.write_row(this_row, 0, my_val)
workbook.close()
push_xls_to_ss(wb_file, app['XLS_ORDER_SUMMARY'])

#
# Write the Detailed order Excel File
# Includes ALL "Interesting" SKU's
#
wb_file = path_to_files + app['XLS_ORDER_DETAIL'] + app['AS_OF_DATE']  + '.xlsx'
workbook = xlsxwriter.Workbook(wb_file)
worksheet = workbook.add_worksheet()
for this_row, my_val in enumerate(order_rows):
    worksheet.write_row(this_row, 0, my_val)
workbook.close()
push_xls_to_ss(wb_file, app['XLS_ORDER_DETAIL'])

#
# Write TA Customer List to a local excel workbook
#
# Insert a header row before writing
top_row = ['erp_customer_name','end_customer_ultimate_name']
customer_list.insert(0,top_row)

wb_file = path_to_files + app['XLS_CUSTOMER'] + app['AS_OF_DATE']  + '.xlsx'
workbook = xlsxwriter.Workbook(wb_file)
worksheet = workbook.add_worksheet()
for this_row, my_val in enumerate(customer_list):
    worksheet.write(this_row, 0, my_val[0])
    worksheet.write(this_row, 1, my_val[1])
workbook.close()
push_xls_to_ss(wb_file, app['XLS_CUSTOMER'])

exit()