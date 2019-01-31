from settings import app
from build_customer_list import build_customer_list
from open_wb import open_wb
from sheet_map import sheet_map
from build_coverage_dict import build_coverage_dict
from build_sku_dict import build_sku_dict
from build_sheet_map import build_sheet_map
from build_renewals_dict import build_renewals_dict
from cleanup_orders import cleanup_orders
from find_team import find_team
from push_list_to_xls import push_list_to_xls
from push_xls_to_ss import push_xls_to_ss


# Go to Smartsheets and build these two dicts to use reference lookups
# team_dict: {'sales_levels 1-6':[('PSS','TSA')]}
# sku_dict: {sku : [sku_type, sku_description]}
team_dict = build_coverage_dict()
sku_dict = build_sku_dict()

#
# Open up the renewals and bookings excel workbooks
#
wb_renewals, sheet_renewals = open_wb(app['XLS_RENEWALS'])
wb_bookings, sheet_bookings = open_wb(app['XLS_BOOKINGS'])

# From the renewals file get renewal dates for lookup
# {erp_customer_name:[renewal_date,monthly_charge]}
renewals_dict = build_renewals_dict(wb_renewals, sheet_renewals)

# From the current up to date bookings file build a simple dict
# that describes the format of the output file we are creating
# and the columns we need to add (ie PSS, TSA, Renewal Dates)

sheet_map = build_sheet_map(app['XLS_BOOKINGS'], sheet_map, 'XLS_BOOKINGS')
# sheet_map = build_sheet_map(app['XLS_RENEWALS'], sheet_map, 'XLS_RENEWALS')
sheet_map = build_sheet_map(app['SS_COVERAGE'], sheet_map, 'SS_COVERAGE')

#
# init a bunch a variable we need for the main loop
#
customer_list = []
order_top_row = []
order_rows = []
order_row = []
customer_col_num = -1
sku_col_num = -1
pss_col_num = -1
tsa_col_num = -1

# Build the column titles top row
# Also grab
# 1. sku column number
# 2. PSS and TSA column numbers
for idx, val in enumerate(sheet_map):
    src_col_name = val[0]  # Source Sheet Column Name
    src_col_num = val[2]  # Source sheet column number
    order_top_row.append(src_col_name)

    if src_col_name == 'ERP End Customer Name':
        customer_col_num = src_col_num
    elif src_col_name == 'Bundle Product ID':
        sku_col_num = src_col_num
    elif src_col_name == 'pss':
        pss_col_num = idx
    elif src_col_name == 'tsa':
        tsa_col_num = idx
    elif src_col_name == 'Product Type':
        prod_type_col_num = idx
    elif src_col_name == 'Sensor Count':
        sensor_cnt_col_num = idx

order_rows.append(order_top_row)
print(order_rows)

#
# Main loop of bookings excel data
#
for i in range(sheet_bookings.nrows):

    # Is this SKU of interest ?
    sku = sheet_bookings.cell_value(i, sku_col_num)

    if sku in sku_dict:
        # Let's make a row for this order
        # Since it has an "interesting" sku
        customer = sheet_bookings.cell_value(i, customer_col_num)
        order_row = []
        sales_level = ''
        sales_level_cntr = 0

        # Grab SKU data from the SKU dict
        sku_type = sku_dict[sku][0]
        sku_desc = sku_dict[sku][1]
        sku_sensor_cnt = sku_dict[sku][2]

        # Walk across the bookings_dict columns
        # to build this output row cell by cell
        for val in sheet_map:
            col_name = val[0]  # Source Sheet Column Name
            col_idx = val[2]   # Source Sheet Column Number

            # If this is a 'Sales Level X' column then
            # Capture it's value until we get to level 6
            # then do a team lookup
            if col_name[:-2] == 'Sales Level':
                sales_level = sales_level + sheet_bookings.cell_value(i, col_idx) + ','
                sales_level_cntr += 1
                if sales_level_cntr == 6:
                    # We have collected all 6 sales levels
                    # Now go to find_team to do the lookup
                    sales_level = sales_level[:-1]
                    sales_team = find_team(team_dict, sales_level)
                    pss = sales_team[0]
                    tsa = sales_team[1]
                    order_row[pss_col_num] = pss
                    order_row[tsa_col_num] = tsa

            if col_idx != -1:
                # OK we have a cell we need so grab it
                order_row.append(sheet_bookings.cell_value(i, col_idx))
            elif col_name == 'Product Description':
                # Add in the Product Description
                order_row.append(sku_desc)
            elif col_name == 'Product Type':
                # Add in the Product Type
                order_row.append(sku_type)
            elif col_name == 'Sensor Count':
                # Add in the Sensor Count
                order_row.append(sku_sensor_cnt)
            elif col_name == 'Product Bookings':
                order_row.append('renew_revenue')
            elif col_name == 'Renewal Date':
                order_row.append('renew_date')
                # Add in the Renewal Date if there is one
                # Else just add a blank string
                # if customer in renewals_dict:
                #     renewal_date = renewals_dict[customer]
                #     order_row.append(renewal_date[0])
                # else:
                #     order_row.append('')
            else:
                # this cell is assigned a -1 in the bookings_dict
                # so assign a blank as a placeholder for now
                order_row.append('')

        # Done with all the columns in this row
        # Log this row for BOTH customer names and orders
        # Go to next row of the raw bookings data
        order_rows.append(order_row)

print('len ', len(order_rows))
#
# End of main loop
#
# OK we now have a full list (order_rows) of just the SKUs we are interested in
# As determined by the sku_dict


# Now we build a an order dict
# Let's organize as this
# order_dict: {cust_name:[[order1],[order2],[orderN]]}
order_dict = {}
orders = []
order = []

for idx, order_row in enumerate(order_rows):
    if idx == 0:
        continue
    customer = order_row[0]
    orders = []

    # Is this customer in the order dict ?
    if customer in order_dict:
        orders = order_dict[customer]
        orders.append(order_row)
        order_dict[customer] = orders
    else:
        orders.append(order_row)
        order_dict[customer] = orders

# Create a simple customer_list
# Contains a full set of unique sorted customer names
# Example: customer_list = [[erp_customer_name,end_customer_ultimate], [CustA,CustA]]
customer_list = build_customer_list()
print('customers ', len(customer_list))

# Clean up order_dict to remove:
# 1.  +/- zero sum orders
# 2. zero revenue orders
order_dict = cleanup_orders(customer_list, order_dict, sheet_map)

#
# Create a summary order file out of the order_dict
#
summary_order_rows = [order_top_row]
for key, val in order_dict.items():
    for my_row in val:
        summary_order_rows.append(my_row)
print('summary len ', len(summary_order_rows))

#
# Push our lists to an excel file
#
push_list_to_xls(summary_order_rows, app['XLS_ORDER_SUMMARY'])
push_list_to_xls(order_rows, app['XLS_ORDER_DETAIL'])
push_list_to_xls(customer_list, app['XLS_CUSTOMER'])

exit()
#
# Push our lists to a smart sheet
#
# push_xls_to_ss(wb_file, app['XLS_ORDER_SUMMARY'])
# push_xls_to_ss(wb_file, app['XLS_ORDER_DETAIL'])
# push_xls_to_ss(wb_file, app['XLS_CUSTOMER'])
exit()
