from settings import *
import xlsxwriter
from build_coverage_dict import *
from build_sku_dict import *
from build_bookings_dict import *
from build_renewals_dict import *
from scrub_orders import *
from find_team import *
# from push_list_to_ss import *





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
# Open up the renewals and bookings workbooks
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
master_dict = {}
customer_list = []
csv_top_row = []
csv_rows = []
csv_row = []
sku_col_num = -1
col_pss_num = -1
col_tsa_num = -1

# Build the column titles top row \
# Also grab
# 1. sku column number
# 2. PSS and TSA column numbers
for idx, val in enumerate(bookings_dict['col_info']):
    col_name = val[0]
    col_num = val[1]
    csv_top_row.append(val[0])
    if col_name == 'Bundle Product ID':
        sku_col_num = col_num
    if col_name == 'PSS':
        col_pss_num = idx
    if col_name == 'TSA':
        col_tsa_num = idx

csv_row = csv_top_row
csv_rows.append(csv_row)

#
# Main loop of bookings data
#
for i in range(sheet_bookings.nrows):

    # SKU of interest ?
    sku = sheet_bookings.cell_value(i,sku_col_num)

    if sku in sku_dict :
        # Let's make a row for this order
        # Since it has an "interesting" sku
        csv_row = []
        sales_level = ''
        sales_level_cntr = 0
        sku_desc = sku_dict[sku][1]

        # Loop across the bookings_dict
        # to build this output row
        for val in bookings_dict['col_info']:
            col_name = val[0]
            col_idx = val[1]

            # Capture both of the Customer names
            if col_name == 'ERP End Customer Name':
                customer_name_erp = sheet_bookings.cell_value(i, col_idx)
            if col_name == 'End Customer Global Ultimate Name':
                customer_name_end = sheet_bookings.cell_value(i, col_idx)

            # Lookup the PSS/TSA team for this order
            if col_name[:-2] == 'Sales Level':
                sales_level = sales_level + sheet_bookings.cell_value(i, col_idx) +','
                sales_level_cntr += 1
                if sales_level_cntr == 6:
                    sales_level = sales_level[:-1]
                    sales_team = find_team(team_dict,sales_level)
                    pss = sales_team[0]
                    tsa = sales_team[1]
                    csv_row[col_pss_num] = pss
                    csv_row[col_tsa_num] = tsa

            if col_idx != -1:
                csv_row.append(sheet_bookings.cell_value(i, col_idx))
            elif col_name == 'Product Description':
                # Add in the Product Description
                csv_row.append(sku_desc)
            elif col_name == 'Renewal Date(s)':
                # Add in the Renewal Date if there is one
                if customer_name_erp in renewals_dict:
                    renewal_date = renewals_dict[customer_name_erp]
                    csv_row.append(renewal_date[0])
                else:
                    csv_row.append('')

            else:
                csv_row.append('')

        # Done with this row
        # Log this row and
        # Go to next row of the raw bookings data
        # customer_list.append(customer_name_erp)
        customer_list.append((customer_name_erp, customer_name_end))
        csv_rows.append(csv_row)

#
# End
#

# OK we now have a full list (csv_rows) of just the SKUs we are interested in
# As determined by the sku_dict

# Now we build a Customer Summary/Detail
# master_dict: {cust_name:[[order1],[order2],[orderN]]}

# Let's organize and summarize
orders = []
order = []
x = 0
for csv_row in csv_rows:
    customer = csv_row[0]
    if x==0:
        x += 1
        continue

    # Is this in the master dict ?
    if customer in master_dict:
        orders = []
        for order in master_dict[customer]:
            orders.append(order)

        orders.append(csv_row)
        master_dict[customer] = orders

    else:
        orders = []
        orders.append(csv_row)
        master_dict[customer]=orders


# we now create a simple customer_list list
# to contain a full set of unique customer names

# Create a unique SET of Customers
customer_list = set(customer_list)

# Convert the SET to a LIST so we can sort it
customer_list = list(customer_list)

# Sort the LIST
# customer_list.sort()
customer_list.sort(key=lambda tup: tup[0])
print('We have: ', len(customer_list), ' customers')


# Clean up orders to remove:
# 1.  +/- zero sum orders
# 2. zero revenue orders
master_dict = scrub_orders(customer_list,master_dict,csv_top_row)

# Create a csv file out of the master_dict
scrubbed_csv_rows=[]
scrubbed_csv_rows.append(csv_top_row)
for key,val in master_dict.items():
    for my_row in val:
        scrubbed_csv_rows.append(my_row)

#
# Write the CSV file
#
# print(csv_rows)
# print(customer_list)


workbook = xlsxwriter.Workbook(path_to_files + 'scrubbed' + app['AS_OF_DATE'] + '.xlsx')
worksheet = workbook.add_worksheet()

for this_row, my_val in enumerate(scrubbed_csv_rows):
    worksheet.write_row(this_row, 0, my_val)

workbook.close()

#
#
#
workbook = xlsxwriter.Workbook(path_to_files + 'master' + app['AS_OF_DATE']  + '.xlsx')
worksheet = workbook.add_worksheet()

for this_row, my_val in enumerate(csv_rows):
    worksheet.write_row(this_row, 0, my_val)

workbook.close()

exit()

#
# Push Unique Customer List to SmartSheets
#
ss_rows = []
for my_row in customer_list:
    ss_rows.append(list(my_row))

# Set the SmartSheet Column names
ss_cols = [{'primary': True, 'title': 'ERP Customer Name', 'type': 'TEXT_NUMBER'},
            {'title': 'End Customer Ultimate Name', 'type': 'TEXT_NUMBER'}]

# This first row MUST have the column names
ss_rows.insert(0, ['ERP Customer Name','End Customer Ultimate Name'])
push_list_to_ss('Unique TA Customer Names',ss_cols, ss_rows)


#
# Push Unique Customer List to a local excel workbook
#
workbook = xlsxwriter.Workbook('unique_customers.xlsx')
worksheet = workbook.add_worksheet()

for this_row, my_val in enumerate(customer_list):
    worksheet.write(this_row, 0, my_val[0])
    worksheet.write(this_row, 1, my_val[1])

workbook.close()


exit()

