import xlrd
import xlsxwriter
from settings import app
from build_bookings_dict import build_bookings_dict


def build_customer_list():
    #
    # Get settings for file locations and names
    #
    home = app['HOME']
    working_dir = app['WORKING_DIR']
    bookings_file = app['XLS_BOOKINGS']  # Master bookings has 9958 rows as of 12-15-18
    path_to_files = home + '\\' + working_dir + '\\'
    path_to_bookings = path_to_files + bookings_file

    wb_bookings = xlrd.open_workbook(path_to_bookings)
    sheet_bookings = wb_bookings.sheet_by_index(0)

    bookings_dict = build_bookings_dict(path_to_bookings, sheet_bookings)
    customer_list = []
    col_num_end = 0
    col_num_erp = 0

    #
    # First find the column numbers for these column names in the book
    #
    for val in bookings_dict['col_info']:
        if val[0] == 'ERP End Customer Name':
            col_num_erp = val[1]
        elif val[0] == 'End Customer Global Ultimate Name':
            col_num_end = val[1]

    #
    # Main loop of bookings excel data
    #
    for i in range(sheet_bookings.nrows):
        # Capture both of the Customer names
        if i == 0:
            continue

        customer_name_erp = sheet_bookings.cell_value(i, col_num_erp)
        customer_name_end = sheet_bookings.cell_value(i, col_num_end)
        customer_list.append((customer_name_erp, customer_name_end))

    # Create a simple customer_list list of tuples
    # Contains a full set of unique sorted customer names
    # customer_list = [(erp_customer_name,end_customer_ultimate), (CustA,CustA)]
    customer_list = set(customer_list)

    # Convert the SET to a LIST so we can sort it
    customer_list = list(customer_list)

    # Sort the LIST
    customer_list.sort(key=lambda tup: tup[0])

    #
    # Write TA Customer List to a local excel workbook
    #
    # Insert a header row before writing
    top_row = ['erp_customer_name', 'end_customer_ultimate_name']
    customer_list.insert(0, top_row)

    wb_file = path_to_files + app['XLS_CUSTOMER'] + app['AS_OF_DATE'] + '.xlsx'
    workbook = xlsxwriter.Workbook(wb_file)
    worksheet = workbook.add_worksheet()
    for this_row, my_val in enumerate(customer_list):
        worksheet.write(this_row, 0, my_val[0])
        worksheet.write(this_row, 1, my_val[1])
    workbook.close()

    return customer_list


if __name__ == "__main__":
    our_customers = build_customer_list()
    print('We have: ', len(our_customers), ' customers')
