#
# This takes the .xls bookings sheet
#

def build_bookings_dict(path_to_bookings, sheet):
    # Return a dict (my_dict) with bookings file info

    my_dict = {'pathname' : path_to_bookings,
               'rows': sheet.nrows,
               'cols': sheet.ncols,
               'col_info' : [
                    ['ERP End Customer Name', -1],
                    ['End Customer Global Ultimate Name', -1],
                    ['PSS', -1],
                    ['TSA', -1],
                    ['AS PM', -1],
                    ['AS CSE', -1],
                    ['CX Contact', -1],
                    ['Renewal Date(s)', -1],
                    ['Total Bookings', -1],
                    ['Product Type', -1],
                    ['Bundle Product ID', -1],
                    ['Product Description', -1],
                    ['Sensor Count', -1],
                    ['Sales Level 1', -1],
                    ['Sales Level 2', -1],
                    ['Sales Level 3', -1],
                    ['Sales Level 4', -1],
                    ['Sales Level 5', -1],
                    ['Sales Level 6', -1]]}

    # Loop across all column headings in the bookings file and
    # Find the column number that matches the col_name in my_dict
    for bookings_col_num in range(sheet.ncols):
        # Loop across my_dict to find a match from the bookings file
        for idx, val in enumerate(my_dict['col_info']):
            col_name = val[0]
            if col_name == sheet.cell_value(0, bookings_col_num):
                my_dict['col_info'][idx][1] = bookings_col_num

    return my_dict

