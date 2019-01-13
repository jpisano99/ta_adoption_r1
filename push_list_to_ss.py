from Ssheet_class import Ssheet

#
# Get the SmartSheet Coverage as a simple list
#
# EXAMPLE USAGE:
# ss_cols = [{'primary': True, 'title': 'ERP Customer Name', 'type': 'TEXT_NUMBER'},
#             {'title': 'End Customer Ultimate Name', 'type': 'TEXT_NUMBER'}]
#
# # This first row of ss_rows MUST have the column names
#
# ss_rows.insert(0, ['ERP Customer Name','End Customer Ultimate Name'])
# push_list_to_ss('Thurs Test',ss_cols, ss_rows)
#


def push_list_to_ss(ss_name, my_cols, my_rows):

    my_ss = Ssheet(ss_name)

    if my_ss.id != -1:
        print(my_ss.id, ss_name + ' Already Exists')
        exit()
    else:
        # Assumption is we are CREATING this SmartSheet first
        my_ss.create_sheet(ss_name, my_cols)
        my_ss.refresh()

        # Grab the col_names from the FIRST row of my_rows
        my_col_names = my_rows[0]

        # Get this dict which has {col_name:col_id}
        this_ss_col_names = my_ss.col_name_idx

        rows_to_add = []

        # Go down the row
        for row_num, row in enumerate(my_rows):
            cells_to_add = []

            if row_num == 0:
                continue

            # Go across the columns to add cells for this row
            for col_num, cell_value in enumerate(row):
                this_col_name = my_col_names[col_num]
                this_col_id = this_ss_col_names[this_col_name]

                cells_to_add.append({"strict": False, "columnId": this_col_id, "value": cell_value})

            rows_to_add.append(cells_to_add)

        # Send to SmartSheets to ADD
        my_ss.add_rows(rows_to_add)
    return
