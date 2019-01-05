from get_list_from_ss import *


cx_sheet = 'Tetration Engaged Customer Report'
folder_to_bookings = 'C:/Users/jim/PycharmProjects/renewals'

rows = get_list_from_ss(cx_sheet)
for row in rows:
    print (row[1])


