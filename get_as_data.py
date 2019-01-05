from get_list_from_ss import *


as_sheet = 'Tetration Shipping Notification & Invoicing Status'
folder_to_bookings = 'C:/Users/jim/PycharmProjects/renewals'

rows = get_list_from_ss(as_sheet)
for row in rows:
    print (row[0],row[3])