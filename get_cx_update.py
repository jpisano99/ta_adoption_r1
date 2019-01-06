from get_list_from_ss import *
from settings import app


def get_cx_update():
    cx_sheet = app['SS_CX']

    rows = get_list_from_ss(cx_sheet)
    cx_dict = {}
    for i, row in enumerate(rows):
        if i == 0 :
            continue
        customer_name = row[1]
        cx_contact = row[20]
        cx_status = row[16]
        cx_dict[customer_name] = [cx_contact,cx_status]

    return cx_dict


