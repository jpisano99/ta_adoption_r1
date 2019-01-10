from get_list_from_ss import *
from settings import app

def get_as_update():
    as_sheet = app['SS_AS']

    rows = get_list_from_ss(as_sheet)

    as_dict = {}
    for i, row in enumerate(rows):
        if i == 0:
            continue
        customer_name = row[0]
        as_contact = row[8]
        as_status = row[5]
        as_dict[customer_name] = [as_contact, as_status]

    return as_dict


if __name__ == "__main__":
    get_as_update()