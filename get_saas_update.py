from get_list_from_ss import *
from settings import app

def get_saas_update():
    saas_sheet = app['SS_SAAS']

    rows = get_list_from_ss(saas_sheet)

    saas_dict = {}
    for i, row in enumerate(rows):
        if i == 0:
            continue
        customer_name = row[1]
        saas_contact = row[9]
        saas_status = row[4]
        saas_dict[customer_name] = [saas_contact, saas_status]

    return saas_dict


if __name__ == "__main__":
    print(get_saas_update())









# from get_list_from_ss import *
# import xlrd
# import xlsxwriter
#
# from fuzzywuzzy import fuzz
#
#
# saas_sheet = 'SaaS customer tracking'
# path_to_customer = 'C:/Users/jim/PycharmProjects/renewals/unique_customers.xlsx'
#
#
# wb_customers = xlrd.open_workbook(path_to_customer)
# master_list = wb_customers.sheet_by_index(0)
# customer1_dict={}
# customer2_dict={}
#
# # Build the customers dicts based on
# # customer1_dict = ERP Customer Name
# # customer2_dict = End Customer Ultimate Name
# for i in range(master_list.nrows):
#     customer1_dict[master_list.cell_value(i, 0)] = i
#
# for i in range(master_list.nrows):
#     customer2_dict[master_list.cell_value(i, 1)] = i
#
# print(customer1_dict)
# print(customer2_dict)
# print (len(customer1_dict), len(customer2_dict))
#
# rows = get_list_from_ss(saas_sheet)
# print (len(rows),len(customer1_dict))
#
# for row in rows:
#     saas_customer = row[1]
#     if saas_customer not in customer1_dict:
#         print('---------------------------------')
#         print ('No exact match for:', saas_customer)
#
#         best_fuzz = 0
#         best_customer = ''
#         for this_customer in customer1_dict:
#
#             fuzz_num = fuzz.ratio(this_customer, saas_customer)
#
#             if fuzz_num >= best_fuzz:
#                 best_fuzz = fuzz_num
#                 best_customer = this_customer
#
#         print('\t\tBEST fuzzy match: ', best_customer, ' \t',best_fuzz)
