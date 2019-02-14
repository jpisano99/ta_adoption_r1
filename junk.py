import win32com.client as win32
from Ssheet_class import Ssheet
from push_list_to_xls import push_list_to_xls

sh = Ssheet('TA Unified Adoption Dashboard_as_of_01_31_2019')
print (sh.id)
exit()
jim = [['col1', 'col2'], ['stan', 'blanche']]

push_list_to_xls(jim, 'ang')

# cell_format = wb.add_format()
# cell_format.set_bold()
# cell_format.set_font_color('red')

exit()
#
# excel = win32.gencache.EnsureDispatch('Excel.Application')
# my_path = "C:/Users/jpisano/PycharmProjects/ta_adoption_r1/jim.xlsx"
# wb = excel.Workbooks.Open(my_path)
#
# ws = wb.Worksheets("Sheet1")
# ws.Columns("C:C").AutoFit()
#
# wb.Save()
# excel.Application.Quit()