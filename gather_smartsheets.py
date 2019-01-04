from settings import *
from get_list_from_ss import *
import smartsheet

#
# Pull down the Smartsheets we need to build
#

home = app['HOME']
working_dir = app['WORKING_DIR']
download_path = home +'\\' + working_dir  + '\\'


my_cx = Ssheet(app['SS_CX'])
my_saas = Ssheet(app['SS_SAAS'])
my_as = Ssheet(app['SS_AS'])


ss_token = passwords['SS_TOKEN']
ss = smartsheet.Smartsheet(ss_token)

# Download CX Data
ss.Sheets.get_sheet_as_excel(my_cx.id, download_path)
file = download_path + '\\'   + app['SS_CX'] + '.xlsx'
file1 = download_path + '\\tmp_CX_data' + app['AS_OF_DATE'] + '.xlsx'
os.rename(file,file1)

# Download AS Data
ss.Sheets.get_sheet_as_excel(my_as.id, download_path)
file = download_path + '\\'   + app['SS_AS'] + '.xlsx'
file1 = download_path + '\\tmp_AS_data' + app['AS_OF_DATE'] + '.xlsx'
os.rename(file,file1)

# Download SAAS Data
ss.Sheets.get_sheet_as_excel(my_saas.id, download_path)
file = download_path + '\\'   + app['SS_SAAS'] + '.xlsx'
file1 = download_path + '\\tmp_SAAS_data' + app['AS_OF_DATE'] + '.xlsx'
os.rename(file,file1)