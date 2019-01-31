#
# [Source Column Name, Which SS/Excel sheet ?, Source Column Num, NEW Column name]
#
sheet_map = [['ERP End Customer Name', 'XLS_BOOKINGS', -1, 'Customer Name'],
             ['End Customer Global Ultimate Name', 'XLS_BOOKINGS', -1, 'Customer Alias'],
             ['pss', 'SS_COVERAGE', -1, 'PSS'],
             ['tsa', 'SS_COVERAGE', -1, 'TSA'],
             ['AS PM', 'SS_AS', -1, ''],
             ['AS CSE', 'SS_AS', -1, ''],
             ['Project Status/PM Completion', 'SS_AS', -1, 'AS Status'],
             ['Delivery Comments', 'SS_AS', -1, 'AS Comments'],
             ['Provisioning completed', 'SS_SAAS', -1, 'SAAS Status'],
             ['CuSM Name', 'SS_CX', -1, 'CX Contact'],
             ['Next Action', 'SS_CX', -1, 'CX Next Steps'],
             # ['Renewal Date', 'XLS_RENEWALS', -1, 'Renewal Date(s)'],
             # ['Product Bookings', 'XLS_RENEWALS', -1, 'Renewal Revenue'],
             ['Orders Found', '', -1, ''],
             ['Total Bookings', 'XLS_BOOKINGS', -1, ''],
             ['Service Bookings', '', -1, ''],
             ['Product Type', 'SS_SKU', -1, '*DELETE*'],
             ['Bundle Product ID', 'XLS_BOOKINGS', -1, '*DELETE*'],
             ['Product Description', 'SS_SKU', -1, 'Platform Type'],
             ['Sensor Count', 'SS_SKU', -1, ''],
             ['Active Sensors', '', -1, ''],
             ['Sales Level 1', 'XLS_BOOKINGS', -1, ''],
             ['Sales Level 2', 'XLS_BOOKINGS', -1, ''],
             ['Sales Level 3', 'XLS_BOOKINGS', -1, ''],
             ['Sales Level 4', 'XLS_BOOKINGS', -1, ''],
             ['Sales Level 5', 'XLS_BOOKINGS', -1, ''],
             ['Sales Level 6', 'XLS_BOOKINGS', -1, '']]

sheet_keys = [['XLS_BOOKINGS', 'ERP End Customer Name', -1],
              ['SS_CX', 'Account Name', -1],
              ['SS_SAAS', 'Customer name', -1],
              ['SS_AS', 'Customer Name', -1]]
