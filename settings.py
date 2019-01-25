__author__ = 'jpisano'

from datetime import datetime
from my_secrets import passwords
import os

# database configuration settings
database = dict(
    DATABASE="cust_ref_db",
    USER="root",
    PASSWORD=passwords["DB_PASSWORD"],
    HOST="localhost"
)

# Smart sheet Config settings
ss_token = dict(
    SS_TOKEN=passwords["SS_TOKEN"]
)

# application predefined constants
app = dict(
    VERSION=1.0,
    GITHUB="{url}",
    HOME=os.path.expanduser("~"),
    WORKING_DIR='desktop\TA Adoption Data',
    XLS_RENEWALS='TA Renewal Dates as of 12-10-18.xlsx',
    XLS_BOOKINGS='TA Master Bookings as of 12-15-18.xlsx',
    XLS_CUSTOMER='TA Customer List',
    XLS_ORDER_DETAIL='TA Order Details',
    XLS_ORDER_SUMMARY='TA Order Summary',
    SS_SAAS='SaaS customer tracking',
    SS_CX='Tetration Engaged Customer Report',
    SS_AS='Tetration Shipping Notification & Invoicing Status',
    SS_COVERAGE='Tetration Coverage Map',
    SS_SKU='Tetration SKUs',
    SS_CUSTOMERS='TA Customer List',
    SS_DASHBOARD='TA Unified Adoption Dashboard',
    SS_WORKSPACE='Tetration Customer Adoption Workspace',
    AS_OF_DATE=datetime.now().strftime('_as_of_%m_%d_%Y')
)
