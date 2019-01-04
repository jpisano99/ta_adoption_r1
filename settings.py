__author__ = 'jpisano'

from datetime import datetime
from my_secrets import passwords
import os

#database configuration settings
database = dict(
    DATABASE = "cust_ref_db",
    USER     = "root",
    PASSWORD = passwords["DB_PASSWORD"],
    HOST     = "localhost"
)

#Smartsheet Config settings
smartsheet = dict(
    SS_TOKEN = passwords["SS_TOKEN"]
)

#application predefined constants
app = dict(
    VERSION   = 1.0,
    GITHUB    = "{url}",
    HOME = os.path.expanduser("~"),
    WORKING_DIR = 'desktop\TA Adoption Data',
    RENEWALS = 'TA Renewal Dates as of 12-10-18.xlsx',
    BOOKINGS ='TA Master Bookings as of 12-15-18.xlsx',
    SS_SAAS = 'SaaS customer tracking',
    SS_CX='Tetration Engaged Customer Report',
    SS_AS='Tetration Shipping Notification & Invoicing Status',
    SS_COVERAGE='Tetration Coverage Map',
    SS_SKU='Tetration SKUs',
    SS_CUSTOMERS = 'Tetration Customer Names',
    SS_MASTER = 'TA Unified Adoption Planner',
    SS_WORKSPACE = 'Tetration Customer Adoption Workspace',
    AS_OF_DATE = datetime.now().strftime('_as_of_%m_%d_%Y')
)