"""Module to define common functions and variable for reporting the backlog."""

# Import Modules
import pandas as pd
from IPython.display import display
import spreadsheets

# Data for google sheets spreadsheet ID's.
distribution_spreadsheets = {
    'Backlog Breakdown': '135RHrwhgMmnkN7NjEjQKdts4JTus90IKGEFIrt7VHmc',
    'Carrier Push': '1zwTfoqVzZDKikoNaYDb-ArsGrV21d1QRgrzR0oyY1Zs'
}

# Sifting criteria Criteria
ok_consumer = 'data["DISTRIBUTION_CHANNEL_1"] == "OK Consumer"'
no = 'data["ORG"] == "NO"'
ok = 'data["ORG"] == "OK"'

backlog = 'data["SHIP_DATE_CATEGORY"] == "Late"'

other = 'data["SALES_CHANNEL"] == "OTHER"'
international = 'data["SALES_CHANNEL"] == "INTERNATIONAL"'
ace = 'data["DISTRIBUTION_CHANNEL_2"] == "ACE HDW"'
kilgore = 'data["SHIP_TO_NAME"] == "ORGILL INC - KILGORE - W 006"'
orgill = 'data["DISTRIBUTION_CHANNEL_2"] == "ORGILL"'
not_kilgore = 'data["SHIP_TO_NAME"] != "ORGILL INC - KILGORE - W 006"'
fsd = 'data["DISTRIBUTION_CHANNEL_2"] == "DISTRIBUTORS&FIELD SALES"'
blish = 'data["BILL_TO_NAME"] == "BLISH-MIZE CO"'
ecommerce = 'data["DISTRIBUTION_CHANNEL_2"] == "ECOMMERCE"'
lowes = 'data["DISTRIBUTION_CHANNEL_2"] == "LOWES"'
home_depot = 'data["DISTRIBUTION_CHANNEL_2"] == "HOME DEPOT"'
home_depot_dotcom = 'data["DISTRIBUTION_CHANNEL_2"] == "HOME DEPOT.COM"'
menards = 'data["DISTRIBUTION_CHANNEL_2"] == "MENARDS"'

shortage = 'data["DISTRIBUTION_STATUS"] == "Shortage"'
covered_ready = 'data["DISTRIBUTION_STATUS"] == "Covered - Ready"'
covered_not_ready = 'data["DISTRIBUTION_STATUS"] == "Covered - Not Ready"'
picked = 'data["DISTRIBUTION_STATUS"] == "Picked"'
released = 'data["DISTRIBUTION_STATUS"] == "Released"'
customer_transportation = 'data["CARRIER_STATUS"] != "None"'
overage = 'data["CARRIER_STATUS"] == "Overage"'
carrier_delay = 'data["CARRIER_STATUS"] == "Carrier Delay"'
t_m_delay = 'data["CARRIER_STATUS"] == "Transportation Management Delay"'


def read(report_name, report_location):
    """Reads standard M-D reports from .xlsx into DataFrame.
    
        Args:
            report_name(str): Name of report, e.g., 'Backlog_Report', 'Open_Orders_Extract'.
            report_location(str): File location.
            
        Returns:
            result(pandas.DataFrame): Report data.
    """
    
    # Backlog_Report no longer a valid report
    if report_name == 'Backlog_Report':
        result = pd.read_excel(report_location, sheet_name='ReportOutput', header=0, skiprows=2, index_col=0)
        result = result.fillna(0)
        result.drop(['SALES_CHANNEL', 'BUSINESS_CAT', 'Group Total', 'GrandTotal'], inplace=True)
        result = result[result.index.notnull()]
        result['PASTDUE'] = result[['PASTDUE 9+', 'PASTDUE 4 - 8', 'PASTDUE 1 - 3']].astype(float).sum(axis=1)
        
    # Open orders extract contains most info
    elif report_name == 'Open_Orders_Extract':
        result = pd.read_excel(report_location, sheet_name='ReportOutput', header=0, skiprows=1)
        result = result.fillna(0)
        
    # Backlog report contains dropship orders
    elif report_name == 'Backlog':
        result = pd.read_excel(report_location, sheet_name='ReportOutput', header=0, skiprows=1)
        result = result.fillna(0)
    else:
        raise Exception('No such report: \'{0}\''.format(report_name))
    return result

# Used in other functions
def sift(data, *args):
    """Sifts data for records that match args criteria."""
    
    data = data.copy()
    
    for criteria in args:
        data = data[eval(criteria)]
        
    return data

# Used in other functions
def sift_dollars(*args):
    """Sums dollars sifted through the sift method."""
    
    return sum(sift(*args)['DOLLARS'])

# Used to categorize consumer orders to the best of my knowledge, I don't know the actual logic, I'm sure its in oracle somewhere.
def calc_distribution_channel_1(data):
    """Determines the distribution channel of a line."""
    
    if data['ORG'] != 'OK':
        return 'Other'
    elif data['SALES_CHANNEL'] not in ["DISTRIBUTORS", "FIELD SALES", "GIANTS" , "ECOMMERCE", "INTERNATIONAL", "OTHER"]:
        return 'Other'
    elif data['BILL_TO_NAME'] == 'PARAMIT MALAYSIA SDN BHD.':
        return 'Other'
    else:
        return 'OK Consumer'

# Used to convert SALES_CHANNEL to the executive sales channel categories to the best of my knowledge, I don't know the actual logic, I'm sure its in oracle somewhere.
def calc_distribution_channel_2(data):
    """Determines the distribution channel of a line."""
    
    if data['BILL_TO_NAME'] == 'LOWES COMPANIES INC':
        return "LOWES"
    elif data['BILL_TO_NAME'] == "HOME DEPOT.COM":
        return "HOME DEPOT.COM"
    elif data['BILL_TO_NAME'] == 'HOME DEPOT':
        return 'HOME DEPOT'
    elif 'MENARDS' in data['BILL_TO_NAME']:
        return "MENARDS"
    elif data['BILL_TO_NAME'] == 'ACE HDW CORP':
        return 'ACE HDW'
    elif data['BILL_TO_NAME'] == 'ORGILL INC':
        return 'ORGILL'
    elif data['SALES_CHANNEL'] in ["DISTRIBUTORS", "FIELD SALES", "OTHER", "INTERNATIONAL"]:
        if data['BILL_TO_NAME'] == 'PARAMIT MALAYSIA SDN BHD.':
            return 'Not Consumer'
        else:
            return 'DISTRIBUTORS&FIELD SALES'
    elif data['SALES_CHANNEL'] == "ECOMMERCE":
        return 'ECOMMERCE'
    else:
        return 'Not Consumer'
        
    return result

# Categorizes lines by thier distribution status, depends on info in google sheets, not in Oracle.
def calc_distribution_status(data):
    """Determines the distribution status of a line."""

    carrier_status = data['CARRIER_STATUS']
    shortage_category = data['SHORTAGE_CATEGORY']
    line_status = data['LINE_STATUS']
    
    if carrier_status in ['Carrier Delay', 'Overage', 'Transportation Management Delay']:
        return carrier_status    
    elif (shortage_category == "Short" and line_status not in ["Released", "Picked"]) or line_status == "Awaiting":
        return 'Shortage'
    elif shortage_category == "Covered" and line_status not in ["Ready", "Released", "Picked", "Awaiting"]:
        return 'Covered - Not Ready'
    elif shortage_category == "Covered" and line_status == "Ready":
        return 'Covered - Ready'
    elif line_status == "Released":
        return 'Released'
    elif line_status == "Picked":
        return 'Picked'
    else:
        return 'Other'

# Get data that is in google sheets, not in Oracle
def get_carrier_status():
    """Gets the carrier_status table from google sheets."""
    
    # Get Data
    values = spreadsheets.get(distribution_spreadsheets['Carrier Push'], 'Carrier Push')['values']
    result = spreadsheets.df(values, index='ORDER_NUMBER')
    result.index = result.index.astype(int)
    
    # Remove Duplicate Indicies
    index = result.index
    is_duplicate = index.duplicated(keep="first")
    not_duplicate = ~is_duplicate
    result = result[not_duplicate]

    return result

# Get data that is in google sheets, not in Oracle
def calc_carrier_status(data, carrier_status):
    """Assigns a carrier status to data according to carrier_status table"""

    if data['ORDER_NUMBER'] in list(carrier_status.index):
        return carrier_status['CARRIER_STATUS'][data['ORDER_NUMBER']]
    else:
        return 'None'

# Used in program_1
def process_1(report_location='Backlog_Report.xlsx'):
    """Read Backlog_Report and write data into Backlog Breakdown."""
    
    yesterday_kwargs = {
                        'spreadsheetId': distribution_spreadsheets['Backlog Breakdown'],
                        'range': 'B3:B47'
                       }
    
    yesterday = spreadsheets.service().spreadsheets().values().get(**yesterday_kwargs).execute()['values']
      
    yesterday_request_kwargs = {
                                'spreadsheetId': distribution_spreadsheets['Backlog Breakdown'],
                                'range': 'J3:J47',
                                'valueInputOption': 'USER_ENTERED',
                                'body': {'values': yesterday}                                
                               }
    
    yesterday_request = spreadsheets.service().spreadsheets().values().update(**yesterday_request_kwargs).execute()
    
    return yesterday_request

# Used in program_2
def process_2(report_location='Backlog.xlsx'):
    """Calculate Late Dropships."""
    
    process_2_data = read('Backlog', report_location)
    result = {}
    
    late_dropships = process_2_data.query('`Shipping Org` == "OK"')
    late_dropships = late_dropships.query('`Days Late` > 0')
    late_dropships = late_dropships.query('`Order Type` == "VENDOR DROPSHIP"')
    
    orgill_data = late_dropships.query('`Bill To Customer` == "ORGILL INC"')
    result['Orgill'] = sum(orgill_data['Amount'])
    
    fsd_data = late_dropships.query('(`Sales Channel` == "Distributors") and `Bill To Customer` != "ORGILL INC"')
    result['FSD'] = sum(fsd_data['Amount'])
    
    return result

# Used in program_1
def process_3(report_location='Open_Orders_Extract.xlsx'):
    """Lists late trucks in order by dollar value"""
    
    oorex = read('Open_Orders_Extract', report_location)
    select_data = oorex.query('SHIP_DATE_CATEGORY == "Late"')
    select_data = select_data.query('ORG == "OK"')
    select_data = select_data.query('SHIPPING_CATEGORY == "TRUCK"')
    select_data = select_data.query('SALES_CHANNEL in ["DISTRIBUTORS", "FIELD SALES", "GIANTS" , "ECOMMERCE"]')
    pivot = pd.pivot_table(select_data, values='DOLLARS', index=['ORDER_NUMBER', 'BILL_TO_NAME', 'SHIP_TO_NAME'], aggfunc=sum)
    
    return pivot.sort_values('DOLLARS', ascending=False)

# Used in program_2
def process_4(report_location='Open_Orders_Extract.xlsx'):
    """Reads Open Orders report, categorizes lines, exports and returns data"""
    
    data = read('Open_Orders_Extract', report_location)
    carrier_status = get_carrier_status()
    data['DISTRIBUTION_CHANNEL_1'] = data.apply(calc_distribution_channel_1, axis=1)
    data['DISTRIBUTION_CHANNEL_2'] = data.apply(calc_distribution_channel_2, axis=1)
    data['CARRIER_STATUS'] = data.apply(calc_carrier_status, axis=1, carrier_status=carrier_status)
    data['DISTRIBUTION_STATUS'] = data.apply(calc_distribution_status, axis=1)
    select_columns = ['DISTRIBUTION_CHANNEL_1',
                      'DISTRIBUTION_CHANNEL_2',
                      'DISTRIBUTION_STATUS',
                      'SHIP_DATE_CATEGORY', 
                      'ORDER_NUMBER',
                      'DOLLARS',
                      'CASES',
                      'ORG',
                      'SALES_CHANNEL',
                      'BILL_TO_NAME',
                      'SHIP_TO_NAME',
                      'SHIP_DATE',
                      'SHIPPING_METHOD',
                      'SHIPPING_CATEGORY',
                      'LINE_STATUS',
                      'CARRIER_STATUS',
                      'SHORTAGE_CATEGORY',
                      'ITEM_NO',
                      'ITEM_DESCRIPTION',
                      'PIECE_QTY',
                      'Open Orders',
                      'DISTRIBUTION_ONHAND',
                      'RESERVED_QUANTITY']
    
    data = data[select_columns]
    data.to_excel('Line Detail.xlsx', sheet_name='Line Detail', index=False)

    return data

def program_1():
    """Runs processes 1 and 3."""
    
    p1_request = process_1()
    late_trucks = process_3()
    
    display(p1_request)
    display(late_trucks.head(60))

def program_2():
    """Runs processes 2 and 4 and updates Backlog Breakdown Google Sheet."""
    
    dropships = process_2()
    p4_data = process_4()
    
    data = [
            {'range': 'Current!B6:B10',
             'values': [[sift_dollars(p4_data, ok_consumer, backlog, ace, shortage)],
                        [sift_dollars(p4_data, ok_consumer, backlog, ace, covered_ready)],
                        [sift_dollars(p4_data, ok_consumer, backlog, ace, picked)],
                        [sift_dollars(p4_data, ok_consumer, backlog, ace, released)],
                        [sift_dollars(p4_data, ok_consumer, backlog, ace, customer_transportation)]
                       ]},
            {'range': 'Current!B12:B21',
             'values': [[sift_dollars(p4_data, ok_consumer, backlog, kilgore, shortage)],
                        [sift_dollars(p4_data, ok_consumer, backlog, kilgore, covered_ready)],
                        [sift_dollars(p4_data, ok_consumer, backlog, kilgore, picked)],
                        [sift_dollars(p4_data, ok_consumer, backlog, kilgore, released)],
                        [sift_dollars(p4_data, ok_consumer, backlog, kilgore, customer_transportation)],
                        [sift_dollars(p4_data, ok_consumer, backlog, orgill, not_kilgore, shortage)],
                        [sift_dollars(p4_data, ok_consumer, backlog, orgill, not_kilgore, covered_ready)],
                        [sift_dollars(p4_data, ok_consumer, backlog, orgill, not_kilgore, picked)],
                        [sift_dollars(p4_data, ok_consumer, backlog, orgill, not_kilgore, released)],
                        [sift_dollars(p4_data, ok_consumer, backlog, orgill, not_kilgore, covered_not_ready) + dropships['Orgill']]
                       ]},
            {'range': 'Current!B23:B28',
             'values': [[sift_dollars(p4_data, ok_consumer, backlog, fsd, shortage)],
                        [sift_dollars(p4_data, ok_consumer, backlog, fsd, covered_ready)],
                        [sift_dollars(p4_data, ok_consumer, backlog, fsd, picked)],
                        [sift_dollars(p4_data, ok_consumer, backlog, fsd, released)],
                        [sift_dollars(p4_data, ok_consumer, backlog, fsd, customer_transportation)],
                        [sift_dollars(p4_data, ok_consumer, backlog, fsd, covered_not_ready) + dropships['FSD']]
                       ]},
            {'range': 'Current!B30',
             'values': [[sift_dollars(p4_data, ok_consumer, backlog, ecommerce, shortage)]
                       ]},            
            {'range': 'Current!B33:35',
             'values': [[sift_dollars(p4_data, ok_consumer, backlog, lowes, overage)],
                        [sift_dollars(p4_data, ok_consumer, backlog, lowes, carrier_delay)],
                        [sift_dollars(p4_data, ok_consumer, backlog, lowes, t_m_delay)],
                       ]},
            {'range': 'Current!B39:42',
             'values': [[sift_dollars(p4_data, ok_consumer, backlog, home_depot, overage)],
                        [sift_dollars(p4_data, ok_consumer, backlog, home_depot, carrier_delay)],
                        [sift_dollars(p4_data, ok_consumer, backlog, home_depot, t_m_delay)],
                       ]},
            {'range': 'Current!B44',
             'values': [[sift_dollars(p4_data, ok_consumer, backlog, menards, shortage)]
                       ]}        
           ]
    
    body = {
            'valueInputOption': 'USER_ENTERED',
            'data': data
           }

    kwargs = {
              'spreadsheetId': distribution_spreadsheets['Backlog Breakdown'],
              'body': body
             }

    spreadsheets.service().spreadsheets().values().batchUpdate(**kwargs).execute()
    
    select_data = sift(p4_data, ok_consumer, backlog)
    table1 = pd.pivot_table(select_data, values='DOLLARS', index='DISTRIBUTION_CHANNEL_2', aggfunc=sum)
    spreadsheets.update(
        distribution_spreadsheets['Backlog Breakdown'],
        'OK Consumer Backlog!A3',
        spreadsheets.values(table1, index=True)
    )
    
    select_data = sift(p4_data, ok_consumer, backlog)
    table1 = pd.pivot_table(select_data, values='DOLLARS', index='DISTRIBUTION_STATUS', aggfunc=sum)
    spreadsheets.update(
        distribution_spreadsheets['Backlog Breakdown'],
        'OK Consumer Backlog!A19',
        spreadsheets.values(table1, index=True)
    )

# Not in use.
# def breakdown(report_location='Backlog_Report.xlsx'):
#    """Fills out Backlog Breakdown."""
#       
#    backlog_report = read('Backlog_Report', report_location)
#
#    backlog_channels = {'HOME DEPOT-OK':'Breakdown!B6',
#                        'HOME DEPOT.COM-OK':'Breakdown!B15',
#                        'LOWES-OK':'Breakdown!B24',
#                        'MENARDS-OK':'Breakdown!B33',
#                        'ACE HDW-OK':'Breakdown!B42',
#                        'ORGILL-OK':'Breakdown!B51',
#                        'DISTRIBUTORS&FIELD SALES-OK':'Breakdown!B60',
#                        'ECOMMERCE-OK':'Breakdown!B69'
#                       }
#    
#    other_sites = ['HOME DEPOT-AR', 'HOME DEPOT-MDT', 'HOME DEPOT-WB', 'HOME DEPOT.COM-MDT', 'HOME DEPOT.COM-WB',
#                   'LOWES-AR']
#    
#    other_sites_backlog = 0
#    for x in other_sites:
#        try:
#            other_sites_backlog = other_sites_backlog + backlog_report['PASTDUE'][x]
#        except:
#            pass
#    
#    data = []
#    for index, value in backlog_channels.items():
#        upload_value = backlog_report['PASTDUE'][index]
#        data.append({'range': value, 'values': [[upload_value]]})
#    data.append({'range': 'Breakdown!B78', 'values': [[other_sites_backlog]]})
#     
#    body = {
#            'valueInputOption': 'USER_ENTERED',
#            'data': data
#           }
#    
#    request_kwargs = {
#                      'spreadsheetId': distribution_spreadsheets['Backlog Breakdown'],
#                      'body': body
#                     }
#
#    request = spreadsheets.service().spreadsheets().values().batchUpdate(**request_kwargs).execute()
    
def main():
    program_1()
    program_2()

if __name__ == "__main__":
    main()
