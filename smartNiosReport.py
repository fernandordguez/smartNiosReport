from csv import DictWriter
import csv
import xmltodict
from collections import defaultdict
import gspread
import argparse
from argparse import RawDescriptionHelpFormatter
import bloxone
from mod import read_b1_ini,verify_api_key
import sys

def pastecsv(data, gsheet, wksname):      #When gsheet exists, this function is used to import data without
                                                # deleting all the existing worksheets (unlike import_csv function)
    gsheet.add_worksheet(wksname, len(data) + 10, len(data[0]) + 10)
    wksheet = gsheet.worksheet(wksname)
    body = {'requests': [{'pasteData': {'coordinate': {'sheetId': wksheet.id,'rowIndex': 0,'columnIndex': 0},'data': data,'type': 'PASTE_NORMAL','delimiter': ','}}]}
    gsheet.batch_update(body)
    return None


def get_args():  ## Handles the arguments passed to the script from the command line
    # Parse arguments
    usage = ' -c b1config.ini -i {"wapi", "xml"} -r {"log","csv","gsheet"} [ --delimiter {",",";"} ] [ --yaml <yaml file> ] [ --help ]'
    description = 'This script gets DHCP leases from NIOS (via WAPI or from a Grid Backup), collects BloxOne DHCP leases from B1DDI API and compares network by network the number of leases on each platform'
    epilog = ''' sample b1config.ini 
                [BloxOne]
                url = 'https://csp.infoblox.com'
                api_version = 'v1'
                api_key = 'API_KEY'''
    par = argparse.ArgumentParser(formatter_class=RawDescriptionHelpFormatter, description=description, add_help=False,usage='%(prog)s' + usage, epilog=epilog)
    # Required Argument(s)
    required = par.add_argument_group('Required Arguments')
    req_grp = required.add_argument
    req_grp('-c', '--config', action="store", dest="config", help="Path to ini file with API key", required=True)
    req_grp('-i', '--interface', action="store", dest="interface", help="source from where NIOS data will be imported",choices=['wapi', 'xml'], required=True)
    req_grp('-r', action="store", dest="report", help="Defines the type of reporting that will be produced",choices=['log', 'csv', 'gsheet'], required=True, default='log')
    # Optional Arguments(s)
    optional = par.add_argument_group('Optional Arguments')
    opt_grp = optional.add_argument
    opt_grp('--delimiter', action="store", dest="csvdelimiter", help="Delimiter used in CSV data file", choices=[',', ';'])
    opt_grp('--yaml', action="store", help="Alternate yaml file for supported objects", default='objects.yaml')
    opt_grp('--debug', action='store_true', help=argparse.SUPPRESS, dest='debug')
    opt_grp('-f', '--filter', action='store_true', help='Excludes networks with 0 leases from the report',dest='filter')
    # opt_grp('--version', action='version', version='%(prog)s ' + __version__)
    opt_grp('-h', '--help', action='help', help='show this help message and exit')
    return par.parse_args(args=None if sys.argv[1:] else ['-h'])


def csvtogsheet(sheetname,wksname,conf):  # Opens (if exists) or Creates (if doesn´t) a Gsheet
    myibmail = conf['ib_email']
    gc = gspread.service_account(conf['ib_service_acc'])
    # Email account is important, otherwise user will not be allowed to access or update the gsheet (it's created by a service account')
    try:  # Sheet exists, we just open the document. We cannot import the CSV because it will delete all other wksheets
        sheet = gc.open(sheetname)
        wks = pastecsv(csvContents, sheet, wksname)  # This function does not delete any existing worksheets
        with open(conf['csvfile'], 'r') as f:
            csvContents = f.read()
        #formatGsheet(wks)
    except gspread.exceptions.SpreadsheetNotFound:  # Sheet does not exists --> New sheet, we can just import the csv
        try:
            sheet = gc.create(sheetname)
            # Adapt as required. By default, it will share the document with the service account email (mandatory) and with
            # the email address introduced previously with read/write permission. Anyone with the link can access in R/O
            sheet.share(gc.auth.service_account_email, role='writer', perm_type='user')
            sheet.share(myibmail, role='writer', perm_type='user')
            sheet.share('', role='reader', perm_type='anyone')
            gc.import_csv(sheet.id,data)  # deletes any existing worksheets & imports the data in sh.sheet1
            sheet = gc.open_by_key(sheet.id)  # use import_csv method only with new gsheets (to keep history record)
            wks = sheet.sheet1
            wks.update_title(wksname)
            #formatGsheet(wks)
        except gspread.exceptions.GSpreadException:
            print("error while creating the Gsheet")
    print("Gsheet available on the URL", sheet.url)  # Returns the URL to the Gsheet
    print('Filter option was enabled so output is reduced to networks with any leases \n')
    return None

def parseniosdb():
    listobjects = []
    xmlfile = '/Users/fernandorguez/onedb.xml'
    shortlistedtypes = ['.com.infoblox.dns.bind_ns', '.com.infoblox.dns.bind_soa', '.com.infoblox.dns.member_views_item', '.com.infoblox.dns.network', '.com.infoblox.dns.network_container']
    fxmml = open(xmlfile, 'r')
    xml_content = fxmml.read()
    objects = xmltodict.parse(xml_content)
    objects = objects['DATABASE']['OBJECT']
    for obj in objects:
        anobject = {}
        for item in (obj['PROPERTY']):
            anobject[item['@NAME']] = item['@VALUE']
        listobjects.append(anobject)
    return listobjects

args = get_args()
conf = read_b1_ini(args.config)
listobjects= parseniosdb()
report = defaultdict()
for ob in listobjects:
    if '__type' in ob:
        report.setdefault(ob.pop('__type'), []).append(ob)

csvname = ''
shortlistedtypes = ['.com.infoblox.dns.bind_ns', '.com.infoblox.dns.bind_soa', '.com.infoblox.dns.member_views_item', '.com.infoblox.dns.network', '.com.infoblox.dns.network_container']
for key in shortlistedtypes:
    csvname = key + '.csv'
    with open(csvname, 'w', newline='') as csvfile:
        cols = list(report[key][0].keys())
        writer = csv.DictWriter(csvfile, fieldnames = cols)
        writer.writeheader()
        writer.writerows(report[key])
    

#csvtogsheet('Smart NIOS report', key, conf, datatoexport)


#wksheet = pastecsv(datatoexport, 'gsheet', 'DNS Report')