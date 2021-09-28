#!/usr/bin/python3
'''
------------------------------------------------------------------------
 Usage:
   python3 smartNiosReport.py -c config.ini -r [excel | gsheet]
 Description:
   Python script that reads Grid DB, parses the data and generates several
    reports based on the corresponding yaml file. This file that can be
    used to define the reports to be created and the fields that must be
    present on each one of them
 Requirements:
   Python3 with collections, xmltodict, argparse, sys,
   gspread, csv, gspread_formatting, time, pandas
   Used two custom modules (dblib and mod) to reuse few functions
 Author: Fernando Rodriguez
    (some code has been from my colleagues Chris Marrison & John Neerdael
    has been re-used)
 Date Last Updated: 20210928

 Copyright (c) 2021 Fernando Rodriguez / Infoblox
 Redistribution and use in source and binary forms,
 with or without modification, are permitted provided
 that the following conditions are met:
 1. Redistributions of source code must retain the above copyright
 notice, this list of conditions and the following disclaimer.
 2. Redistributions in binary form must reproduce the above copyright
 notice, this list of conditions and the following disclaimer in the
 documentation and/or other materials provided with the distribution.
 THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS
 "AS IS" AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT
 LIMITED TO, THE IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS
 FOR A PARTICULAR PURPOSE ARE DISCLAIMED. IN NO EVENT SHALL THE
 COPYRIGHT HOLDER OR CONTRIBUTORS BE LIABLE FOR ANY DIRECT, INDIRECT,
 INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING,
 BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES;
 LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER
 CAUSED AND ON ANY THEORY OF LIABILITY, WHetreeHER IN CONTRACT, STRICT
 LIABILITY, OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN
 ANY WAY OUT OF THE USE OF THIS SOFTWARE, EVEN IF ADVISED OF THE
 POSSIBILITY OF SUCH DAMAGE.
------------------------------------------------------------------------
'''
__version__ = '0.5.0'
__author__ = 'Fernando Rodriguez'
__author_email__ = 'frodriguez@infoblox.com'

from collections import defaultdict
import dblib
import xmltodict
import argparse
from argparse import RawDescriptionHelpFormatter
import sys
import gspread
import csv
from mod import read_niosdb_ini
from gspread_formatting import *
import time
import pandas as pd

def formatGsheet(wks, colcount):  ## Applies a bit of formatting to the Google Sheet document created
    body = {"requests": [{"autoResizeDimensions": {
        "dimensions": {"sheetId": wks.id, "dimension": "COLUMNS", "startIndex": 0, "endIndex": colcount}}}]}
    wks.spreadsheet.batch_update(body)
    set_frozen(wks, rows=1)
    borderitem = {"style": "SOLID"}
    fmt = cellFormat(textFormat=textFormat(bold=True, foregroundColor=color(1, 1, 1)), horizontalAlignment='CENTER',
                     backgroundColor=color(0, 0, 1), borders={"top": borderitem,"right": borderitem,"bottom": borderitem,"left": borderitem})
    rangefmt = 'A1:' + gspread.utils.rowcol_to_a1(1,colcount)
    format_cell_ranges(wks, [(rangefmt, fmt)])
    return None

def pastecsv(csvContents, sheet, wksname):  # When gsheet exists, this function is used to import data without
    try:                                        # deleting all the existing worksheets (unlike import_csv function)
        wksheet = sheet.add_worksheet(wksname, len(csvContents), len(csvContents[0]))
        body = {'requests': [{'pasteData': {'coordinate': {'sheetId': wksheet.id, 'rowIndex': 0, 'columnIndex': 0},
                                         'data': csvContents, 'type': 'PASTE_NORMAL', 'delimiter': ','}}]}
        
        try:
            wksheet = sheet.batch_update(body)
            return wksheet
        except gspread.exceptions.APIError as err:
            print('Error updating the worksheet:',str(err))
            return
    except gspread.exceptions.GSpreadException or gspread.exceptions.APIError as errpastecsv:
        try:
            wksheet = sheet.worksheet(wksname)
            body = {'requests': [{'pasteData': {'coordinate': {'sheetId': wksheet.id, 'rowIndex': 0, 'columnIndex': 0},
                                            'data': csvContents, 'type': 'PASTE_NORMAL', 'delimiter': ','}}]}
            wksheet.clear()
            try:
                wksheet = sheet.batch_update(body)
                return wksheet
            except gspread.exceptions.APIError as err:
                print('Error updating the worksheet:', str(err))
                print(wksname, ' report has not been completed')
                return
            return wksheet
        except gspread.exceptions.WorksheetNotFound as ewnf:
            print('Worksheet not found', str(ewnf))
            print(wksname,' report has not been completed')
    return

def csvtogsheet(conf, wksname, timenow):  # Opens (if exists) or Creates (if doesn´t) a Gsheet
    sheetname = ''
    myibmail = conf['ib_email']
    gc = gspread.service_account(conf['ib_service_acc'])
    sheetname = 'Smart Report' # + timenow
    with open(wksname, 'r') as f:
        csvContents = f.read()
    # Email account is important, otherwise user will not be allowed to access or update the gsheet (it's created by a service account')
    try:  # Sheet exists, we just open the document. We cannot import the CSV because it will delete all other wksheets
        sheet = gc.open(sheetname)
        wks = pastecsv(csvContents, sheet, wksname)  # This function does not delete any existing worksheets
        #if wks is not None:
            #formatGsheet(wks)
    except gspread.exceptions.SpreadsheetNotFound:  # Sheet does not exists --> New sheet, we can just import the csv
        try:
            sheet = gc.create(sheetname)
            # Adapt as required. By default, it will share the document with the service account email (mandatory) and with
            # the email address introduced previously with read/write permission. Anyone with the link can access in R/O
            sheet.share(gc.auth.service_account_email, role='writer', perm_type='user')
            sheet.share(myibmail, role='writer', perm_type='user')
            sheet.share('', role='reader', perm_type='anyone')
            gc.import_csv(sheet.id,csvContents)  # deletes any existing worksheets & imports the data in sh.sheet1
            sheet = gc.open_by_key(sheet.id)  # use import_csv method only with new gsheets (to keep history record)
            wks = sheet.sheet1
            wks.update_title(wksname)
            #formatGsheet(wks)
        except gspread.exceptions.GSpreadException:
            print("error while creating the Gsheet")
            return None
    return sheet

def export2gsheet(listcsvs, conf):
    # Export to a single Excel file with multiple tabs
    timenow = time.strftime('%Y/%m/%d - %H:%M')
    for k in listcsvs:
        sheet = csvtogsheet(conf, k, timenow)
    #if sheet is not None:
    if isinstance(sheet,gspread.models.Spreadsheet):
        for w in sheet.worksheets():
            formatGsheet(w, w.col_count)
        print("Gsheet available on the URL", sheet.url)  # Returns the URL to the Gsheet
    return None

def export2excel(listcsvs):
    # Export to a single Excel file with multiple tabs
    with pd.ExcelWriter('Smart Report.xlsx', engine='xlsxwriter') as writer:
        for k in listcsvs:
            dftemp = pd.read_csv(k)
            dftemp.to_excel(writer, sheet_name=k.split('.')[0])
    return None

def processreports(listobjects, yamlObjects):
    vnodeids = {}
    netviews = {}
    tempobject = {}
    report = defaultdict()
    for ob in listobjects:
        if '__type' in ob and ob['__type'] in yamlObjects.objects():    #Read first virtual node ids and network views
            if ob['__type'] == '.com.infoblox.one.virtual_node':        #to improve readibility
                vnodeids[ob['virtual_oid']] = ob['host_name']
            elif ob['__type'] == '.com.infoblox.dns.network_view':
                netviews[ob['id']] = ob['name']
    for ob in listobjects:
        if '__type' in ob and ob['__type'] in yamlObjects.objects():
            if 'properties' in yamlObjects.obj_keys(ob['__type']) and yamlObjects.properties(ob['__type']) is not None:
                tempobject = {}
                for field in yamlObjects.properties(ob['__type']):  #Created new object with only the fields in yaml dict
                    if isinstance(field, str) and field in ob.keys():
                        if field == 'network_view':             #Translate network view ids to view names for readibility
                            tempobject['network_view'] = netviews[ob['network_view']]
                        elif field == 'virtual_node':           #Translate vnode ids to member names for readibility
                            tempobject['virtual_node'] = vnodeids[ob['virtual_node']]
                        else:
                            tempobject[field] = ob[field]
                report.setdefault(ob['__type'], []).append(tempobject)  #Appends all objects of the same type
    listcsvs = []
    for key in yamlObjects.objects():
        if yamlObjects.obj_type(key) is not None and key != 'object' and key in report.keys():
            csvname = yamlObjects.obj_type(key) + '.csv'
            listcsvs.append(csvname)
            with open(csvname, 'w', newline='') as csvfile:
                cols = report[key][0].keys()
                writer = csv.DictWriter(csvfile, fieldnames=cols, extrasaction='ignore')
                nooutput = writer.writeheader()
                writer.writerows(report[key])
    return report, listcsvs

def parseniosdb(xmlfile):
    listobjects = []
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

def get_args():  ## Handles the arguments passed to the script from the command line
    # Parse arguments
    usage = ' -c b1config.ini -r {"excel","gsheet"} [ --delimiter {",",";"} ] [ --yaml <yaml file> ] [ --help ]'
    description = 'This script generates different reports from NIOS DB. These are defined by a yaml file'
    par = argparse.ArgumentParser(formatter_class=RawDescriptionHelpFormatter, description=description, add_help=False,usage='%(prog)s' + usage)
    # Required Argument(s)
    required = par.add_argument_group('Required Arguments')
    req_grp = required.add_argument
    req_grp('-c', '--config', action="store", dest="config", help="Path to ini file with multiple settings", required=True)
    req_grp('-r', action="store", dest="report", help="Defines the type of reporting that will be produced", choices=['excel', 'gsheet'], required=True, default='excel')
    # Optional Arguments(s)
    optional = par.add_argument_group('Optional Arguments')
    opt_grp = optional.add_argument
    opt_grp('--delimiter', action="store", dest="csvdelimiter", help="Delimiter used in CSV csvContents file", choices=[',', ';'])
    opt_grp('--yaml', action="store", help="Alternate yaml file for supported objects", default='objects.yaml')
    opt_grp('-h', '--help', action='help', help='show this help message and exit')
    return par.parse_args(args=None if sys.argv[1:] else ['-h'])

def main():
    listobjects = []
    report = defaultdict()
    args = get_args()
    conf = read_niosdb_ini(args.config)
    yamlObjects = dblib.DBCONFIG(conf['yaml'])
    listobjects = parseniosdb(conf['dbfile'])
    report, listcsvs = processreports(listobjects, yamlObjects)
    if args.report == 'excel':
        export2excel(listcsvs)
    elif args.report == 'gsheet':
        export2gsheet(listcsvs, conf)
        
### Main ###
if __name__ == '__main__':
    main()
    sys.exit()
