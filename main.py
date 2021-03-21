from __future__ import print_function
import pickle
import os.path
import argparse
import sys
import json
import logging
import csv
import time
import yaml
from functools import cmp_to_key
from pprint import pformat
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from google.auth.transport.requests import AuthorizedSession

#from consts import *

# If modifying these scopes, delete the file token.pickle.
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
DISCOVERY_DOC = ('https://docs.googleapis.com/$discovery/rest?'
                 'version=v1')

def args_parser():
    parser = argparse.ArgumentParser()
    parser.add_argument('-d', '--document', help='The ID of the spreadsheet into which orders are imported', type=str)
    parser.add_argument('-f', '--file',
                        help='The filename of the CSV containing orders from FundTeam',
                        type=str, required=True)
    return parser

def parse_csv(filename):
    rows = []
    with open(filename, mode='r') as csvfile:
        reader = csv.DictReader(csvfile)
        for row in reader:
            rows.append(row)            
    return rows

def load_credentials():
    log = logging.getLogger('load_credentials')
    creds = None
    # The file token.pickle stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file('credentials.json', SCOPES)
            creds = flow.run_console()
        # Save the credentials for the next run
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)
    return creds

def build_google_api(service, version):
    log = logging.getLogger('build_google_api')

    log.info('getting credentials')
    creds = load_credentials()   
    log.info('building service object') 
    svc = build(service, version, credentials=creds, discoveryServiceUrl=DISCOVERY_DOC)
    return svc

def make_tickets(rows, target_doc):
    log = logging.getLogger('make_tickets')
    log.debug('getting service')
    docs = build_google_api('docs', 'v1')

    log.info('Creating tickets')
    idx = 0
    requests = []
    pageBreakReq = { 'insertPageBreak': { 'location': { 'index': 1 } } } 
    counter = 1

    for r in rows:
        log.info('Creating for ' + r[ORDER_NUMBER])
        requests.append(insert_ticket(r, target_doc, docs))        
        idx = idx + 1

        if idx % 4 == 0:
            requests.append(pageBreakReq)

        log.info('Sending requests')
        log.debug(pformat(requests))
        result = docs.documents().batchUpdate(documentId=target_doc, body={'requests': requests}).execute()
        log.debug(pformat(result))
        requests = []
        log.info("{} / {}".format(counter, len(rows)))
        counter = counter + 1
        log.info('Sleeping for 1 second to avoid quota')
        time.sleep(1)

def row_to_value(row):
    return ([ row['Invoice'], row['Date'], row['Name'], row['Email'], 
    row['Phone'], row['Address'], row['Zip'], row['Gate Code'], row['Validate'], 
    row['Pickup'], row['Steep'], row['Merchandise'], row['1-Bag: Hardwood'], 
    row['Pallet: Hardwood'], row['1-Bag: Black'], row['Pallet: Black'], 
    row['Instructions'] ])

def import_rows(sheet, rows):
    log = logging.getLogger('import_rows')
    log.debug("Getting service")
    sheets = build_google_api('sheets', 'v4')
    values = []
    real_rows = [ x for x in rows if x['Invoice'].startswith('21') ]

    for idx, row in enumerate(real_rows):
        values.append(row_to_value(row))
        if (idx > 0 and idx % 50 == 0) or idx == len(real_rows) - 1:
            log.info(idx)
            resource = { "majorDimension": "ROWS", "values": values }    
            sheets.spreadsheets().values().append(
            spreadsheetId=sheet,
            range="Import!A:A",
            body=resource,
            valueInputOption="USER_ENTERED"
            ).execute()
            values = []

def main():
    with open("config.yml", 'r') as stream:
        config = yaml.safe_load(stream)

    logging.basicConfig(level=logging.INFO)    
    logging.getLogger('googleapiclient').setLevel(logging.ERROR)
    log = logging.getLogger('main')
    parser = args_parser()
    args = parser.parse_args()

    log.debug('Default spreadsheet: ' + config['default_spreadsheet'])
    rows = parse_csv(args.file)
    sheetId = args.document or config['default_spreadsheet']
    import_rows(sheetId, rows)
    # for row in [ x for x in rows if x['Merchandise'] != '0.0' ]:
    #     if row['Pickup'] != '':
    #         log.debug(row['Invoice'] + "," + row['Address'] + "," + row['Instructions'])
    # rows = sort_rows(rows)
    # # for row in rows:
    # #     log.debug("{},{},{}".format(row[ZIP], row[TOTAL_BAGS], row[DELIVERY_ZONE]))
    # make_tickets(rows, args.document)

if __name__ == '__main__':
    main()
