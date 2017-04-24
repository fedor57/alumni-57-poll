#!/usr/bin/env python
# vim: set fileencoding=utf-8 :

# This file is based on Google Sheets API Python Quickstart example at
# https://developers.google.com/sheets/api/quickstart/python see there for
# more explanations and notably about how to create the client_secret.json
# file.
#
# Also consider running the script with --noauth_local_webserver command line
# argument to perform authentication in another browser (possibly on another
# machine).

import httplib2
import os
import sys

from apiclient import discovery
from oauth2client import client
from oauth2client import tools
from oauth2client.file import Storage

import argparse
flags = argparse.ArgumentParser(parents=[tools.argparser]).parse_args()

# If modifying these scopes, delete your previously saved credentials
SCOPES = 'https://www.googleapis.com/auth/spreadsheets.readonly'
APPLICATION_NAME = 'Vote Tally for School 57 Council Elections'

home_dir = os.path.expanduser('~')
google_sheet_api_dir = os.path.join(home_dir, '.google_sheet_api')

def get_credentials():
    """Gets valid user credentials from storage.

    If nothing has been stored, or if the stored credentials are invalid,
    the OAuth2 flow is completed to obtain the new credentials.

    Returns:
        Credentials, the obtained credential.
    """
    if not os.path.exists(google_sheet_api_dir):
        print 'Пожалуйста создайте директорию "{}" с файлом "client_secret.json".'.format(google_sheet_api_dir)
        sys.exit(1)

    credential_path = os.path.join(google_sheet_api_dir, 'info.alumni57-elections201704')

    store = Storage(credential_path)
    credentials = store.get()
    if not credentials or credentials.invalid:
        flow = client.flow_from_clientsecrets(os.path.join(google_sheet_api_dir, 'client_secret.json'), SCOPES)
        flow.user_agent = APPLICATION_NAME
        credentials = tools.run_flow(flow, store, flags)
        print 'Storing credentials to ' + credential_path
    return credentials

def get_raw_candidates():
    credentials = get_credentials()
    http = credentials.authorize(httplib2.Http())
    discoveryUrl = ('https://sheets.googleapis.com/$discovery/rest?version=v4')
    service = discovery.build('sheets', 'v4', http=http, discoveryServiceUrl=discoveryUrl)

    values = service.spreadsheets().values()

    spreadsheetId = '1KdLqYCIAo4bOcFfPBRkL5OyzIdeN2hERlgX0ZV_e--w'
    rangeName = 'K2:K9999'
    result = values.get(spreadsheetId=spreadsheetId, range=rangeName).execute()
    data = result.get('values', [])

    if not data:
        raise 'Failed to retrieve data.'

    candidates = []
    for row in data:
        names = row[0]
        candidates = candidates + names.split(', ')
    return candidates


candidates = get_raw_candidates()

import collections
counts = collections.Counter(candidates)

print 'Кандидаты в убывающем порядке голосов:\n'
print 'Голосов  Кандидат'
print '------------------------------------------------------------------------'

for c in counts.most_common():
    print '{:>7}  {}'.format(c[1], c[0].encode('utf-8'))
