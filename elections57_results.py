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
    """Returns a dict containing the voter ID as keys and the list of
    the candidates voted for as values.

    Each returned value is a comma-separated list of candidates.
    """
    credentials = get_credentials()
    http = credentials.authorize(httplib2.Http())
    discoveryUrl = ('https://sheets.googleapis.com/$discovery/rest?version=v4')
    service = discovery.build('sheets', 'v4', http=http, discoveryServiceUrl=discoveryUrl)

    values = service.spreadsheets().values()

    spreadsheetId = '1KdLqYCIAo4bOcFfPBRkL5OyzIdeN2hERlgX0ZV_e--w'
    rangeName = 'A2:M9999'
    result = values.get(spreadsheetId=spreadsheetId, range=rangeName).execute()
    data = result.get('values', [])

    if not data:
        raise Exception('Failed to retrieve data.')

    # Symbolic names for spreadsheet columns:
    col_timestamp, col_has_code, col_code, col_name, col_class, \
    col_fbpage, col_vkpage, col_email, col_sub_news, col_pub_dir, \
    col_candidates, col_bylaws, col_comment = range(13)

    candidates = {}
    row_num = 0
    for row in data:
        row_num += 1

        # If we have the unique code, use its class and name parts as voter ID.
        code = row[col_code]
        if code:
            # The code format is "57-1234v-wyxz-0123456789abcdef" except that
            # there can be 5 or even 6 letters depending on transliteration
            # vagaries, so we can't rely on its index.
            start = 3
            end = code.find('-', start + 9)
            if end == -1:
                print >> sys.stderr,\
                    'Skipping row {}: invalid code format "{}"'.format(row_num, code)
                next

            voter_id = code[start:end]
        else:
            # Make up an ID from the class and the name.
            voter_id = row[col_class] + '-'
            if row[col_name]:
                voter_id += row[col_name]
            elif row[col_email]:
                voter_id += row[col_email]
            else:
                print >> sys.stderr,\
                    'Skipping row {}: no voter identification'.format(row_num)
                next

        candidates[voter_id] = row[col_candidates]
    return candidates


candidates = get_raw_candidates()
all_candidates = [name for c in candidates.values() for name in c.split(', ')]

print 'Результаты голосования {} выборщиков.\n'.format(len(candidates))

print 'Кандидаты в убывающем порядке голосов:\n'
import collections
counts = collections.Counter(all_candidates)

print 'Голосов  Кандидат'
print '------------------------------------------------------------------------'

for c in counts.most_common():
    print u'{:>7}  {}'.format(c[1], c[0])

print '\nВсего: {} голосов.'.format(len(all_candidates))
