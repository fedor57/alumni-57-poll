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

import cPickle
import requests

class code_cache:
    # Just hard-code it for now.
    cache_file_name = '57codes.cache'

    code_to_voter = {}

    """Allows to lookup a voter by their code and caches them in a file to
    speed up the next execution of the script.
    """

    def __init__(self):
        """Try to load the cached from file."""
        try:
            with open(self.cache_file_name, 'rb') as f:
                self.code_to_voter = cPickle.load(f)
        except IOError:
            # Ignore the absence of the cache.
            pass

    def __del__(self):
        """Save the cache when this object is destroyed."""
        if self.code_to_voter:
            try:
                with open(self.cache_file_name, 'wb') as f:
                    cPickle.dump(self.code_to_voter, f)
            except IOError as e:
                print >> sys.stderr, 'Failed to save code cache:', e

    def get_voter_id(self, code):
        """Returns the voter id of the form "<class>-<name>" for the given
        code.

        Throws if the code is invalid.
        """
        if code in self.code_to_voter:
            voter_id = self.code_to_voter[code]
        else:
            voter_id = self._get_id_from_code(code)
            self.code_to_voter[code] = voter_id
        return voter_id

    def _get_id_from_code(self, code):
        res = requests.post('http://auth.alumni57.ru/api/v1/check_code', data={'code': code})
        res.raise_for_status()
        d = res.json()
        if d['status'] != 'ok':
            raise Exception('Bad code {}: status: {}'.format(code, d['status']))
        return u'{}{}-{}'.format(d['year'], d['letter'], d['full_name'])

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
    cc = code_cache()
    row_num = 0
    for row in data:
        row_num += 1

        # If we have the unique code, obtain the voter ID from it.
        code = row[col_code]
        if code:
            # Code should normally be in "temp_code (masked_perm_code)" form,
            # but accept not masked permanent codes too, just in case we have
            # any bugs preventing the code replacing them with temporary codes
            # from running again.
            end_temp_code = code.find(' (')
            if end_temp_code != -1:
                code = code[0:end_temp_code]
            voter_id = cc.get_voter_id(code)
        else:
            if not row[col_class]:
                print >> sys.stderr,\
                    'Skipping row {}: no class specified'.format(row_num)
                next

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
