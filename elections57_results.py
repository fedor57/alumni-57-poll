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

import collections
import httplib2
import os
import sys
import bcrypt

from apiclient import discovery
from oauth2client import client
from oauth2client import tools
from oauth2client.file import Storage

import argparse
argparser = argparse.ArgumentParser(description='Show election results', parents=[tools.argparser])
argparser.add_argument('operation', nargs='?',
    help='Operation to perform, show results by default',
    choices=['results', 'dump', 'year_stats'], default='results',)
argparser.add_argument('secret', nargs='?', help='Secret for bcrypt hashing')
cmdline_args = argparser.parse_args()

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
        credentials = tools.run_flow(flow, store, cmdline_args)
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

# Symbolic names for spreadsheet columns:
col_timestamp, col_has_code, col_code, col_name, col_class, \
col_fbpage, col_vkpage, col_email, col_sub_news, col_pub_dir, \
col_candidates, col_bylaws, col_comment = range(13)

def get_dedup_data():
    """Returns a dict containing the voter ID as keys and the final values
    of the vote for this voter as values and total number of voters (not
    necessarily unique).
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

    dedup_data = {}
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
                continue

            # Make up an ID from the class and the name.
            voter_id = row[col_class].upper() + '-'
            if row[col_name]:
                voter_id += row[col_name]
            elif row[col_email]:
                voter_id += row[col_email]
            else:
                print >> sys.stderr,\
                    'Skipping row {}: no voter identification'.format(row_num)
                continue

        dedup_data[voter_id] = row
    return dedup_data, row_num

data, total_votes = get_dedup_data()

if cmdline_args.operation == 'results':
    candidates = {}
    for v in data:
        candidates[v] = data[v][col_candidates]
    all_candidates = [name for c in candidates.values() for name in c.split(', ')]

    print 'Результаты голосования {} уникальных избирателей ({} всего).\n'.\
        format(len(data), total_votes)

    print 'Кандидаты в убывающем порядке голосов:\n'
    counts = collections.Counter(all_candidates)

    print 'Голосов  Кандидат'
    print '------------------------------------------------------------------------'

    for c in counts.most_common():
        print u'{:>7}  {}'.format(c[1], c[0])

    print '\nВсего: {} голосов.'.format(len(all_candidates))
elif cmdline_args.operation == 'year_stats':
    def get_year(v):
        return int(v[0:4])

    def get_class(v):
        return v[4]

    classes = []
    votes_by_year = collections.defaultdict(lambda: collections.defaultdict(int))
    for v in data.keys():
        y, c = get_year(v), get_class(v)
        if c not in classes:
            classes.append(c)

        votes_by_year[y][c] += 1

    classes = sorted(classes)
    min_year = min(votes_by_year.keys())
    max_year = max(votes_by_year.keys())
    print 'x,' + ','.join(classes)
    for y in range(min_year, max_year + 1):
        votes = votes_by_year[y]
        votes_by_class = []
        for c in sorted(classes):
            votes_by_class.append(votes[c])
        print '{},{}'.format(y, ','.join([str(x) for x in votes_by_class]))
elif cmdline_args.operation == 'dump':
    from datetime import datetime

    def parse_timestamp(s):
        return datetime.strptime(s, '%Y-%m-%d %H:%M:%S')

    # See http://stackoverflow.com/questions/6832445/how-can-bcrypt-have-built-in-salts
    salt = bcrypt.gensalt()

    # Dump in vote order, i.e. sort by timestamp.
    for d in sorted(data.items(), key=lambda x: parse_timestamp(x[1][col_timestamp])):
        print u'{}\t{}\t{}'.format(
                bcrypt.hashpw(d[0] + cmdline_args.secret, salt), 
                # Dump it in ISO format instead of C one used in raw data.
                parse_timestamp(d[1][col_timestamp]),
                # Omit columns containing information identifying the voter.
                '\t'.join(d[1][col_candidates:col_bylaws])
            )
else:
    print >> sys.stderr,\
        'Unknown operation "{}", see help.'.format(cmdline_args.operation)
    sys.exit(1)
