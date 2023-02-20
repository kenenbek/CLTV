# Copyright 2018 Google LLC
#
# Licensed under the Apache License, Version 2.0 (the "License");
# you may not use this file except in compliance with the License.
# You may obtain a copy of the License at
#
#     http://www.apache.org/licenses/LICENSE-2.0
#
# Unless required by applicable law or agreed to in writing, software
# distributed under the License is distributed on an "AS IS" BASIS,
# WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
# See the License for the specific language governing permissions and
# limitations under the License.

# [START sheets_quickstart]
from __future__ import print_function

import os.path

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# If modifying these scopes, delete the file token.json.
SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly']

# The ID and range of a sample spreadsheet.
from typing import List, Dict
from collections import defaultdict

import locale
locale.setlocale(locale.LC_ALL, 'ru_RU.UTF-8')

class SpreadSheet:
    def __init__(self):
        creds = None
        # The file token.json stores the user's access and refresh tokens, and is
        # created automatically when the authorization flow completes for the first
        # time.
        if os.path.exists('token.json'):
            creds = Credentials.from_authorized_user_file('token.json', SCOPES)
        # If there are no (valid) credentials available, let the user log in.
        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                creds.refresh(Request())
            else:
                flow = InstalledAppFlow.from_client_secrets_file(
                    'credentials.json', SCOPES)
                creds = flow.run_local_server(port=0)
            # Save the credentials for the next run
            with open('token.json', 'w') as token:
                token.write(creds.to_json())
        sheet = None
        try:
            service = build('sheets', 'v4', credentials=creds)

            # Call the Sheets API
            sheet = service.spreadsheets()

        except HttpError as err:
            print(err)

        self.sheet = sheet

    def make_request(self, spreadsheetId, range_name):
        result = self.sheet.values().get(spreadsheetId=spreadsheetId,
                                         range=range_name).execute()
        values = result.get('values', [])
        return values

    def make_request_two_column(self, spreadsheetId, range1, range2):
        two_columns = {}
        result = self.sheet.values().batchGet(spreadsheetId=spreadsheetId,
                                              ranges=[range1, range2]).execute()
        values = result.get('valueRanges', [])

        min_len = min(len(values[0]['values']), len(values[1]['values']))
        for i in range(min_len):
            if len(values[0]['values'][i]) == 0 or len(values[1]['values'][i]) == 0:
                continue
            two_columns[values[0]['values'][i][0].strip().lower()] = values[1]['values'][i][0]

        return two_columns


class Metrics:
    def __init__(self):
        self.clients_duration = defaultdict(int)  # client - amount of month duration
        self.company_people_amount_dict = defaultdict(int)
        self.client_incomes = defaultdict(list)

    def clean_column(self, values: List[str]):
        new_values = []
        for value in values:
            new_value = value[0].strip().lower()
            new_values.append(new_value)
        return new_values

    def clean_money_strings(self, moneys: List[str]):
        new_values = []
        for value in moneys:
            new_value = int(value[0].strip().lower().replace(u'\xa0', u''))
            new_values.append(new_value)
        return new_values

    def remove_empty_lines(self, clients: List[str], moneys: List[str]):
        new_clients = []
        new_moneys = []

        for client, money in zip(clients, moneys):
            if len(client) == 0 or len(money) == 0:
                continue
            new_clients.append(client)
            new_moneys.append(money)

        return new_clients, new_moneys

    def update_client_duration(self, client_money_one_month: Dict[str, int]):
        for client in client_money_one_month:
            self.clients_duration[client] += 1

    def update_income_from_clients(self, client_money_one_month: Dict[str, int]):
        for client, money in client_money_one_month.items():
            self.client_incomes[client].append(int(money))

    def sort_duration_dict(self):
        self.clients_duration = dict(sorted(self.clients_duration.items(), key=lambda x: x[1], reverse=True))

    def sort_clients_duration_money_dict(self):
        self.client_incomes = dict(sorted(self.client_incomes.items(), key=lambda x: len(x[1]), reverse=True))

    def clean_company_people_amount(self, company_people_amount_dict: Dict):
        for k, v in company_people_amount_dict.items():
            company_people_amount_dict[k.lower().strip()] = int(v)
        self.company_people_amount_dict = company_people_amount_dict
        return company_people_amount_dict

    def clean_client_money_one_month(self, client_money_one_month: Dict):
        for k, v in client_money_one_month.items():
            client_money_one_month[k] = locale.atof(v.strip().lower().replace(u'\xa0', u''))
        return client_money_one_month
