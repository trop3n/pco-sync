import os
import requests
from datetime import datetime
import schedule
import time
from dotenv import load_dotenv
from msal import ConfidentialClientApplication

# Load environment variables
load_dotenv()

PCO_APP_ID = os.getenv('PCO_APP_ID')
PCO_SECRET = os.getenv('PCO_SECRET')
MICROSOFT_TENANT_ID = os.getenv('MICROSOFT_TENANT_ID')
MICROSOFT_CLIENT_ID = os.getenv('MICROSOFT_CLIENT_ID')
MICROSOFT_CLIENT_SECRET = os.getenv('MICROSOFT_CLIENT_SECRET')
SHARED_CALENDAR_ID = os.getenv('SHARED_CALENDAR_ID')
SYNC_INTERVAL_MINUTES = 60

class CalenderSync:
    def __init__(self):
        self.pco_auth = (os.getenv('PCO_APP_ID'), os.getenv('PCO_SECRET'))
        self.graph_token = self._get_microsoft_token()
        self.headers = {
            'Authorization': f'Bearer {self.graph_token}',
            'Content-Type': 'application/json'
        }

        self.calender_id = os.getenv('SHARED_CALENDER_ID')
        self.existing_events = self._get_existing_outlook_events()

    def _get_microsoft_token(self):
        # get Microsoft Graph API token using client credentials
        authority = f'https://login.microsoftonline.com/{os.getenv(MICROSOFT_TENANT_ID)}'
        app = ConfidentialClientApplication(
            client_id=os.getenv("MICROSOFT_CLIENT_ID"),
            client_credential=os.getenv('MICROSOFT_CLIENT_SECRET'),
            authority=authority
        )
        return app.acquire_token_for_client(scopes=['https://graph.microsoft.com/.default'])['access_token']
    
    def _get_pco_events(self):
        # Retrieve operator events from Planning Center Online
        url = 'https://api.planningcenteronline.com/services/v2/service_types/.../plans'
        params = {
            'include': 'team_members',
            'where[team_name]': 'Operator',
            'per_page': 100
        }

        events = []
        response = requests.get(url, auth=self.pco_auth, params=params)

        for plan in response.json().get('data', []):
            for member in plan.get('relationships', {}).get('team_members', {}).get('data', []):
                events.append({
                    'subject': f"Operator Shift: {member.get('attributes', {}).get('name')}",
                    'start': {
                        'dateTime': plan['attributes']['starts_at'],
                        'timeZone': 'America/Chicago'
                    },
                    'end': {
                        'dateTime': plan['attributes']['ends_at'],
                        'timeZone': 'America/Chicago'
                    },
                    'body': {
                        'content': f"Service {plan['attributes']['title']}\nPerson: {member['attributes']['name']}"
                    },
                    'singleValueExtendedProperties': [{
                        'id': 'String {66f5a359-4659-4830-9070-000000000000} Name PCO_ID',
                        'value': f"{plan['id']}|{member['id']}"
                    }]
                })
            return events
        
        def _get_existing_outlook_events(self):
            # retrieve existing outlook events with PCO identifiers
            events = {}
            url = f'https://graph.microsoft.com/v1.0/users/{self.calender_id}/events'
            params = {
                '$select': 'id',
                '$expand': 'singleValueExtendedProperties($filter=id eq \'String {66f5a359-4659-4830-9070-000000000000} Name PCO_ID\')'
            }