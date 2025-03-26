import os
import requests
from datetime import datetime, timedelta
import schedule
import time
from dotenv import load_dotenv
from msal import ConfidentialClientApplication

# load environment variables
load_dotenv()

# Config
PCO_APP_ID = os.getenv('PCO_APP_ID')
PCO_SECRET = os.getenv('PCO_SECRET')
MICROSOFT_TENANT_ID = os.getenv('MICROSOFT_TENANT_ID')
MICROSOFT_CLIENT_ID = os.getenv('MICROSOFT_CLIENT_ID')
MICROSOFT_CLIENT_SECRET = os.getenv('MICROSOFT_CLIENT_SECRET')
SHARED_CALENDER_ID = os.getenv('SHARED_CALENDER_ID')
SYNC_INTERVAL_MINUTES = 60

class CalenderSync:
    def __init__(self):
        self.pco_auth = (PCO_APP_ID, PCO_SECRET)
        self.graph_token = self._get_microsoft_token()
        self.headers = {
            'Authorization': f'Bearer {self.graph.token}',
            'Content-Type': 'application/json'
        }
    def _get_microsoft_token(self):
        authority = f'https://login.microsoftonline.com/{MICROSOFT_TENANT_ID}'
        app = ConfidentialClientApplication(
            client_id=MICROSOFT_CLIENT_ID,
            client_credential=MICROSOFT_CLIENT_SECRET,
            authority=authority
        )
        result = app.acquire_token_for_client(scopes=['https://graph.microsoft.com/.default'])
        return result['access_token']
    
    def _get_pco_operator_events(self):
        url = 'https://api.planningcenteronline.com/services/v2/service_types/.../plans'
        params = {
            'include': 'team_members',
            'where[team_name]': 'Operator'
            'per_page': 100
        }

        response = requests.get(url, auth=self.pco_auth, params=params)
        events = []

        for plan in response.json().get('data', []):
            for member in plan.get('relationships', {}).get('team_members', {}).get('data', []):
                event = {
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
                        'content': f"Service: {plan['attributes']['title']}\nPerson: {member['attributes']['name']}"
                    }
                }
                events.append(event)
            return events
        
        