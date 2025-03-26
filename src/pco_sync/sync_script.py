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
