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