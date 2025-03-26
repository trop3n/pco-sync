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

class SafeCalenderSync:
    def __init__(self):
        self.pco_auth = (PCO_APP_ID, PCO_SECRET)
        self.graph_token = self._get_microsoft_token()
        self.headers = {
            'Authorization': f'Bearer {self.graph.token}',
            'Content-Type': 'application/json'
        }
        self.calendar_id = os.getenv('SHARED_CALENDER_ID')
        self.existing_events = self._get_existing_outlook_events()

    def _get_microsoft_token(self):
        authority = f'https://login.microsoftonline.com/{MICROSOFT_TENANT_ID}'
        app = ConfidentialClientApplication(
            client_id=MICROSOFT_CLIENT_ID,
            client_credential=MICROSOFT_CLIENT_SECRET,
            authority=authority
        )
        return app.acquire_token_for_client(scopes=['https://graph.microsoft.com/.default'])['access_token']
    
    
    def _get_pco_events(self):
        url = 'https://api.planningcenteronline.com/services/v2/service_types/.../plans'
        params = {
            'include': 'team_members',
            'where[team_name]': 'Operator',
            'per_page': 100
        }

        response = requests.get(url, auth=self.pco_auth, params=params)
        pco_events = []
        for event in pco_events:
            event['singleValueExtendedProperties'] = [{
                'id': 'String {66f5a359-4659-4830-9070-000000000000} Name PCO_ID',
                'value': f"{plan_id}|{member_id}"
            }]
        return pco_events
        
    def _get_existing_outlook_events(self):
        """Get existing events with PCO identifiers"""
        events = {}
        url = f'https://graph.microsoft.com/v1.0/users/{self.calender_id}/events'
        params = {
            '$select': 'id,subject,start,end',
            '$expand': 'singleValueExtendedProperties($filter=id eq \'String {66f5a359-4659-4830-9070-000000000000} Name PCO_ID\')'
        }

        while url:
            response = requests.get(url, headers=self.headers, params=params)
            for event in response.json().get('value', []):
                for prop in event.get('singleValueExtendedProperties, []'):
                    if prop['id'] == 'String {66f5a359-4659-4830-9070-000000000000} Name PCO_ID':
                        events[prop['value']] = event['id']
            url = response.json().get('@odata.nextLink')
        return events
        
    def _sync_events(self, pco_events):
        """Safe sync with update/create operations"""
        new_counts = 0
        update_count = 0

        for event in pco_events:
            pco_id = next(p['value'] for p in event['singleValueExtendedProperties'])
            if pco_id in self.existing_events:
                # update existing event if needed
                if self._needs_update(self.existing_events[pco_id], event):
                    self._update_event(self.existing_events[pco_id], event)
                    update_count += 1
            else:
                # Create new event
                self._create_event(event)
                new_count += 1

        print(f"Added {new_count} new events, updated {update_count} existing events")

    def _needs_update(self, event_id, new_event):
        """Check if existing event needs updating"""
        existing_event = requests.get(
            f'https://graph.microsoft.com/v1.0/users/{self.calender_id}/events/{event_id}',
            headers=self.headers
        ).json()

        return(
            existing_event['start']['dateTime'] != new_event['start']['dateTime'] or
            existing_event['end']['dateTime'] != new_event['end']['dateTime'] or
            existing_event['subject'] != new_event['subject']
        )
    
    def _create_event(self, event):
        response = requests.post(
            f'https://graph.microsoft.com/v1.0/users/{self.calendar_id}/events',
            headers=self.headers,
            json=event
        )
        if response.status_code in [200, 201]:
            self.existing_events[
                next(p['value'] for p in event['singleValueExtendedProperties'])
            ] = response.json()['id']

    def _update_event(self, event_id, new_data):
        requests.patch(
            f'https://graph.microsoft.com/v1.0/users/{self.calendar_id}/events/{event_id}',
            headers=self.headers,
            json={
                'start': new_data['start'],
                'end': new_data['end']
                'subject': new_data['subject'],
                'body': new_data.get('body', {})
            }
        )

    def sync(self):
        print(f"Starting sync at {datetime.now()}")
        try:
            events = self._get_pco_events()
            self._sync_events(pco_events)
            print(f"Safe sync completed successfully")
        except Exception as e:
            print(f"Sync failed: {str(e)}")
        finally:
            # refresh existing events for next sync
            self.existing_events = self._get_existing_outlook_events()

    def start_scheduler(self):
        schedule.every(int(os.getenv('SYNC_INTERVAL_MINUTES', 60))).minutes.do(self.sync)
        while True:
            schedule.run_pending()
            time.sleep(1)

if __name__ == 'main':
    syncer = SafeCalenderSync()
    syncer.sync() # Initial sync
    syncer.start_scheduler()