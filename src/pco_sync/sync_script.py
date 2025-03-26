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

        while url:
            response = requests.get(url, headers=self.headers, params=params)
            for event in response.json().get('value', []):
                for prop in event.get('singleValueExtendedProperties', []):
                    if prop['id'] == 'String {66f5a359-4659-4830-9070-000000000000} Name PCO_ID':
                        events[prop['value']] = event['id']
            url = response.json().get('@odata.nextLink')
        return events
    
    def _sync_events(self, pco_events):
        # perform full sync with create/update/delete operations
        current_pco_ids = set()
        new_count = 0
        update_count = 0
        delete_count = 0

        # process PCO events
        for event in pco_events:
            pco_id = next(p['value'] for p in event['singleValueExtendedProperties'])
            current_pco_ids.add(pco_id)

            if pco_id in self.existing_events:
                if self._needs_update(self.existing_events[pco_id], event):
                    self._update_event(self.existing_events[pco_id], event)
                    update_count += 1
            else:
                # create new event
                self._create_event(event)
                new_count += 1

        # delete removed events
        for pco_id in set(self.existing_events.keys()) - current_pco_ids:
            self._delete_event(self.existing_events[pco_id])
            delete_count += 1

        print(f"Created {new_count}, updated {update_count}, deleted {delete_count} events")

    def _needs_update(self, event_id, new_event):
        # check if an event needs updating
        existing_event = requests.get(
            f'https://graph.microsoft.com/v1.0/users/{self.calender_id}/events/{event_id}', headers=self.headers
        ).json()

        return (
            existing_event['start']['dateTime'] != new_event['start']['dateTime']
            or existing_event['end']['dateTime'] != new_event['end']['dateTime']
            or existing_event['subject'] != new_event['subject']
        )
    
    def _create_event(self, event):
        # create new Outlook event
        response = requests.post(
            f'https://graph.microsoft.com/v1.0/users/{self.calender_id}/events',
            headers=self.headers,
            json=event
        )
        if response.status_code in [200, 201]:
            self.existing_events[
                next(p['value'] for p in event['singleValueExtendedProperties'])
            ] = response.json()['id']

    def _update_event(self, event_id, new_data):
        # update existing outlook event
        requests.patch(
            f'https://graph.microsoft.com/v1.0/users/{self.calender_id}/events/{event_id}',
            headers=self.headers,
            json={
                'start': new_data['start'],
                'end': new_data['end'],
                'subject': new_data['subject'],
                'body': new_data.get('body', {})
            }
        )
    
    def _delete_event(self, event_id):
        # delete Outlook event
        response = requests.delete(
            f'https://graph.microsoft.com/v1.0/users/{self.calender_id}/events/{event_id}'
            headers=self.headers
        )
        if response.status_code == 204:
            del self.existing_events[pco_id]

    def sync(self):
        # main sync operations
        print(f"Starting sync at {datetime.now()}")
        try:
            pco_events = self._get_pco_events()
            self._sync_events(pco_events)
            print(f"Sync complete. Total PCO events: {len(pco_events)}")
        except Exception as e:
            print(f"Sync failed: {str(e)}")
        finally:
            self.existing_events = self._get_existing_outlook_events()

    def start_scheduler(self):
        # start scheduled syncs
        schedule.every(int(os.getenv('SYNC_INTERVAL_MINUTES', 60))).minutes.do(self.sync)
        while True:
            schedule.run_pending()
            time.sleep(1)

if __name__ == '__main__':
    syncer = CalenderSync()
    syncer.sync()
    syncer.start_scheduler()