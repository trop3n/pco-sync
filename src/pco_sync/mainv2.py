import os
import requests
from datetime import datetime
import schedule
import time
from dotenv import load_dotenv
from msal import ConfidentialClientApplication

load_dotenv()

PCO_APP_ID = os.getenv('PCO_APP_ID')
PCO_SECRET = os.getenv('PCO_SECRET')
MICROSOFT_TENANT_ID = os.getenv('MICROSOFT_TENANT_ID')
MICROSOFT_CLIENT_ID = os.getenv('MICROSOFT_CLIENT_ID')
MICROSOFT_CLIENT_SECRET = os.getenv('MICROSOFT_CLIENT_SECRET')
SHARED_CALENDAR_ID = os.getenv('SHARED_CALENDAR_ID')
SYNC_INTERVAL_MINUTES = 60

class CalendarSync:
    def __init__(self):
        self.pco_auth = (os.getenv('PCO_APP_ID'), os.getenv('PCO_SECRET'))
        self.graph_token = self._get_microsoft_token()
        self.headers = {
            'Authorization': f'Bearer {self.graph_token}',
            'Content-Type': 'application/json'
        }
        self.calendar_id = os.getenv('SHARED_CALENDAR_ID')
        
        # Store existing Outlook events with PCO identifiers
        self.existing_events = self._get_existing_outlook_events()

    def _get_microsoft_token(self):
        # ... (same as previous token retrieval) ...

    def _get_pco_events(self):
        # ... (same PCO API call as before) ...
        # Add PCO identifiers to events
        for plan in response.json().get('data', []):
            for member in plan.get('relationships', {}).get('team_members', {}).get('data', []):
                event = {
                    # Add PCO identifiers as extended properties
                    'singleValueExtendedProperties': [{
                        'id': 'String {66f5a359-4659-4830-9070-000000000000} Name PCO_ID',
                        'value': f"{plan['id']}|{member['id']}"
                    }],
                    # ... rest of event data ...
                }
                events.append(event)
        return events

    def _get_existing_outlook_events(self):
        """Get existing events with PCO identifiers"""
        events = {}
        url = f'https://graph.microsoft.com/v1.0/users/{self.calendar_id}/events'
        params = {
            '$select': 'id,subject,start,end',
            '$expand': 'singleValueExtendedProperties($filter=id eq \'String {66f5a359-4659-4830-9070-000000000000} Name PCO_ID\')'
        }

        while url:
            response = requests.get(url, headers=self.headers, params=params)
            for event in response.json().get('value', []):
                for prop in event.get('singleValueExtendedProperties', []):
                    if prop['id'] == 'String {66f5a359-4659-4830-9070-000000000000} Name PCO_ID':
                        events[prop['value']] = {
                            'event_id': event['id'],
                            'start': event['start']['dateTime'],
                            'end': event['end']['dateTime']
                        }
            url = response.json().get('@odata.nextLink')
        return events

    def _sync_events(self, pco_events):
        """Smart sync with update/delete/create operations"""
        current_pco_ids = set()
        
        for event in pco_events:
            pco_id = next(p['value'] for p in event['singleValueExtendedProperties'])
            current_pco_ids.add(pco_id)

            # Check if event exists
            if pco_id in self.existing_events:
                # Update if changes detected
                existing = self.existing_events[pco_id]
                if (event['start']['dateTime'] != existing['start'] or
                    event['end']['dateTime'] != existing['end']):
                    self._update_event(existing['event_id'], event)
            else:
                # Create new event
                self._create_event(event)

        # Delete removed events
        for pco_id in set(self.existing_events.keys()) - current_pco_ids:
            self._delete_event(self.existing_events[pco_id]['event_id'])

    def _create_event(self, event):
        response = requests.post(
            f'https://graph.microsoft.com/v1.0/users/{self.calendar_id}/events',
            headers=self.headers,
            json=event
        )
        if response.status_code in [200, 201]:
            print(f"Created event: {response.json()['id']}")

    def _update_event(self, event_id, new_data):
        response = requests.patch(
            f'https://graph.microsoft.com/v1.0/users/{self.calendar_id}/events/{event_id}',
            headers=self.headers,
            json={
                'subject': new_data['subject'],
                'start': new_data['start'],
                'end': new_data['end'],
                'body': new_data['body']
            }
        )
        if response.status_code == 200:
            print(f"Updated event: {event_id}")

    def _delete_event(self, event_id):
        response = requests.delete(
            f'https://graph.microsoft.com/v1.0/users/{self.calendar_id}/events/{event_id}',
            headers=self.headers
        )
        if response.status_code == 204:
            print(f"Deleted event: {event_id}")

    def sync(self):
        print(f"Starting sync at {datetime.now()}")
        try:
            pco_events = self._get_pco_events()
            self._sync_events(pco_events)
            print(f"Sync complete. Total events: {len(pco_events)}")
        except Exception as e:
            print(f"Sync failed: {str(e)}")
        finally:
            # Refresh existing events cache
            self.existing_events = self._get_existing_outlook_events()

if __name__ == '__main__':
    syncer = CalendarSync()
    syncer.sync()  # Initial sync
    syncer.start_scheduler()