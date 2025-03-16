from office365.graph_client import GraphClient
from office365.directory.users.collection import UserCollection

from dtMsalO365Wrapper._token_auth_session import TokenAuthSession
from dtMsalO365Wrapper.users.user import User
from office365.runtime.paths.resource_path import ResourcePath
import logging
import datetime

class Subscriptions:

    def __init__(self, graph_client: GraphClient, token_auth_session: TokenAuthSession):
        self._graph_client = graph_client
        self._token_auth_session = token_auth_session
        self._graph_client.subscriptions.get_all().execute_query()
        self._subscriptions = self._graph_client.subscriptions

    def add_subscription(self, resource: str, notification_url: str, change_type: str,
                         expiration_date_time: datetime.datetime):
        subscription_payload = {
            "changeType": change_type,
            "notificationUrl": notification_url,
            "resource": resource,
            "expirationDateTime": expiration_date_time.isoformat()
        }

        resp = self._token_auth_session.request("POST", "/subscriptions", json=subscription_payload)
        if resp.status_code != 201:
            logging.error(f'Failed to add Subscription: {resp.content}')
            raise Exception(f'Failed to add Subscription: {resp.content}')

        return resp.json()

    def add_messages_subscription(self, user: User, notification_url: str, change_type: str ='created',
                                  expiration_date_time = datetime.datetime.now(datetime.UTC) + datetime.timedelta(hours=24)):
        resource = f"users/{user.id}/messages"
        return self.add_subscription(resource, notification_url, change_type, expiration_date_time)

    def update_subscription(self, subscription_id: str, notification_url: str = None,
                                  expiration_date_time = datetime.datetime.now(datetime.UTC) + datetime.timedelta(hours=24)):
        subscription_payload = {}
        if notification_url:
            subscription_payload["notificationUrl"] = notification_url
        if expiration_date_time:
            subscription_payload["expirationDateTime"] = expiration_date_time.isoformat()

        resp = self._token_auth_session.request("PATCH", f"/subscriptions/{subscription_id}",
                                                json=subscription_payload)
        if resp.status_code != 200:
            logging.error(f'Failed to update Subscription: {resp.content}')
            raise Exception(f'Failed to update Subscription: {resp.content}')

        return resp.json()

    def delete_subscription(self, subscription_id: str):
        resp = self._token_auth_session.request("DELETE", f"/subscriptions/{subscription_id}")
        if resp.status_code != 204:
            logging.error(f'Failed to delete Subscription: {resp.content}')
            raise Exception(f'Failed to delete Subscription: {resp.content}')

        return resp.json()