from office365.graph_client import GraphClient
from office365.directory.users.collection import UserCollection
from dtMsalO365Wrapper._token_auth_session import TokenAuthSession
from dtMsalO365Wrapper.messages.message import Message

import logging

class Messages:

    def __init__(self, graph_client: GraphClient, token_auth_session: TokenAuthSession):
        self._graph_client = graph_client
        self._token_auth_session = token_auth_session

    def get_message(self, user, message_id):
        resp = self._token_auth_session.request("GET", f"/users/{user.id}/messages/{message_id}")
        if resp.status_code != 200:
            logging.error(f'Failed to get Message: {resp.content}')
            raise Exception(f'Failed to get Message: {resp.content}')

        return Message(self._graph_client, self._token_auth_session, user, resp.json())