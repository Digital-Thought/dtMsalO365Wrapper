from office365.graph_client import GraphClient
from office365.teams.team import Team as O365Team

from dtMsalO365Wrapper._token_auth_session import TokenAuthSession
from dtMsalO365Wrapper.users.user import User
from office365.runtime.paths.resource_path import ResourcePath
import logging
import datetime

class Channel:

    def __init__(self, channel_detail, team, graph_client: GraphClient, token_auth_session: TokenAuthSession, power_automate):
        self._graph_client = graph_client
        self._token_auth_session = token_auth_session
        self._team = team
        self._power_automate = power_automate
        self._channel_detail = channel_detail

    @property
    def display_name(self):
        return self._channel_detail.get('displayName')

    @property
    def description(self):
        return self._channel_detail.get('description')

    @property
    def id(self):
        return self._channel_detail.get('id')

    @property
    def created(self):
        return datetime.datetime.fromisoformat(self._channel_detail.get('createdDateTime'))

    @property
    def email(self):
        return self._channel_detail.get('email')

    @property
    def url(self):
        return self._channel_detail.get('webUrl')

    @property
    def membership_type(self):
        return self._channel_detail.get('membershipType')

    @property
    def is_archived(self):
        return self._channel_detail.get('isArchived')

    def send_message(self, power_automate_teams_webhook_url: str, message: str):
        self._power_automate.send_teams_message(self._team, self, power_automate_teams_webhook_url, message)