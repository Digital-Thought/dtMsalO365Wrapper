from office365.graph_client import GraphClient
from office365.teams.team import Team as O365Team

from dtMsalO365Wrapper._token_auth_session import TokenAuthSession
from dtMsalO365Wrapper.teams.channel import Channel
from office365.outlook.mail.item_body import ItemBody
import logging
import datetime

class Team:

    def __init__(self, team_detail: O365Team, graph_client: GraphClient, token_auth_session: TokenAuthSession, power_automate):
        self._graph_client = graph_client
        self._token_auth_session = token_auth_session
        self._team_detail = team_detail
        self._power_automate = power_automate

    @property
    def display_name(self):
        return self._team_detail.display_name

    @property
    def description(self):
        return self._team_detail.description

    @property
    def id(self):
        return self._team_detail.id

    def get_channels(self):
        resp = self._token_auth_session.request('GET', f'/teams/{self._team_detail.id}/allChannels')
        if resp.status_code != 200:
            raise RuntimeError(f'{resp.text} (Team ID: {self._team_detail.id})')

        return [Channel(i, self, self._graph_client, self._token_auth_session, self._power_automate) for i in resp.json()['value']]


