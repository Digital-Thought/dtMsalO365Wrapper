from office365.graph_client import GraphClient
from office365.directory.users.collection import UserCollection

from dtMsalO365Wrapper._token_auth_session import TokenAuthSession
from dtMsalO365Wrapper.users.user import User
from office365.runtime.paths.resource_path import ResourcePath
from dtMsalO365Wrapper.teams.team import Team
import logging
import datetime

class Teams:

    def __init__(self, graph_client: GraphClient, token_auth_session: TokenAuthSession, power_automate):
        self._graph_client = graph_client
        self._token_auth_session = token_auth_session
        self._power_automate = power_automate

    def get_joined_teams(self, user: User = None):
        if user is None:
            t = self._graph_client.me.joined_teams.get().paged().execute_query()
        else:
            t = user._user.joined_teams.get().paged().execute_query()

        return [Team(i, self._graph_client, self._token_auth_session, self._power_automate) for i in t]

    def get_all(self):
        t = self._graph_client.teams.get_all().paged().execute_query()
        _l = []
        for team in t:
            if team.id:
                _l.append(Team(team, self._graph_client, self._token_auth_session, self._power_automate))
        return _l

    def get_by_query(self, query: str):
        for t in self._graph_client.teams.filter(query).get().execute_query():
            yield Team(t, self._graph_client, self._token_auth_session, self._power_automate)