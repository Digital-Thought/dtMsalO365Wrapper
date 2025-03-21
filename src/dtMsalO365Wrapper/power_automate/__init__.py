from office365.graph_client import GraphClient
from office365.teams.team import Team as O365Team

from dtMsalO365Wrapper._token_auth_session import TokenAuthSession
from dtMsalO365Wrapper.teams.channel import Channel
from dtMsalO365Wrapper.teams import Team

class PowerAutomate:

    def __init__(self, power_automate_token_auth_session: TokenAuthSession):
        self._power_automate_token_auth_session = power_automate_token_auth_session

    def send_teams_message(self, team: Team, channel: Channel, power_automate_teams_webhook_url: str, message: str):
        payload = {
            "message": message,
            "team_id": team.id,
            "channel_id": channel.id
        }
        resp = self._power_automate_token_auth_session.request('POST', power_automate_teams_webhook_url,
                                                        json=payload)

        if resp.status_code != 202:
            raise Exception(f'Failed to send Teams Message to Power Automate, Team Webhook: {power_automate_teams_webhook_url})')
