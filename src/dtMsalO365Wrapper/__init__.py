import time
import logging
import datetime
import msal
from azure.identity import CertificateCredential
from office365.graph_client import GraphClient
from dtMsalO365Wrapper._token_auth_session import TokenAuthSession

from dtMsalO365Wrapper.users import Users
from dtMsalO365Wrapper.communications import Communications
from dtMsalO365Wrapper.subscriptions import Subscriptions


class MsalO365Client:
    """
    Represents a client for interacting with Microsoft Graph API using Microsoft's authentication libraries.

    This class is used to handle authentication and provide access to Microsoft Graph services
    (e.g., Sites, Users) through delegated or app-only permissions. It supports both using
    client secrets and certificate-based authentication for secure token acquisition. The class
    also facilitates session management and token renewal to maintain seamless access to
    Microsoft Graph resources.

    :ivar _tenant_id: Azure AD tenant ID for authentication.
    :type _tenant_id: str
    :ivar _client_id: Azure AD application (client) ID.
    :type _client_id: str
    :ivar _client_secret: Client secret used for authenticating the app. None if using certificates.
    :type _client_secret: str | None
    :ivar _certificate_path: Path to the certificate file for authentication. None if using client secrets.
    :type _certificate_path: str | None
    :ivar _certificate_password: Password for the certificate file. None if not required or using client secrets.
    :type _certificate_password: str | None
    :ivar _access_token: Access token used for authentication with Microsoft Graph API. None initially.
    :type _access_token: dict | None
    :ivar _token_expiry: Expiry time of the current access token. None initially.
    :type _token_expiry: datetime.datetime | None
    :ivar graph_client: Instance of GraphClient for interacting with Microsoft Graph API.
    :type graph_client: GraphClient
    :ivar token_auth_session: Instance of TokenAuthSession for token-based API session management.
    :type token_auth_session: TokenAuthSession
    """
    def __init__(self, tenant_id, client_id, client_secret=None, certificate_path=None, certificate_password=None):
        self._tenant_id = tenant_id
        self._client_id = client_id
        self._client_secret = client_secret
        self._certificate_path = certificate_path
        self._certificate_password = certificate_password
        self._access_token = None
        self._token_expiry = None
        self.graph_client = GraphClient(self._acquire_token)
        self.token_auth_session = TokenAuthSession(self._acquire_token)
        root_site = self.graph_client.sites.root.get().execute_query()
        logging.info(f'Successfully Authenticated: {root_site.web_url}')

    @classmethod
    def with_client_id_secret(cls, tenant_id, client_id, client_secret):
        """
        Creates an instance of the class using the provided client ID and client secret. This method
        is a convenience constructor that allows initializing the class with the tenant ID, client
        ID, and client secret.

        :param tenant_id: The identifier representing the tenant.
        :type tenant_id: str
        :param client_id: The client ID used for authentication.
        :type client_id: str
        :param client_secret: The client secret associated with the client ID.
        :type client_secret: str
        :return: A new instance of the class initialized with the provided credentials.
        :rtype: cls
        """
        return cls(tenant_id=tenant_id, client_id=client_id, client_secret=client_secret)

    @classmethod
    def with_client_id_certificate(cls, tenant_id, client_id, certificate_path, certificate_password=None):
        """
        Creates an instance of the class using client ID and certificate-based authentication. This method
        initializes the object with specified credentials needed for connection or processing tasks.

        :param tenant_id: Tenant ID or directory identifier for the associated authentication process.
        :type tenant_id: str
        :param client_id: Client ID representing the application requesting authentication.
        :type client_id: str
        :param certificate_path: File path to the certificate used for authentication.
        :type certificate_path: str
        :param certificate_password: (Optional) Password for securing the certificate used in
            the authentication process. Default is None.
        :type certificate_password: Optional[str]
        :return: An instance of the class initialized with the provided client ID and certificate details.
        :rtype: cls
        """
        return cls(tenant_id=tenant_id, client_id=client_id, certificate_path=certificate_path,
                   certificate_password=certificate_password)

    def _acquire_token(self):
        """
        Acquires and returns an access token for authentication with Microsoft Graph API. If an access
        token is already available and valid, it reuses the token; otherwise, it retrieves a new token
        using either client secret credentials or a certificate. The token includes expiry information
        that helps determine token validity for future requests.

        :raises ValueError: If no valid credential (client secret or certificate) is available
            for authentication.

        :raises RuntimeError: If the token acquisition process fails due to invalid configuration,
            authentication failure or network-related issues.

        :return: An authentication token dictionary containing elements such as 'access_token',
            'expires_in', 'token_type', 'ext_expires_in', and 'token_source'.
        :rtype: dict
        """
        if self._access_token is None or datetime.datetime.now() > self._token_expiry:
            authority_url = "https://login.microsoftonline.com/{0}".format(self._tenant_id)

            if self._client_secret is not None:
                app = msal.ConfidentialClientApplication(
                    authority=authority_url,
                    client_id=self._client_id,
                    client_credential=self._client_secret,
                )
                self._access_token = app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])
            else:
                creds = CertificateCredential(tenant_id=self._tenant_id,
                                      client_id=self._client_id,
                                      certificate_path=self._certificate_path,
                                      password=self._certificate_password)
                token = creds.get_token("https://graph.microsoft.com/.default")
                self._access_token = {
                    "access_token": token.token,
                    "expires_in": int(token.expires_on - time.time()),
                    "token_type": "Bearer",
                    "ext_expires_in": int(token.expires_on - time.time()),
                    "token_source": 'identity_provider'
                }

            self._token_expiry = datetime.datetime.now() + datetime.timedelta(seconds=self._access_token["expires_in"])

        return self._access_token

    def users(self) -> Users:
        """
        Provides access to the Users functionality within the application by returning
        an instance of the `Users` class. This enables interaction with specific user-related
        operations or services through the graph client and authenticated session token.

        :return: Returns an instance of the `Users` class to interact with user-related
            functionality.
        :rtype: Users
        """
        return Users(self.graph_client, self.token_auth_session)


    def communications(self) -> Communications:
        """
        This method initializes and returns an instance of the Communications class.
        It revolves around handling communication-related functionalities by utilizing
        the provided `graph_client` and `token_auth_session`. The class being returned
        is expected to encapsulate precise methods and attributes for performing
        communications-based operations.

        :return: An instance of the Communications class
        :rtype: Communications
        """
        return Communications(self.graph_client, self.token_auth_session)

    def subscriptions(self) -> Subscriptions:
        return Subscriptions(self.graph_client, self.token_auth_session)