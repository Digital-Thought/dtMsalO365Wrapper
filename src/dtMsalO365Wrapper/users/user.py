from office365.graph_client import GraphClient
from office365.directory.users.user import User as Office365User
from office365.directory.users.user import Presence

from dtMsalO365Wrapper.messages import Messages
from dtMsalO365Wrapper._token_auth_session import TokenAuthSession

class User:
    """
    Represents a user in the context of a Microsoft Graph API.

    Provides functionality to manage and access information about a user, such as
    their personal details, job title, contact information, and presence data. This
    class utilizes an instance of a `GraphClient` to facilitate communication with
    the API and an `Office365User` object representing the user. Ensures data is
    lazily loaded only when required.

    :ivar _graph_client: Instance of the GraphClient used for API interactions.
    :type _graph_client: GraphClient
    :ivar _user: Instance of Office365User that this User object is associated with.
    :type _user: Office365User
    :ivar _loaded: Boolean flag indicating whether user data has been fully loaded.
    :type _loaded: bool
    """
    def __init__(self, graph_client: GraphClient, token_auth_session: TokenAuthSession, user: Office365User):
        self._graph_client = graph_client
        self._user: Office365User = user
        self._token_auth_session = token_auth_session
        self._loaded = False

    def get_loaded_user(self):
        """
        Retrieves the currently loaded user object. If the user is not yet loaded, it triggers
        a query execution to load the user and sets the loaded status accordingly.

        :return: The loaded user object.
        :rtype: object
        """
        if not self._loaded:
            self._graph_client.load(self._user).execute_query()
            self._loaded = True
        return self._user

    @property
    def presence(self) -> Presence:
        """
        Retrieves and returns the user's presence status.

        The method accesses the loaded user's presence information and processes it
        using the graph client. It ensures the presence data is fetched and available
        after executing the query.

        :return: The user's presence status.
        :rtype: Presence
        """
        _p = self.get_loaded_user().presence
        self._graph_client.load(_p).execute_query()
        return _p

    @property
    def id(self):
        """
        Gets the unique identifier of the user.

        This property retrieves the ID associated with the user instance. The ID is
        expected to uniquely identify the user within the context of the application.

        :rtype: int
        :return: The unique ID of the user instance.
        """
        return self._user.id

    @property
    def user_principal_name(self):
        """
        Provides access to the `user_principal_name` property of the internal `_user` object.

        This property retrieves the user principal name associated with the internal
        user object. It is a read-only property.

        :rtype: str
        :return: The user principal name of the internal `_user` object.
        """
        return self._user.user_principal_name

    @property
    def display_name(self):
        """
        Retrieve the display name of the user.

        This property fetches the `displayName` attribute from the `_user`'s
        `properties` dictionary. If the `displayName` is not present, the
        property returns `None`.

        :return: The display name of the user or `None` if not available
        :rtype: str or None
        """
        return self._user.properties.get("displayName", None)

    @property
    def given_name(self):
        """
        Retrieve the given name of the user.

        This property fetches the user's first or given name from their properties
        attribute if available. It returns None if the "givenName" key is not present
        in the user's properties.

        :return: The given name of the user or None if not found
        :rtype: Optional[str]
        """
        return self._user.properties.get("givenName", None)

    @property
    def job_title(self):
        """
        Retrieves the job title of the user.

        This property accesses the `jobTitle` field from the user's properties
        and returns its current value if it exists. If the `jobTitle` field is
        not present in the user's properties, the property will return `None`.

        :rtype: Optional[str]
        :return: The job title of the user if it exists, otherwise `None`.
        """
        return self._user.properties.get("jobTitle", None)

    @property
    def mail(self):
        """
        Retrieves the email address of the user.

        This property fetches the email address associated with the user
        from the user's properties. If no email address is available, it
        returns None.

        :return: The user's email address if available, otherwise None
        :rtype: str or None
        """
        return self._user.properties.get("mail", None)

    @property
    def mobile_phone(self):
        """
        Provides access to the user's mobile phone number stored in properties,
        if available. Returns `None` if the mobile phone number is not set.

        :return: The mobile phone number associated with the user or `None` if
            not available.
        :rtype: Optional[str]
        """
        return self._user.properties.get("mobilePhone", None)

    @property
    def office_location(self):
        """
        Retrieves the office location of the user.

        This property fetches the value associated with the "officeLocation"
        key from the `_user` object's properties. If the key is not present,
        it returns `None`.

        :rtype: Optional[str]
        :return: The office location of the user if available, otherwise `None`.
        """
        return self._user.properties.get("officeLocation", None)

    @property
    def surname(self):
        """
        Gets the surname of the user from the user properties.

        This property retrieves the surname of a user from the "properties" dictionary of the
        user object. If the "surname" key does not exist in the dictionary, it returns None.

        :return: The user's surname or None if it does not exist in the properties.
        :rtype: Optional[str]
        """
        return self._user.properties.get("surname", None)

    @property
    def preferred_language(self):
        """
        Gets the preferred language of the user from their properties.

        This property retrieves the value associated with the key
        ``"preferredLanguage"`` in the user's properties dictionary. If the key
        does not exist in the dictionary, it returns None.

        :return: The user's preferred language if it exists, or None.
        :rtype: str or None
        """
        return self._user.properties.get("preferredLanguage", None)

    def set_property(self, property_name, value):
        """
        Sets a property with the specified name and value for the user.

        This method allows updating the property of the user by associating a specific
        property name with its corresponding value. After setting the property, the
        method updates and synchronizes the changes to the server.

        :param property_name: The name of the property to set.
        :type property_name: str
        :param value: The value to be associated with the property.
        :type value: Any
        :return: None
        """
        self._user.set_property(property_name, value).update().execute_query()

    def get_message(self, message_id):
        return Messages(self._graph_client, self._token_auth_session).get_message(self, message_id)