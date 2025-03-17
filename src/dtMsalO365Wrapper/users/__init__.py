from office365.graph_client import GraphClient
from office365.directory.users.collection import UserCollection

from dtMsalO365Wrapper.users.user import User
from dtMsalO365Wrapper._token_auth_session import TokenAuthSession

import logging

class Users:

    DEFAULT_SELECT_FIELDS = ['id','userPrincipalName','accountEnabled','assignedLicenses','assignedPlans','businessPhones','city','companyName','country','createdDateTime','department','displayName','givenName','jobTitle','mail','officeLocation']

    """
    Manages and interacts with user-related data and functionalities.

    This class provides utilities to access and manage user information stored in a
    graph client, as well as fetch additional details such as user presence. It is
    designed to work with large datasets efficiently by supporting operations in
    batches, thus optimizing memory usage and API interactions.

    :ivar _graph_client: The GraphClient instance used to interact with the data source.
    :type _graph_client: GraphClient
    :ivar _token_auth_session: The TokenAuthSession instance for handling authenticated API requests.
    :type _token_auth_session: TokenAuthSession
    :ivar _users: A UserCollection object providing direct access to user records in the graph client.
    :type _users: UserCollection
    """
    def __init__(self, graph_client: GraphClient, token_auth_session: TokenAuthSession):
        self._graph_client = graph_client
        self._token_auth_session = token_auth_session
        self._users: UserCollection = self._graph_client.users

    def get(self, query_filter, select_fields: list = DEFAULT_SELECT_FIELDS):
        """
        Retrieves a list of user objects based on the provided query filter and selected fields.

        This method queries the underlying user data source using the given `query_filter`
        to narrow down the results. It allows specifying specific fields to include in
        the output by using `select_fields`. The method fetches a maximum of 100 user
        records and maps each result to a `User` object for further manipulation or
        usage.

        :param query_filter: Filter criteria to be applied when fetching user data.
        :param select_fields: Optional list of specific fields to include in the
            output. Defaults to `DEFAULT_SELECT_FIELDS` if not specified.
        :return: A list of `User` objects matching the specified query conditions.
        :rtype: list
        """
        _users = []
        for u in self._users.filter(query_filter).get_all(100).select(select_fields).execute_query():
            _users.append(User(self._graph_client, u))
        return _users

    def get_enabled_accounts(self, select_fields: list = DEFAULT_SELECT_FIELDS):
        """
        Retrieve accounts that are currently enabled.

        This method queries accounts where the 'accountEnabled' property is set
        to true and retrieves specified fields for each matching account.

        :param select_fields: A list of strings indicating the fields to be
            retrieved for each enabled account. Defaults to `DEFAULT_SELECT_FIELDS`.
        :type select_fields: list
        :return: The result of the query, typically a list or database objects
            matching the filtering criteria with specified fields included.
        :rtype: Any
        """
        return self.get("accountEnabled eq true", select_fields)

    def get_guest_accounts(self, select_fields: list = DEFAULT_SELECT_FIELDS):
        """
        Retrieves guest user accounts based on a specific query and returns the response with the desired
        fields selected. This function fetches user accounts that belong to the 'guest' type by applying a
        filtering expression while retrieving the data, allowing for customization of returned fields.

        :param select_fields: A list of fields to be included in the response. Defaults are based on the
            `DEFAULT_SELECT_FIELDS`.
            - Must be a list of strings representing valid field names.
        :return: The result of the query containing guest user accounts with the selected fields populated.
        :rtype: Any
        """
        return self.get("userType eq 'guest'", select_fields)

    def get_member_accounts(self, select_fields: list = DEFAULT_SELECT_FIELDS):
        """
        Retrieves accounts associated with member users.

        This method fetches accounts where the user type is defined as 'member'
        and returns them with the specified fields or default selected fields,
        if no arguments are provided.

        :param select_fields: List of fields to be included in the response. If
            not provided, the predefined default select fields will be used.
        :type select_fields: list
        :return: The accounts for member users with the specified fields.
        :rtype: list
        """
        return self.get("userType eq 'member'", select_fields)

    def count(self, count_filter: str = None):
        """
        Retrieves the count of users, optionally filtered by a provided condition. The
        filter is applied using an OData query syntax. The response uses eventual
        consistency level.

        :param count_filter: The filter condition as a string in OData query syntax.
            If not provided, no filter is applied. Defaults to None.
        :return: The count of users as an integer if the request is successful.
            Returns 0 if an error occurs during the request.
        """
        if count_filter:
            count_filter = f"?$filter={count_filter}"
        else:
            count_filter = ''
        response = self._token_auth_session.request(
            "GET",
            f"/users/$count{count_filter}",
            headers={"ConsistencyLevel": "eventual"}
        )

        if response.status_code == 200:
            return response.json()
        else:
            logging.error(f"Error retrieving count: {response.status_code} / {response.content}")
            return 0

    def get_all(self, select_fields: list = DEFAULT_SELECT_FIELDS):
        """
        Retrieves all User instances by querying the underlying data source.

        This generator function fetches user data in batches of 100 records and yields
        User instances, which are initialized using the graph client and the individual
        user data from the data source. This is useful for iterating over a potentially
        large number of users without loading all of them into memory at once.

        :return: A generator that yields User instances based on the data source.
        :rtype: Iterator[User]
        """
        _users = []
        for u in self._users.get_all(100).select(select_fields).execute_query():
            _users.append( User(self._graph_client, u))
        return _users

    def get_top(self, top):
        """
        Yields the top users from the query results of the internal users' collection.

        This method retrieves a specified number of top users by accessing the internal
        users' collection through a series of query operations. The retrieved users are
        then wrapped in `User` objects and yielded one by one. The number of users
        to retrieve is controlled by the `top` parameter. This function is a generator,
        which means it uses lazy evaluation and yields results as they are processed.

        :param top: The number of top users to retrieve.
        :type top: int

        :return: A generator yielding `User` objects representing the top users.
        :rtype: Iterator[User]
        """
        for u in self._users.get().top(top).execute_query():
            yield User(self._graph_client, u)
