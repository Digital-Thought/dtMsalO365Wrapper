from office365.graph_client import GraphClient
from office365.directory.users.collection import UserCollection

from dtMsalO365Wrapper._token_auth_session import TokenAuthSession

import logging

class Communications:

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
    """
    def __init__(self, graph_client: GraphClient, token_auth_session: TokenAuthSession):
        self._graph_client = graph_client
        self._token_auth_session = token_auth_session


    def get_presence(self, users, batch_size = 650):
        """
        Fetches the presence of users by processing them in batches and associating the results
        with the corresponding user objects. A batch size can be specified to determine the
        number of users processed in each request. Effective for handling large sets of user data.

        The method makes a POST request to a specific endpoint to retrieve user presence
        information. It logs the progress of processing each batch and handles error cases
        by logging the corresponding responses.

        :param users: A list of user objects that contain the details of users whose presence
                      needs to be fetched.
        :type users: list

        :param batch_size: The number of users to process per batch.
        :type batch_size: int, optional, default is 650

        :return: A consolidated list of presence data with associated user information.
        :rtype: list
        """
        _users = users
        _ids = [u.id for u in _users]
        consolidated_results = []  # Store all results

        # Process users in batches
        for i in range(0, len(_ids), batch_size):
            batch_ids = _ids[i:i + batch_size]
            logging.info(f"Processing batch {i // batch_size + 1} ({len(batch_ids)} users)...")

            response = self._token_auth_session.request(
                "POST",
                "/communications/getPresencesByUserId",
                json={"ids": batch_ids}
            )

            if response.ok:
                consolidated_results.extend(response.json()['value'])  # Store successful results
            else:
                logging.error(f"Error processing batch {i // batch_size + 1}: {response.status_code} -> {response.content}")
        for a in consolidated_results:
            for u in _users:
                if a['id'] == u.id:
                    a['user'] = u
                    break
        return consolidated_results