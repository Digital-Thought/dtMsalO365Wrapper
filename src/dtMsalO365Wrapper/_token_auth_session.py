import time
import requests
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

class TokenAuthSession(requests.Session):
    """
    Provides a session with token-based authentication and built-in retry
    mechanism for transient HTTP errors.

    TokenAuthSession is designed to simplify API interactions by automatically
    managing authentication tokens and retrying requests in case of transient
    errors. It is particularly useful for interacting with APIs that require
    bearer token authentication and may impose rate limits.

    :ivar token_func: Function to retrieve an access token. This function should
        return a token response that includes the 'access_token' key.
    :type token_func: Callable[[], dict]
    :ivar root_url: Base URL for the API. This will be prefixed to all requested
        URLs to construct the complete endpoint path.
    :type root_url: str
    """
    def __init__(self, token_func, scope, root_url='https://graph.microsoft.com/v1.0'):
        super().__init__()
        self.token_func = token_func
        self.root_url = root_url
        self.scope = scope

        # Configure retries for transient errors (except 429, which we handle separately)
        retries = Retry(
            total=5,
            backoff_factor=2,  # Exponential backoff (2^retry seconds)
            status_forcelist=[500, 502, 503],  # Retry on these HTTP errors
            allowed_methods={"GET", "POST", "PUT", "DELETE", "PATCH"}  # Methods to retry
        )
        self.mount("https://", HTTPAdapter(max_retries=retries))

    def get_token(self):
        """
        Fetches the access token using a token retrieval function.

        This function calls `self.token_func()` and returns the 'access_token' from
        the response. It is assumed that `self.token_func()` provides a dictionary-
        formatted response containing the required token under the key 'access_token'.

        :raises KeyError: If the key 'access_token' is not present in the response.

        :return: The access token as a string.
        :rtype: str
        """
        resp = self.token_func(self.scope)
        return resp['access_token']

    def request(self, method, url, **kwargs):
        """
        Sends an HTTP request with an authorization token and handles rate-limiting responses
        automatically. This method dynamically fetches a new token for each request, utilizes
        it in the request headers, and retries the request if rate limits are hit.

        :param method: The HTTP method to use for the request (e.g., "GET", "POST").
        :type method: str
        :param url: The relative URL path for the request to be appended to the root URL.
        :type url: str
        :param kwargs: Additional keyword arguments to pass to the request, such as payload
            or custom headers.
        :type kwargs: dict
        :return: The HTTP response object returned after a successful request or when
            errors other than rate-limiting are encountered.
        :rtype: requests.Response
        """
        # Get a fresh token for each request
        token = self.get_token()
        kwargs.setdefault("headers", {})
        kwargs["headers"]["Authorization"] = f"Bearer {token}"

        while True:
            if url.startswith('https://'):
                response = super().request(method, f'{url}', **kwargs)
            else:
                response = super().request(method, f'{self.root_url}{url}', **kwargs)

            if response.status_code == 429:  # Handle Rate Limiting
                retry_after = int(response.headers.get("Retry-After", 5))  # Default to 5s if not provided
                print(f"Rate limited! Retrying after {retry_after} seconds...")
                time.sleep(retry_after)
                continue  # Retry the request

            return response  # Return successful response or other non-retry errors