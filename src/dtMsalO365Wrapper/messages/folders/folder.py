from office365.graph_client import GraphClient
from office365.directory.users.user import User as Office365User
from office365.directory.users.user import Presence

from dtMsalO365Wrapper._token_auth_session import TokenAuthSession

import logging

class Folder:

    def __init__(self, graph_client: GraphClient, token_auth_session: TokenAuthSession, user, folder_detail: dict):
        self._graph_client = graph_client
        self._folder_detail: dict = folder_detail
        self._token_auth_session = token_auth_session
        self.user = user

    def get_parent_folder(self):
        parent_folder_id = self._folder_detail['parentFolderId']
        resp = self._token_auth_session.request("GET", f"/users/{self.user.id}/mailFolders/{parent_folder_id}")
        if resp.status_code != 200:
            logging.error(f'Failed to get folder: {resp.content}')
            raise Exception(f'Failed to get folder: {resp.content}')

        return Folder(self._graph_client, self._token_auth_session, self.user, resp.json())

    @property
    def folder_path(self):
        path = [self.display_name]
        current_folder_id = self.id
        folder = self.get_parent_folder()
        while True:
            if folder.id == current_folder_id:
                break
            current_folder_id = folder.id
            path.insert(0, folder.display_name)
            folder = folder.get_parent_folder()

        return '/'.join(path)


    @property
    def id(self):
        return self._folder_detail['id']

    @property
    def display_name(self):
        return self._folder_detail['displayName']

    @property
    def unread_item_count(self):
        return self._folder_detail['unreadItemCount']

    @property
    def child_folder_count(self):
        return self._folder_detail['childFolderCount']

    @property
    def size_in_bytes(self):
        return self._folder_detail['sizeInBytes']

    @property
    def total_item_count(self):
        return self._folder_detail['totalItemCount']

    @property
    def hidden(self):
        return self._folder_detail['isHidden']
