from office365.graph_client import GraphClient
from office365.directory.users.user import User as Office365User
from office365.directory.users.user import Presence

from dtMsalO365Wrapper.messages.folders.folder import Folder
from dtMsalO365Wrapper._token_auth_session import TokenAuthSession

import logging
import datetime

class Message:

    def __init__(self, graph_client: GraphClient, token_auth_session: TokenAuthSession, user, message_detail: dict):
        self._graph_client = graph_client
        self._message_detail: dict = message_detail
        self._token_auth_session = token_auth_session
        self.user = user

    def get_parent_folder(self):
        parent_folder_id = self._message_detail['parentFolderId']
        resp = self._token_auth_session.request("GET", f"/users/{self.user.id}/mailFolders/{parent_folder_id}")
        if resp.status_code != 200:
            logging.error(f'Failed to get folder: {resp.content}')
            raise Exception(f'Failed to get folder: {resp.content}')

        return Folder(self._graph_client, self._token_auth_session, self.user, resp.json())

    @property
    def id(self):
        return self._message_detail['id']

    @property
    def created(self):
        return datetime.datetime.fromisoformat(self._message_detail['createdDateTime'])

    @property
    def last_modified(self):
        return datetime.datetime.fromisoformat(self._message_detail['lastModifiedDateTime'])

    @property
    def categories(self):
        return self._message_detail['categories']

    @property
    def received(self):
        return datetime.datetime.fromisoformat(self._message_detail['receivedDateTime'])

    @property
    def sent(self):
        return datetime.datetime.fromisoformat(self._message_detail['sentDateTime'])

    @property
    def has_attachments(self):
        return self._message_detail['hasAttachments']

    @property
    def internet_message_id(self):
        return self._message_detail['internetMessageId']

    @property
    def subject(self):
        return self._message_detail['subject']

    @property
    def body_preview(self):
        return self._message_detail['bodyPreview']

    @property
    def importance(self):
        return self._message_detail['importance']

    @property
    def conversation_id(self):
        return self._message_detail['conversationId']

    @property
    def conversation_index(self):
        return self._message_detail['conversationIndex']

    @property
    def is_read(self):
        return self._message_detail['isRead']

    @property
    def is_draft(self):
        return self._message_detail['isDraft']

    @property
    def body(self):
        return self._message_detail['body']

    @property
    def sender(self):
        return self._message_detail['sender']

    @property
    def from_(self):
        return self._message_detail['from']

    @property
    def to_recipients(self):
        return self._message_detail['toRecipients']

    @property
    def cc_recipients(self):
        return self._message_detail['ccRecipients']

    @property
    def bcc_recipients(self):
        return self._message_detail['bccRecipients']

    @property
    def reply_to(self):
        return self._message_detail['replyTo']

    @property
    def flag(self):
        return self._message_detail['flag']

