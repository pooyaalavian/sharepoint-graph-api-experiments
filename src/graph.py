import os
from azure.identity import DeviceCodeCredential
from msgraph import GraphServiceClient
from msgraph.generated.users.item.user_item_request_builder import UserItemRequestBuilder
from msgraph.generated.users.item.mail_folders.item.messages.messages_request_builder import (
    MessagesRequestBuilder)
from msgraph.generated.users.item.send_mail.send_mail_post_request_body import (
    SendMailPostRequestBody)
from msgraph.generated.models.message import Message
from msgraph.generated.models.item_body import ItemBody
from msgraph.generated.models.body_type import BodyType
from msgraph.generated.models.recipient import Recipient
from msgraph.generated.models.email_address import EmailAddress

class Graph:
    device_code_credential: DeviceCodeCredential
    user_client: GraphServiceClient

    def __init__(self):
        self.client_id = os.getenv('CLIENT_ID')
        self.tenant_id = os.getenv('TENANt_ID')
        self.graph_scopes = os.getenv('GRAPH_SCOPES').split(' ')

        self.device_code_credential = DeviceCodeCredential(self.client_id, tenant_id = self.tenant_id)
        self.user_client = GraphServiceClient(self.device_code_credential, self.graph_scopes)
        
    async def get_user_token(self, scopes=None):
        if scopes is None:
            scopes = self.graph_scopes
        access_token = self.device_code_credential.get_token(scopes)
        return access_token.token
    
    async def get_user(self):
        # Only request specific properties using $select
        query_params = UserItemRequestBuilder.UserItemRequestBuilderGetQueryParameters(
            # select=['displayName', 'mail', 'userPrincipalName']
        )

        request_config = UserItemRequestBuilder.UserItemRequestBuilderGetRequestConfiguration(
            query_parameters=query_params
        )

        user = await self.user_client.me.get(request_configuration=request_config)
        return user