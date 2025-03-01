# dtMsalO365Wrapper

## Overview
**dtMsalO365Wrapper** is a Python library that simplifies authentication and integration with Microsoft 365 services using **Microsoft Authentication Library (MSAL)** and **Office365-REST-Python-Client**. It provides an abstraction layer for interacting with **Azure Active Directory (Azure AD)** and **Microsoft 365 APIs** such as SharePoint, Outlook, OneDrive, and Teams.

## Features
- 🔐 **Authentication**: Simplifies authentication using MSAL with support for client credentials, authorization code, and device code flows.
- 🔄 **Token Management**: Automatically handles token retrieval, caching, and refreshing.
- 📂 **Microsoft 365 API Access**: Pre-built functions for interacting with SharePoint, Outlook, OneDrive, and Teams.
- ⚙️ **Developer-Friendly**: Provides easy-to-use methods to interact with Microsoft services without low-level API complexity.

## Installation

Ensure you have **Python 3.8+** installed. You can install the library via pip:

```sh
pip install dtMsalO365Wrapper
```

## Quick Start

### 1️⃣ **Authenticate with Azure AD**

```python
from dtMsalO365Wrapper import MsalAuth

# Initialize authentication
auth = MsalAuth(
    client_id="your-client-id",
    client_secret="your-client-secret",
    tenant_id="your-tenant-id"
)

token = auth.get_access_token()
print("Access Token:", token)
```

### 2️⃣ **Interact with Microsoft 365 Services**

#### Retrieve User Profile from Microsoft Graph API

```python
from dtMsalO365Wrapper import O365Client

o365 = O365Client(auth)
user_profile = o365.get_user_profile()
print(user_profile)
```

#### List SharePoint Sites

```python
sites = o365.list_sharepoint_sites()
print(sites)
```

#### Send an Email via Outlook API

```python
o365.send_email(
    recipient="user@example.com",
    subject="Test Email",
    body="Hello from dtMsalO365Wrapper!"
)
```

## Configuration
### Environment Variables
You can also configure authentication using environment variables:

```sh
export MSAL_CLIENT_ID="your-client-id"
export MSAL_CLIENT_SECRET="your-client-secret"
export MSAL_TENANT_ID="your-tenant-id"
```

## Supported Authentication Flows
| Authentication Flow  | Supported |
|----------------------|-----------|
| Client Credentials  | ✅ Yes |
| Authorization Code  | ✅ Yes |
| Device Code         | ✅ Yes |
| Interactive MFA     | ⏳ Planned |

## Roadmap
🚀 Upcoming features:
- ✅ **Expanded Microsoft Graph API Support**
- ✅ **Power Automate Integration**
- ⏳ **MFA & Interactive Authentication Support**

## Contributing
Contributions are welcome! Please follow these steps:
1. Fork the repository
2. Create a new feature branch (`git checkout -b feature-xyz`)
3. Commit changes and push (`git commit -m "Add new feature" && git push`)
4. Submit a pull request

## License
This project is licensed under the MIT License.

## Contact
For support or inquiries, please contact **dev@digital-thought.org**.