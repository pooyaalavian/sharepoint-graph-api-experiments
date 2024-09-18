# Cognitive Search + Sharepoint

## Env Setup
```sh
# Create a new Python virtual environment
python3 -m venv venv

# Activate the virtual environment
# On Windows, use:
venv\Scripts\activate
# On Unix or MacOS, use:
source venv/bin/activate

# Step 3: Install the packages listed in the requirements.txt file
pip install -r requirements.txt
```

## Configure App Registration

### On Azure
1. Create App registration
2. Create a secret (not needed)
3. Under "Authentication" menu, enable "Allow Public Client Workflows".
4. On your device, run powershell as admin and then:
```ps1
.\Create-SelfSignedCertificate.ps1 -CommonName "MyCompanyName" -StartDate 2023-11-01 -EndDate 2024-11-01
``` 
5. Permissions:
    - Graph:
    - SharePoint:
        - `Sites.Manage.All` with **Application**


### On SharePoint
1. Go to sharepoint site page:
```
https://{tenant}.sharepoint.com/sites/{your-site}/_layouts/15/appinv.aspx
```
2. Look up the App Id (Client Id) from previous section.
3. Set Domain to `localhost`, redirect url to `http://localhost`.
4. Use this template for permission:
```xml
<AppPermissionRequests AllowAppOnlyPolicy="true">  
    <AppPermissionRequest Scope="http://sharepoint/content/sitecollection" Right="FullControl" />

</AppPermissionRequests>
``` 
(for more info see 
https://learn.microsoft.com/en-us/sharepoint/dev/solution-guidance/security-apponly-azureacs
and 
https://learn.microsoft.com/en-us/sharepoint/dev/sp-add-ins/add-in-permissions-in-sharepoint
)
5. Submit. If asked for tenant administrator approval, visit:
```
https://{tenant}-admin.sharepoint.com/_layouts/15/appinv.aspx
#               ^^^^^^              ^^^
```

## Run the code
### Set up environment variables
Make sure you set these environment variables:
```sh
TENANT_ID="a7053deb-..."
CLIENT_ID="8458a0df-..."
CLIENT_SECRET="e5U8Q~..."
CLIENT_CERT_THUMBPRINT="2A55C92..."
CLIENT_CERT_PATH="MyCompanyName.pfx"
SHAREPOINT_TENANT="adient"
SHAREPOINT_SITE="samplesite01"
```
Run `python main.py`.
