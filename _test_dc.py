import msal

CLIENT_ID = "de8bc8b5-d9f9-48b1-a8ad-b748da725064"   # Graph Explorer
AUTHORITY = "https://login.microsoftonline.com/organizations"
SCOPES    = ["https://graph.microsoft.com/Team.ReadBasic.All"]

app = msal.PublicClientApplication(CLIENT_ID, authority=AUTHORITY)
flow = app.initiate_device_flow(scopes=SCOPES)

if "user_code" in flow:
    print("OK")
    print("user_code:", flow["user_code"])
    print("url      :", flow["verification_uri"])
else:
    print("FAIL:", flow.get("error_description", flow))
