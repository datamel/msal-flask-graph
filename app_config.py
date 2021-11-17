# Contains settings for Web App

AUTHORITY = "https://login.microsoftonline.com/common" # This is for multi-tenant, alternatively set equal to "https://login.microsoftonline.com/<TENANT_ID>" for a single tenant
SCOPE = ["ChannelSettings.Read.All",
"ChannelMember.ReadWrite.All",
"ChannelMessage.Send",
"Team.ReadBasic.All",
"TeamMember.ReadWrite.All",
"Notes.ReadWrite.All",
"User.ReadBasic.All"]
SESSION_TYPE = "filesystem"  # So the token cache will be stored in a server-side session
REDIRECT_PATH = "http://localhost:5000/getAToken"
CLIENT_SECRET = "<ENTER CLIENT SECRET FROM Azure Active Directory HERE>"
CLIENT_ID = "<ENTER CLIENT ID FROM Azure Active Directory HERE>"
TEAM_ID = "<ENTER TEAM ID FROM TEAMS or GRAPH HERE>" 
