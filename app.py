import uuid
import requests
from flask import Flask, render_template, session, request, redirect, url_for, send_file
from flask_session import Session
import msal
from markdownify import markdownify as md

import app_config

# Code based on https://github.com/Azure-Samples/ms-identity-python-webapp

app = Flask(__name__)
app.config.from_object(app_config)
Session(app)
app.debug = True

@app.route("/")
def index():
    if not session.get("user"):
        return redirect(url_for("login"))
    return render_template('index.html')


@app.route("/msal-demo")
def msal_demo():
    if not session.get("user"):
        return redirect(url_for("login"))
    token = get_token(app_config.SCOPE)
    graph_data = requests.get(  # Use token to call downstream service
        'https://graph.microsoft.com/v1.0/me',
        headers={'Authorization': 'Bearer ' + token['access_token']},
    ).json()        
    return render_template('msal-demo/index.html', result=graph_data, user=session["user"])

###############################################################################

#                       ONENOTE DEMO FUNCTIONS                                #

###############################################################################

@app.route("/onenote-demo")
def onenote_demo():
    if not session.get("user"):
        return redirect(url_for("login"))
    token = get_token(app_config.SCOPE)
    one_notes = requests.get(  # Use token to call downstream service
        'https://graph.microsoft.com/v1.0/me/onenote/pages',
        headers={'Authorization': 'Bearer ' + token['access_token']},
    ).json()        
    return render_template('onenote-demo/index.html', result=one_notes)


@app.route("/onenotepage")
def fetch_onenote_page():
    token = get_token(app_config.SCOPE)
    page = requests.get(  # Use token to call downstream service
        request.args['page_url'],
        headers={'Authorization': 'Bearer ' + token['access_token']},
    )
    return page.text # returns HTML


@app.route("/onenotepagemd")
def fetch_onenote_page_md():
    token = get_token(app_config.SCOPE)
    page = requests.get(  # Use token to call downstream service
        request.args['page_url'],
        headers={'Authorization': 'Bearer ' + token['access_token']},
    )
    markdown = md(page.text, heading_style='ATX')
    return markdown

###############################################################################

#                       TEAMS DEMO FUNCTIONS                                  #

###############################################################################

@app.route("/teams-demo")
def teams_demo():
    if not session.get("user"):
        return redirect(url_for("login"))
    team = _get_team(app_config.TEAM_ID)
    teamMembers = _get_team_members(app_config.TEAM_ID)
    return render_template('teams-demo/index.html',  team = team, teamMembers = teamMembers.get('value'))

@app.route("/status-update", methods=["GET", "POST"])
def status_update():
    if not session.get("user"):
        return redirect(url_for("login"))
    if request.method == "POST":  
        statusUpdate = request.form.get('status')
        additionalMessage = request.form.get('message')
        channelId = request.form.get('channelId')
        token = get_token(app_config.SCOPE)
        requests.post(f"https://graph.microsoft.com/v1.0/teams/{app_config.TEAM_ID}/channels/{channelId}/messages", json={
           "body": {
            "content": f"{statusUpdate} - {additionalMessage}"
            }},
            headers={'Authorization': 'Bearer ' + token['access_token'], 'Content-type': 'application/json'}).json()        
    channel = _get_channel(app_config.TEAM_ID, channelId)
    channelMembers = _get_channel_members(app_config.TEAM_ID, channelId)
    return render_template('teams-demo/channel_mgt.html', channel = channel, channelMembers = channelMembers.get('value'))

@app.route("/create-channel", methods=["GET", "POST"])
def create_channel():
    if not session.get("user"):
        return redirect(url_for("login"))
    if request.method == "POST":    
        channelName = f"IncidentChannel-{request.form.get('channelName')}"
        incidentDescription = request.form.get('incidentDescription')
        members = request.form.getlist('members')
        teamId = app_config.TEAM_ID
        members_list = _build_members_list(members)
        token = get_token(app_config.SCOPE)
        channel = requests.post(f"https://graph.microsoft.com/v1.0/teams/{teamId}/channels", json={
            "displayName": channelName,
            "description": incidentDescription,
            "membershipType": "private",
              "members": members_list
                },
            headers={'Authorization': 'Bearer ' + token['access_token'], 'Content-type': 'application/json'}).json()
        channelMembers = _get_channel_members(teamId, channel.get('id'))
        return render_template('teams-demo/channel_mgt.html', channel = channel, channelMembers = channelMembers.get('value'))

def _build_members_list(members):
    members_list = []
    for memberId in members:
        members_list.append(
                    {
                    "@odata.type":"#microsoft.graph.aadUserConversationMember",
                    "user@odata.bind":f"https://graph.microsoft.com/v1.0/users('{memberId}')", # add authenticated user
                    "roles":["owner"]
                    })
    return members_list

def _get_channel(teamId, channelId):
    token = get_token(app_config.SCOPE)        
    return requests.get(f"https://graph.microsoft.com/v1.0/teams/{teamId}/channels/{channelId}",
        headers={'Authorization': 'Bearer ' + token['access_token']}).json()

def _get_team(id):
    token = get_token(app_config.SCOPE)        
    return requests.get(f"https://graph.microsoft.com/v1.0/teams/{id}",
        headers={'Authorization': 'Bearer ' + token['access_token']}).json()

def _get_team_members(teamId):
    token = get_token(app_config.SCOPE)        
    return requests.get(f"https://graph.microsoft.com/v1.0/teams/{teamId}/members",
    headers={'Authorization': 'Bearer ' + token['access_token']}).json()


def _get_channel_members(teamId, channelId):
    token = get_token(app_config.SCOPE)        
    return requests.get(f"https://graph.microsoft.com/v1.0/teams/{teamId}/channels/{channelId}/members",
        headers={'Authorization': 'Bearer ' + token['access_token']}).json()

###############################################################################

#                       TOKEN CACHING AND AUTH FUNCTIONS                      #

###############################################################################

# Its absolute URL must match your app's redirect_uri set in AAD
@app.route("/getAToken")
def authorized():
    if request.args['state'] != session.get("state"):
        return redirect(url_for("login"))
    cache = _load_cache()
    result = _build_msal_app(cache).acquire_token_by_authorization_code(
        request.args['code'],
        scopes=app_config.SCOPE,
        redirect_uri=url_for("authorized", _external=True))
    if "error" in result:
        return "Login failure: %s, %s" % (
            result["error"], result.get("error_description"))
    session["user"] = result.get("id_token_claims")
    _save_cache(cache)
    return redirect(url_for("index"))


def _load_cache():
    cache = msal.SerializableTokenCache()
    if session.get("token_cache"):
        cache.deserialize(session["token_cache"])
    return cache


def _save_cache(cache):
    if cache.has_state_changed:
        session["token_cache"] = cache.serialize()


def _build_msal_app(cache=None, authority=None):
    return msal.ConfidentialClientApplication(
        app_config.CLIENT_ID, authority=authority or app_config.AUTHORITY,
        client_credential=app_config.CLIENT_SECRET, token_cache=cache)


def _get_token_from_cache(scope=None):
    cache = _load_cache()  # This web app maintains one cache per session
    cca = _build_msal_app(cache)
    accounts = cca.get_accounts()
    if accounts:  # So all accounts belong to the current signed-in user
        result = cca.acquire_token_silent(scope, account=accounts[0])
        _save_cache(cache)
        return result


def get_token(scope):
    token = _get_token_from_cache(scope)
    if not token:
        return redirect(url_for("login"))
    return token

###############################################################################

#                       LOGN/LOGOUT FUNCTIONS                                 #

###############################################################################

@app.route("/login")
def login():
    session["state"] = str(uuid.uuid4())
    auth_url = _build_msal_app().get_authorization_request_url(
        app_config.SCOPE,
        state=session["state"],
        redirect_uri=url_for("authorized", _external=True))
    return "<a href='%s'>Login with Microsoft Identity</a>" % auth_url


@app.route("/logout")
def logout():
    session.clear()  # Wipe out the user and the token cache from the session
    return redirect(  # Also need to log out from the Microsoft Identity platform
        "https://login.microsoftonline.com/common/oauth2/v2.0/logout"
        "?post_logout_redirect_uri=" + url_for("index", _external=True))


if __name__ == "__main__":
    app.run()
