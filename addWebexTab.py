import requests
import json
import configparser
from datetime import datetime
import calendar, time
import pytz

config = configparser.ConfigParser(comment_prefixes='/', allow_no_value=True)
config.read('config.ini')

token = config['TOKEN']['token']


def getTeamsList():
    url = "https://graph.microsoft.com/v1.0/groups?$select=id,displayName,resourceProvisioningOptions,createdDateTime"

    payload = {}
    headers = {
      'Authorization': 'Bearer ' + token
    }

    response = requests.request("GET", url, headers=headers, data = payload)
    if response.status_code == 200:
        data = response.json()
        values = data['value']

        for value in values:
            id = value['id']
            displayName = value['displayName']
            resourceProvisioningOptions = value['resourceProvisioningOptions']
            createdDateTime = value['createdDateTime']
            if len(resourceProvisioningOptions) >= 1:
                checkIfNewer(id,displayName,createdDateTime)

        resetLastCheck()
    else:
        print("getTeamsList; failed with error code: " + str(response.status_code))



def checkIfNewer(teamId,displayName,createdDateTime):
    utc=pytz.UTC

    lastCheckString = config['GENERAL']['lastCheck']
    lastCheckDate = datetime.strptime(lastCheckString, '%Y-%m-%dT%H:%M:%SZ')
    lastCheckTz = utc.localize(lastCheckDate)

    createdDate = datetime.strptime(createdDateTime, '%Y-%m-%dT%H:%M:%SZ')
    channelCreated = utc.localize(createdDate)

    if channelCreated > lastCheckTz:
        print("Newer channel than lastCheck. Adding " + displayName)
        getChannelList(teamId)
    else:
        print("Older channel than lastCheck. Ignoring " + displayName)



def getChannelList(teamId):
    url = "https://graph.microsoft.com/v1.0/teams/" + teamId + "/channels"

    payload = {}
    headers = {
      'Authorization': 'Bearer ' + token
    }

    response = requests.request("GET", url, headers=headers, data = payload)
    if response.status_code == 200:
        data = response.json()

        values = data['value']

        for value in values:
            channelId = value['id']
            displayName = value['displayName']
            if displayName == "General":
                addAppToTeams(teamId,channelId)
    else:
        print("getChannelList; failed with error code: " + str(response.status_code))


def addAppToTeams(teamId,channelId):
    url = "https://graph.microsoft.com/v1.0/teams/" + teamId + "/installedApps"
    myBody = """{
        'teamsApp@odata.bind':'https://graph.microsoft.com/v1.0/appCatalogs/teamsApps/80f3c320-e55f-434f-98e8-d798dfcbe182\'
    }"""

    headers = {
      'Content-Type': 'application/json',
      'Authorization': 'Bearer ' + token
    }

    response = requests.request("POST", url, headers=headers, data = myBody)
    if response.status_code == 201:
        print("Response AddAppToTeams")
        print(response.status_code)
        addTabToChannel(teamId,channelId)
    else:
        print("addAppToTeams; failed with error code: " + str(response.status_code))



def addTabToChannel(teamId,channelId):
    url = "https://graph.microsoft.com/v1.0/teams/" + teamId + "/channels/" + channelId + "/tabs"
    myBody = """{
        'displayName': 'Webex',
        'teamsApp@odata.bind' : 'https://graph.microsoft.com/v1.0/appCatalogs/teamsApps/80f3c320-e55f-434f-98e8-d798dfcbe182',
        'configuration': {
            'entityId': 'notes_by_channel',
            'contentUrl': 'https://teams.webexapps.com/teams/tabs/dash',
            'websiteUrl': 'https://teams.webexapps.com/teams/tabs/web?teamId=""" + channelId + """',
            'removeUrl': 'https://teams.webexapps.com/teams/tabs/remove'
        }
    }"""

    headers = {
      'Content-Type': 'application/json',
      'Authorization': 'Bearer ' + token
    }

    response = requests.request("POST", url, headers=headers, data = myBody)
    if response.status_code == 201:
        print("Response AddTabToChannel")
        print(response.status_code)
    else:
         print("addTabToChannel; failed with error code: " + str(response.status_code))


def resetLastCheck():
    tz_GMT = pytz.timezone('Etc/GMT')
    datetime_GMT = datetime.now(tz_GMT)
    dateTimeString = datetime_GMT.strftime("%Y-%m-%dT%H:%M:%SZ")
    print("Updated lastCheck: ", dateTimeString)

    config['GENERAL']['lastCheck'] = dateTimeString
    with open('config.ini', 'w') as configfile:
        config.write(configfile)




def checkToken():
    timeNow = calendar.timegm(time.gmtime())

    tokenExpiry = config['TOKEN']['tokenexpire']
    tokenExpInt = int(tokenExpiry)

    sec90Days = 60 * 60 * 24 * 90 # 7,776,000
    exp90Day = tokenExpInt + sec90Days

    if timeNow < tokenExpInt:
        # Token still active
        print('Token Active')
        getTeamsList()
    elif timeNow < exp90Day:
        # Token expired but still able to refresh
        print('Token Expired, but refreshable')
        refreshToken()
    else:
        # Token expired beyond 90 days. Need to re-generate.
        print('Token full expired')
        generateToken()



def refreshToken():
    client_id = config['USER']['clientid']
    client_secret = config['USER']['clientsecret']
    tenant_id = config['USER']['tenantid']
    username = config['USER']['username']
    password = config['USER']['password']
    refresh_token = config['TOKEN']['refreshtoken']

    token_url = 'https://login.microsoftonline.com/' + tenant_id + '/oauth2/token'
    token_data = {
     'grant_type': 'refresh_token',
     'client_id': client_id,
     'client_secret': client_secret,
     'scope':'https://graph.microsoft.com',
     'refresh_token': refresh_token
    }

    token_r = requests.post(token_url, data=token_data)

    token = token_r.json().get('access_token')
    refresh = token_r.json().get('refresh_token')
    expires_on = token_r.json().get('expires_on')

    writeTokensToConfig(token,refresh,expires_on)
    #getTeamsList()

    #writeToConfig = writeTokensToConfig(token,refresh,expires_on)
    #if writeToConfig:
        #getTeamsList()



def generateToken():
    client_id = config['USER']['clientid']
    client_secret = config['USER']['clientsecret']
    tenant_id = config['USER']['tenantid']
    username = config['USER']['username']
    password = config['USER']['password']

    token_url = 'https://login.microsoftonline.com/' + tenant_id + '/oauth2/token'
    token_data = {
     'grant_type': 'password',
     'client_id': client_id,
     'client_secret': client_secret,
     'resource': 'https://graph.microsoft.com',
     'scope':'https://graph.microsoft.com',
     'username': username,
     'password': password
    }

    token_r = requests.post(token_url, data=token_data)

    token = token_r.json().get('access_token')
    refresh = token_r.json().get('refresh_token')
    expires_on = token_r.json().get('expires_on')

    writeTokensToConfig(token,refresh,expires_on)
    #getTeamsList()
    #writeToConfig = writeTokensToConfig(token,refresh,expires_on)
    #if writeToConfig:
        #getTeamsList()



def writeTokensToConfig(localToken,refresh,expires_on):
    global token

    print('setting values in config.ini')

    config['TOKEN']['token'] = localToken
    config['TOKEN']['refreshToken'] = refresh
    config['TOKEN']['tokenExpire'] = expires_on

    with open('config.ini', 'w') as configfile:
        config.write(configfile)
    token = localToken
    
    getTeamsList()
    #return True


checkToken()
