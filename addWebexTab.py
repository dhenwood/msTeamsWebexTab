import requests
import json
import configparser
from datetime import datetime
import pytz

config = configparser.ConfigParser()
config.read('config.ini')
token = config['DEFAULT']['token']


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

    lastCheckString = config['DEFAULT']['lastCheck']
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
                #print("breakpoint")
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

    config['DEFAULT']['lastCheck'] = dateTimeString
    with open('config.ini', 'w') as configfile:
        config.write(configfile)


getTeamsList()
