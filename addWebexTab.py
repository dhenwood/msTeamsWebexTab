import requests
import json

token = "validO365OAuthToken"

def getTeamsList():
    url = "https://graph.microsoft.com/v1.0/groups?$select=id,displayName,resourceProvisioningOptions"

    payload = {}
    headers = {
      'Authorization': 'Bearer ' + token
    }

    response = requests.request("GET", url, headers=headers, data = payload)

    data = response.json()
    values = data['value']

    for value in values:
        id = value['id']
        displayName = value['displayName']
        resourceProvisioningOptions = value['resourceProvisioningOptions']
        if len(resourceProvisioningOptions) >= 1:
            choice = input("Team: " + displayName + ". Add Webex Tab (y/n)?")
            if choice == "y":    
                getChannelList(id)

def getChannelList(teamId):
    url = "https://graph.microsoft.com/v1.0/teams/" + teamId + "/channels"

    payload = {}
    headers = {
      'Authorization': 'Bearer ' + token
    }

    response = requests.request("GET", url, headers=headers, data = payload)
    data = response.json()

    values = data['value']

    for value in values:
        channelId = value['id']
        displayName = value['displayName']
        if displayName == "General":
            addAppToTeams(teamId,channelId)

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
    #data = response.json()
    #print(data)
    addTabToChannel(teamId,channelId)

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
    data = response.json()
    print(data)


getTeamsList()
