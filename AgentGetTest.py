"""
*
-Josh
6/20/2019
*
"""

import requests
import time
import csv
import sys
import xmltodict

#get WebServiceKey and agentID from user
webServiceKey = input("Web Service Key: ")
agentID = input("Agent ID: ")
print("\n\n")
#use input to call API
url = "https://api.mozenda.com/rest?WebServiceKey=" + webServiceKey + "&Service=Mozenda10&Operation=Agent.Get&AgentID=" + agentID
#double parse the agentGetResponse XML
agentXML = xmltodict.parse(xmltodict.parse(requests.get(url).content)["AgentGetResponse"]["Definition"])
#make list of actions from XML
actionList = agentXML["Agent"]["AgentDefinition"]["ActionList"]["Action"]
holder = []
clickHolder = []
pageListHolder = []
for item in actionList:
    elementType = item['ActionType']
    if elementType == "GetElementValue":#check if action is a capture action
        if not "IsCaptureOptional" in item or item['IsCaptureOptional'] == "True":#check if the capture has "ignore DataNotCaptured errors" selected
            print(dict(item))#prints the action
            print("\n - ignores data not captured errors\n")
            holder.append(item['ActionID'])#adds the action to a list
    elif elementType == "ClickElement":
        clickHolder.append(item['ActionID'])
        print(dict(item))#prints the action
        print("\n - is a click item\n")
    elif elementType == "PageList":
        pageListHolder.append(item['ActionID'])
        print(dict(item))#prints the action
        print("\n - is a page list\n")
if not holder:#If holder list is empty, no capture actions had "ignore DataNotCaptured errors" selected
    print("No items were found with 'ignore data not captured errors' selected. Nice job. \n")
else:#Prints the holder list with ActionIDs
    print("Items with 'ignore data not captured errors' selected: \n")
    for index in holder:
        print(index)
if not clickHolder:
    print("\nNo click item actions were found.\n")
else: 
    print("\nClick item actions found:\n")
    for index in clickHolder:
        print(index)
if not pageListHolder:
    print("\nThis agent does not have a page list action.\n")
else: 
    print("\nPage list actions found:\n")
    for index in pageListHolder:
        print(index)
print("\n\nSee other information for these items printed above. \n\n\n")
agentInfo = agentXML["Agent"]
requestBlocking = str(agentInfo['RequestBlockLevel'])
geolocation = agentInfo['Geolocation']
if requestBlocking == "-1":
    requestBlocking = "None"
print("\nRequest blocking level: \s" % requestBlocking)
print("\nGeolocation: \s" % geolocation)

