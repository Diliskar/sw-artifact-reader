from __future__ import print_function
import json
import pickle
import os.path

from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

import usersettings

SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

spreadsheet_id = usersettings.spreadsheet_id
sheet_id = usersettings.sheet_id
sheet_name = usersettings.sheet_name
json_location = usersettings.json_location

# Ranges for Mon / Artifacts are fixed
mon_range = sheet_name + '!A2:D'
mainartifact_range = sheet_name + '!F2:G'
subartifact_range = sheet_name + '!N2:U'
header_range = sheet_name + '!A1:U'

# First Setup for Google API
creds = None
# The file token.pickle stores the user's access and refresh tokens, and is
# created automatically when the authorization flow completes for the first
# time.
if os.path.exists('token.pickle'):
    with open('token.pickle', 'rb') as token:
        creds = pickle.load(token)
# If there are no (valid) credentials available, let the user log in.
if not creds or not creds.valid:
    if creds and creds.expired and creds.refresh_token:
        creds.refresh(Request())
    else:
        flow = InstalledAppFlow.from_client_secrets_file(
            'credentials.json', SCOPES)
        creds = flow.run_local_server(port=0)
    # Save the credentials for the next run
    with open('token.pickle', 'wb') as token:
        pickle.dump(creds, token)

service = build('sheets', 'v4', credentials=creds)


def main():
    with open(f"{json_location}", "rb") as jsonFile:
        jsonData = json.load(jsonFile)

    # Mon / Stat IDs
    with open('mapping.json', "rb") as mapJson:
        mapper = json.load(mapJson)

    # Artifact Sub stats IDs
    with open('artifactMap.json', "rb") as artiJson:
        artiMap = json.load(artiJson)

    # Unique Ids
    try:
        uniqueId_txt = open(r"uniqueIds.txt", "r+")
        unique_ids = uniqueId_txt.read().split("\n")
        uniqueId_txt.close()
    except FileNotFoundError:
        uniqueId_txt = open(r"uniqueIds.txt", "w+")
        uniqueId_txt.close()
        uniqueId_txt = open(r"uniqueIds.txt", "r+")
        unique_ids = uniqueId_txt.read().split("\n")
        uniqueId_txt.close()

    values = []

    result = service.spreadsheets().values().get(
        spreadsheetId=spreadsheet_id, range=sheet_name + "!A:A").execute()
    rows = result.get('values', [])
    mon_range = sheet_name + "!A" + str(len(rows)+1) + ":D"

    # Checks if the unit is already in the sheet via the unique ID
    # If the unit is in there, itll be skipped, if not, itll be appended to the sheet & .txt
    for unit in range(len(jsonData['unit_list'])):

        mon = jsonData['unit_list'][unit]

        com2us_id = mon['unit_master_id']
        unique_id = mon['unit_id']
        name = mapper[str(com2us_id)]['name']
        element = mapper[str(com2us_id)]['element']
        archetype = mapper[str(com2us_id)]['archetype']
        date = mon["create_time"]

        if mon['class'] == 6:
            if str(unique_id) not in unique_ids:
                uniqueId_txt = open(r"uniqueIds.txt", "a")
                uniqueId_txt.write(str(unique_id) + "\n")
                uniqueId_txt.close()
                values.append([name, date, element, archetype])

    jsonFile.close()
    mapJson.close()
    artiJson.close()

    # Appends the missing mons which arent in the sheet yet

    body = {
        'values': values
    }

    result = service.spreadsheets().values().append(spreadsheetId=spreadsheet_id, range=mon_range,
                                                    valueInputOption="RAW", body=body).execute()
    print('{0} cells appended.'.format(result \
                                       .get('updates') \
                                       .get('updatedCells')))

    mon_range = sheet_name + '!A2:D'

    # Sorts after Column B (Date)

    requests = [
        {
            "sortRange": {
                "range": {
                    "sheetId": sheet_id,
                    "startRowIndex": 1,
                    "endRowIndex": 999,
                    "startColumnIndex": 0,
                    "endColumnIndex": 22
                },
                "sortSpecs": [
                    {
                        "dimensionIndex": 1,
                        "sortOrder": "ASCENDING"
                    }
                ]
            }
        }
    ]
    body = {
        'requests': requests
    }

    result = service.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id, body=body).execute()
    print('Sorted Sheet - Column B - Ascending')

    values = []
    values2 = []

    # Updates the artifacts + artifact subs

    for unit in range(len(jsonData['unit_list'])):

        mon = jsonData['unit_list'][unit]

        if mon['class'] == 6:

            mon = jsonData['unit_list'][unit]
            artifacts = mon['artifacts']

            artifact1 = None
            artifact2 = None

            art1sub1 = None
            art1sub2 = None
            art1sub3 = None
            art1sub4 = None
            art2sub1 = None
            art2sub2 = None
            art2sub3 = None
            art2sub4 = None

            for i in range(len(mon['artifacts'])):
                if artifacts[i]['slot'] == 1:
                    try:
                        artifact1 = artiMap['artifact']['effectTypes']['main'][str(artifacts[i]['pri_effect'][0])]
                        art1sub1 = artiMap['artifact']['effectTypes']['sub'][str(artifacts[i]['sec_effects'][0][0])]
                        art1sub2 = artiMap['artifact']['effectTypes']['sub'][str(artifacts[i]['sec_effects'][1][0])]
                        art1sub3 = artiMap['artifact']['effectTypes']['sub'][str(artifacts[i]['sec_effects'][2][0])]
                        art1sub4 = artiMap['artifact']['effectTypes']['sub'][str(artifacts[i]['sec_effects'][3][0])]
                    except IndexError:
                        None
                elif artifacts[i]['slot'] == 2:
                    try:
                        artifact2 = artiMap['artifact']['effectTypes']['main'][str(artifacts[i]['pri_effect'][0])]
                        art2sub1 = artiMap['artifact']['effectTypes']['sub'][str(artifacts[i]['sec_effects'][0][0])]
                        art2sub2 = artiMap['artifact']['effectTypes']['sub'][str(artifacts[i]['sec_effects'][1][0])]
                        art2sub3 = artiMap['artifact']['effectTypes']['sub'][str(artifacts[i]['sec_effects'][2][0])]
                        art2sub4 = artiMap['artifact']['effectTypes']['sub'][str(artifacts[i]['sec_effects'][3][0])]
                    except IndexError:
                        None

            # Appends the main artifact stats to values

            values.append([artifact1, artifact2])
            values2.append([art1sub1, art1sub2, art1sub3, art1sub4, art2sub1, art2sub2, art2sub3, art2sub4])

    data = [
        {
            'range': mainartifact_range,
            'values': values
        },
        {
            'range': subartifact_range,
            'values': values2
        }
    ]

    body = {
        'valueInputOption': 'RAW',
        'data': data
    }

    result = service.spreadsheets().values().batchUpdate(
        spreadsheetId=spreadsheet_id, body=body).execute()
    print('{0} cells updated.'.format(result.get('totalUpdatedCells')))

    # add data validation

    # Sorts after Column C (Element)

    requests = [
        {
            "sortRange": {
                "range": {
                    "sheetId": sheet_id,
                    "startRowIndex": 1,
                    "endRowIndex": 999,
                    "startColumnIndex": 0,
                    "endColumnIndex": 22
                },
                "sortSpecs": [
                    {
                        "dimensionIndex": 2,
                        "sortOrder": "DESCENDING"
                    }
                ]
            }
        }
    ]
    body = {
        'requests': requests
    }

    results = service.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id, body=body).execute()
    print('Sorted Sheet - Column C - DESCENDING')


def setupsheet():

    # Checks if the setup.txt has been changed, 1 = skip setup, 0 = execute setup
    try:
        setup = open(r"setup.txt", "r+")
        text = setup.read()
        setup.close()
    except FileNotFoundError:
        setup = open(r"setup.txt", "w+")
        setup.close()
        text = "0"

    if text == "0":
        setup = open(r"setup.txt", "w+")
        setup.write("2")
        setup.close()
        print('~~~ First time sheet setup ~~~')
    elif text == "1":
        setup = open(r"setup.txt", "w+")
        setup.write("2")
        setup.close()
        print('~~~ Updating Sheet Layout ~~~')
    else:
        return



    # Writes Header row
    values = [
        [
            'Monster', 'Date', 'Element', 'Type', 'Main Stat', 'Attribute Arti', 'Type Arti', 'Logic',
            'Note Attribute Arti', 'Note Attribute Arti 2', 'Note Type Arti', 'Note Type Arti 2',
            None, 'Art 1 Sub 1', 'Art 1 Sub 2', 'Art 1 Sub 3', 'Art 1 Sub 4',
            'Art 2 Sub 1', 'Art 2 Sub 2', 'Art 2 Sub 3', 'Art 2 Sub 4'
        ]
    ]
    body = {
        'values': values
    }
    result = service.spreadsheets().values().update(
        spreadsheetId=spreadsheet_id, range=header_range,
        valueInputOption='RAW', body=body).execute()
    print('{0} cells updated.'.format(result.get('updatedCells')))

    # Request to freeze 1st row and change column sizes
    requests = [
        {
            # Freezes the 1st row
            'updateSheetProperties': {
                'properties': {
                    'sheetId': sheet_id,
                    'gridProperties': {
                        'frozenRowCount': 1
                    }
                },
                "fields": "gridProperties.frozenRowCount"
            }
        },
        # Following changes the size of the columns
        # A
        {

            'updateDimensionProperties': {
                'range': {
                    'sheetId': sheet_id,
                    'dimension': "COLUMNS",
                    'startIndex': 0,
                    'endIndex': 1
                },
                'properties': {
                    'pixelSize': 135
                },
                'fields': 'pixelSize'
            }
        },
        # B
        {
            'updateDimensionProperties': {
                'range': {
                    'sheetId': sheet_id,
                    'dimension': "COLUMNS",
                    'startIndex': 1,
                    'endIndex': 2
                },
                'properties': {
                    'pixelSize': 32
                },
                'fields': 'pixelSize'
            }
        },
        # C:D
        {
            'updateDimensionProperties': {
                'range': {
                    'sheetId': sheet_id,
                    'dimension': "COLUMNS",
                    'startIndex': 2,
                    'endIndex': 4
                },
                'properties': {
                    'pixelSize': 100
                },
                'fields': 'pixelSize'
            }
        },
        # E:H
        {
            'updateDimensionProperties': {
                'range': {
                    'sheetId': sheet_id,
                    'dimension': "COLUMNS",
                    'startIndex': 4,
                    'endIndex': 8
                },
                'properties': {
                    'pixelSize': 85
                },
                'fields': 'pixelSize'
            }
        },
        # I:L
        {
            'updateDimensionProperties': {
                'range': {
                    'sheetId': sheet_id,
                    'dimension': "COLUMNS",
                    'startIndex': 8,
                    'endIndex': 12
                },
                'properties': {
                    'pixelSize': 260
                },
                'fields': 'pixelSize'
            }
        }
    ]
    body = {
        'requests': requests
    }

    results = service.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id, body=body).execute()

    # Adds Data Validation (Check Box - Drop Down List)

    checkbox_range = {
        'sheetId': sheet_id,
        'startRowIndex': 1,
        'endRowIndex': 999,
        'startColumnIndex': 7,
        'endColumnIndex': 8
    }
    values_dropdown_attribute = [
        {
            "userEnteredValue": "ATK Increased Proportional to Lost HP"
        },
        {
            "userEnteredValue": "DEF Increased Proportional to Lost HP"
        },
        {
            "userEnteredValue": "SPD Increased Proportional to Lost HP"
        },
        {
            "userEnteredValue": "ATK Increasing Effect"
        },
        {
            "userEnteredValue": "DEF Increasing Effect"
        },
        {
            "userEnteredValue": "SPD Increasing Effect"
        },
        {
            "userEnteredValue": "Damage Dealt by Counterattack"
        },
        {
            "userEnteredValue": "Damage Dealt by Attacking Together"
        },
        {
            "userEnteredValue": "Bomb Damage"
        },
        {
            "userEnteredValue": "Received Crit DMG"
        },
        {
            "userEnteredValue": "Life Drain"
        },
        {
            "userEnteredValue": "HP when Revived"
        },
        {
            "userEnteredValue": "Attack Bar when Revived"
        },
        {
            "userEnteredValue": "Additional Damage by % of HP"
        },
        {
            "userEnteredValue": "Additional Damage by % of ATK"
        },
        {
            "userEnteredValue": "Additional Damage by % of DEF"
        },
        {
            "userEnteredValue": "Additional Damage by % of SPD"
        },
        {
            "userEnteredValue": "Damage Dealt on Fire"
        },
        {
            "userEnteredValue": "Damage Dealt on Water"
        },
        {
            "userEnteredValue": "Damage Dealt on Wind"
        },
        {
            "userEnteredValue": "Damage Dealt on Light"
        },
        {
            "userEnteredValue": "Damage Dealt on Dark"
        },
        {
            "userEnteredValue": "Damage Received from Fire"
        },
        {
            "userEnteredValue": "Damage Received from Water"
        },
        {
            "userEnteredValue": "Damage Received from Wind"
        },
        {
            "userEnteredValue": "Damage Received from Light"
        },
        {
            "userEnteredValue": "Damage Received from Dark"
        }
    ]
    values_dropdown_type = [
        {
            "userEnteredValue": "ATK Increased Proportional to Lost HP"
        },
        {
            "userEnteredValue": "DEF Increased Proportional to Lost HP"
        },
        {
            "userEnteredValue": "SPD Increased Proportional to Lost HP"
        },
        {
            "userEnteredValue": "ATK Increasing Effect"
        },
        {
            "userEnteredValue": "DEF Increasing Effect"
        },
        {
            "userEnteredValue": "SPD Increasing Effect"
        },
        {
            "userEnteredValue": "Damage Dealt by Counterattack"
        },
        {
            "userEnteredValue": "Damage Dealt by Attacking Together"
        },
        {
            "userEnteredValue": "Bomb Damage"
        },
        {
            "userEnteredValue": "Received Crit DMG"
        },
        {
            "userEnteredValue": "Life Drain"
        },
        {
            "userEnteredValue": "HP when Revived"
        },
        {
            "userEnteredValue": "Attack Bar when Revived"
        },
        {
            "userEnteredValue": "Additional Damage by % of HP"
        },
        {
            "userEnteredValue": "Additional Damage by % of ATK"
        },
        {
            "userEnteredValue": "Additional Damage by % of DEF"
        },
        {
            "userEnteredValue": "Additional Damage by % of SPD"
        },
        {
            "userEnteredValue": "Skill 1 CRIT DMG"
        },
        {
            "userEnteredValue": "Skill 2 CRIT DMG"
        },
        {
            "userEnteredValue": "Skill 3 CRIT DMG"
        },
        {
            "userEnteredValue": "Skill 4 CRIT DMG"
        },
        {
            "userEnteredValue": "Skill 1 Recovery"
        },
        {
            "userEnteredValue": "Skill 2 Recovery"
        },
        {
            "userEnteredValue": "Skill 3 Recovery"
        },
        {
            "userEnteredValue": "Skill 1 Accuracy"
        },
        {
            "userEnteredValue": "Skill 2 Accuracy"
        },
        {
            "userEnteredValue": "Skill 3 Accuracy"
        }
    ]

    requests = [
        # Checkbox Formula
        {
            'repeatCell': {
                'range': checkbox_range,
                'cell': {
                    'userEnteredValue': {
                        'formulaValue': '=IF(NOT(ISNA(AND(AND(HLOOKUP(I2,N2:Q2,1,False)=I2, HLOOKUP(J2,N2:Q2,1,False)=J2)=TRUE,AND(HLOOKUP(K2,R2:U2,1,False)=K2, HLOOKUP(L2,R2:U2,1,False)=L2)=TRUE)=TRUE)), TRUE,FALSE)'
                    }
                },
                'fields': 'userEnteredValue'
            }
        },
        # Apply the checkboxes
        {
            'setDataValidation': {
                'range': checkbox_range,
                'rule': {
                    'condition': {
                        'type': "BOOLEAN",
                    }
                }
            }
        },
        # Attribute Artifact Drop Down List
        {
            'setDataValidation': {
                'range': {
                    'sheetId': sheet_id,
                    'startRowIndex': 1,
                    'endRowIndex': 999,
                    'startColumnIndex': 8,
                    'endColumnIndex': 10
                },
                'rule': {
                    'condition': {
                        'type': 'ONE_OF_LIST',
                        'values': values_dropdown_attribute
                    },
                    'showCustomUi': True,
                    'strict': True
                }
            }
        },
        # Type Artifact Drop Down List
        {
            'setDataValidation': {
                'range': {
                    'sheetId': sheet_id,
                    'startRowIndex': 1,
                    'endRowIndex': 999,
                    'startColumnIndex': 10,
                    'endColumnIndex': 12
                },
                'rule': {
                    'condition': {
                        'type': 'ONE_OF_LIST',
                        'values': values_dropdown_type
                    },
                    'showCustomUi': True,
                    'strict': True
                }
            }
        }
    ]
    body = {
        'requests': requests
    }

    results = service.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id, body=body).execute()

    if text == "1":
        return

    # Add conditional formatting

    element_range = [
        {
            'sheetId': sheet_id,
            'startColumnIndex': 2,
            'endColumnIndex': 3,
            'startRowIndex': 1,
            'endRowIndex': 999
        }
    ]
    attrArtifact_range = [
        {
            'sheetId': sheet_id,
            'startColumnIndex': 8,
            'endColumnIndex': 10,
            'startRowIndex': 1,
            'endRowIndex': 999
        }
    ]
    typeArtifact_range = [
        {
            'sheetId': sheet_id,
            'startColumnIndex': 10,
            'endColumnIndex': 12,
            'startRowIndex': 1,
            'endRowIndex': 999
        }
    ]

    requests = [
        # Wind
        {
            # Wind
            'addConditionalFormatRule': {
                'rule': {
                    'ranges': element_range,
                    'booleanRule': {
                        'condition': {
                            'type': "TEXT_CONTAINS",
                            'values': [
                                {
                                    'userEnteredValue': 'Wind'
                                }
                            ]
                        },
                        'format': {
                            'backgroundColor': {
                                'green': 0.949,
                                'red': 1,
                                'blue': 0.8,
                            }
                        }
                    }
                },
                'index': 0
            }
        },
        # Fire
        {
            'addConditionalFormatRule': {
                'rule': {
                    'ranges': element_range,
                    'booleanRule': {
                        'condition': {
                            'type': "TEXT_CONTAINS",
                            'values': [
                                {
                                    'userEnteredValue': 'Fire'
                                }
                            ]
                        },
                        'format': {
                            'backgroundColor': {
                                'red': 0.956,
                                'green': 0.8,
                                'blue': 0.8,
                            }
                        }
                    }
                },
                'index': 1
            }
        },
        # Water
        {
            'addConditionalFormatRule': {
                'rule': {
                    'ranges': element_range,
                    'booleanRule': {
                        'condition': {
                            'type': "TEXT_CONTAINS",
                            'values': [
                                {
                                    'userEnteredValue': 'Water'
                                }
                            ]
                        },
                        'format': {
                            'backgroundColor': {
                                'red': 0.815,
                                'green': 0.878,
                                'blue': 0.890,
                            }
                        }
                    }
                },
                'index': 2
            }
        },
        # Light
        {
            'addConditionalFormatRule': {
                'rule': {
                    'ranges': element_range,
                    'booleanRule': {
                        'condition': {
                            'type': "TEXT_CONTAINS",
                            'values': [
                                {
                                    'userEnteredValue': 'Light'
                                }
                            ]
                        },
                        'format': {
                            'backgroundColor': {
                                'red': 0.952,
                                'green': 0.952,
                                'blue': 0.952,
                            }
                        }
                    }
                },
                'index': 3
            }
        },
        # Dark
        {
            'addConditionalFormatRule': {
                'rule': {
                    'ranges': element_range,
                    'booleanRule': {
                        'condition': {
                            'type': "TEXT_CONTAINS",
                            'values': [
                                {
                                    'userEnteredValue': 'Dark'
                                }
                            ]
                        },
                        'format': {
                            'backgroundColor': {
                                'red': 0.6,
                                'green': 0.6,
                                'blue': 0.6,
                            }
                        }
                    }
                },
                'index': 4
            }
        },
        # Attribute Artifact Color if in subs
        {
            'addConditionalFormatRule': {
                'rule': {
                    'ranges': attrArtifact_range,
                    'booleanRule': {
                        'condition': {
                            'type': 'CUSTOM_FORMULA',
                            'values': [
                                {
                                    'userEnteredValue': '=HLOOKUP(I2,$N2:$Q2,1,False) = I2'
                                }
                            ]
                        },
                        'format': {
                            'backgroundColor': {
                                'red': 0.717,
                                'green': 0.882,
                                'blue': 0.803,
                            }
                        }
                    }
                },
                'index': 0
            }
        },
        # Type Artifact Color if in subs
        {
            'addConditionalFormatRule': {
                'rule': {
                    'ranges': typeArtifact_range,
                    'booleanRule': {
                        'condition': {
                            'type': 'CUSTOM_FORMULA',
                            'values': [
                                {
                                    'userEnteredValue': '=HLOOKUP(K2,$R2:$U2,1,False) = K2'
                                }
                            ]
                        },
                        'format': {
                            'backgroundColor': {
                                'red': 0.717,
                                'green': 0.882,
                                'blue': 0.803,
                            }
                        }
                    }
                },
                'index': 0
            }
        }
    ]
    body = {
        'requests': requests
    }

    results = service.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id, body=body).execute()

    if text == "0":
        print('~~~ First time setup completed ~~~ \n')
    else:
        print('~~~ Layout update completed ~~~ \n')


if __name__ == '__main__':

    setupsheet()

    print('~~~ Updating Sheet ~~~')
    main()
    print('~~~ Sheet updated ~~~')
