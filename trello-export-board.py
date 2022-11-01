#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Exports Trello Board, Cards and attachments.

Requirements:
- Python 3
- docxtpl (https://pypi.org/project/docxtpl/)
- python-dateutil (https://pypi.org/project/python-dateutil/)
- requests (https://pypi.org/project/requests/)
- XlsxWriter (https://pypi.org/project/XlsxWriter/)

Configure:
`config.ini`
- `[Dates]`: time zone and date format
- `[TrelloApi]`: api key, token, url
- `[Proxy]`: proxy configuration
- `[Labels]`: custom titles
- `[Template]`: docx template

Usage:
`python3 trello-export-board.py`
"""

import configparser
import json
import os
import re
import requests
import shutil
import sys
import unicodedata
import xlsxwriter

from dateutil.parser import parse
from dateutil import tz
from datetime import datetime
from docxtpl import DocxTemplate
from io import BytesIO
from os.path import exists as path_exists
from os import remove


def get_config_from_ini():
    """
    Read configuration file (config.ini)
    :return: configparser.ConfigParser
    """
    if not os.path.exists('config.ini'):
        shutil.copyfile('config-sample.ini', 'config.ini')
        #print("error : config.ini missing\n")
        #sys.exit()
    config = configparser.ConfigParser()
    config.read('config.ini', encoding='utf-8')
    return config

def lists_get_name(listId, haystack):
    """
    Search for a given list name in list (haystack)
    :param listId: String list id (Pattern: ^[0-9a-fA-F]$)
    :param haystack: 2D Array lists (lists[[id],[name],[pos]])
        haystack[i][0] : 'id'
        haystack[i][1] : 'name'
        haystack[i][2] : 'pos'
    :return: String list name or empty string
    """
    for row in haystack:
        if row[0] == listId:
            return str(row[1])
    return ''

def sanitize_filename(filename):
    """
    Turn a string into a valid filename
    :param filename: String string to sanitize
    :return: String filename
    """
    filename = unicodedata.normalize('NFKD', filename).encode('ascii', 'ignore').decode('utf8')
    filename = re.sub(r'[^\.\w\s-]', '', filename)
    filename = re.sub(r'[-_\s]+', '-', filename)
    filename = re.sub(r'[\.]+', '.', filename)
    filename = filename.strip('-_.')
    return filename

def escape_xml(str_input):
    """
    Escape XML chars : '&', '>', '<'
    :param str_input: String
    :return: String
    """
    return str_input.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")

def convert_UTC_to_Local_Datetime(strDate):
    """
    Convert UTC date to local datetime
    :param strDate: String date
    :return: datetime
    """
    from_zone = tz.gettz(config['Dates']['tz_from_zone'])
    to_zone = tz.gettz(config['Dates']['tz_to_zone'])
    utc = parse(strDate)
    utc = utc.replace(tzinfo=from_zone)
    local = utc.astimezone(to_zone)
    #return local.replace(tzinfo=None)
    return local

def remove_if_exists(fileName):
    """
    Check if fileName exixts. Remove fileName if exists
    :param fileName: String
    """
    if path_exists(fileName):
        try:
            remove(fileName)
        except:
            print("Oops, an error occurred.")
            print(f"'{fileName}' already exists.\nYou should rename or delete '{boardFileName}.")
            sys.exit()



# Configure
# ------------------------------
# Read config
config = get_config_from_ini()
# Set proxies
proxies = ''
if str(config['Proxy']['use']) == 'True':
    proxies = {
       'http': config['Proxy']['http'],
       'https': config['Proxy']['https'],
    }
# Set headers
headers={
    'Cache-Control': 'no-cache',
    'Content-Type': 'application/json; charset=utf-8',
    'Accept': 'application/json'
}
attachmentsHeaders = {
    "Authorization": f"OAuth oauth_consumer_key=\"{config['TrelloApi']['apiKey']}\", oauth_token=\"{config['TrelloApi']['apiToken']}\""
}
# Get base_url
base_url = config['TrelloApi']['apiUrl']
# Set paths
# TODO : set path in config.ini
path = './exports/'
pathToCards = path + 'cards/'
pathToArchived = path + 'archived/'
pathToAttachments = path + 'attachments/'
if not path_exists(path):
    os.makedirs(path)
if not path_exists(pathToCards):
    os.makedirs(pathToCards)
if not path_exists(pathToArchived):
    os.makedirs(pathToArchived)
if not path_exists(pathToAttachments):
    os.makedirs(pathToAttachments)



# Step 1 : Get boards
# ------------------------------
url = base_url + "members/me/boards"
query = {
   'key': config['TrelloApi']['apiKey'],
   'token': config['TrelloApi']['apiToken'],
   'fields': 'id,name,desc'
}
boards = []
try:
    response = requests.request("GET",url,proxies=proxies,params=query,headers=headers)
except:
    # TODO : add most common reason (Internet cnx, proxy...)
    print("Oops, an error occurred.")
    sys.exit()
if response.status_code == 200:
    json_data = json.loads(response.text)
    # To be able to sort results by name, we create a List
    for i in range(len(json_data)):
        boards.append([
            json_data[i]['id'],
            json_data[i]['name'],
            json_data[i]['desc'],
        ])
    boards = sorted(boards, key=lambda x: x[1])
else:
    print("Oops, an error occurred.")
    print(f"Cannot retrieve boards that you are member of.")
    print(f"Response code : {response.status_code}")
    sys.exit()



# Step 2 : Select a board
# ------------------------------
num = 0
if len(boards) > 1:
    print("------------------------------")
    for i in range(len(boards)):
        print( f'{i:4}: {boards[i][1]}')
    print("------------------------------")
    while True:
        try:
            num = int(input("Select a board : "))
            if num < 0 or num >= len(boards):
                raise ValueError()
        except ValueError:
            print("This is not a valid board number.")
            continue
        else:
            break
boardId   = boards[num][0]
boardName = boards[num][1]
boardDesc = boards[num][2]



# Step 3 : Get lists
# ------------------------------
url = base_url + f"boards/{boardId}/lists"
query = {
   'key': config['TrelloApi']['apiKey'],
   'token': config['TrelloApi']['apiToken'],
   'fields': 'id,name,pos'
}
response = requests.request("GET",url,proxies=proxies,params=query,headers=headers)
lists = []
if response.status_code == 200:
    json_data = json.loads(response.text)
    # To be able to sort results (by pos), we create a List
    for i in range(len(json_data)):
        lists.append([
            json_data[i]['id'],
            json_data[i]['name'],
            json_data[i]['pos'],
        ])
    lists = sorted(lists, key=lambda x: x[2])
else:
    print("Oops, an error occurred.")
    print(f"Cannot retrieve Lists on '{boardName}'")
    print(f"Response code : {response.status_code}")
    sys.exit()



# Step 4 : Export board to XLSX
# ------------------------------
print(f"Exporting board '{boardName}'...")

url = base_url + f"boards/{boardId}/cards/all"
query = {
   'key': config['TrelloApi']['apiKey'],
   'token': config['TrelloApi']['apiToken'],
   'fields': 'id,name,desc,idList,labels,start,due,dueComplete,dateLastActivity,closed,idShort,shortUrl'
}
response = requests.request("GET",url,proxies=proxies,params=query,headers=headers)
if response.status_code == 200:
    cards = json.loads(response.text)
else:
    print("Oops, an error occurred.")
    print(f"Cannot retrieve Board '{boardName}'")
    print(f"Response code : {response.status_code}")
    sys.exit()

if len(cards) == 0:
    print("No Card on this Board.")
    sys.exit()

# Set board filename
boardFileName = sanitize_filename(boardName)
boardFileName = boardFileName[:250] + ".xlsx"

# Remove boardFileName if exists
remove_if_exists(path + boardFileName)

# Create workbook
workbook = xlsxwriter.Workbook(path + boardFileName)
worksheet1 = workbook.add_worksheet(config['Labels']['sheet_opened_cards'])
worksheet2 = workbook.add_worksheet(config['Labels']['sheet_archived_cards'])

# Worksheet setup
cellHeight = 15
worksheet1.set_portrait()
worksheet1.set_default_row(cellHeight)
worksheet1.set_column('A:A', 20)
worksheet1.set_column('B:B', 45)
worksheet1.set_column('C:C', 50)
worksheet1.set_column('D:D', 16)
worksheet1.set_column('E:E', 16)
worksheet1.set_column('F:F', 16)
worksheet1.set_column('G:G', 20)
worksheet1.set_column('H:H', 10)
worksheet1.set_column('I:I', 30)

worksheet2.set_portrait()
worksheet2.set_default_row(cellHeight)
worksheet2.set_column('A:A', 20)
worksheet2.set_column('B:B', 45)
worksheet2.set_column('C:C', 50)
worksheet2.set_column('D:D', 16)
worksheet2.set_column('E:E', 16)
worksheet2.set_column('F:F', 16)
worksheet2.set_column('G:G', 20)
worksheet2.set_column('H:H', 10)
worksheet2.set_column('I:I', 30)

# Cell Format
cell_text_title = workbook.add_format()
cell_text_title.set_align('left')
cell_text_title.set_align('top')
cell_text_title.set_bold()
cell_text_title.set_text_wrap()

cell_text_left = workbook.add_format()
cell_text_left.set_align('left')
cell_text_left.set_align('top')
cell_text_left.set_text_wrap()

cell_text_center = workbook.add_format()
cell_text_center.set_align('center')
cell_text_center.set_align('top')
cell_text_center.set_text_wrap()

# Set headers
wsHeaders = [
    config['Labels']['ws_header_list'],
    config['Labels']['ws_header_title'],
    config['Labels']['ws_header_description'],
    config['Labels']['ws_header_start_date'],
    config['Labels']['ws_header_due_date'],
    config['Labels']['ws_header_last_activity_date'],
    config['Labels']['ws_header_labels'],
    config['Labels']['ws_header_num'],
    config['Labels']['ws_header_url']
    ]
for row_num, data in enumerate(wsHeaders):
    worksheet1.write(0, row_num, data, cell_text_title)
    worksheet2.write(0, row_num, data, cell_text_title)

worksheet1cellRow = 1
worksheet2cellRow = 1
for i in range(len(cards)):
    listName = lists_get_name(cards[i]['idList'], lists)
    cardStart = ''
    if cards[i]['start'] is not None:
        cardStart = convert_UTC_to_Local_Datetime(cards[i]['start']).strftime(config['Dates']['str_date_format'])
    cardDue = ''
    if cards[i]['due'] is not None:
        cardDue = convert_UTC_to_Local_Datetime(cards[i]['due']).strftime(config['Dates']['str_datetime_format'])
    cardLastActivity = ''
    if cards[i]['dateLastActivity'] is not None:
        cardLastActivity = convert_UTC_to_Local_Datetime(cards[i]['dateLastActivity']).strftime(config['Dates']['str_datetime_format'])
    cardLabels = ''
    if len(cards[i]['labels']) > 0:
        arrLabels = []
        for j in range(len(cards[i]['labels'])):
            if cards[i]['labels'][j]['name'] != "":
                arrLabels.append(cards[i]['labels'][j]['name'])
        cardLabels = ", ".join(arrLabels)

    if str(cards[i]['closed']) == 'False':
        # Card is open
        worksheet1.write(worksheet1cellRow, 0, listName, cell_text_left)
        worksheet1.write(worksheet1cellRow, 1, cards[i]['name'], cell_text_left)
        worksheet1.write(worksheet1cellRow, 2, cards[i]['desc'], cell_text_left)
        worksheet1.write(worksheet1cellRow, 3, cardStart, cell_text_center)
        worksheet1.write(worksheet1cellRow, 4, cardDue, cell_text_center)
        worksheet1.write(worksheet1cellRow, 5, cardLastActivity, cell_text_center)
        worksheet1.write(worksheet1cellRow, 6, cardLabels, cell_text_left)
        worksheet1.write(worksheet1cellRow, 7, cards[i]['idShort'], cell_text_center)
        worksheet1.write(worksheet1cellRow, 8, cards[i]['shortUrl'], cell_text_left)
        worksheet1cellRow += 1
    else:
        # Card was archived
        worksheet2.write(worksheet2cellRow, 0, listName, cell_text_left)
        worksheet2.write(worksheet2cellRow, 1, cards[i]['name'], cell_text_left)
        worksheet2.write(worksheet2cellRow, 2, cards[i]['desc'], cell_text_left)
        worksheet2.write(worksheet2cellRow, 3, cardStart, cell_text_center)
        worksheet2.write(worksheet2cellRow, 4, cardDue, cell_text_center)
        worksheet2.write(worksheet2cellRow, 5, cardLastActivity, cell_text_center)
        worksheet2.write(worksheet2cellRow, 6, cardLabels, cell_text_left)
        worksheet2.write(worksheet2cellRow, 7, cards[i]['idShort'], cell_text_center)
        worksheet2.write(worksheet2cellRow, 8, cards[i]['shortUrl'], cell_text_left)
        worksheet2cellRow += 1

worksheet1.autofilter(0, 0, worksheet1cellRow-1, 7)
worksheet2.autofilter(0, 0, worksheet2cellRow-1, 7)
workbook.close()



# Step 4 : Export cards to DOCX
# ------------------------------
print("Exporting cards...")

# For each card
for i in range(len(cards)):
    url = base_url + f"cards/{cards[i]['id']}"
    query = {
       'key': config['TrelloApi']['apiKey'],
       'token': config['TrelloApi']['apiToken'],
       'fields': 'id,closed,dueComplete,dateLastActivity,desc,due,idList,idShort,labels,name,start',
       'actions': 'commentCard,updateCheckItemStateOnCard',
       'attachments': 'true',
       'checklists': 'all'
    }
    response = requests.request("GET",url,proxies=proxies,params=query,headers=headers)
    if response.status_code == 200:
        card = json.loads(response.text)
    else:
        print("Oops, an error occurred.")
        sys.exit()

    print(f"[{i+1}/{len(cards)}] Card #{card['idShort']} '{card['name']}'")
    document = DocxTemplate('./templates/' + config['Template']['template'])

    # Labels
    cardLabels = ''
    if len(card['labels']) > 0:
        arrLabels = []
        for i in range(len(card['labels'])):
            if card['labels'][i]['name'] != "":
                arrLabels.append( escape_xml(card['labels'][i]['name']) )
        cardLabels = ", ".join(arrLabels)
    
    # Dates
    cardStart = ''
    if card['start'] is not None:
        cardStart = convert_UTC_to_Local_Datetime(card['start']).strftime(config['Dates']['str_date_format'])
    cardDue = ''
    if card['due'] is not None:
        cardDue = convert_UTC_to_Local_Datetime(card['due']).strftime(config['Dates']['str_datetime_format'])
    cardLastActivity = ''
    if card['dateLastActivity'] is not None:
        cardLastActivity = convert_UTC_to_Local_Datetime(card['dateLastActivity']).strftime(config['Dates']['str_datetime_format'])
    
    # Checklists
    checklists = []
    if 'checklists' in card:
        if len(card['checklists']) > 0:
            checklists = []
            # get checklists
            for i in range(len(card['checklists'])):
                if card['checklists'][i]['name'] != "":
                    checkItems = []
                    pcentComplet = ''
                    # get checkItems
                    if card['checklists'][i]['checkItems']:
                        if len(card['checklists'][i]['checkItems']) > 0:
                            checklistComplete = 0
                            for j in range(len(card['checklists'][i]['checkItems'])):
                                # Get last 'updateCheckItemStateOnCard' date
                                updateCheckItemStateDate = ''
                                if 'actions' in card:
                                    if len(card['actions']) > 0:
                                        for k in range(len(card['actions'])):
                                            if card['actions'][k]['type'] == 'updateCheckItemStateOnCard':
                                                if card['actions'][k]['data']['checkItem']['id'] == card['checklists'][i]['checkItems'][j]['id']:
                                                    updateCheckItemStateDate = convert_UTC_to_Local_Datetime(card['actions'][k]['date']).strftime(config['Dates']['str_date_format'])
                                                    break
                                # checkItem status
                                if card['checklists'][i]['checkItems'][j]['state'] == 'complete':
                                    checklistComplete += 1
                                # push checkItem
                                checkItems.append([
                                    escape_xml(card['checklists'][i]['checkItems'][j]['name']),
                                    card['checklists'][i]['checkItems'][j]['pos'],
                                    card['checklists'][i]['checkItems'][j]['state'],
                                    updateCheckItemStateDate
                                ])
                            # order checkItems by position (pos)
                            if len(checkItems) > 0:
                                checkItems =  sorted(checkItems, key=lambda x: x[1])    
                            # pcent complete
                            pcentComplet =  str(int(round((checklistComplete / len(card['checklists'][i]['checkItems'])) * 100, 0))) + '%'
                    # push checklist
                    checklists.append([
                        escape_xml(card['checklists'][i]['name']),
                        card['checklists'][i]['pos'],
                        checkItems,
                        pcentComplet
                    ])
                # order checklists by position (pos)
                if len(checklists) > 0:
                    checklists = sorted(checklists, key=lambda x: x[1])

    # Actions
    actions = []
    if 'actions' in card:
        if len(card['actions']) > 0:
            for i in range(len(card['actions'])):
                if card['actions'][i]["type"] == "commentCard":
                    actions.append([
                        convert_UTC_to_Local_Datetime(card['actions'][i]['date']).strftime(config['Dates']['str_datetime_format']),
                        escape_xml(card['actions'][i]['memberCreator'][config['Labels']['user_name']]),
                        escape_xml(card['actions'][i]["data"]["text"])
                    ])

    # Attachments
    attachments = []
    if 'attachments' in card:
        if len(card['attachments']) > 0:
            for i in range(len(card['attachments'])):
                # Download attachment
                attachmentFileName = sanitize_filename( str(card['idShort']) + "-" + str(card['attachments'][i]['name']) )
                remove_if_exists(pathToAttachments + attachmentFileName)
                url = card['attachments'][i]['url']
                print(f"  Downloading '{attachmentFileName}'")
                try:
                    response = requests.request("GET",url,headers=attachmentsHeaders)
                except:
                    print("  Oops, an error occurred.")
                    print(f"  Cannot download '{attachmentFileName}'")
                    print(f"  Response code : {response.status_code}")
                    continue
                bytesio_object = BytesIO(response.content)
                # Write the stuff
                with open(pathToAttachments + attachmentFileName, "wb") as f:
                    f.write(bytesio_object.getbuffer())
                # push attachment
                attachments.append([
                    attachmentFileName,
                    convert_UTC_to_Local_Datetime(card['attachments'][i]['date']).strftime(config['Dates']['str_date_format'])
                ])

    # Template context             
    context = {}
    context["title"] = escape_xml(card['name'])
    context["list"] = escape_xml(lists_get_name(card['idList'], lists))
    context["labels"] = escape_xml(cardLabels)
    context["startDate"] = cardStart
    context["dueDate"] = cardDue
    context["lastActivityDate"] = cardLastActivity
    context["description"] = escape_xml(card['desc'])
    context["checklists"] = checklists
    context["actions"] = actions
    context["attachments"] = attachments

    # Render
    document.render(context)
    outputFileName = sanitize_filename(card['name'])
    outputFileName = outputFileName[:250] + ".docx"

    # Save
    if str(card['closed']) == 'False':
        remove_if_exists(pathToCards + outputFileName)
        document.save(pathToCards + outputFileName)
    else:
        remove_if_exists(pathToArchived + outputFileName)
        document.save(pathToArchived + outputFileName)


print("Done.")
