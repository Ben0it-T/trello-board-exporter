# Trello Board Exporter

Exports Trello Board, Cards and attachments.

With a Trello 'free plan' we can only export a board or card to JSON... nothing else.
For my personnal use, I wrote this python script to be able to easily extract boards, cards and attachments.

This script,
- exports Board to XLSX document
- exports all cards on the Board to DOCX documents
- exports all attachments

## Requirements

- Python 3
- docxtpl (https://pypi.org/project/docxtpl/)
- python-dateutil (https://pypi.org/project/python-dateutil/)
- requests (https://pypi.org/project/requests/)
- XlsxWriter (https://pypi.org/project/XlsxWriter/)

## Configure

`config.ini` is a basic configuration file containing:
- `[Dates]`: time zone and date format
- `[TrelloApi]`: api key, token, url
- `[Proxy]`: proxy configuration
- `[Labels]`: custom titles
- `[Template]`: docx template

### Get Trello api key and token
- Get your api key : https://trello.com/app-key
- Generate a (read only) token : https://trello.com/1/authorize?expiration=30days&scope=read&response_type=token&name=PersonalToken&key={YourAPIKey}

To revoke a token : https://trello.com/my/account

### Create `config.ini`
Copy the `config-sample.ini` to `config.ini`
- add your api key and token
- customize dates, proxies, labels and template document

### Create your card DOCX template.
See templates in `templates/`

## Usage

Simply run 
```
python3 trello-export-board.py
```
Then select a board and enjoy !

