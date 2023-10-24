# Python SOW maker

Interact with the Google Sheets API to fetch values from a spreadsheet.


### Installation

It is recommended to use this project in a virtual environment.
```python
git clone https://github.com/misirov/python-sow-maker
cd python-sow-maker
python -m venv .
pip install --upgrade requirements.txt
```


### Sources

- [API Python quickstart](https://developers.google.com/sheets/api/quickstart/python)
- [Testing Spreadsheet](https://docs.google.com/spreadsheets/d/1tmQe3KLCsw1RwR325s_u1ADGOtaBCUQJOa0yn2OU7V0/edit#gid=0)


### To Do's
- [x] Create a Google Workspace project  
    - [x] create oauth
    - [x] download `credentials.json`
    - [x] authenticate
- [x] Test basic functionality
    - [x] fetch values from spreadsheet
- [x] Populate SOW template
    - [x] spreadsheet: create values used in SOW
    - [x] create SOW template
    - [x] populate SOW
- [ ] Create Templates
    - [x] Cantina Managed
    - [x] Guild
    - [ ] vCISO
- [ ] Send populated SOW via Discord
    - [ ] create command
    - [ ] send document over discord
    - [ ] deploy 
