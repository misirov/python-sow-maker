# Python SOW maker

Interact with the Google Sheets API to fetch values from a spreadsheet.


### Installation

It is recommended to use this project in a virtual environment.
```python
python -m venv python-sow-maker
cd python-sow-maker
git clone https://github.com/misirov/python-sow-maker
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
- [ ] Populate SOW template
    - [ ] Spreadsheet: create values used in SOW
    - [ ] Create SOW template
    - [ ] populate SOW
- [ ] Send populated SOW via Discord
