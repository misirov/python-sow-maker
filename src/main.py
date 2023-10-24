
import os
from dotenv import load_dotenv
from sow_generator import SOWGenerator


if __name__ == "__main__":
    # E.g: https://docs.google.com/spreadsheets/d/<SPREADSHEET_ID>/
    SPREADSHEET_ID = "1tmQe3KLCsw1RwR325s_u1ADGOtaBCUQJOa0yn2OU7V0"
    
    # SOWGenerator.cantina_managed(SPREADSHEET_ID)
    SOWGenerator.guild(SPREADSHEET_ID)