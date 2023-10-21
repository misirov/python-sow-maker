from sow_filler import SOWFiller
from spreadsheet import Spreadsheet
from datetime import datetime, timedelta


def main():
    # Spreadsheet ID
    SPREADSHEET_ID = '1tmQe3KLCsw1RwR325s_u1ADGOtaBCUQJOa0yn2OU7V0'
    # Range of values to retrieve. E,g., range of values from cell A1 until cell B16 
    RANGE_NAME = 'Sheet1!A1:B18'
    # Initialize class
    spreadsheet = Spreadsheet(SPREADSHEET_ID, RANGE_NAME)
    # Retrieve values. Returns a list of lists
    data = spreadsheet.fetch_data()
    # Convert list of lists into a dictionary and format according to the docx placeholder template
    replacements = {item[0]: item[1] for item in data if len(item) == 2}
    formatted_replacements = {"{" + key + "}": value for key, value in replacements.items()}
    
    print(formatted_replacements)


    # today = datetime.today()
    # day_issued = today.strftime("%m-%d-%Y")
    # five_days_later = today + timedelta(days=5)
    # day_expires = five_days_later.strftime("%m-%d-%Y")


    template_path = "templates/sow_template.docx"
    sow_filler = SOWFiller(template_path)
    sow_filler.replace_placeholders(formatted_replacements)





if __name__ == "__main__":
    main()