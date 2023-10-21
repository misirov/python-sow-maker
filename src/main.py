from sow_filler import SOWFiller
from spreadsheet import Spreadsheet

def main():
    # Spreadsheet ID
    SPREADSHEET_ID = '1tmQe3KLCsw1RwR325s_u1ADGOtaBCUQJOa0yn2OU7V0'
    # Range of values to retrieve. E.g., range of values from cell A1 until cell B16 
    RANGE_NAME = 'Sheet1!A1:E27'
    # Initialize class
    spreadsheet = Spreadsheet(SPREADSHEET_ID, RANGE_NAME)
    # Retrieve values. Returns a list of lists
    data = spreadsheet.fetch_data()
    print("This is the list containing the data in the spreadsheet\n\n", data)
    
    # Convert list of lists into a dictionary and format according to the docx placeholder template
    replacements = {}
    for row in data:
        for i in range(len(row)-1):  # No need to step by 2 this time
            key = row[i]
            value = row[i+1]
            if key and value:  # Ensure the key and value are not empty
                replacements[key] = value
    formatted_replacements = {"{" + key + "}": value for key, value in replacements.items()}
    
    print("\nThis is the created dictionary\n\n", formatted_replacements)

    template_path = "templates/sow_template.docx"
    sow_filler = SOWFiller(template_path)
    sow_filler.replace_placeholders(formatted_replacements)

if __name__ == "__main__":
    main()
