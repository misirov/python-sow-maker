from spreadsheet import Spreadsheet



def main():
    SAMPLE_SPREADSHEET_ID = '1tmQe3KLCsw1RwR325s_u1ADGOtaBCUQJOa0yn2OU7V0'
    SAMPLE_RANGE_NAME = 'Sheet1!A1:B16'

    spreadsheet = Spreadsheet(SAMPLE_SPREADSHEET_ID, SAMPLE_RANGE_NAME)
    data = spreadsheet.fetch_data()

    print('Printing values:')
    for row in data:
        print('%s, %s' % (row[0], row[1]))





if __name__ == '__main__':
    main()
