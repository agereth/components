import httplib2
import apiclient.discovery
from oauth2client.service_account import ServiceAccountCredentials
import datetime
import sys

CREDENTIALS_FILE = 'LSComponents.json'
spreadsheetId = '1NQdFG3qc20w3qTyEUpKVSsf03qzGgRsVUvyigEr-kak'
indexes = {
    'Sound': [1, 3, 5, 7, 8, 9, 11, 12, 13, 14, 19, 21, 25, 26, 27, 28, 29],
    'RGB': [11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22],
    'Simple': [13, 14, 15, 16, 17, 18, 19, 20]
}


def get_total_for_row(row: list, quantity: int, type: str)->int:
    """
    get total amount of components: last_total_amount - quantity_per_board*number_of_boards
    :param row: row with data (last_total_amount, price (unused), last operation (unused), quantity_per_board
    :return: new total of current component
    """
    if row[1].isdigit():
        amount_for_board = 0
        total = int(row[1])
        if row[4].isdigit():
            amount_for_board = int(row[4])
        else:
            amounts = row[4].split(', ')
            for amount in amounts:
                if type in amount:
                    amount_for_board = int(amount.split(type)[1].split(',')[0])
        total -= quantity*amount_for_board
    else:
        total = row[1]
    return total

def get_number_of_rows (type: str) -> int:
    """

    :param type: Sound, RGB or Simple
    :return: rows number
    """
    if type == 'Sound':
        return 32
    if type == 'RGB':
        return 25
    if type == Simple:
        return 23


def main(quantity, type):
    credentials = ServiceAccountCredentials.from_json_keyfile_name(CREDENTIALS_FILE, ['https://www.googleapis.com/auth/spreadsheets',
                                                                                      'https://www.googleapis.com/auth/drive'])
    httpAuth = credentials.authorize(httplib2.Http())
    service = apiclient.discovery.build('sheets', 'v4', http=httpAuth)

    answer = service.spreadsheets().values().get(spreadsheetId=spreadsheetId, range='a3:e17').execute()
    values = answer.get('values', [])
    now = datetime.datetime.now()

    result = []
    for row in values:
        total = get_total_for_row(row, int(quantity), type)
        new_row = [row[0], str(total), row[2], row[3]+" ,%i плат %d.%d.%d" % (int(quantity), now.day, now.month, now.year), row[4]]
        result.append(new_row)

    request_body = {"valueInputOption" : "RAW","data" : [{"range":"a3:e17", "values":result}]}
    request = service.spreadsheets().values().batchUpdate(spreadsheetId=spreadsheetId, body=request_body)
    _ = request.execute()

    answer = service.spreadsheets().values().get(spreadsheetId=spreadsheetId, range='%s!a3:e%i'
                                                                % (type, get_number_of_rows(type)) ).execute()
    values = answer.get('values', [])

    for i in indexes[type]:
        row = values[i]
        total = get_total_for_row(row, int(quantity), type)
        new_row = [row[0], str(total), row[2], row[3] + " ,%i плат %d.%d.%d" % (int(quantity), now.day, now.month, now.year),
                   row[4]]
        request_body = {"valueInputOption" : "RAW", "data" : [{"range" : "Sound!a%i:e3%i" % (i+3, i+3), "values" : [new_row]}]}
        request = service.spreadsheets().values().batchUpdate(spreadsheetId=spreadsheetId, body=request_body)
        _ = request.execute()

if __name__ == '__main__':
    if len(sys.argv) > 2:
        main(sys.argv[1], sys.argv[2])
    else:
        print("Number of boards and board type required")