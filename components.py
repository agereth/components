import httplib2
import apiclient.discovery
from oauth2client.service_account import ServiceAccountCredentials
import datetime
import sys

CREDENTIALS_FILE = 'LSComponents.json'




def main(spreadsheetId, quantity):
    quantity = int(quantity)
    credentials = ServiceAccountCredentials.from_json_keyfile_name(CREDENTIALS_FILE, ['https://www.googleapis.com/auth/spreadsheets',
                                                                                      'https://www.googleapis.com/auth/drive'])
    httpAuth = credentials.authorize(httplib2.Http())
    service = apiclient.discovery.build('sheets', 'v4', http=httpAuth)

    answer = service.spreadsheets().values().get(spreadsheetId=spreadsheetId, range='a1:h60').execute()
    values = answer.get('values', [])
    now = datetime.datetime.now()

    result = []
    if 'Total' in values[0]:
        i_total = values[0].index('Total')
    if 'Кол-во' in values[0]:
        i_total = values[0].index('Кол-во')
    if 'На плату' in values[0]:
        i_num = values[0].index('На плату')
    if 'Quantity' in values[0]:
        i_num = values[0].index('Quantity')
    if 'Комментарий' in values[0]:
        i_comm = values[0].index('Комментарий')
    if 'Notes' in values[0]:
        i_comm = values[0].index('Notes')

    for row in values[1:]:
        new_row = row
        new_row[i_total] = str(int(row[i_total]) - quantity*int(row[i_num]))
        new_row[i_comm] = row[i_comm] +" ,%i плат %d.%d.%d" % (quantity, now.day, now.month, now.year)
        result.append(new_row)

    request_body = {"valueInputOption" : "RAW","data" : [{"range":"a2:h60", "values":result}]}
    request = service.spreadsheets().values().batchUpdate(spreadsheetId=spreadsheetId, body=request_body)
    _ = request.execute()

if __name__ == '__main__':
    if len(sys.argv) > 1:
        main(sys.argv[1], sys.argv[2])
    else:
        print("Shreadsheet id and Number of boards required")