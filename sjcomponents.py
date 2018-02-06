import httplib2
import apiclient.discovery
from oauth2client.service_account import ServiceAccountCredentials
import datetime
import sys

CREDENTIALS_FILE = 'LSComponents.json'
spreadsheetId = '1B9rRrszIAgPAuOgDN_Ao67_bmMs3QnVK188PDGVOx-M'



def main(rxquantity, txquantity):
    rxquantity = int(rxquantity[2:])
    txquantity = int(txquantity[2:])

    credentials = ServiceAccountCredentials.from_json_keyfile_name(CREDENTIALS_FILE, ['https://www.googleapis.com/auth/spreadsheets',
                                                                                      'https://www.googleapis.com/auth/drive'])
    httpAuth = credentials.authorize(httplib2.Http())
    service = apiclient.discovery.build('sheets', 'v4', http=httpAuth)

    answer = service.spreadsheets().values().get(spreadsheetId=spreadsheetId, range='a1:h60').execute()
    values = answer.get('values', [])
    now = datetime.datetime.now()

    result = []
    i_num_rx = 0
    i_num_tx = 0
    if 'Total' in values[0]:
        i_total = values[0].index('Total')
    if 'Кол-во' in values[0]:
        i_total = values[0].index('Кол-во')
    if 'На плату RX' in values[0]:
        i_num_rx = values[0].index('На плату RX')
    if 'На плату TX' in values[0]:
        i_num_tx = values[0].index('На плату RX')
    if 'Quantity RX' in values[0]:
        i_num_rx = values[0].index('Quantity RX')
    if 'Quantity TX' in values[0]:
        i_num_tx = values[0].index('Quantity TX')
    if 'Комментарий' in values[0]:
        i_comm = values[0].index('Комментарий')
    if 'Notes' in values[0]:
        i_comm = values[0].index('Notes')

    for row in values[1:]:
        new_row = row
        new_row[i_total] = str(int(row[i_total]) - rxquantity*int(row[i_num_rx]) - txquantity*int(row[i_num_tx]))
        if rxquantity != 0:
            new_row[i_comm] = row[i_comm] +" ,%i плат RX %d.%d.%d" % (rxquantity, now.day, now.month, now.year)
        if txquantity != 0:
            new_row[i_comm] = row[i_comm] + " ,%i плат TX %d.%d.%d" % (txquantity, now.day, now.month, now.year)
        result.append(new_row)

    request_body = {"valueInputOption" : "RAW","data" : [{"range":"a2:h60", "values":result}]}
    request = service.spreadsheets().values().batchUpdate(spreadsheetId=spreadsheetId, body=request_body)
    _ = request.execute()

if __name__ == '__main__':
    if len(sys.argv) > 1:
        main(sys.argv[1], sys.argv[2])
    else:
        print("Number of RX boards required, Number of TX boards required")