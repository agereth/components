import httplib2
import apiclient.discovery
from oauth2client.service_account import ServiceAccountCredentials
import datetime
import sys

CREDENTIALS_FILE = 'LSComponents.json'
spreadsheetId = '1sLBlQ4CB5Jg_2u6PHM8TWyguRj8-oWZGTSkgAXmszFw'

diode_indexes = {'алюминий':1, 'медь':2, 'green':3, 'red':4, 'royal blue': 5, 'orange': 6, 'arctic blue':7, 'deep red': 8, 'pc amber':9,
                 'white': 10}

color_codes = {'green': ('gr', 'green', 'g'),
               'red': ('red', 'r'),
               'royal blue': ('rb', 'royal blue', 'royalblue'),
               'orange':('o', 'orange'),
               'arctic blue': ('ab', 'arctic', 'arcticblue', 'arctic blue'),
               'deep red': ('dr', 'deepred', 'deep red'),
               'pc amber': ('pc', 'amber', 'pcamber', 'pc amber'),
               'white':('w', 'white')}

threshold_board = 20
threshold_crystal = 10


def update_crystals(diode: list, values: list, now:datetime, service)->bool:
    """
    :param diode:
    :return:
    """
    data = diode.split()
    count = int(data[-1])
    board_type = data[-2]
    crystals = data[:-2]

    row = values[diode_indexes[board_type]]
    new_count = (int(row[1]) - count)
    row[1] = str(new_count)
    row[3] = row[3] + ', %s  %d.%d.%d' % (diode, now.day, now.month, now.year)
    if new_count <= threshold_board:
        if new_count == 0:
            print('Подложки %s кончились!' % board_type)
            row[2] = 'НЕТ!'
        else:
            print("Подложки %s кончаются, осталось %i" % (board_type, new_count))
            row[2] = "МАЛО!"
    else:
        row[2] = ""

    request_body = {"valueInputOption": "RAW", "data": [{"range": "a%i:d%i" % (diode_indexes[board_type]+1, diode_indexes[board_type]+1),
                                                         "values": [row]}]}
    request = service.spreadsheets().values().batchUpdate(spreadsheetId=spreadsheetId, body=request_body)
    _ = request.execute()

    crystals = [crystal.lower().strip() for crystal in crystals]
    if crystals == ['rgb']:
        colors =('red', 'green', 'royal blue')
    else:
        colors = []
        crystals = (' ').join(crystals)
        crystals = crystals.split('-')
        crystals = [crystal.strip() for crystal in crystals]
        for crystal in crystals:
            for (color, codes) in color_codes.items():
                if crystal in codes:
                    colors.append(color)

    if len(colors) < 3:
        print ("Wrong diode %s" % data)



    for color in colors:
        row = values[diode_indexes[color]]
        new_count = (int(row[1]) - count)
        row[1] = str(new_count)
        row[3] = row[3] + ', %s  %d.%d.%d' % (diode, now.day, now.month, now.year)
        if new_count <= threshold_crystal:
            print("Кристаллы %s кончаются" % color)
            row[2] = "МАЛО!"
        else:
            row[2] = ""
        request_body = {"valueInputOption": "RAW",
                        "data": [{"range": "a%i:d%i" % (diode_indexes[color]+1, diode_indexes[color]+1),
                                  "values": [row]}]}
        request = service.spreadsheets().values().batchUpdate(spreadsheetId=spreadsheetId, body=request_body)
        _ = request.execute()








def main (diodes: str):
    """

    :param file:
    :return:
    """

    credentials = ServiceAccountCredentials.from_json_keyfile_name(CREDENTIALS_FILE,
                                                                   ['https://www.googleapis.com/auth/spreadsheets',
                                                                    'https://www.googleapis.com/auth/drive'])
    httpAuth = credentials.authorize(httplib2.Http())
    service = apiclient.discovery.build('sheets', 'v4', http=httpAuth)

    answer = service.spreadsheets().values().get(spreadsheetId=spreadsheetId, range='a1:d11').execute()
    values = answer.get('values', [])
    diodes = diodes.split(', ')
    now = datetime.datetime.now()
    for diode in diodes:
        update_crystals(diode, values, now, service)


if __name__ == '__main__':
    if len(sys.argv) > 1:
        main(sys.argv[1])
    else:
        print ("Diode list expected")
