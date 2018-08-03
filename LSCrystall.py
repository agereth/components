import httplib2
import apiclient.discovery
from oauth2client.service_account import ServiceAccountCredentials
import datetime
import sys

CREDENTIALS_FILE = 'LSComponents.json'
spreadsheetId = '1sLBlQ4CB5Jg_2u6PHM8TWyguRj8-oWZGTSkgAXmszFw'

diode_indexes = {'алюминий':1, 'медь':2, 'green':3, 'red':4, 'royal blue': 5, 'orange': 6, 'arctic blue':7, 'deep red': 8, 'pc amber':9,
                 'white': 10}

color_picker = {'g':'green', 'gr':'green', 'green':'green',
                'r': 'red', 'red': 'red',
                'rb': 'royal blue', 'royalblue': 'royal blue', 'royal blue': 'royal blue',
                'o': 'orange', 'orange':'orange',
                'ab': 'arctic blue', 'arctic': 'arctic blue', 'arcticblue': 'arctic blue', 'arctic blue': 'arctic blue',
                'deep red': 'deep red', 'dr':'deep red', 'deepred': 'deep red',
                'pc amber': 'pc amber', 'pc': 'pc amber', 'pcamber':'pcamber',
                'w':'white', 'white':'white'}

threshold_board = 20
threshold_crystal = 10

def check_quantity(new_count: int, component: str,  threshold: int, component_type: str)->str:
    """
    check if new count of component is less then threshold (20 for boards, 10 for crystals)
    :param new_count: current component quantity
    :param component: type of crystal or board
    :param threshold: threshold
    :param component_type: boards ('подложки') or crystals 'кристаллы'
    :return: nothing if there is enough, "МАЛО!" if 0 < new_count < threshold, "НЕТ!" if  new_count = 0
    """

    if new_count <= threshold:
        if new_count == 0:
            print('%s %s кончились!' % (component_type, component))
            return 'НЕТ!'
        print("%s %s кончаются, осталось %i" % (component_type, component, new_count))
        return "МАЛО!"
    return ""


def update_crystals(diode: str, values: list, now: datetime, service: apiclient.discovery.Resource) -> bool:
    """
    function for decreasing quantity of available crystals and boards
    it gets string with diode information (three colours of crystals, type of board, quantity
    and modifies all necessary rows in googlesheet

    :param diode: string with type of dioide and quantity
    :param values: quantity of diodes from spreadsheet
    :param now: current date and time
    :param service: service for work with  googlesheets
    :return: true if updated successfully
    """

    data = diode.split()
    count = int(data[-1])
    board_type = data[-2]
    crystals = data[:-2]

    try:
        row = values[diode_indexes[board_type]]
    except IndexError:
        print ("Wrong board type")
        return False

    new_count = (int(row[1]) - count)
    row[1] = str(new_count)
    row[3] = row[3] + ', %s  %d.%d.%d' % (diode, now.day, now.month, now.year)
    row[2] = check_quantity(new_count, board_type, threshold_board, "Подложки")

    request_body = {"valueInputOption": "RAW", "data": [{"range": "a%i:d%i" % (diode_indexes[board_type]+1, diode_indexes[board_type]+1),
                                                         "values": [row]}]}
    request = service.spreadsheets().values().batchUpdate(spreadsheetId=spreadsheetId, body=request_body)
    _ = request.execute()

    crystals = [crystal.lower().strip() for crystal in crystals]
    if crystals == ['rgb']:
        colors =('red', 'green', 'royal blue')
    else:
        colors = []
        crystals = ' '.join(crystals)
        crystals = crystals.split('-')
        crystals = [crystal.strip() for crystal in crystals]
        for crystal in crystals:
            try:
                colors.append(color_picker[crystal])
            except IndexError:
                print("Can't recognize %s crystal" % crystal)
                return False

    for color in colors:
        row = values[diode_indexes[color]]
        new_count = (int(row[1]) - count)
        row[1] = str(new_count)
        row[3] = row[3] + ', %s  %d.%d.%d' % (diode, now.day, now.month, now.year)
        row[2] = check_quantity(new_count, color, threshold_crystal, "Кристаллы")

        request_body = {"valueInputOption": "RAW",
                        "data": [{"range": "a%i:d%i" % (diode_indexes[color]+1, diode_indexes[color]+1),
                                  "values": [row]}]}
        request = service.spreadsheets().values().batchUpdate(spreadsheetId=spreadsheetId, body=request_body)
        _ = request.execute()

    return True


def main (diodes: str):
    """
    updates spreadsheet with crystals and boards quantity
    :param diodes: list of diodes
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
