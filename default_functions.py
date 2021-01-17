import time, requests, json
from termcolor import colored
import colorama


def eternity_cycle_deep():
    while True:
        input()


def print_col(message, color):
    colorama.init()
    print(colored(message, color))


def to_digit(symbol):
    if len(symbol) == 2:
        symbol1 = symbol
        symbol = symbol[0:1]
    else:
        symbol1 = 'No'
    symbol = symbol \
            .replace('A', '0') \
            .replace('B', '1') \
            .replace('C', '2') \
            .replace('D', '3') \
            .replace('E', '4') \
            .replace('F', '5') \
            .replace('G', '6') \
            .replace('H', '7') \
            .replace('I', '8') \
            .replace('J', '9') \
            .replace('K', '10') \
            .replace('L', '11') \
            .replace('M', '12') \
            .replace('N', '13') \
            .replace('O', '14') \
            .replace('P', '15') \
            .replace('Q', '16') \
            .replace('R', '17') \
            .replace('S', '18') \
            .replace('T', '19') \
            .replace('U', '20') \
            .replace('V', '21') \
            .replace('W', '22') \
            .replace('X', '23') \
            .replace('Y', '24') \
            .replace('Z', '25')
    if symbol1 == 'No':
        return int(symbol)
    if True:
        symbol1 = symbol1[1::] \
            .replace('A', '0') \
            .replace('B', '1') \
            .replace('C', '2') \
            .replace('D', '3') \
            .replace('E', '4') \
            .replace('F', '5') \
            .replace('G', '6') \
            .replace('H', '7') \
            .replace('I', '8') \
            .replace('J', '9') \
            .replace('K', '10') \
            .replace('L', '11') \
            .replace('M', '12') \
            .replace('N', '13') \
            .replace('O', '14') \
            .replace('P', '15') \
            .replace('Q', '16') \
            .replace('R', '17') \
            .replace('S', '18') \
            .replace('T', '19') \
            .replace('U', '20') \
            .replace('V', '21') \
            .replace('W', '22') \
            .replace('X', '23') \
            .replace('Y', '24') \
            .replace('Z', '25')
        return 26 * (int(symbol) + 1) + int(symbol1)


def update(vin):
    vin = vin.replace('У', 'Y').replace('А', 'A').replace('Т', 'T').\
        replace('Н', 'H').replace('В', 'B').replace('К', 'K')
    vin = vin.replace('Х', 'X').replace('О', '0').replace('O', '0').\
        replace('С', 'C').replace('М', 'M').replace('Р', 'P').replace('Е', 'E')
    vin = vin.replace('text:', '').replace('"', '').replace("'", '')
    return vin


def AntiCaptcha(site_key, url, key):
    try:
        start = time.time()
        r = requests.post('http://api.anti-captcha.com/createTask', json={"clientKey":key, "task":{"type":"NoCaptchaTaskProxyless", "websiteURL":url, "websiteKey":site_key}})
        r = json.loads(r.text)

        taskId = r['taskId']
        for i in range(12):
            time.sleep(10)
            r = requests.post('http://api.anti-captcha.com/getTaskResult', json={"clientKey": key, "taskId": taskId})
            r = json.loads(r.text)
            try:
                if r['status'] =='ready':
                    delta_time = time.time() - start
                    token = r['solution']['gRecaptchaResponse']

                    print(f'Токен найден за {round(delta_time)} сек.')
                    return token
            except Exception:
                pass
        token = None
        return token
    except Exception:
        print('Неверно введен Anticaptcha key или он отсутствует.')


def get_data(cells, index):
    try:
        content = cells[index].text_content()
    except Exception:
        content = '-'
    return content

