import ast
import multiprocessing
import time, json
from multiprocessing.pool import Pool
import colorama
import requests
import xlrd
from lxml import html
from xlutils.copy import copy
from default_functions import get_data, AntiCaptcha, print_col, update, eternity_cycle_deep


def getInfoForVIN(data):
    vincode = data['vin']
    date = data['date']
    antiCaptchaKey = data['aniCaptchaKey']
    site_key = '6Lf2uycUAAAAALo3u8D10FqNuSpUvUXlfP7BzHOk'
    url = 'https://dkbm-web.autoins.ru/dkbm-web-1.0/policy.htm?vin=' + vincode + '&licensePlate=&bodyNumber=' \
                                                                                 '&chassisNumber=&date=' + date
    captchaKey = None
    while captchaKey == None:
        captchaKey = AntiCaptcha(site_key, url, antiCaptchaKey)

    r1 = requests.post("https://dkbm-web.autoins.ru/dkbm-web-1.0/policyInfo.htm", data={"bsoseries": "ССС", "bsonumber": "", "requestDate": date, "vin": vincode, "licensePlate": "", "bodyNumber": "", "chassisNumber": "", "isBsoRequest": "false", "captcha": captchaKey}, verify=False)
    res = json.loads(r1.text)

    if res["validCaptcha"]:
        processId = res["processId"]
        time.sleep(2)
        r2 = requests.get("https://dkbm-web.autoins.ru/dkbm-web-1.0/checkPolicyInfoStatus.htm", params=[("processId", processId),("_", int(time.time()*1000))], verify=False)
        res = json.loads(r2.text)

        if(res["RequestStatusInfo"]["RequestStatusCode"] == 3):
            r3 = requests.post("https://dkbm-web.autoins.ru/dkbm-web-1.0/policyInfoData.htm", data={"processId": processId, "bsoseries": "ССС", "bsonumber": "", "vin": vincode, "licensePlate": "", "bodyNumber": "", "chassisNumber": "", "requestDate": date, "g-recaptcha-response": ""}, verify=False)
        
            result = []

            for row in html.document_fromstring(r3.text).xpath('//tr[@class="data-row"]'):
                cells = row.xpath('.//td')
                sp = []
                for i in range(1, 30):
                    sp.append(get_data(cells, i))
                elems = sp[4].split('\n')
                elems_need = []
                for i in elems:
                    if i.strip() != '':
                        elems_need.append(i.strip())
                znak, mark, vinc, n_k, m_d = '-', '-', '-', '-', '-'
                for i in range(len(elems_need)):
                    if elems_need[i] == 'Государственный регистрационный знак' and i < len(elems_need) - 1:
                        znak = elems_need[i + 1]

                    if elems_need[i] == 'Марка и модель транспортного средства' and i < len(elems_need) - 2:
                        mark = elems_need[i + 2]

                    if elems_need[i] == 'VIN' and i < len(elems_need) - 1:
                        vinc = elems_need[i + 1]

                    if elems_need[i] == 'Номер кузова' and i < len(elems_need) - 1:
                        n_k = elems_need[i + 1]

                    if elems_need[i] == 'Мощность двигателя для категории' and i < len(elems_need) - 2:
                        m_d = elems_need[i + 2]
                sp_reverse = sp[::-1]
                z = 0
                for i in range(len(sp_reverse)):
                    if sp_reverse[i] == '-':
                        continue
                    z += 1
                    if z == 1:
                        bms = sp_reverse[i]
                    if z == 2:
                        own = sp_reverse[i]
                    if z == 3:
                        polhold = sp_reverse[i]
                    if z == 4:
                        lim = sp_reverse[i]
                    if z == 5:
                        use_pr = sp_reverse[i]
                    if z == 6:
                        w_t = sp_reverse[i]
                    if z == 7:
                        go_pl = sp_reverse[i]
                result = [{
                    "OSAGO": sp[0],
                    "nameCompany": sp[1],
                    "statusContract": sp[2],
                    "validity": sp[3],
                    "carInfo": {
                        "model": mark,
                        "stateRegistrationMark": znak,
                        "VIN": vinc,
                        "bodyNumber": n_k,
                        "enginePower": m_d,
                    },
                    "goPlace": go_pl,
                    "withTrailer": w_t,
                    "usePurpose": use_pr,
                    "limits": lim,
                    "policyholder": polhold,
                    "owner": own,
                    "BMS": bms
                }]
            with open('log_rsa.txt', 'a', encoding='utf-8') as log:
                log.write(str(result))
                log.write("\n")
            answer_sp = []
            answer__sp = open('log_rsa.txt', 'r', encoding='UTF-8').read().split('\n')
            for i in range(len(answer__sp)):
                try:
                    if i == len(answer__sp) - 1:
                        break
                    try:
                        answer_sp.append(ast.literal_eval(answer__sp[i]))
                    except Exception:
                        pass
                except Exception:
                    pass
            print(f'Обработано: {len(answer_sp)}.')
            return 0
        else:
            print(f'[ERROR] {res}')
    else:
        getInfoForVIN(data)


def get_data_exel(file_name, POS_VIN, date, aniCaptchaKey, delta_del):
    try:
    #if True:
        excel_file = xlrd.open_workbook(file_name)
        sheet = excel_file.sheet_by_index(0)
        row_number = sheet.nrows
    except Exception:
        print_col('Ошибка чтения из файла! Перезапустите программу.', 'red')
        eternity_cycle_deep()

    vin_codes = []
    delta = 0
    super_vin_codes = []
    for row in range(1, row_number):
        try:
            delta += 1
            if delta % delta_del == 0:
                super_vin_codes.append(vin_codes)
                vin_codes = []
                delta = 0
            vin = str((sheet.row(row)[POS_VIN]))
            vin = update(vin)
            vin_codes.append({'vin': vin, 'date': date, 'aniCaptchaKey': aniCaptchaKey})
        except Exception:
            pass
    super_vin_codes.append(vin_codes)
    return super_vin_codes


def mult(data, pool_col):
    with open('log_rsa.txt', 'w', encoding='utf-8') as log:
        log.write('')

    for d in data:
        with Pool(pool_col) as p:
           try:
                p.map(getInfoForVIN, d)
           except Exception:
                pass
        p.close()
        print_col('Остановка на 30 секунд.', 'blue')
        time.sleep(30)
        print_col('Перезагрузка завершена.', 'magenta')


def write(file_name, POS):
    answer_sp = []
    answer__sp = open('log_rsa.txt', 'r', encoding='UTF-8').read().split('\n')
    for i in range(len(answer__sp)):
        try:
            if i == len(answer__sp) - 1:
                break
            try:
                answer_sp.append(ast.literal_eval(answer__sp[i]))
            except Exception:
                pass
        except Exception:
            pass
    exel_file = xlrd.open_workbook(file_name)
    wb = copy(exel_file)
    sheet = exel_file.sheet_by_index(0)
    name_sheet = sheet.name
    col_number = sheet.ncols
    row_number = sheet.nrows
    ws1 = wb.get_sheet(name_sheet)
    for i in range(0, len(answer_sp)):
        answer_sp[i] = answer_sp[i][-1]
        osago = answer_sp[i]['OSAGO']
        nameCompany = answer_sp[i]['nameCompany']
        statusContract = answer_sp[i]['statusContract']
        validity = answer_sp[i]['validity']
        model = answer_sp[i]['carInfo']['model']
        stateRegistrationMark = answer_sp[i]['carInfo']['stateRegistrationMark']
        VIN = answer_sp[i]['carInfo']['VIN']
        bodyNumber = answer_sp[i]['carInfo']['bodyNumber']
        enginePower = answer_sp[i]['carInfo']['enginePower']
        goPlace = answer_sp[i]['goPlace']
        withTrailer = answer_sp[i]['withTrailer']
        usePurpose = answer_sp[i]['usePurpose']
        limits = answer_sp[i]['limits']
        policyholder = answer_sp[i]['policyholder']
        owner = answer_sp[i]['owner']
        BMS = answer_sp[i]['BMS']
        vin = update(VIN)

        for row in range(1, row_number):
            try:
                win = update(str(sheet.row(row)[POS]))
                if vin == win:
                    ws1.write(row, col_number, osago.split()[0])
                    ws1.write(row, col_number + 1, osago.split()[1])
                    ws1.write(row, col_number + 2, nameCompany)
                    ws1.write(row, col_number + 3, statusContract)
                    ws1.write(row, col_number + 4, validity)
                    ws1.write(row, col_number + 5, model)
                    ws1.write(row, col_number + 6, stateRegistrationMark)
                    ws1.write(row, col_number + 7, bodyNumber)
                    ws1.write(row, col_number + 8, enginePower)
                    ws1.write(row, col_number + 9, goPlace)
                    ws1.write(row, col_number + 10, withTrailer)
                    ws1.write(row, col_number + 11, usePurpose)
                    ws1.write(row, col_number + 12, limits)
                    ws1.write(row, col_number + 13, policyholder)
                    ws1.write(row, col_number + 14, owner)
                    ws1.write(row, col_number + 15, BMS)
                    break
            except Exception:
                pass
    ws1.write(0, col_number, 'Серия договора ОСАГО')
    ws1.write(0, col_number + 1, 'Номер договора ОСАГО')
    ws1.write(0, col_number + 2, 'Наименование страховой организации')
    ws1.write(0, col_number + 3, 'Статус договора ОСАГО')
    ws1.write(0, col_number + 4, 'Срок действия и период использования транспортного средства договора ОСАГО')
    ws1.write(0, col_number + 5, 'Марка и модель транспортного средства')
    ws1.write(0, col_number + 6, 'Государственный регистрационный знак')
    ws1.write(0, col_number + 7, 'Номер кузова')
    ws1.write(0, col_number + 8, 'Мощность двигателя')
    ws1.write(0, col_number + 9, 'Цель использования транспортного средства')
    ws1.write(0, col_number + 10, 'Договор ОСАГО с ограничениями/без ограничений лиц, допущенных к управлению транспортным средством')
    ws1.write(0, col_number + 11, 'Сведения о страхователе транспортного средства')
    ws1.write(0, col_number + 12, 'Сведения о собственнике транспортного средства')
    ws1.write(0, col_number + 13, 'КБМ по договору ОСАГО')
    ws1.write(0, col_number + 14, 'Транспортное средство используется в регионе')
    ws1.write(0, col_number + 15, 'Страховая премия')
    wb.save(file_name)


def starting(file_name, POS_VIN, date, antiCaptchaKey, delta, pool_col):
    print_col('\nДля продолжения нажмите Enter', 'blue')
    input()
    data = get_data_exel(file_name, POS_VIN, date, antiCaptchaKey, delta)
    print_col('Данные из EXEL таблицы получены!\nЗапуск потоков и начало обработки.', 'magenta')
    print_col(f'Разделено на {len(data)} частей.', 'yellow')
    mult(data, pool_col)
    print_col("Обработка завершена. Запись...", 'green')
    write(file_name, POS_VIN)
    print_col('Запись завершена.', 'green')
