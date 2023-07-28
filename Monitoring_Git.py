# для запуска браузера html
import pandas as pd
from selenium import webdriver
# добавления параметров работы браузера при автоматическом запуске
from selenium.webdriver.edge.options import Options
from selenium.webdriver.edge.service import Service
# для работы с временем
import time
# для чтения html
from bs4 import BeautifulSoup
# работа с датой и временем
from datetime import datetime
# запуск приложений, в частности OutLook
import win32com.client as win32
# для чтения файлов
import os
from tabulate import tabulate

#_____________________________________________________________________________
#БЛОК ОБЪЯВЛЕНИЯ ФУНКЦИЙ
#функция для формирования сообщения
def send_email(to_email, subject, text,  path_template):
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItemFromTemplate(path_template)
    account = None
    for acc in mail.Session.Accounts:
        if acc.SmtpAddress == "Адресат":
            account = acc
            break
    mail._oleobj_.Invoke(*(64209, 0, 8, 0, account))
    mail.To = to_email
    mail.HTMLBody = text + mail.HTMLBody
    mail.Subject = subject
    mail.Send()

# функция форматирования даты и времени из логов
def to_date_my(date_old):
    year = date_old[:4]
    month = date_old[4:6]
    day = date_old[6:8]
    hour = date_old[9:11]
    minutes = date_old[11:13]
    seconds = date_old[13:15]
    return f'{year}.{month}.{day} {hour}:{minutes}:{seconds}'

#функция для парсинга html
def parsing_html(link):
    s = Service(driver)
    browser = webdriver.Edge(service=s, options=edge_options)
    # открытие окна с целевой web-страницей
    browser.get(link)
    # уменьшение масштаба до минимума, чтобы влезли все строки, так как выдает только видимые элементы
    browser.execute_script("document.body.style.zoom='25%'")
    # тайм-аут для того, чтобы точно успел прогрузить всю страницу
    time.sleep(10)
    try:
        # переключение блока отображения элементов на "отобразить все" для AP
        browser.find_element('xpath', '//select[@id="Pagesize"]/option[@value="0"]').click()
        # тайм-аут для загрузки всех элементов
        time.sleep(10)
    except:
        pass
    # получение html
    source = browser.page_source
    # парсинг html
    html_soup_object = BeautifulSoup(source, 'html.parser')
    # возврат масштаба, пока не придумала как иначе
    browser.execute_script("document.body.style.zoom='100%'")
    # закрытие окна браузера
    browser.quit()
    return html_soup_object

def qs_html_perfect_data(html_qs, server):
    rows = html_qs.find_all('tr')
    for row in rows:
        cols = row.find_all('td')
        cols = [ele.text.strip() for ele in cols]
        qs_apps.append([ele for ele in cols if ele])
        return qs_apps

#_____________________________________________________________________________

# БЛОК ОБЪЯВЛЕНИЯ ВСЕХ ПЕРЕМЕННЫХ ДЛЯ РАБОТЫ
# создание переменных для перечня данных по приложениям
qs_prod_apps = []
qs_test_apps = []
ap_qv_prod_apps = []
ap_qv_test_apps = []
list_of_name_qvd = []
qvw_app_data = []
total_inf = pd.DataFrame(columns=['1', '2', '3', '4', '5', '6'])

# начальное значение переменной с сообщением об ошибке подключения к серверу
message_error_prod_qs = 'Подключение к ресурсу прошло успешно.'
message_error_prod_qv = 'Подключение к ресурсу прошло успешно.'
message_error_test_qs = 'Подключение к ресурсу прошло успешно.'
message_error_test_qv = 'Подключение к ресурсу прошло успешно.'
message_error_sbx_qv = 'Подключение к ресурсу прошло успешно.'
message_error_test_log = 'Подключение к файловому ресурсу  прошло успешно.'
message_error_qvd = 'Подключение к файловому ресурсу слоев 2 и 3 прошло успешно.'

# перебор всех приложение и запись в списки последней даты обновления и наименование приложения
updates_prod = []
title_prod = []
updates_test = []
title_test = []
# списки приложений для отправки и их длительность - тестовый ландшафт
target_app_test = []
target_duration_test = []
last_time_update = []

# текущая дата в системе
current_datetime = str(datetime.now())
current_time = datetime.now().time()
current_date = datetime.now().date()
current_date = current_date.strftime("%d.%m.%Y")
current_time_for_check = datetime.now().time().replace(microsecond=0, second=0)
current_minutes = int(datetime.now().strftime("%M"))

# определение критических меток времени
extract_prod_start = datetime.strptime('06:00', '%H:%M').time()
extract_prod_end = datetime.strptime('09:00', '%H:%M').time()

extract_test_start = datetime.strptime('6:00', '%H:%M').time()
extract_test_end = datetime.strptime('10:00', '%H:%M').time()

ap_prod_technology_start = datetime.strptime('7:00', '%H:%M').time()
ap_prod_technology_end = datetime.strptime('10:00', '%H:%M').time()

ap_prod_smotr_start = datetime.strptime('6:00', '%H:%M').time()
ap_prod_smotr_end = datetime.strptime('10:00', '%H:%M').time()

ap_prod_technology2_start = datetime.strptime('8:00', '%H:%M').time()
ap_prod_technology2_end = datetime.strptime('10:00', '%H:%M').time()

ap_test_technology2_start = datetime.strptime('8:00', '%H:%M').time()
ap_test_technology2_end = datetime.strptime('10:00', '%H:%M').time()

ap_test_smotr_start = datetime.strptime('6:00', '%H:%M').time()
ap_test_smotr_end = datetime.strptime('10:00', '%H:%M').time()

ap_test_iceberg_start = datetime.strptime('9:00', '%H:%M').time()
ap_test_iceberg_end = datetime.strptime('15:30', '%H:%M').time()

extract_0194_logt_start = datetime.strptime('07:30', '%H:%M').time()
extract_0194_logt_end = datetime.strptime('10:00', '%H:%M').time()

model_tech2_logt_start = datetime.strptime('08:00', '%H:%M').time()
model_tech2_logt_end = datetime.strptime('10:00', '%H:%M').time()

qvd_prod_tier2_start = datetime.strptime('06:30', '%H:%M').time()
qvd_prod_tier2_end = datetime.strptime('07:30', '%H:%M').time()

qvd_prod_tier3_start = datetime.strptime('07:00', '%H:%M').time()
qvd_prod_tier3_end = datetime.strptime('08:00', '%H:%M').time()

limit_to_reload = 120
limit_to_model = 60

time_check_1 = datetime.strptime('09:00', '%H:%M').time()
time_check_2 = datetime.strptime('12:30', '%H:%M').time()
time_check_3 = datetime.strptime('15:00', '%H:%M').time()

# переменная, определяющая какое сообщение направить
flag_to_send_prod = 0
flag_to_send_test = 0
flag_to_send_test_ice = 0
flag_to_send_qvd = 0
flag_to_send_model = 0
flag_to_send_model_log = 0
flag_to_send_extract = 0
# переменная, хранящая разницу между датой последнего запуска Extract и текущей датой системы
duration_row = 0
# переменная хранящая разницу между датой последнего запуска целевого Extract и текущей датой системы
duration_target = 0

first_date_month = datetime.now().replace(day=1).strftime("%Y%m%d")
month = datetime.now().strftime("%Y%m")

# путь к директории, содержащей файлы моделей на тестовом контуре
path_model = r'path'
# путь к директории, содержащей файлы экстракторов на тестовом контуре
path_extract = r'path'
# имя модели
name_models = ['1', '2']
# имя экстрактора
name_extracts = ['1', '2']

# имя приложения
app_name = 'name'
# путь к директории, содержащей qvd файлы второго слоя продуктивного ландшафта
prod_second_tier = r'path'
# путь к директории, содержащей qvd файлы второго слоя продуктивного ландшафта
prod_third_tier = r'path'
# список списков наименования целевых файлов второго и третьего слоев
file_names = ['список списков файлов']

# БЛОК ПАРСИНГА HTML
#webdriver для Chrome
driver = ("path")
# запуск Chrome
edge_options = Options()
# запуск Chrome в фоне, скрыто, чтобы не мешало работе
edge_options.add_argument("--headless")

# product QV
try: ap_qv_prod_html = parsing_html("link1").find('ul', attrs={"id": "appList"})
except: message_error_prod_qv = 'Ошибка подключения к серверу product QV. Считать данные обновления экстракторов невозможно.'
# test QV
try: ap_qv_test_html = parsing_html("link2").find('ul', attrs={"id": "appList"})
except: message_error_test_qv = 'Ошибка подключения к серверу test QV. Считать данные обновления экстракторов невозможно.'
# product QS
try: hub_qs_prod_html = parsing_html("link3").find('tbody')
except: message_error_prod_qs = 'Ошибка подключения к серверу product QS. Считать данные обновления экстракторов невозможно.'
# test QS
try: hub_qs_test_html = parsing_html("link4").find('tbody')
except: message_error_test_qs = 'Ошибка подключения к серверу test QS. Считать данные обновления экстракторов невозможно.'
# SBX
try: hub_qs_test_html = parsing_html("link5").find('tbody')
except: message_error_test_qs = 'Ошибка подключения к серверу test QS. Считать данные обновления экстракторов невозможно.'
# sbx QV - проверка доступности
try: parsing_html("link6").find_all('li')
except: message_error_sbx_qv = 'Ошибка подключения к серверу sandbox QV. Недоступность сервера.'

# БЛОК QLIKSENSE ПРИЛОЖЕНИЕ МОНИТОРИНГА ЭКСТРАКТОРОВ
# перебор всех строк таблицы и запись их в список списков (каждый элемент списка содержит список с значениями ячеек одной строки)
# product QS
try:
    rows = hub_qs_prod_html.find_all('tr')
    for row in rows:
        cols = row.find_all('td')
        cols = [ele.text.strip() for ele in cols]
        qs_prod_apps.append([ele for ele in cols if ele])
except:
    message_error_prod_qs = 'Ошибка подключения к серверу product QS. Считать данные обновления экстракторов невозможно.'

# test QS
try:
    rows = hub_qs_test_html.find_all('tr')
    for row in rows:
        cols = row.find_all('td')
        cols = [ele.text.strip() for ele in cols]
        qs_test_apps.append([ele for ele in cols if ele])
except:
    message_error_test_qs = 'Ошибка подключения к серверу test QS. Считать данные обновления экстракторов невозможно.'

# БЛОК ACCESS POINT
# product QV
try:
    rows = ap_qv_prod_html.find_all('li')
    for row in rows:
        date = row.find('span', attrs={'class': 'gridOnly appDates'})  # поиск последней даты обновления
        name = row.find_all('a', attrs={'class': 'name'})  # поиск наименования приложения
        if date != None: updates_prod.append(datetime.strptime(date.text[13:], '%Y-%m-%d %H:%M').strftime(
            '%d.%m.%Y %H:%M:%S'))  # точно вытащить строку из html нельзя, поэтому есть пустые строки и для каждой даты обнвления есть начало строки в виде "Last updated "
        if name != []: title_prod.append(
            name[1].text)  # точно вытащить строку из html нельзя, поэтому есть пустые строки

    # создание списка списков [наименование, последняя дата обновления]
    ap_qv_prod_apps = [list(i) for i in zip(title_prod, updates_prod)]

except:
    message_error_prod_qv = 'Ошибка подключения к серверу product QV. Считать данные обновления витрин невозможно.'

# test QV
try:
    rows = ap_qv_test_html.find_all('li')
    for row in rows:
        date = row.find('span', attrs={'class': 'gridOnly appDates'})  # поиск последней даты обновления
        name = row.find_all('a', attrs={'class': 'name'})  # поиск наименования приложения
        if date != None: updates_test.append(datetime.strptime(date.text[13:], '%Y-%m-%d %H:%M').strftime(
            '%d.%m.%Y %H:%M:%S'))  # точно вытащить строку из html нельзя, поэтому есть пустые строки и для каждой даты обнвления есть начало строки в виде "Last updated "
        if name != []: title_test.append(
            name[1].text)  # точно вытащить строку из html нельзя, поэтому есть пустые строки

    # создание списка списков [наименование, последняя дата обновления]
    ap_qv_test_apps = [list(i) for i in zip(title_test, updates_test)]
except:
    message_error_test_qv = 'Ошибка подключения к серверу product QV. Считать данные обновления витрин невозможно.'

# объединение двух списков продуктивного ландшафта
data_prod = qs_prod_apps + ap_qv_prod_apps
# объединение двух списков тестового ландшафта
data_test = qs_test_apps + ap_qv_test_apps

# списки приложений для отправки и их длительность - продуктивный ландшафт
target_app_prod = []
target_duration_prod = []
# поиск списка с целевым Extract
for row in data_prod:
    duration_row = datetime.strptime(current_datetime, "%Y-%m-%d %H:%M:%S.%f") - datetime.strptime(row[1],
                                                                                                   "%d.%m.%Y %H:%M:%S")
    duration_target = duration_row.total_seconds() / 60
    if (row[0] == 'name_app' or row[0] == 'name_app' or
        row[0] == 'name_app') and duration_target > limit_to_reload and current_time >= extract_prod_start and current_time <= extract_prod_end and current_minutes != 30 \
            or ((row[0] == 'name_app' or row[0] == 'name_app') and current_time >= ap_prod_technology2_start and current_time <= ap_prod_technology2_end
            or row[0] == 'name_app' and current_time >= ap_prod_technology_start and current_time <= ap_prod_technology_end and current_minutes != 30
            or row[0] == 'name_app' and current_time >= ap_prod_smotr_start and current_time <= ap_prod_smotr_end and current_minutes != 30
    ) and duration_target > limit_to_model:
        target_app_prod.extend([[row[0], 'prod', int(duration_target), row[1], '', 'Web']])

for row in data_test:
    duration_row = datetime.strptime(current_datetime, "%Y-%m-%d %H:%M:%S.%f") - datetime.strptime(row[1], "%d.%m.%Y %H:%M:%S")
    duration_target = duration_row.total_seconds() / 60
    date_row = datetime.strptime(row[1], "%d.%m.%Y %H:%M:%S")
    date_row = date_row.strftime("%d.%m.%Y")
    # условие для отправки сообщения по экстракторам
    if (row[0] == 'name_app' or row[0] == 'name_app' or row[
        0] == 'name_app') and duration_target > limit_to_reload and current_time >= extract_test_start and current_time <= extract_test_end and current_minutes != 30 or (
            (row[0] == 'name_app' or row[0] == 'name_app') and current_time >= ap_test_technology2_start and current_time <= ap_test_technology2_end
            or row[0] == 'name_app' and current_time >= ap_test_smotr_start and current_time <= ap_test_smotr_end and current_minutes != 30
    ) and duration_target > limit_to_model:
        target_app_test.extend([[row[0], 'test', int(duration_target), row[1], '', 'Web']])
    elif (row[0] == 'name_app' or row[0] == 'name_app') and current_time >= ap_test_iceberg_start and current_time <= ap_test_iceberg_end \
            and date_row != current_date and (current_time_for_check == time_check_1 or current_time_for_check == time_check_2 or current_time_for_check == time_check_3):
        flag_to_send_test_ice = 1
        last_time_update.append(row[0] + ' ('+str(datetime.strptime(row[1], "%d.%m.%Y %H:%M:%S"))+')')

# БЛОК ЛОГИ ПРИЛОЖЕНИЙ
try:
    for name_model in name_models:
        # определение текущей директории для дальнейшей работы
        file_log_model = os.chdir(path_model)
        # открытие файла логов
        file_log_model = open(name_model + '.log', encoding='utf-8-sig')

        # получение даты последнего изменения файлов приложений для дальнейшего определения ошибки сохранения приложений (дата лога больше даты изменения модели)
        date_file_model = os.path.getmtime(path_model + '\\' + name_model)
        date_file_model = datetime.fromtimestamp(date_file_model)

        # чтение файлов логов
        file_rows_model = file_log_model.read().split('\n')

        # получение даты окончания загрузки
        end_date_model = to_date_my([x for x in file_rows_model[-2].split(' ') if x][0])

        # получение статуса и ошибки загрузки, если обновление не прошло успешно
        success_model = (''.join([x for x in file_rows_model[-2].split(' ') if x][1:]) == 'Executionfinished.' and ''.join(
            [x for x in file_rows_model[-3].split(' ') if x][1:]) != 'ExecutionFailed')
        if success_model:
            reason_model = ''
        else:
            reason_model = ''.join([x for x in file_rows_model[-4].split(' ') if x][1:])

        # получение разницы времени между текущей датой системы и последним обновлением приложения
        duration_model = datetime.strptime(current_datetime, "%Y-%m-%d %H:%M:%S.%f") - datetime.strptime(end_date_model,
                                                                                                         "%Y.%m.%d %H:%M:%S")
        # перевод времени в минуты - целое
        duration_target_model = int(duration_model.total_seconds() / 60)

        dif_duration_model = date_file_model - datetime.strptime(end_date_model, "%Y.%m.%d %H:%M:%S")
        dif_duration_model = int(dif_duration_model.total_seconds() / 60)
        if (duration_target_model > limit_to_model or success_model == False) and current_time >= model_tech2_logt_start and current_time <= model_tech2_logt_end:
            qvw_app_data.extend([[name_model, 'test', duration_target_model, end_date_model, reason_model, 'Local']])
        if dif_duration_model > 5 and current_time >= model_tech2_logt_start and current_time <= model_tech2_logt_end:
            qvw_app_data.extend([[name_model, 'test', dif_duration_model, end_date_model, 'Ошибка сохранения приложения - логи старше модели', 'Local']])
except:
    message_error_test_log = 'Ошибка подключения или чтения файлов ИР'

try:
    for name_extract in name_extracts:
        file_log_extract = os.chdir(path_extract)
        file_log_extract = open(name_extract + '.log', encoding='utf-8-sig')

        # чтение файлов логов
        file_rows_extract = file_log_extract.read().split('\n')

        # получение даты окончания загрузки
        end_date_extract = to_date_my([x for x in file_rows_extract[-2].split(' ') if x][0])

        # получение статуса и ошибки загрузки, если обновление не прошло успешно
        success_extract = (
                ''.join([x for x in file_rows_extract[-2].split(' ') if x][1:]) == 'Executionfinished.' and ''.join(
            [x for x in file_rows_extract[-3].split(' ') if x][1:]) != 'ExecutionFailed')
        if success_extract:
            reason_extract = ''
        else:
            reason_extract = ''.join([x for x in file_rows_extract[-4].split(' ') if x][1:])

        duration_extract = datetime.strptime(current_datetime, "%Y-%m-%d %H:%M:%S.%f") - datetime.strptime(
            end_date_extract,
            "%Y.%m.%d %H:%M:%S")
        duration_target_extract = int(duration_extract.total_seconds() / 60)
        if (duration_target_extract > limit_to_model or success_extract == False) and current_time >= extract_0194_logt_start and current_time <= extract_0194_logt_end:
            qvw_app_data.extend([[name_extract, 'test', duration_target_extract, end_date_extract, reason_extract, 'Local']])
except:
    message_error_test_log = 'Ошибка подключения или чтения файлов ИР'

#БЛОК ПРОВЕРКИ QVD 2 И 3 СЛОЕВ ПРОДУКТИВНОГО ЛАНДШАФТА
try:
    for file in file_names[0]:
        date_file = datetime.fromtimestamp(os.path.getmtime(prod_second_tier + '\\' + file + '.qvd')).strftime(
            "%Y-%m-%d %H:%M:%S.%f")
        dif_duration_qvd_min = int((datetime.strptime(current_datetime, "%Y-%m-%d %H:%M:%S.%f") - datetime.strptime(
            date_file, "%Y-%m-%d %H:%M:%S.%f")).total_seconds() / 60)
        if dif_duration_qvd_min > limit_to_reload and current_time >= qvd_prod_tier2_start and current_time <= qvd_prod_tier2_end:
            list_of_name_qvd.extend([[ist_of_name_qvd + '\nDataTier2: ' + file, 'prod', dif_duration_qvd_min, date_file, '', 'qvd file']])

    for file in file_names[1]:
        date_file = datetime.fromtimestamp(os.path.getmtime(prod_third_tier + '\\' + file + '.qvd')).strftime(
            "%Y-%m-%d %H:%M:%S.%f")
        dif_duration_qvd_min = int((datetime.strptime(current_datetime, "%Y-%m-%d %H:%M:%S.%f") - datetime.strptime(
            date_file, "%Y-%m-%d %H:%M:%S.%f")).total_seconds() / 60)
        if dif_duration_qvd_min > limit_to_reload and current_time >= qvd_prod_tier3_start and current_time <= qvd_prod_tier3_end:
            list_of_name_qvd.extend([[ist_of_name_qvd + '\nDataTier3: ' + file, 'prod', dif_duration_qvd_min, date_file, '', 'qvd file']])
except:
    message_error_qvd = 'Ошибка подключения или чтения файлов ИР'


text_message = message_error_prod_qs + '\n' + message_error_test_qs + '\n' + message_error_prod_qv + '\n' + message_error_test_qv + '\n' + message_error_sbx_qv + '\n' + message_error_test_log + '\n' + message_error_qvd

target_app_prod = pd.DataFrame(target_app_prod, columns=total_inf.columns)
target_app_test = pd.DataFrame(target_app_test, columns=total_inf.columns)
qvw_app_data = pd.DataFrame(qvw_app_data, columns=total_inf.columns)
list_of_name_qvd = pd.DataFrame(list_of_name_qvd, columns=total_inf.columns)
total_inf = pd.concat([total_inf, target_app_prod, target_app_test, qvw_app_data, list_of_name_qvd])

total_inf = total_inf.iloc[1:].rename_axis(None, axis=1)
total_inf.set_index('Название',inplace = True)
head_table = """
            <table class=MsoTableGrid border=1 cellspacing=0 cellpadding=0 style='border-collapse:collapse;border:none;mso-border-alt:solid windowtext .5pt;mso-yfti-tbllook:1184;mso-padding-alt:0cm 5.4pt 0cm 5.4pt'><tr style='mso-yfti-irow:0;mso-yfti-firstrow:yes'><td width=116 valign=top style='width:150pt;border:solid windowtext 1.0pt;mso-border-alt:solid windowtext .5pt;padding:0cm 5.4pt 0cm 5.4pt'><p class=MsoNormal align=center style='text-align:center'><b style='mso-bidi-font-weight:normal'><span style='font-size:10.0pt;font-family:"Times New Roman",serif'>Название<o:p></o:p></span></b></p></td><td width=90 valign=top style='width:77.95pt;border:solid windowtext 1.0pt;border-left:none;mso-border-left-alt:solid windowtext .5pt;mso-border-alt:solid windowtext .5pt;padding:0cm 5.4pt 0cm 5.4pt'><p class=MsoNormal align=center style='text-align:center'><b style='mso-bidi-font-weight:normal'><span style='font-size:10.0pt;font-family:"Times New Roman",serif'>Ландшафт<o:p></o:p></span></b></p></td><td width=120 valign=top style='width:120.5pt;border:solid windowtext 1.0pt;border-left:none;mso-border-left-alt:solid windowtext .5pt;mso-border-alt:solid windowtext .5pt;padding:0cm 5.4pt 0cm 5.4pt'><p class=MsoNormal align=center style='text-align:center'><b style='mso-bidi-font-weight:normal'><span style='font-size:10.0pt;font-family:"Times New Roman",serif'>Время с последнего обновления (мин)<o:p></o:p></span></b></p></td><td width=120 valign=top style='width:120.5pt;border:solid windowtext 1.0pt;border-left:none;mso-border-left-alt:solid windowtext .5pt;mso-border-alt:solid windowtext .5pt;padding:0cm 5.4pt 0cm 5.4pt'><p class=MsoNormal align=center style='text-align:center'><b style='mso-bidi-font-weight:normal'><span style='font-size:10.0pt;font-family:"Times New Roman",serif'>Дата и время последнего обновления<o:p></o:p></span></b></p></td><td width=129 valign=top style='width:127.55pt;border:solid windowtext 1.0pt;border-left:none;mso-border-left-alt:solid windowtext .5pt;mso-border-alt:solid windowtext .5pt;padding:0cm 5.4pt 0cm 5.4pt'><p class=MsoNormal align=center style='text-align:center'><b style='mso-bidi-font-weight:normal'><span style='font-size:10.0pt;font-family:"Times New Roman",serif'>Сообщение об ошибке (для локальных приложений)<o:p></o:p></span></b></p></td><td width=132 valign=top style='width:148.85pt;border:solid windowtext 1.0pt;border-left:none;mso-border-left-alt:solid windowtext .5pt;mso-border-alt:solid windowtext .5pt;padding:0cm 5.4pt 0cm 5.4pt'><p class=MsoNormal align=center style='text-align:center'><b style='mso-bidi-font-weight:normal'><span lang=EN-US style='font-size:10.0pt;font-family:"Times New Roman",serif;mso-ansi-language:EN-US'>Local/Web</span></b><b style='mso-bidi-font-weight:normal'><span style='font-size:10.0pt;font-family:"Times New Roman",serif'><o:p></o:p></span></b></p></td></tr> 
"""
table_rows = """"""
for row in total_inf.itertuples():
    table_rows = table_rows + """<tr style='mso-yfti-irow:1;mso-yfti-lastrow:yes'>"""
    for elem in row:
        table_rows = table_rows + """
            <td width=116 valign=top style='width:150pt;border:solid windowtext 1.0pt;border-top:none;mso-border-top-alt:solid windowtext .5pt;mso-border-alt:solid windowtext .5pt;padding:0cm 5.4pt 0cm 5.4pt'><p class=MsoNormal><span class=SpellE><span style='font-size:10.0pt;font-family:"Times New Roman",serif'>
            """+str(elem)+"""
            </span></span><span style='font-size:10.0pt;font-family:"Times New Roman",serif'><o:p></o:p></span></p></td>
            """
    table_rows = table_rows + """</tr>"""

end_table = """</table>"""
print(table_rows)

if table_rows != '':
    table_message = head_table + table_rows + end_table
else: table_message = ''

#ОТПРАВКА СООБЩЕНИЯ
if flag_to_send_test_ice == 1: send_email('Кому отправляем', 'Тема', 'Статус: ' + str(last_time_update), r"path")

if current_time <= ap_prod_smotr_end: send_email('Кому отправляем', 'Тема', '<br>'+text_message.replace('\n', '<br>')+table_message, r"path")

