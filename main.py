from configparser import ConfigParser
from cryptography.fernet import Fernet
from cryptography.hazmat.primitives import hashes
from cryptography.hazmat.primitives.kdf.pbkdf2 import PBKDF2HMAC
from datetime import datetime
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import PySimpleGUI as sg
import base64
import codecs
import datetime as dt
import fnmatch
import os
import re
import smtplib
import sys
import time
import xlsxwriter

version = 'v0.23.12.1'
ini_file = 'logs2xlsx.config'
args = {'sendmail': '--sendmail', 'debug': '--debug'}

# TRIAL usage
trial, exp_date = False, dt.date(2021, 7, 10)
trial_expired = True if trial is True and dt.date.today() >= exp_date else False

config_defaults = {'schedule_logs_path': '', 'schedule_log_start': '06:00', 'schedule_log_end': '06:01',
                   'host': 'smtp.gmail.com', 'port': '587', 'mail': '', 'pass': '',
                   'recipients': 'List of emails separated by commas and without spaces', 'subject_title': 'LOGs from PyLogParser',
                   'body': 'New log file in attachment', 'schedule_log_name': 'log.txt'}

sendmail, debug = False, False
if len(sys.argv) > 1:
    sendmail = sys.argv[1] if sys.argv[1] == args.get('sendmail') else False
    debug = sys.argv[1] if sys.argv[1] == args.get('debug') else False
    if len(sys.argv) == 3:
        debug = sys.argv[2] if sys.argv[2] == args.get('debug') else False


def debug_log(data):
    if debug:
        debug_file = codecs.open('logs2xlsx.log', 'a', 'utf_8_sig')
        debug_file.write(f"\n{datetime.now().strftime('%H:%M:%S')} - {data}")
        debug_file.close()


debug_log('START')
debug_log(f'sendmail: {sendmail}')
debug_log(f'debug: {debug}')


def config_get(section):
    config = ConfigParser()
    config.read(ini_file)
    return config[section]


def password_crypt(mode, password):
    kdf = PBKDF2HMAC(algorithm=hashes.SHA256(), length=32, salt=os.urandom(0), iterations=1)
    key = base64.urlsafe_b64encode(kdf.derive(b"J$?{Mf[xKzK`*JV,&_(#~{)x7jMQ>Y"))
    f = Fernet(key)
    token = f.encrypt(password)
    if mode == 'encrypt':
        return token
    else:
        return f.decrypt(password)


def config_set(settings):
    config = ConfigParser()
    if settings.get('pass') == '********':
        config.read(ini_file)
        email_pass = config['SMTP']['pass']
    elif settings.get('pass') == '':
        email_pass = ''
    else:
        email_pass = password_crypt('encrypt', bytes(settings.get('pass'), encoding='utf-8')).decode('utf-8')

    config['LOGS'] = {
        'path': settings.get('schedule_logs_path'),
        'start': settings.get('schedule_log_start'),
        'end': settings.get('schedule_log_end')
    }
    config['SMTP'] = {
        'host': settings.get('host'),
        'port': settings.get('port'),
        'mail': settings.get('mail'),
        'pass': email_pass
    }
    config['EMAIL'] = {
        'recipients': settings.get('recipients'),
        'subject_title': settings.get('subject_title'),
        'body': settings.get('body'),
        'log_name': settings.get('schedule_log_name')
    }

    with open(ini_file, 'w') as conf:
        config.write(conf)


if not os.path.exists(ini_file):
    config_set(config_defaults)


def calc_month(date_format='%Y-%m', this_date=dt.date.today(), delta='prev'):
    first_day = this_date.replace(day=1)
    if delta == 'next':
        month = first_day + dt.timedelta(days=31)
        month = month.replace(day=1)
    else:
        month = first_day - dt.timedelta(days=1)

    return month.strftime(date_format)


def main(settings):
    if settings.get('logs_path') == '':
        print('Необходимо выбрать папку с логами')
        return

    # settings section
    logs_path = settings.get('logs_path')
    xlsx_path = settings.get('xlsx_path') + '/'
    file_log = xlsx_path + 'full_log.tmp'

    file_xlsx = settings.get('file_xlsx')
    if file_xlsx == '':
        file_xlsx = calc_month('%Y-%m') + '_logs'
    file_xlsx = xlsx_path + file_xlsx + '.xlsx'

    worksheet_name = settings.get('ws_name')
    if worksheet_name == '':
        worksheet_name = 'logs'

    worksheet_header = [settings.get('col1'), settings.get('col2'), settings.get('col3'), settings.get('col4')]

    files_month = datetime.strptime(settings.get('month'), '%Y %B')
    logs_date = files_month.strftime('%Y-%m')

    offset_minus = False
    if settings.get('offset_sign') == '-':
        offset_minus = True

    time_offset = settings.get('time_offset')

    string_clip_start = settings.get('start')
    if string_clip_start == '':
        string_clip_start = 'CLIP START'

    string_clip_end = settings.get('end')
    if string_clip_end == '':
        string_clip_end = 'CLIP STOP'

    # IDs of commercial advertising block
    ad_start = settings.get('ad_start')
    ad_end = settings.get('ad_end')
    if ad_start == '':
        ad_start = 'ID_REC_DTMF_In'
    if ad_end == '':
        ad_end = 'ID_REC_DTMF_Out'

    ad_start = string_clip_start + '\t' + ad_start
    ad_end = string_clip_start + '\t' + ad_end
    # END settings section

    def search_string(lines, string, ad_ending=False):
        line_num = 0
        result = []
        for line in lines:
            line_num += 1
            if ad_ending and (string in line or ad_ending in line):
                result.append(line_num)
            elif string in line:
                result = (line.split())

        if len(result) > 0:
            return result
        else:
            return

    def format_row(data):
        date = data[0]
        date_time = datetime.strptime(date + data[1], '%Y-%m-%d%H:%M:%S')
        if not data[3]:
            timing = string_clip_end + ' NOT FOUND'
        else:
            time_end = datetime.strptime(date + data[3], '%Y-%m-%d%H:%M:%S')
            timing = time_end - date_time

        if offset_minus:
            datetime_shifted = date_time - dt.timedelta(hours=int(time_offset))
        else:
            datetime_shifted = date_time + dt.timedelta(hours=int(time_offset))

        date_time_list = str(datetime_shifted).split()
        clip_date = datetime.strptime(date_time_list[0], '%Y-%m-%d')

        if clip_date.strftime('%Y-%m') == logs_date:
            return data[2], date_time_list[0], date_time_list[1], str(timing)
        else:
            return

    def add_row_to_xlsx(row_content, row_number):
        ws.write_row(row_number, 0, row_content)

    # start
    start_time = time.time()

    # preparing files
    files_prev_month = calc_month('%Y-%m-%d', files_month)
    files_next_month = calc_month('%Y-%m-%d', files_month, 'next')

    files_pattern = ['*' + files_prev_month + '*',
                     '*' + logs_date + '*',
                     '*' + files_next_month + '*']
    files_pattern = r'|'.join([fnmatch.translate(x) for x in files_pattern])

    logs_list = []
    for dirpath, dirs, files in os.walk(logs_path):
        files = [f for f in files if re.match(files_pattern, f)]

        for file_name in sorted(files):
            logs_list.append(file_name)

    if os.path.exists(file_log):
        os.remove(file_log)
    full_log = codecs.open(file_log, 'a', 'utf_8_sig')

    for filename in sorted(logs_list):
        with codecs.open(logs_path + '/' + filename, 'rb', 'utf_8_sig') as log:
            full_log.write(log.read())
    full_log.close()

    full_log = codecs.open(file_log, 'r', 'utf_8_sig')
    linelist = full_log.readlines()

    with open(file_xlsx, 'a+'):
        now = datetime.now()
        print(now.strftime('%H:%M:%S') + ' - START')

    wb = xlsxwriter.Workbook(file_xlsx)
    ws = wb.add_worksheet(worksheet_name)

    # finding and grouping line numbers to pairs(advertising start to end)
    ad_lines_list = search_string(linelist, ad_start, ad_end)
    try:
        len(ad_lines_list)
    except TypeError:
        print(f"{now.strftime('%H:%M:%S')} - Не найдены записи с заданными настройками,"
              f"\nпроверьте пути к файлам, даты и ID роликов")
        return

    ad_lines_list_len = len(ad_lines_list)
    if ad_lines_list_len % 2 == 0 and ad_lines_list is not None:
        ad_start_end_pairs = []
        for i in range(0, ad_lines_list_len, 2):
            ad_start_end_pairs.append((ad_lines_list[i], ad_lines_list[i + 1]))
    else:
        full_log.close()
        print(f"{now.strftime('%H:%M:%S')} - ID рекламных отбивок указаны неправильно!")
        print(f'Количество рекламных блоков: {ad_lines_list_len / 2}')
        return

    ws.write_row(0, 0, worksheet_header)
    xlsx_row = 1
    # ad strings processing
    for i in range(0, len(ad_start_end_pairs)):
        ad_start_line = ad_start_end_pairs[i][0] + 1
        ad_end_line = ad_start_end_pairs[i][1] - 1
        skip_num = False
        for num in range(ad_start_line, ad_end_line):
            if num == skip_num:
                skip_num = False
                continue
            line_data = [linelist[num]]
            ad = search_string(line_data, string_clip_start)
            if ad is not None:
                ad_end = (search_string([linelist[num + 1]], string_clip_end))
                if ad_end is not None and ad[6] == ad_end[6]:
                    ad_tuple = (ad[0], ad[1], ad[5], ad_end[1])
                    skip_num = num + 1
                else:
                    ad_tuple = (ad[0], ad[1], ad[5], False)

                formatted_row = format_row(ad_tuple)
                if formatted_row is not None:
                    add_row_to_xlsx(formatted_row, xlsx_row)
                    xlsx_row += 1

    # xlsx formatting
    ws.set_column(0, 0, 40)
    ws.set_column(1, 3, 13)

    full_log.close()
    os.remove(file_log)
    wb.close()
    print(f'Строк записано: {xlsx_row}')
    print('Выполнено за %.2f сек.' % (time.time() - start_time))


def scheduled_log():
    config_logs, config_email = config_get('LOGS'), config_get('EMAIL')
    logs_path = config_logs['path']
    log_name = config_email['log_name'].replace('{date_yesterday}',
                                                (dt.date.today() - dt.timedelta(1)).strftime('%d.%m.%Y'))
    now_date_str = dt.date.today().strftime('%Y-%m-%d')
    yesterday_date_str = (dt.date.today() - dt.timedelta(1)).strftime('%Y-%m-%d')

    files_pattern = r'|'.join([fnmatch.translate(x) for x in
                              ['*' + now_date_str + '*', '*' + yesterday_date_str + '*']])

    logs_list = []
    for dirpath, dirs, files in os.walk(logs_path):
        files = [f for f in files if re.match(files_pattern, f)]
        for file_name in sorted(files):
            logs_list.append(file_name)

    if os.path.exists(log_name):
        os.remove(log_name)
    schedule_file_log = codecs.open(log_name, 'a', 'utf_8_sig')

    start_datetime = datetime.strptime(f"{yesterday_date_str} {config_logs['start']}:00", '%Y-%m-%d %H:%M:%S')
    end_datetime = datetime.strptime(f"{now_date_str} {config_logs['end']}:00", '%Y-%m-%d %H:%M:%S')

    for filename in sorted(logs_list):
        with codecs.open(logs_path + '/' + filename, 'rb', 'utf_8_sig') as log:
            for line in log:
                line_datetime = datetime.strptime('\t'.join(line.split('\t')[:3]), '\t%Y-%m-%d\t%H:%M:%S')
                if start_datetime < line_datetime < end_datetime:
                    schedule_file_log.write(line)

    schedule_file_log.close()


def gui_interface():
    schedule_logs, schedule_smtp, schedule_email = config_get('LOGS'), config_get('SMTP'), config_get('EMAIL')
    schedule_smtp['pass'] = '********' if schedule_smtp['pass'] else ''

    version_text = f'{version} - Пробный период до {exp_date}' if trial else version
    version_text_color = '#6a0000' if trial else '#ffffff'

    trial_expired_layout = [
        [sg.Text(f'{version}\nПробный период использования закончился {exp_date}',
                 text_color=version_text_color, justification='center')],
        [sg.Button('Закрыть', key='Exit', size=(10, 1))]]

    main_layout = [
        [sg.Text('Папка c логами', size=(13, 1)),
         sg.Input(readonly=True, size=(36, 1), key='logs_path', enable_events=True),
         sg.FolderBrowse()],
        [sg.Text('Папка для отчёта', size=(13, 1)),
         sg.Input(readonly=True, size=(36, 1), key='xlsx_path'),
         sg.FolderBrowse()],
        [sg.Text('Сформировать за', size=(13, 1)),
         sg.InputText(default_text=calc_month('%Y %B'), readonly=True, justification='right', size=(16, 1),
                      key='month'),
         sg.CalendarButton('Календарь', format='%Y %B', no_titlebar=False, title='Календарь')],
        [sg.Text('Имя файла', size=(13, 1)),
         sg.InputText(default_text=calc_month('%Y-%m') + '_logs', justification='right', size=(16, 1), key='file_xlsx'),
         sg.Text('.xslx', justification='left', size=(4, 1)),
         sg.Text('Имя листа', justification='right', size=(9, 1)),
         sg.InputText(default_text='1', size=(8, 1), key='ws_name')],
        [sg.Text('Колонки', size=(13, 1),),
         sg.InputText(default_text='Ролик', size=(9, 1), key='col1'),
         sg.InputText(default_text='Дата', size=(9, 1), key='col2'),
         sg.InputText(default_text='Время', size=(8, 1), key='col3'),
         sg.InputText(default_text='Хронометраж', size=(13, 1), key='col4')],
        [sg.Text('_' * 63)],
        [sg.Text('Настройки:')],
        [sg.Text('Cмещение времени'),
         sg.Combo(['+', '-'], default_value='+', readonly=True, size=(2, 1), key='offset_sign'),
         sg.Spin([i for i in range(0, 13)], initial_value=2, readonly=True, size=(3, 1), key='time_offset')],
        [sg.Text('Начало рекламного блока', size=(26, 1)),
         sg.Text('Окончание рекламного блока', size=(22, 1))],
        [sg.InputText(default_text='ID_REC_DTMF_In', size=(30, 1), key='ad_start'),
         sg.InputText(default_text='ID_REC_DTMF_Out', size=(30, 1), key='ad_end')],
        [sg.Text('Начало ролика', size=(26, 1)),
         sg.Text('Окончание ролика', size=(22, 1))],
        [sg.Input(default_text='CLIP START', size=(30, 1), key='start'),
         sg.InputText(default_text='CLIP STOP', size=(30, 1), key='end')],
        [sg.Text('_' * 63)],
        [sg.Text('Лог выполнения:', size=(17, 1))],
        [sg.Output(size=(60, 3))],
        [sg.Text(version_text, size=(37, 1), text_color=version_text_color),
         sg.Button('Сформировать отчёт', key='submit')]
    ]

    scheduler = [
        [sg.Text('Ключ запуска', size=(15, 1)),
         sg.Input(default_text=args.get('sendmail'), key=args.get('sendmail'), disabled=True, size=(43, 1))],
        [sg.Text('Папка c логами', size=(15, 1)),
         sg.Input(default_text=schedule_logs['path'], readonly=True, size=(34, 1),
                  key='schedule_logs_path', enable_events=True),
         sg.FolderBrowse()],
        [sg.Text('Формировать отчёт', size=(15, 1)),
         sg.Combo(['вчера', ], default_value='вчера', readonly=True, size=(8, 1)),
         sg.Combo(['05:59', '06:00'], default_value=schedule_logs['start'], readonly=True, size=(5, 1),
                  key='schedule_log_start'),
         sg.Text('– ', size=(1, 1)),
         sg.Combo(['сегодня', ], default_value='сегодня', readonly=True, size=(8, 1)),
         sg.Combo(['06:00', '06:01'], default_value=schedule_logs['end'], readonly=True, size=(5, 1),
                  key='schedule_log_end')],
        [sg.Text('_' * 63)],
        [sg.Text('Настройки почты:')],
        [sg.Text('SMTP сервер', size=(15, 1)),
         sg.Input(default_text=schedule_smtp['host'], key='host', size=(43, 1))],
        [sg.Text('SMTP порт', size=(15, 1)),
         sg.Input(default_text=schedule_smtp['port'], key='port', size=(43, 1))],
        [sg.Text('EMAIL отправителя', size=(15, 1)),
         sg.Input(default_text=schedule_smtp['mail'], key='mail', size=(43, 1))],
        [sg.Text('Пароль отправителя', size=(15, 1)),
         sg.Input(default_text=schedule_smtp['pass'], key='pass', size=(43, 1))],
        [sg.Text('_' * 63)],
        [sg.Text('Настройки письма:')],
        [sg.Text('Получатели', size=(15, 1)),
         sg.Input(default_text=schedule_email['recipients'], key='recipients', size=(43, 1))],
        [sg.Text('Заголовок', size=(15, 1)),
         sg.Input(default_text=schedule_email['subject_title'], key='subject_title', size=(43, 1))],
        [sg.Text('Текст письма', size=(15, 1)),
         sg.Input(default_text=schedule_email['body'], key='body', size=(43, 1))],
        [sg.Text('Имя файла', size=(15, 1)),
         sg.InputText(default_text=schedule_email['log_name'], key='schedule_log_name', size=(43, 1))],
        [sg.Text(' ' * 63, size=(1, 2))],
        [sg.Text(version_text, size=(28, 1), text_color=version_text_color),
         sg.Text('', key='success_text', size=(7, 1), justification='right'),
         sg.Button('Сохранить настройки', key='save_config', size=(16, 1)),
         sg.Button('OK', key='OK', size=(4, 1), visible=False)]
    ]

    tabs = [
        [sg.TabGroup(
            [[sg.Tab('Парсер логов в xlsx', main_layout), sg.Tab('Настройки для рассылки txt лога', scheduler)]],
            tab_location='centertop',
            title_color='Black',
            tab_background_color='Gray',
            border_width=1)]
    ]

    logo_icon = os.path.dirname(os.path.realpath(__file__)) + '/logo.ico'

    if trial_expired:
        window = sg.Window('PyLogParser', icon=logo_icon,
                           element_justification='c').Layout(trial_expired_layout)
    else:
        window = sg.Window('PyLogParser', icon=logo_icon).Layout(tabs)

    while True:
        event, values = window.read()
        if event in (None, 'OK', 'Exit', 'Cancel'):
            break
        if event == 'logs_path':
            gui_xlsx_path = values.get('logs_path')
            window['xlsx_path'].Update('/'.join(gui_xlsx_path.split('/')[:-1]))
            main_layout += [[sg.Button('Save'), sg.Button('Exit')]]
        if event == 'save_config':
            config_set(values)
            window['pass'].Update('********')
            window['save_config'].Update(visible=False)
            window['OK'].Update(visible=True)
            window.FindElement('success_text').Update('Настройки сохранены')
            window['success_text'].set_size((19, 1))
        if event == 'submit':
            main(values)
    window.close()


if not trial_expired and sendmail:
    debug_log('sendmail')
    scheduled_log()

    smtp, email = config_get('SMTP'), config_get('EMAIL')

    date_yesterday = dt.date.today() - dt.timedelta(days=1)
    email_title = email['subject_title'].replace('{date_yesterday}', date_yesterday.strftime('%d.%m.%Y'))
    # logfile = email['log_name']
    logfile = email['log_name'].replace('{date_yesterday}', date_yesterday.strftime('%d.%m.%Y'))

    mail = MIMEMultipart()
    mail['From'] = smtp['mail']
    mail['To'] = email['recipients']
    mail['Subject'] = email_title
    mail.attach(MIMEText(email['body']))
    fp = open(logfile, 'rb')
    attachment = MIMEBase('multipart', 'plain')
    attachment.set_payload(fp.read())
    fp.close()
    encoders.encode_base64(attachment)
    attachment.add_header('Content-Disposition', 'attachment', filename=logfile)
    mail.attach(attachment)

    debug_log('attach')

    session = smtplib.SMTP(smtp['host'], smtp['port'])
    session.starttls()
    session.login(smtp['mail'], password_crypt('decrypt', bytes(smtp['pass'], encoding='utf-8')).decode('utf-8'))
    session.sendmail(smtp['mail'], email['recipients'].split(','), mail.as_string())
    session.quit()

    # if os.path.exists(logfile):
    #     os.remove(logfile)

    debug_log('session')
else:
    debug_log('GUI')
    gui_interface()

debug_log('END')
