'''
Author: Gerald Liu
'''

import datetime
import time
import win32com.client
import pandas as pd

DATA_FILENAME = 'DATA.XLSX'
EMAIL_DATE = '20220705'
START_TIME = '09:00'
END_TIME = '11:00'
INTERVAL = 600 # seconds
EMAIL_TOFIND_SUBJECT_PREFIX = '今日数据已准备-'
EMAIL_TOSEND_RECIPIENT = 'gerald.w.liu@gmail.com'
EMAIL_TOSEND_SUBJECT_PREFIX = 'FT TRDADE'


def has_subject(mail, subject):
    return mail.Subject == subject


"""
folder_name = {'inbox', 'sent'}
"""
def search_folder(ol, folder_name, criteria_func, criteria_func_kwargs):
    mapi = ol.GetNamespace('MAPI')

    if folder_name == 'inbox':
        folder_id = 6
    elif folder_name == 'sent':
        folder_id = 5

    folder = mapi.GetDefaultFolder(folder_id)
    folder_mails = folder.Items

    for mail in folder_mails:
        criteria_func_kwargs['mail'] = mail
        if criteria_func(**criteria_func_kwargs):
            return True
    
    return False


def send_mail(ol, to, subject, html_body):
    mail = ol.CreateItem(0)
    mail.To = to
    mail.Subject = subject
    mail.HTMLBody = html_body

    mail.Send()


"""
timestamp: pd.Timestamp, e.g. 2022-07-05 18:55:10
date: str %Y%m%d, e.g. 20220705
"""
def date_equals(timestamp, date):
    date_str = timestamp.date().strftime('%Y%m%d')
    return date_str == date


'''
ATTENTION: filtering standard is hard-coded
date: str %Y%m%d, e.g. 20220705
'''
def get_data(filename, date):
    df = pd.read_excel(filename)
    
    # 筛选标准1：某某实体 为 CXCCFT
    mask_entity = (df['某某实体'] == 'CXCCFT')
    # 筛选标准2：子账户名称 不为 CXCC
    mask_subac = (df['子账户名称'] != 'CXCC')

    # 新开合约筛选标准：创建日期为指定日期
    mask_new_1 = df['合约创建时间（系统信息）'].apply(lambda x: date_equals(x, date))
    mask_new = mask_entity & mask_subac & mask_new_1

    # 终止合约筛选标准：更新日期为指定日期，且状态为已到期
    mask_terminated_1 = (df['合约状态（系统信息）'] == '已到期')
    mask_terminated_2 = df['合约更新时间（系统信息）'].apply(lambda x: date_equals(x, date))
    mask_terminated = mask_entity & mask_subac & mask_terminated_1 & mask_terminated_2

    df_new = df[mask_new]
    df_terminated = df[mask_terminated]

    df_new_out = df_new[['合约名称', '子账户名称']]
    df_terminated_out = df_terminated[['合约名称', '子账户名称']]

    return (df_new_out, df_terminated_out)


'''
Generate the body of the mail, in HTML format
'''
def get_mail_body(df_new, df_terminated):
    # reset index, avoid duplicate column names
    df_new = df_new.reset_index(drop=True)
    df_new.columns = ['合约名称_新开', '子账户名称_新开']
    df_terminated = df_terminated.reset_index(drop=True)
    df_terminated.columns = ['合约名称_终止', '子账户名称_终止']
    
    # Merge dataframes
    merged_df = pd.concat([df_new, df_terminated], axis=1)
    merged_df = merged_df.fillna('')

    # Format merged dataframe as HTML table
    header_1 = '新交易'
    header_2 = '终止合约（包括部分了结）'
    col_1, col_2 = df_new.columns
    col_3, col_4 = df_terminated.columns
    col_names = []
    for col in [col_1, col_2, col_3, col_4]:
        col_names.append(col.split('_')[0])
    
    border_style = 'style="border: 1px solid black;"'
    border_th_style = 'style="border: 1px solid black; width: 25%;"'

    html_table = '<table style="border-collapse: collapse; width: 100%; text-align: center;">'
    html_table += f'''
        <tr>
        <th colspan="2" {border_style}>{header_1}</th>
        <th colspan="2" {border_style}>{header_2}</th>
        </tr>
    '''
    html_table += f'''
        <tr {border_style}>
        <th {border_th_style}>{col_names[0]}</th>
        <th {border_th_style}>{col_names[1]}</th>
        <th {border_th_style}>{col_names[2]}</th>
        <th {border_th_style}>{col_names[3]}</th>
        </tr>
    '''
    for _, row in merged_df.iterrows():
        html_table += f'''
            <tr {border_style}>
            <td {border_style}>{row[col_1]}</td>
            <td {border_style}>{row[col_2]}</td>
            <td {border_style}>{row[col_3]}</td>
            <td {border_style}>{row[col_4]}</td>
            </tr>
        '''
    html_table += '</table>'

    # draft the rest of the email
    date_datetime = datetime.datetime.strptime(EMAIL_DATE, '%Y%m%d')
    month = date_datetime.month
    day = date_datetime.day
    new_count = df_new.shape[0]
    terminated_count = df_terminated.shape[0]

    html_body = f"""
        <html><body style="font-family: DengXian; font-size: 11pt">
        Dear All<br><br>
        {month}/{day}日，新开交易{new_count}笔，终止合约{terminated_count}笔，烦请核对，谢谢<br><br>
        {html_table}<br><br>
        </body></html>
    """

    return html_body


"""
if the mail is found in inbox or sent items, send a mail accordingly 
"""
def func_to_exec(search_inbox_kwargs, search_sent_kwargs):
    if search_folder(**search_inbox_kwargs) or search_folder(**search_sent_kwargs):
        df_new, df_terminated = get_data(DATA_FILENAME, EMAIL_DATE)

        send_mail_kwargs = {
            'ol': outlook,
            'to': EMAIL_TOSEND_RECIPIENT,
            'subject': f'{EMAIL_TOSEND_SUBJECT_PREFIX} {EMAIL_DATE}',
            'html_body': get_mail_body(df_new, df_terminated)
        }
        
        send_mail(**send_mail_kwargs)
        return True
    else:
        return False

""" 
start_time, end_time: %H:%M, e.g. 09:00
interval: in seconds, e.g. 600 (=10min)
func: function to execute. if it returns true, break the loop.
func_kwargs: dict of kwargs
wait: if False, only runs between start_time and end_time. if True, keep running
    and wait till the start_time comes, and will need to be ended mannually.
loop: if True, keep trying for every interval. if False, only try once.
"""
def timed_execution(
    start_time, end_time, interval, func, func_kwargs, wait=False, loop=True
):
    start_datetime = datetime.datetime.strptime(start_time, '%H:%M').time()
    end_datetime = datetime.datetime.strptime(end_time, '%H:%M').time()
    completed = False

    while True:
        now = datetime.datetime.now().time()

        if now < start_datetime:
            print('搜索邮件：还未到开始时间')
            if not wait:
                break
        elif now <= end_datetime:
            completed =  func(**func_kwargs)
            if completed:
                print('搜索邮件：已找到邮件\n发送邮件：已发送')
                break
            else:
                print('搜索邮件: 未找到相应邮件')
                if not loop:
                    break
        else:
            print('搜索邮件：还未到开始时间')
            if not wait:
                break


        time.sleep(interval)


outlook = win32com.client.Dispatch('outlook.application')

has_subject_kwargs = {
    'mail': None,
    'subject': f'{EMAIL_TOFIND_SUBJECT_PREFIX}{EMAIL_DATE}'
}

search_inbox_kwargs = {
    'ol': outlook,
    'folder_name': 'inbox',
    'criteria_func': has_subject,
    'criteria_func_kwargs': has_subject_kwargs
}

search_sent_kwargs = {
    'ol': outlook,
    'folder_name': 'sent',
    'criteria_func': has_subject,
    'criteria_func_kwargs': has_subject_kwargs
}

func_to_exec_kwargs = {
    'search_inbox_kwargs': search_inbox_kwargs,
    'search_sent_kwargs': search_sent_kwargs,
}


timed_execution(
    START_TIME, END_TIME, INTERVAL, func_to_exec, func_to_exec_kwargs,
    wait=False, loop=True
)
