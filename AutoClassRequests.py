import re
import time
import requests
import schedule
import datetime
import warnings
import pandas as pd
warnings.filterwarnings("ignore", category=UserWarning)


class_list = {104: '46', 105: '47',
              202: '48', 203: '49', 204: '50',
              302: '51', 303: '52', 304: '53', 305: '54', 306: '55', 307: '56', 308: '57',
              401: '60', 403: '61', 404: '62', 410: '63', 411: '64', 412: '65', 413: '66', 414: '67'}
url = "http://192.168.16.240/Data/CC"
headers = {
    'Accept': 'text/plain, */*; q=0.01',
    'Accept-Encoding': 'gzip, deflate',
    'Accept-Language': 'zh-CN,zh;q=0.9,en;q=0.8,en-GB;q=0.7,en-US;q=0.6',
    'Connection': 'keep-alive',
    'Content-Type': 'application/x-www-form-urlencoded; charset=UTF-8',
    'Cookie': 'shit=; rm=yuanzhao; u=202; uid=yuanzhao; pwd=111111',
    'Host': '192.168.16.240',
    'Origin': 'http://192.168.16.240',
    'Referer': 'http://192.168.16.240/Home/Index',
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36 Edg/131.0.0.0',
    'X-Requested-With': 'XMLHttpRequest'
}

def get_class_info():
    excel_file = '2024年电教科教学综合楼保障情况记录表.xlsx'
    df = pd.read_excel(excel_file, sheet_name=datetime.datetime.now().strftime("%m.%d"))
    g_column_data_str = [str(item) for item in df.iloc[:, 6].tolist()]
    g_class = set([int(re.search(r'\d+', i).group()) for i in g_column_data_str if re.search(r'\d+', i)])
    auto_classes = set([i for i in g_class if i in [i for i in class_list.keys()]])
    print(f'自动{auto_classes}\n手动{g_class-auto_classes}')
    return auto_classes


def class_begin():
    for i in get_class_info():
        payload = {'ip': f'192.168.16.{class_list[i]}', 'o': '17', 'v': '1', 'id': '0'}
        response = requests.post(url, data=payload, headers=headers)
        print(f'{i}: {response.text}')
        time.sleep(1)


def class_over():
    for i in class_list:
        payload = {'ip': f'192.168.16.{class_list[i]}', 'o': '17', 'v': '0', 'id': '0'}
        response = requests.post(url, data=payload, headers=headers)
        print(f'{i}: {response.text}')
        time.sleep(1)


if __name__ == '__main__':
    schedule.every().day.at("07:00").do(class_begin)  # 07:00
    schedule.every().day.at("22:00").do(class_over)  # 22:00
    while True:
        print(f'\r{datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")}', end='')
        schedule.run_pending()
        time.sleep(1)
