import re
import time
import requests
import schedule
import datetime
import pandas as pd


class_list = {104: '46', 105: '47',
              202: '48', 203: '49', 204: '50',
              302: '51', 303: '52', 304: '53', 305: '54', 306: '55', 307: '56', 308: '57',
              401: '60', 403: '61', 404: '62', 410: '63', 411: '64', 412: '65', 413: '66', 414: '67'}
url_class = "http://192.168.16.240/Data/CC"
headers_class = {
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


def get_school_info():
    today = datetime.datetime.now().strftime("%Y-%m-%d")
    url_school = "http://192.168.254.174/teas/schedule/course/arrange/exportClassOverSchedule"
    params = {'classId': 'e2ffc5a884ad4d78971c312dde62ae67,9e829698891144ddb4236b39339a7425',
        'sDate': today, 'eDate': today, 'publish': '1'}
    headers = {
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,/;q=0.8,application/signed-exchange;v=b3;q=0.7",
        "Accept-Encoding": "gzip, deflate",
        "Accept-Language": "zh-CN,zh;q=0.9,en;q=0.8,en-GB;q=0.7,en-US;q=0.6",
        "Connection": "keep-alive",
        "Cookie": "authorization_token=eyJTRVNTSU9OSUQiOiI1OWJlNTg3NDkxNTI0NGE0ODlmMzAzZmJhYTVjNjA3MyIsInR5cCI6IkpXVCIsImFsZyI6IkhTMjU2In0.eyJpc3MiOiJXZW5KIiwiZXhwIjoxNzM1MzI2NTU3fQ.rUpBJ0dgj5QOyvEcc2w2Po0b3R6Hj18JuTVRy6BuD0M; user_name=%E8%A2%81%E9%92%8A; system_global=; photoRelativePath=%5B%7B%22originalFileName%22%3A%22%E8%A2%81%E9%92%8A.jpg%22%2C%22name%22%3A%22%E8%A2%81%E9%92%8A%22%2C%22suffix%22%3A%22jpg%22%2C%22size%22%3A9962%2C%22contentType%22%3A%22image%2Fjpeg%22%2C%22relativePath%22%3A%22%2Fupload%2Ffiles%2F20241029%2F45ec9a16b57e4e4d8aed2deeceddbf88.jpg%22%2C%22absolutePath%22%3A%22https%3A%2F%2F192.168.254.174%2Fupload%2Ffiles%2F20241029%2F45ec9a16b57e4e4d8aed2deeceddbf88.jpg%22%2C%22uploadDate%22%3A%222024-10-29+15%3A10%3A05%22%2C%22securityLevel%22%3Anull%2C%22id%22%3A%22110b3ae43f9249a0b17cb31f3c47b220%22%7D%5D; branchEnable=0; app_auto_login_key=e286c61dd6a61b9191c454595298ac1e6485bab0aa4ca18022e179d5c18cec16cc5784d9697845af65f1aff71f8cfd1a",
        "Host": "192.168.254.174",
        "Referer": "http://192.168.254.174/dsf5/index.html",
        "Upgrade-Insecure-Requests": "1",
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36 Edg/131.0.0.0"
    }
    response = requests.get(url_school, params=params, headers=headers, stream=True)
    if response.status_code == 200:
        with open(f'{today}课程表.xlsx', 'wb') as f:
            for chunk in response.iter_content(chunk_size=8192):
                if chunk:
                    f.write(chunk)
    try:
        df = pd.read_excel(f'{today}课程表.xlsx', sheet_name='Sheet1')
    except Exception as err:
        print(err)
        return []
    else:
        g_column_data_str = [str(i) for i in df.iloc[:, 6].tolist()]
        filtered_data = [item for item in g_column_data_str if '教学' in item]
        g_we_class = set([int(re.search(r'\d+', i).group()) for i in filtered_data if re.search(r'\d+', i)])
        auto_classes = set([i for i in g_we_class if i in [i for i in class_list.keys()]])
        print(f'智慧校园自动{auto_classes}\n智慧校园手动{g_we_class-auto_classes}')
        return auto_classes


def get_class_info():
    try:
        excel_file = '2024年电教科教学综合楼保障情况记录表.xlsx'
        df = pd.read_excel(excel_file, sheet_name=datetime.datetime.now().strftime("%m.%d"))
    except Exception as err:
        print(err)
        return []
    else:
        g_column_data_str = [str(i) for i in df.iloc[:, 6].tolist()]
        g_class = set([int(re.search(r'\d+', i).group()) for i in g_column_data_str if re.search(r'\d+', i)])
        auto_classes = set([i for i in g_class if i in [i for i in class_list.keys()]])
        print(f'记录表自动{auto_classes}\n记录表手动{g_class-auto_classes}')
        return auto_classes


def class_begin():
    class_info = get_class_info()
    if class_info:
        for i in class_info:
            payload = {'ip': f'192.168.16.{class_list[i]}', 'o': '17', 'v': '1', 'id': '0'}
            response = requests.post(url_class, data=payload, headers=headers_class)
            print(f'{i}: {response.text}')
            time.sleep(1)


def class_over():
    for i in class_list:
        payload = {'ip': f'192.168.16.{class_list[i]}', 'o': '17', 'v': '0', 'id': '0'}
        response = requests.post(url_class, data=payload, headers=headers_class)
        print(f'{i}: {response.text}')
        time.sleep(1)


if __name__ == '__main__':
    schedule.every().day.at("07:00").do(class_begin)  # 07:00
    schedule.every().day.at("22:00").do(class_over)  # 22:00
    while True:
        print(f'\r{datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")}', end='')
        schedule.run_pending()
        time.sleep(1)
