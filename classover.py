import time
import requests

# 关闭所有教室的设备

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


for i in class_list:
    payload = {'ip': f'192.168.16.{class_list[i]}', 'o': '17', 'v': '0', 'id': '0'}
    response = requests.post(url_class, data=payload, headers=headers_class)
    print(f'{i}: {response.text}')
    time.sleep(1)



