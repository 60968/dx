import requests
from datetime import datetime, timedelta
import pandas as pd
import os
import zipfile
import warnings
warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')

# 分析教室控制系统用户登录数据

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


def download_syslog(date_file):


    post_url = "http://192.168.16.240/Data/ExportSysLog"
    requests.post(post_url, data={"dt": date_file}, headers=headers)

    download_url = f"http://192.168.16.240/UserData/202/syslog.xlsx"
    file_response = requests.get(download_url, headers=headers, stream=True)
    with open(f"syslog_{date_file}.xlsx", 'wb') as f:
        for chunk in file_response.iter_content():
            f.write(chunk)
    print(f"文件下载成功: {date_file}")


def merge_excel_files():
    dfs = []
    for file in [f for f in os.listdir() if f.endswith('.xlsx')]:
        try:
            with zipfile.ZipFile(file, 'r') as zip_test:
                zip_test.testzip()
            df = pd.read_excel(file)
            if df.shape[0] > 0:
                dfs.append(df)
                print(f"成功加载: {file}，共{df.shape[0]}行")
            else:
                print(f"跳过空文件: {file}（无数据行）")
        except zipfile.BadZipFile:
            print(f"跳过损坏的文件: {file}")
        except Exception as e:
            print(f"处理文件 {file} 时出错: {e}")
            continue

    result_df = pd.concat(dfs, ignore_index=True)
    result_df.to_excel('all_syslog.xlsx', index=False)


def analyze_data():
    df = pd.read_excel('all_syslog.xlsx')
    second_column = df.iloc[:, 1]
    value_counts = second_column.value_counts()
    print(f"总数数据行数: {len(second_column)}")
    print(f"登录涉及人数： {len(value_counts)}\n")
    with open('login_analysis.txt', 'w', encoding='utf-8') as f:
        f.write(f"总数据行数: {len(second_column)}\n")
        f.write(f"登录涉及人数： {len(value_counts)}\n\n")
        for value, count in value_counts.items():
            f.write(f"{value} 出现次数: {count}\n")


if __name__ == "__main__":
    for file in os.listdir():
        if file.endswith('.xlsx'):
            os.remove(file)
            print(f"已清理: {file}")

    for i in range(1, 11):  # 前1天开始，共前 10 天日志
        download_syslog((datetime.now().date() - timedelta(days=i)).strftime('%Y-%m-%d'))
    merge_excel_files()
    analyze_data()
