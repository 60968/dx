import re
import time
import schedule
import datetime
import warnings
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.edge.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
warnings.filterwarnings("ignore", category=UserWarning)


def get_class_info():
    excel_file = '2024年电教科教学综合楼保障情况记录表.xlsx'
    df = pd.read_excel(excel_file, sheet_name=datetime.datetime.now().strftime("%m.%d"))
    g_column_data_str = [str(item) for item in df.iloc[:, 6].tolist()]
    g_class = [re.search(r'\d+', item).group() for item in g_column_data_str if re.search(r'\d+', item)]
    the_driver, the_list = class_control()
    yes_control_class = [i[1:] for i in the_list.keys()]
    no_control_class = ['101', '102', '103', '106', '309']
    all_class = [i for i in g_class if i in yes_control_class and i not in no_control_class]
    classes = ['_' + cls for cls in set(all_class)]
    print(f'自动{set(all_class)}\n手动{set(g_class)-set(all_class)}')
    return classes


def class_control():
    edge_options = Options()
    edge_options.add_argument('--headless')
    edge_options.add_argument('--disable-gpu')
    driver = webdriver.Edge(options=edge_options)
    driver.get('http://192.168.16.240/Home/Login')
    WebDriverWait(driver, 9).until(
        EC.presence_of_element_located((By.XPATH, '//*[@id="uid"]'))).send_keys('yuanzhao')
    WebDriverWait(driver, 9).until(
        EC.presence_of_element_located((By.XPATH, '//*[@id="pwd"]'))).send_keys('111111')
    WebDriverWait(driver, 9).until(
        EC.presence_of_element_located((By.XPATH, '/html/body/div/div[3]/div[4]'))).click()  # 登录
    WebDriverWait(driver, 9).until(
        EC.presence_of_element_located((By.XPATH, '//*[@id="taskbar"]/div[4]/div[1]'))).click()  # 菜单
    WebDriverWait(driver, 9).until(
        EC.presence_of_element_located((By.XPATH, '/html/body/div[5]/div[4]/div[2]'))).click()  # 设备管控
    WebDriverWait(driver, 9).until(
        EC.presence_of_element_located((By.XPATH, '/html/body/div[6]/div[2]/div[2]'))).click()  # 设备控制

    class_list = {
        # '_101' : '//*[@id="sys6_tree_3_span"]', '_102' : '//*[@id="sys6_tree_4_span"]',
        # '_106' : '//*[@id="sys6_tree_8_span"]', '_103' : '//*[@id="sys6_tree_5_span"]',
        # '_309' : '//*[@id="sys6_tree_19_span"]',
        '_104' : '//*[@id="sys6_tree_6_span"]', '_105' : '//*[@id="sys6_tree_7_span"]',
        '_202' : '//*[@id="sys6_tree_9_span"]', '_203' : '//*[@id="sys6_tree_10_span"]',
        '_204' : '//*[@id="sys6_tree_11_span"]', '_302' : '//*[@id="sys6_tree_12_span"]',
        '_303' : '//*[@id="sys6_tree_13_span"]', '_304' : '//*[@id="sys6_tree_14_span"]',
        '_305' : '//*[@id="sys6_tree_15_span"]', '_306' : '//*[@id="sys6_tree_16_span"]',
        '_307' : '//*[@id="sys6_tree_17_span"]', '_308' : '//*[@id="sys6_tree_18_span"]',
        '_401' : '//*[@id="sys6_tree_20_span"]', '_403' : '//*[@id="sys6_tree_21_span"]',
        '_404' : '//*[@id="sys6_tree_22_span"]', '_410' : '//*[@id="sys6_tree_23_span"]',
        '_411' : '//*[@id="sys6_tree_24_span"]', '_412' : '//*[@id="sys6_tree_25_span"]',
        '_413' : '//*[@id="sys6_tree_26_span"]', '_414' : '//*[@id="sys6_tree_27_span"]',
    }

    return driver, class_list


def class_begin():
    print()
    the_driver, the_list = class_control()
    for k in get_class_info():
        try:
            WebDriverWait(the_driver, 9).until(
                EC.presence_of_element_located((By.XPATH, the_list[k]))).click()
            time.sleep(3)
            WebDriverWait(the_driver, 9).until(
                EC.presence_of_element_located(
                    (By.XPATH, '//*[@id="viewport"]/div/div[2]/div[2]/div[10]/button[1]'))).click()
        except:
            WebDriverWait(the_driver, 9).until(
                EC.presence_of_element_located((By.XPATH, '/html/body/div[7]/div/div/div[1]/span[2]'))).click()
            time.sleep(3)
            print(f'开启{k[1:]}失败')
            continue
        else:
            print(f'已开启{k[1:]}')
    the_driver.quit()


def class_over():
    print()
    the_driver, the_list = class_control()
    for k in the_list:
        try:
            WebDriverWait(the_driver, 9).until(
                EC.presence_of_element_located((By.XPATH, the_list[k]))).click()
            time.sleep(3)
            WebDriverWait(the_driver, 9).until(
                EC.presence_of_element_located(
                    (By.XPATH, '//*[@id="viewport"]/div/div[2]/div[2]/div[10]/button[2]'))).click()
        except:
            WebDriverWait(the_driver, 9).until(
                EC.presence_of_element_located((By.XPATH, '/html/body/div[7]/div/div/div[1]/span[2]'))).click()
            time.sleep(3)
            print(f'关闭{k[1:]}失败')
            continue
        else:
            print(f'已关闭{k[1:]}')
    the_driver.quit()


if __name__ == '__main__':
    schedule.every().day.at("07:00").do(class_begin)  # 07:00
    schedule.every().day.at("22:00").do(class_over)  # 22:00
    while True:
        print(f'\r{datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")}', end='')
        schedule.run_pending()
        time.sleep(1)