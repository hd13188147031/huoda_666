from threading import Thread
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd
import csv
import uuid
import time

def get_mac_address():
    mac = ':'.join(['{:02x}'.format((uuid.getnode() >> elements) & 0xff) for elements in range(2,7)][::-1])
    return mac

mac = get_mac_address()
password_1 = mac[8:10]
password_2 = mac.replace(password_1, '')
# 取出奇数位上的字符
odd_chars = password_2[1::2]  # 前
# 取出偶数位上的字符
even_chars = password_2[0::2]  # 后
key = odd_chars + even_chars
with open('./id.txt', 'w') as f:
    f.write(mac)
try:
    with open('./key.txt', 'r') as r:
        read_secret_key = r.read()
except:
    print('缺少授权文件，请联系管理员')
if key != read_secret_key:
    print('密钥错误，请联系管理员')

num = int(input('输入线程数'))
workbook = pd.read_excel(f'.\微博账号目录.xlsx', sheet_name='Sheet1').dropna(how='any')
workbook = workbook['账号名称']
workbook = [i for i in workbook]
def wb(threadnum):#threadnum 线程数
    # 计算每个部分至少应该有的元素数量
    quotient = len(workbook) // threadnum
    # 创建子列表
    parts_names = []
    start = 0
    for i in range(threadnum):
        end = start + quotient
        # 如果是最后一部分，则包含所有剩余的元素
        if i == threadnum - 1:
            end = len(workbook)
        parts_names.append(workbook[start:end])
        start = end
    return parts_names
parts_names = wb(threadnum = num) #返回一个二维列表

# 定义一个函数，该函数将被多个线程并发执行
def worker(thread_id, data_chunk):
    wb_list = []
    chrome = webdriver.Chrome()
    print(f"线程 {thread_id+1} 扫描站点： {data_chunk}")
    for i in data_chunk:
        for a in range(10):
            time.sleep(2)
            chrome.get(f'https://s.weibo.com/user?q={i}&page={a}')
            try:
                WebDriverWait(chrome, 10).until(EC.presence_of_element_located((By.CLASS_NAME, 'name')))
            except:
                print(f'无响应!!!.\微博账号目录:中 https://s.weibo.com/user?q={i}&page={a}')
            # 获取所有微博的账号名称
            names = chrome.find_elements(By.XPATH, '//a[@class="name"]')
            # 获取所有微博账号的发布单位
            names_unit = chrome.find_elements(By.XPATH, '//div/p[1]')
            # 获取所有微博账号的url
            url_list = chrome.find_elements(By.XPATH, '//div[@class="info"]//a')
            # 获取用户名和发布单位插入微博账号总表
            for o in range(len(names)):
                url = url_list[o].get_attribute('href')
                name = names[o].text
                if name[:1] == '-':
                    name = f' {name}'
                name_unit = names_unit[o].text
                if name_unit[1:3] in '粉丝：':
                    name_unit = ''
                wb_dict = {'账号名称': f'{name}', '发布单位': name_unit, '账号url': url}
                wb_list.append(wb_dict)
                print(wb_dict)
        # 将信息写入到csv文件中
        fieldnames = ['账号名称', '发布单位', '账号url']
        try:
            with open(f'.\微博账号_{i}.csv', 'w', encoding='utf-8', newline='') as f:
                write = csv.DictWriter(f, fieldnames=fieldnames)
                write.writeheader()
                for wb_name in wb_list:
                    write.writerow(wb_name)
        except PermissionError:
            print(f'请关闭excle文档：微博账号_{i}.csv并重新提取关键词{i}的微博账号')
    # 关闭浏览器
    chrome.quit()
# 创建线程列表
threads = []
# 为每个数据块创建一个线程
for i, data_chunk in enumerate(parts_names):# ----->重点部分
    # 创建一个Thread实例，指定worker函数为线程执行的函数
    # 并传递线程ID和数据块作为参数
    thread = Thread(target=worker, args=(i, data_chunk))
    # 将线程添加到线程列表中
    threads.append(thread)
    # 启动线程
    thread.start()

# 等待所有线程完成
for thread in threads:
    thread.join()
print()
print()
print(f'='*14+f'程序结束'+'='*14)
