import pyautogui
import time
from PIL import Image
from selenium import webdriver
from bs4 import BeautifulSoup
import csv
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
import pandas as pd
import pyperclip
import uuid


def get_mac_address():
    mac = ':'.join(['{:02x}'.format((uuid.getnode() >> elements) & 0xff) for elements in range(2,7)][::-1])
    return mac

# 微信点击事件-点击微信输入框
def input_and_search():
    print(f'---------------------开始提取关于： {name} 的公众号-----------------------')
    # 读取要识别的搜索图片
    template_search = Image.open(r'.\search.png')
    # 在屏幕截图中查找搜索图片位置
    location_search = pyautogui.locateOnScreen(template_search)
    if location_search is not None:
        # 如果找到图片，根据搜索图片位置计算中心点坐标并向左偏移到输入框
        search_center_x = location_search.left + (location_search.width / 2)
        search_center_y = location_search.top + (location_search.height / 2)
        print(f'输入框点击成功: {search_center_x-200}, {search_center_y}')
        time.sleep(0.5)
        # 点击输入框
        pyautogui.click(search_center_x-200, search_center_y)
        pyautogui.hotkey('ctrl','a')
        pyautogui.hotkey('ctrl','v')
        # 还原搜索按钮本身的坐标位置并点击
        pyautogui.click(search_center_x, search_center_y)
        print(f"搜索框点击成功: {search_center_x}, {search_center_y}")
        pyautogui.moveTo(search_center_x,search_center_y+50)
        time.sleep(0.5)
    else:
        print("未找到搜索图片")


# 微信点击事件-更多
def more():
    try:
        # 读取要识别的更多图片
        template_more = Image.open(r'.\more.png')
        # 在屏幕截图中查找更多图片位置
        location_more = pyautogui.locateOnScreen(template_more)
        if location_more is not None:
            # 如果找到图片，则计算中心点坐标
            more_center_x = location_more.left + (location_more.width / 2)
            more_center_y = location_more.top + (location_more.height / 2)
            time.sleep(0.25)
            # 点击更多按钮
            pyautogui.click(more_center_x, more_center_y)
            time.sleep(0.5)
            print(f"更多按钮点击成功: {more_center_x}, {more_center_y}")
    except:
        print('未找到更多图片')

# 微信监测事件-翻页+停止
def keep_and_stop():
    while True:
        try:
            pyautogui.hotkey('end')
            template_stop = Image.open(r'.\stop.png')
            location_stop = pyautogui.locateOnScreen(template_stop)
            if location_stop is not None:
                print('翻页结束')
                # stop_center_x = location_stop.left + (location_stop.width / 2)
                # stop_center_y = location_stop.top + (location_stop.height / 2)
                # pyautogui.click(stop_center_x,stop_center_y)
                pyautogui.hotkey('ctrl','a')
                pyautogui.hotkey('ctrl','c')
                break
        except:
            pyautogui.hotkey('end')

'''
###以上是控制微信部分###
==================================================
==================================================
==================================================
###以下是提取微信部分###
'''


if __name__ == '__main__':
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
    else:
        column_name = '公众号名称'
        workbook = pd.read_excel(r'.\公众号目录.xlsx', sheet_name='Sheet1')
        # 去除表格内的空行
        workbook = workbook.dropna(how='any')
        for name in workbook[column_name]:
            # 复制name到剪切板
            pyperclip.copy(name)
            # 调用input_and_search函数
            input_and_search()
            time.sleep(0.5)
            pyautogui.hotkey('end')
            time.sleep(1)
            # 调用more()函数
            more()
            # 调用keep_and_stop函数
            keep_and_stop()
            # url = 'https://ckeditor.com/ckeditor-5/demo/feature-rich/'
            url = 'http://localhost/#/routeing'
            chrome = webdriver.Chrome()
            chrome.get(url)
            time.sleep(1.5)
            # 切换到 iframe
            iframe = chrome.find_element(By.ID, "mce_0_ifr")
            chrome.switch_to.frame(iframe)

            # 定位编辑区域
            edit_area = chrome.find_element(By.ID, "tinymce")

            # 使用 ActionChains 类模拟鼠标点击，并在编辑区域输入文本
            actions = ActionChains(chrome)
            actions.move_to_element(edit_area).click().perform()

            # 模拟键盘复制粘贴
            pyautogui.hotkey('ctrl', 'a')
            time.sleep(0.5)
            pyautogui.hotkey('ctrl', 'v')
            time.sleep(1.5)
            # 获取源码
            data = chrome.page_source
            # 公众号名称
            soup = BeautifulSoup(data, 'html.parser')
            # 公众号名称列表
            new_list = [p.get_text() for p in soup.find_all('span',class_="header-title")]

            # 发布单位
            soup_2 = BeautifulSoup(data, 'html.parser')
            # 发布单位列表
            address_list = [p.get_text() for p in soup.find_all('span', class_="header-source-text")]

            # 发布单位属性
            soup_3 = BeautifulSoup(data, 'html.parser')
            # 发布单位属性列表
            span_tags = soup.find_all('div', class_="header-title-container")
            '''
            发布单位属性列表的高级写法
            address_attribute_list = [p_tag.get_text() if p_tag else "" for span in span_tags for p_tag in [span.find('div')]]
            '''
            address_attribute_list = []
            for span in span_tags:
                p_tag = span.find('div')
                if p_tag:
                    # print(p_tag.get_text())
                    address_attribute_list.append(p_tag.get_text())
                else:
                    p_tag = ""
                    address_attribute_list.append(p_tag)
            # 将公众号名称与发布单位整合到GZH_list
            GZH_list = []
            for i in range(len(new_list)):
                dirs = {'公众号名称': new_list[i], '发布单位': address_list[i],'发布单位属性':address_attribute_list[i]}
                GZH_list.append(dirs)

            # 将信息写入到csv文件中
            fieldnames = ['公众号名称', '发布单位','发布单位属性']
            with open(f'.\公众号_{name}.csv', 'w', encoding='utf-8', newline='') as f:
                write = csv.DictWriter(f, fieldnames=fieldnames)
                write.writeheader()
                for GZH in GZH_list:
                    write.writerow(GZH)
            chrome.quit()
            print(f'关于: {name} 的公众号提取成功!!!')
    print('！！！！！！程序结束！！！！！！')