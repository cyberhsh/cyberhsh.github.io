# cyberhsh.github.io

import random
import re
import time
import openpyxl
from bs4 import BeautifulSoup
from selenium.webdriver.support import expected_conditions as EC
import requests
import xlsxwriter as xw
import openpyxl as op
import json
from lxml import etree
from selenium.webdriver import ActionChains
from selenium.webdriver.chrome.service import Service
from selenium import webdriver
from selenium.webdriver.common.by import By
import datetime
import time

from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver import ChromeOptions
class spider():
    def __init__(self):
        self.num_jishu = 1

    def main(self,word_choose):
        service = Service(executable_path='D:\桌面\淘宝爬取\代码文件\chromedriver-win32\chromedriver.exe')
        driver = webdriver.Chrome(service=service)
        driver.set_window_size(800, 700)  # 设置打开的窗口大小
        headers = {
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/121.0.0.0 Safari/537.36"
        }

        # 点击最新的一篇文章
        url_one = 'https://www.taobao.com/'
        driver.get(url_one)
        time.sleep(0.5)
        login_click = driver.find_element(By.XPATH,'/html/body/div[3]/div[2]/div[2]/div[2]/div[6]/div/div[2]/div[1]/a[1]')
        login_click.click()
        time.sleep(0.5)

        # 转到登录页面
        driver.close()
        windows = driver.window_handles
        driver.switch_to.window(windows[-1])
        login_saoma =driver.find_element(By.XPATH,'//*[@id="login"]/div[1]/i')
        login_saoma.click()
        time.sleep(5)
        print("账号登录成功")
        print("开始爬取中，由于需要筛选，请稍等~")

        # 开始爬取商品
        windows = driver.window_handles
        driver.switch_to.window(windows[-1])
        time.sleep(1)
        shop_input = driver.find_element(By.XPATH,'//*[@id="q"]')
        shop_input.send_keys(word_choose)
        button_search = driver.find_element(By.XPATH,'//*[@id="J_TSearchForm"]/div[1]/button')
        button_search.click()
        time.sleep(1)

        # 第一个关键词--六款商品
        windows = driver.window_handles
        driver.switch_to.window(windows[-1])
        time.sleep(1)
        # 获取商品页中所有商品的链接
        href_list = []
        for num in range(1,49):
            windows = driver.window_handles
            driver.switch_to.window(windows[-1])
            shop_href = driver.find_element(By.XPATH,f'//*[@id="root"]/div/div[2]/div[1]/div[1]/div[2]/div[3]/div/div[{num}]/a')
            href = shop_href.get_attribute('href')
            href_list.append(href)
        for href in href_list:
            # 链接的跳转
            if self.num_jishu <= 6:
                # print(href)
                self.href_go(word_choose,driver, href)
                time.sleep(5)
            else:
                break

    def save_data(self,word_chooose,shop_title, ip_list, time_list, text_list):
        try:
            workbook = openpyxl.load_workbook(f'商品评论{word_choose}.xlsx')
            worksheet = workbook.active
        except FileNotFoundError:
            workbook = openpyxl.Workbook()
            worksheet = workbook.active
            worksheet.append(['标题', 'IP', '时间', '评论商品','评论内容'])  # 添加表头
        # 处理掉时间字段中的多余内容
        another_list = []
        time_new_list = []
        for data in time_list:
            title_one = data.split("&nbsp;")[0]
            another_one = data.split("&nbsp;")[1]
            time_new_list.append(title_one)
            another_list.append(another_one)
        # print(time_new_list,another_list)
        # 处理掉评论内容中的span
        text_new_list = []
        for data in text_list:
            data = data.replace("<span>","").replace("</span>","").replace('<span class="Comment--appendInternal--1i_CMIz">','')
            text_new_list.append(data)
        # print(text_new_list)

        # 追加数据
        for ip, time_new, another_thing,text_new in zip(ip_list, time_new_list, another_list,text_new_list):
            worksheet.append([shop_title, ip, time_new, another_thing,text_new])

        # 保存Excel文件
        workbook.save(f'商品评论{word_chooose}.xlsx')

    def href_go(self,word_choose,driver,href):
        driver.get(href)
        time.sleep(2)
        title_text = driver.page_source
        # 衣服的标题
        html_title = etree.HTML(title_text)
        shop_title = html_title.xpath('//*[@id="root"]/div/div[2]/div[2]/div[1]/div/div[2]/div[1]/h1/text()')[0]
        # print(shop_title)
        # shop_title =re.findall(r'<title>(.*?)</title>',text)[0]
        # print(shop_title)
        # 点击"宝贝评价"
        think_person_button = driver.find_element(By.XPATH, '//*[@id="root"]/div/div[2]/div[2]/div[2]/div[1]/div/div/div[2]/span')
        think_person_button.click()
        time.sleep(2)
        windows = driver.window_handles
        driver.switch_to.window(windows[-1])
        html_text = driver.page_source
        one_ip = re.findall(r'<div class="Comment--userName--2cONG4D">(.*?)</div>', html_text)
        one_time = re.findall(r'<div class="Comment--meta--1MFXGJ1">(.*?)</div>', html_text)
        one_text = re.findall(r'<div class="Comment--content--15w7fKj">(.*?)</div>', html_text)
        # if len(one_ip) == 0 and len(one_time) == 0 and len(one_text) == 0:
        #     print("一条评价都没有")
        # 把数据都保存到一个列表中
        ip_list = []
        time_list = []
        text_list = []
        for ip,a_time,text in zip(one_ip,one_time,one_text):
            ip_list.append(ip)
            time_list.append(a_time)
            text_list.append(text)
        # 实现翻页
        A = True
        while A:
            try:
                next_page = driver.find_element(By.XPATH,'//*[@id="root"]/div/div[2]/div[2]/div[2]/div[2]/div[2]/div/div[3]/div/button[@class="detail-btn Comments--nextBtn--1itIAip"]')
                if next_page.is_enabled():
                    next_page.click()
                else:
                    break
                time.sleep(1)
                windows = driver.window_handles
                driver.switch_to.window(windows[-1])
                html_text = driver.page_source
                one_ip = re.findall(r'<div class="Comment--userName--2cONG4D">(.*?)</div>', html_text)
                one_time = re.findall(r'<div class="Comment--meta--1MFXGJ1">(.*?)</div>', html_text)
                one_text = re.findall(r'<div class="Comment--content--15w7fKj">(.*?)</div>', html_text)
                for ip, a_time, text in zip(one_ip, one_time, one_text):
                    ip_list.append(ip)
                    time_list.append(a_time)
                    text_list.append(text)
            except:
                A = False
        if len(ip_list) >= 60:
            # print(href)
            # for i,j,k in zip(ip_list,time_list,text_list):
            #     print(i,j,k)

            try:
                self.save_data(word_choose,shop_title,ip_list,time_list,text_list)
                print(f"第{self.num_jishu}款商品爬取成功")
                print("**************************************************************")
                self.num_jishu += 1
            except:
                pass

        

if __name__ == '__main__':
    while True:
        word_choose = input("请输入关键词：")
        if word_choose == "stop":
            print("爬取结束，欢迎下次继续")
            break
        else:
            a = spider()
            a.main(word_choose)
