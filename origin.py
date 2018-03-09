import requests
# import bs4
import xlsxwriter
# import time
import re
##################################

# -*- coding: utf-8 -*-
from selenium import webdriver
import bs4
from selenium.webdriver.common.action_chains import ActionChains
import time

url = 'https://www.tripadvisor.co.uk/Airline_Review-d8729020-Reviews-Cheap-Flights-American-Airlines#REVIEWS'
chromedriver_dir = "./chromedriver"
driver = webdriver.Chrome(chromedriver_dir)
driver.get(url)

ele = driver.find_element_by_class_name('ulBlueLinks')

actions = ActionChains(driver).click(ele)
actions.perform()
actions.perform()

time.sleep(1)
bs_test = bs4.BeautifulSoup(driver.page_source, 'lxml')
all_div = bs_test.find_all('div', class_='innerBubble')

j = 0
for i in all_div:
    if j % 2 is not 0:
        print(i.find('div', class_='entry').text)
    j = j + 1

#####################################

wb = xlsxwriter.Workbook('C:\\Users\\K\\Documents\\test180226_01.xlsx')
ws = wb.add_worksheet('sheet1')

ws.write(0, 0, "Title")
ws.write(0, 1, "Reviews")
ws.write(0, 2, "Name")
ws.write(0, 3, "Location")
ws.write(0, 6, "Date")
ws.write(0, 9, "Rate")
# ws.write(0,10,"Tag")
ws.write(0, 13, "NofHelp")

url1 = 'https://www.tripadvisor.co.uk/Airline_Review-d8729020-Reviews-Cheap-Flights'
url2 = '-American-Airlines#REVIEWS'

j = 1

for num in range(10):  ### num 이라는 ??변수를 0부터 99 까지 돌리는 것?

    if num == 0:  ###
        page = ''
    else:
        page = '-or' + str(num * 10)  ### 그냥 + 로 텍스트를 합칠 수 있는 것?
    real_url = url1 + page + url2

    site = requests.get(real_url)
    site_bs = bs4.BeautifulSoup(site.text, 'lxml')

    reviewSelector = site_bs.find_all('div', class_='reviewSelector')
    reviewSelector

    for c in reviewSelector:  ### for문 이해,, reviewSelector를 반복하는데 어떤 값을?

        ############  Title, Reviews, Name, Location###########################################
        Title = c.find("span", class_='noQuotes').text
        Reviews = c.find("p", class_='partial_entry').text
        if c.find("span", dir='auto'):
            Name = c.find("span", dir='auto').text
            # NName=c.find("span",class_=expand inline scrname mbrName BD0C46BB8FD550F5EC7EE84EEE3F1563)
            Location = c.find("div", class_='location').text
        else:
            Name = ''
            Location = ''

        ############  Date, Rate, Tag, ########################################################
        Date = c.find("span", class_='ratingDate').text
        Ratetag = c.find("div", class_='rating reviewItemInline')
        Rate = Ratetag.find('span')['class'][1].split("_")[1]
        Tag_all = c.find_all("span", class_='categoryLabel')
        Tag = []
        for t in Tag_all:
            Tag.append(t.text)

        if c.find('span', class_='numHlp'):
            NofHelp = c.find('span', class_='numHlp').find('span').text
        else:
            NofHelp = '0'

        print("[Title]", Title)
        print("[Reviews]", Reviews)
        print("[Name]", Name)  ### 띄아쓰기 없애기
        print("[Location]", Location)  ### \n 지우는 방법
        print("[Date]", Date)  ### 날짜 형식만 남기기,,
        print("[Rate]", Rate)  ### 10의자리를 1의자리로 바꾸기
        print("[Tag]", Tag)  ### 세번째 태그 경로 " - " 이것도 나누기,
        # print("[# of Help]",NofHelp)
        # print("[# of Reviews]",NofR)
        # print("[# of Total Help]",NofTH)

        ws.write(j, 0, Title)
        ws.write(j, 1, Reviews)
        ws.write(j, 2, Name)
        ws.write(j, 3, Location)
        date2 = Date.split(' ')
        ws.write(j, 6, date2[1] + date2[2])  ### 이걸 숫자만 남기는 방법
        ws.write(j, 9, Rate)
        ws.write_row(j, 10, Tag)  ### Tag 오류
        ws.write(j, 13, NofHelp)
        ws.write(j, 15, NofHelp)
        ws.write(j, 16, NofHelp)
        print("_________________________________")

        j = j + 1
        if num in [100, 200, 300, 400, 500, 600, 700, 800, 900, 1000, 1100]:
            time.sleep(10)

wb.close()