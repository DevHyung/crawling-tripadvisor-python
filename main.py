import requests
import time
import xlsxwriter

from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.common.action_chains import ActionChains

if __name__ == "__main__":
    url = 'https://www.tripadvisor.co.uk/Airline_Review-d8729060-Reviews-Cheap-Flights-Delta-Air-Lines'
    driver = webdriver.Chrome("./chromedriver")
    driver.get(url)
    driver.find_element_by_class_name('ulBlueLinks').click()
    time.sleep(1)
    # 엑셀 포맷 제작
    wb = xlsxwriter.Workbook('./test.xlsx')
    ws = wb.add_worksheet('sheet1')
    ws.write(0, 1, "제목")
    ws.write(0, 2, "전체별점")
    ws.write(0, 3, "게시날짜")
    ws.write(0, 4, "리뷰내용")
    ws.write(0, 5, "여행팁")
    ws.write(0, 6, "여행날짜")
    ws.write(0,7,'세부별점1(Legroom), (1~5)')
    ws.write(0,8,'세부별점2(Seat Comfort), (1~5)')
    ws.write(0,9,'세부별점3(Customer Service), (1~5)')
    ws.write(0,10,'세부별점4(Value for Money), (1~5)')
    ws.write(0,11,'세부별점5(Cleanliness), (1~5)')
    ws.write(0,12,'세부별점6(Check-in and Boarding), (1~5)')
    ws.write(0,13,'세부별점7(Food and Beverage), (1~5)')
    ws.write(0,14,'세부별점8(In-flight entertainment (WiFi, TV, films)), (1~5)')
    ws.write(0,15,'태그1(여행유형)')
    ws.write(0,16,'태그2(서비스유형)')
    ws.write(0,17,'태그3(출발지)')
    ws.write(0,18,'태그(도착지)')
    ws.write(0,19,'리뷰투표수')
    ws.write(0,20,'유저아이디')
    ws.write(0,21,'유저출신(나라)')
    ws.write(0,22,'유저출신(지역)')
    ws.write(0,23,'유저레벨')
    ws.write(0,24,'유저총리뷰수')
    ws.write(0,25,'유저총투표수')
    ws.write(0,26,'도시방문수')
    ws.write(0,27,'사진업로드수')
    ws.write(0,28,'유저리뷰분포(Excellent)')
    ws.write(0,29,'유저리뷰분포(Very good)	')
    ws.write(0,30,'유저리뷰분포(Average)')
    ws.write(0,31,'유저리뷰분포(Poor)')
    ws.write(0,32,'유저리뷰분포(Terrible)')

    j = 1
    for num in range(10):  ### num 이라는 ??변수를 0부터 99 까지 돌리는 것?
        if num == 0:  ###
            page = ''
        else:
            page = '-or' + str(num * 10)  ### 그냥 + 로 텍스트를 합칠 수 있는 것?
        real_url = url1 + page + url2

        site = requests.get(real_url)
        site_bs = BeautifulSoup(site.text, 'lxml')

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
