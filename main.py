import requests
import time
import xlsxwriter
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.common.keys import Keys

if __name__ == "__main__":
    # timeout 에러나
    # socket 에러뜨니까 수정 요망 
    url = 'https://www.tripadvisor.co.uk/Airline_Review-d8729060-Reviews-Cheap-Flights-Delta-Air-Lines'
    ### 드라이버 셋팅
    driver = webdriver.Chrome("./chromedriver")
    driver.maximize_window()
    driver.get(url)
    driver.find_element_by_class_name('ulBlueLinks').click() #more 클릭
    time.sleep(1)
    #### 엑셀 포맷 제작
    now = time.localtime()
    filename = "%04d%02d%02d_%02d%02d%02d추출" % (now.tm_year, now.tm_mon, now.tm_mday, now.tm_hour, now.tm_min, now.tm_sec)
    wb = xlsxwriter.Workbook('./'+filename+'.xlsx')
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
    ws.write(0, 33, '이미지링크')
    ### 파싱시작
    j = 1
    for num in range(100): # 100페이지까지 이부분을 수정하면 원하는페이지 파싱가능
        if num == 0: #원래코드유지
            page = ''
            real_url = url
        else:
            page = '-or' + str(num * 10)
            real_url = url + page
            driver.get(real_url)
            driver.find_element_by_class_name('ulBlueLinks').click()
            time.sleep(1)
        # 페이지를 아래로 한번내려야 모든이미지리소스들이 DOM에 적용이되어서
        # 한번페이지를 다 내린후
        elem = driver.find_element_by_tag_name("body")
        for _ in range(5):
            elem.send_keys(Keys.PAGE_DOWN)
            time.sleep(0.2)
        driver.execute_script('window.scrollTo(0, 0);')
        # driver -> pagesource 받기
        site_bs = BeautifulSoup(driver.page_source, 'lxml')
        reviewSelector = site_bs.find_all('div', class_='reviewSelector')
        for c in reviewSelector:

            imglink = c.find('div',class_='avatar').find('a').find('img',class_='avatar')['src']
            if imglink == 'https://static.tacdn.com/img2/x.gif': # 엑박 이미지 리소스있으면
                imglink =  c.find_all('div', class_='avatar')[-1].find('a').find('img', class_='avatar')['src']
            # uid, src 조합으로 사람 상세페이지 url 생성가능
            uid = c.find('div',class_='memberOverlayLink')['id'].split('-')[0].split('_')[1].strip()
            src = c.find('div', class_='memberOverlayLink')['id'].split('-')[1].split('_')[1].strip()
            Title = c.find("span", class_='noQuotes').text #1
            Ratetag = c.find("div", class_='rating reviewItemInline')#2
            Rate = Ratetag.find('span')['class'][1].split("_")[1]
            try: # title이없이 바로써져있을떄
                Date = c.find("span", class_='ratingDate')['title']  # 3
            except:
                Date = c.find("span", class_='ratingDate').text.replace('Reviewed','').strip()
            try:# date string convert to date
                Date = str(time.strptime(Date, '%d %B %Y').tm_year) + '/' + str(
                    time.strptime(Date, '%d %B %Y').tm_mon) + '/' + str(
                    time.strptime(Date, '%d %B %Y').tm_mday)
                # 리뷰의 엔터 제거
                Reviews = c.find_all("div", class_='entry')[-1].text.strip().replace('\n','') #4
            except:
                print("String to Date 오류 ",Date)
            if c.find('div', class_='reviewItem inlineRoomTip'): #5
                Tip = c.find('div', class_='reviewItem inlineRoomTip').text.split('Travel Tip:')[-1].split('See more travel tips')[0].strip()
            else:
                Tip = ''
            if c.find('span',class_='recommend-titleInline'):  # 6
                TravelDate = c.find('span',class_='recommend-titleInline').text
            else:
                TravelDate = ''
            DeatailRating = [] #7
            DeatailRating.clear()
            try:
                uls = c.find('ul', class_='recommend').find('li').find_all('ul')
                for ul in uls:
                    lis = ul.find_all('li')
                    for li in lis:
                        DeatailRating.append(int(li.find('span')['class'][1].split("_")[-1]))
                        #DeatailRating.append(li.text.strip()+":"+li.find('span')['class'][1].split("_")[-1])
            except:
                pass

            Tag_all = c.find_all("span", class_='categoryLabel')  # 8
            Tag = []
            for t in Tag_all[:3]:
                Tag.append(t.text)
            #출발- 도착 나누기
            depart, end = Tag[2].split('-',maxsplit=1)
            Tag[2] = depart.strip()
            Tag.append(end.strip())
            try: # 9
                numhlp = c.find('span',class_='numHlpIn').text
            except:
                numhlp = 0

            if c.find("span", dir='auto'):
                try:
                    Location,nation = c.find("div", class_='location').text.split(',')#11
                except: # 나라,지역 이아니라 둘장 하나있을경우는 두개 동일시
                    Location = c.find("div", class_='location').text
                    nation = Location
            else:
                Location = ''
                nation = ''
            try:
                level = c.find('span',class_='contribution-count').text # 12
            except:#레벨없는경우
                level = 0
            reviewcnt = c.find('span',class_='badgeText').text.split('reviews')[0].strip()#13
            votecnt = c.find_all('span',class_='badgeText')[-1].text.split('helpful')[0].strip()#14
            if 'review' in votecnt: # 리뷰투표수가 없어서 다른게나올시에는 0으로
                votecnt = 0
            #15.16 유저상세페이지
            code = requests.get(
                'https://www.tripadvisor.co.uk/MemberOverlay?Mode=owa&uid={}&c=&src={}&fus=false&partner=false&LsoId='.format(uid,src))
            bs4 = BeautifulSoup(code.text, 'lxml')
            Name = bs4.find('h3',class_='username reviewsEnhancements').text.strip()
            lis = bs4.find_all('li', class_='countsReviewEnhancementsItem')
            citicnt = 0
            photocnt =0
            for li in lis: # 순서가 뒤죽박죽이라 keyword로 찾음
                if 'Cities' in li.get_text().strip():
                    citicnt =li.find('span', class_='badgeTextReviewEnhancements').text.split(' ')[0].strip()
                if 'Photos' in li.get_text().strip():
                    photocnt = li.find('span', class_='badgeTextReviewEnhancements').text.split(' ')[0].strip()
            #17
            divs = bs4.find_all('span', class_='rowCountReviewEnhancements rowCellReviewEnhancements')
            try:
                review_excellent_cnt = divs[0].get_text().strip()
                review_verygood_cnt = divs[1].get_text().strip()
                review_average_cnt = divs[2].get_text().strip()
                review_poor_cnt = divs[3].get_text().strip()
                review_terrible_cnt = divs[4].get_text().strip()
            except:
                review_excellent_cnt = 0
                review_verygood_cnt= 0
                review_average_cnt= 0
                review_poor_cnt = 0
                review_terrible_cnt = 0
            # print 테스트 쪽, 이부분 필요없을시 지우면 좀더 성능향상
            print("[Title]", Title)
            print("[Reviews]", Reviews)
            print("[Name]", Name)  ### 띄아쓰기 없애기
            print("[Location]", Location)  ### \n 지우는 방법
            print("[Date]", Date)  ### 날짜 형식만 남기기,,
            print("[Rate]", Rate)  ### 10의자리를 1의자리로 바꾸기
            print("[Tag]", Tag)  ### 세번째 태그 경로 " - " 이것도 나누기,
            print("[Tip]", Tip)  ### 세번째 태그 경로 " - " 이것도 나누기,
            print("[TravelDate]", TravelDate)  ### 세번째 태그 경로 " - " 이것도 나누기,
            print("[DeatailRating]", DeatailRating)  ### 세번째 태그 경로 " - " 이것도 나누기,
            print("[numhlp]", numhlp)  ### 세번째 태그 경로 " - " 이것도 나누기,
            print("[level]", level)  ### 세번째 태그 경로 " - " 이것도 나누기,
            print("[reviewcnt]", reviewcnt)  ### 세번째 태그 경로 " - " 이것도 나누기,
            print("[votecnt]", votecnt)  ### 세번째 태그 경로 " - " 이것도 나누기,
            print("[citicnt]", citicnt)  ### 세번째 태그 경로 " - " 이것도 나누기,
            print("[photocnt]", photocnt)  ### 세번째 태그 경로 " - " 이것도 나누기,
            print("[imglink]", imglink)  ### 세번째 태그 경로 " - " 이것도 나누기,
            print("[review_excellent_cnt]", review_excellent_cnt)  ### 세번째 태그 경로 " - " 이것도 나누기,
            print("[review_verygood_cnt]", review_verygood_cnt)  ### 세번째 태그 경로 " - " 이것도 나누기,
            print("[review_average_cnt]", review_average_cnt)  ### 세번째 태그 경로 " - " 이것도 나누기,
            print("[review_poor_cnt]", review_poor_cnt)  ### 세번째 태그 경로 " - " 이것도 나누기,
            print("[review_terrible_cnt]", review_terrible_cnt)  ### 세번째 태그 경로 " - " 이것도 나누기,

            # 엑셀저장
            ws.write(j, 1, Title)
            ws.write(j, 2, int(Rate)/10)
            ws.write(j, 3, Date)
            ws.write(j, 4, Reviews)
            ws.write(j, 5, Tip)
            ws.write(j, 6, TravelDate)
            try:
                ws.write(j, 7, DeatailRating[4]) # legroom
                ws.write(j, 8, DeatailRating[0]) #seatcompfort
                ws.write(j, 9, DeatailRating[1]) #customer service
                ws.write(j, 10, DeatailRating[6]) # value for moeny
                ws.write(j, 11, DeatailRating[2]) #clean
                ws.write(j, 12, DeatailRating[-1])#check in and boarding
                ws.write(j, 13, DeatailRating[3]) # food
                ws.write(j, 14, DeatailRating[5])# in -flight
            except: # 리뷰가없
                ws.write(j, 7, 0)  # legroom
                ws.write(j, 8, 0)  # seatcompfort
                ws.write(j, 9, 0)  # customer service
                ws.write(j, 10, 0)  # value for moeny
                ws.write(j, 11, 0)  # clean
                ws.write(j, 12, 0)  # check in and boarding
                ws.write(j, 13, 0)  # food
                ws.write(j, 14, 0)  # in -flight
            ws.write(j, 15, Tag[0])
            ws.write(j, 16, Tag[1])
            ws.write(j, 17, Tag[2])
            ws.write(j, 18, Tag[3])
            ws.write(j, 19, numhlp)
            ws.write(j, 20, Name)
            ws.write(j, 21, nation)
            ws.write(j, 22, Location)
            ws.write(j, 23, level)
            ws.write(j, 24, reviewcnt)
            ws.write(j, 25, votecnt)
            ws.write(j, 26, citicnt)
            ws.write(j, 27, photocnt)
            ws.write(j, 28, review_excellent_cnt)
            ws.write(j, 29, review_verygood_cnt)
            ws.write(j, 30, review_average_cnt)
            ws.write(j, 31, review_poor_cnt)
            ws.write(j, 32, review_terrible_cnt)
            ws.write(j, 33, imglink)
            print("_________________________________")
            j = j + 1
    wb.close()
