from bs4 import BeautifulSoup
from selenium import webdriver
import time
import sys
import re
import math
import numpy  
import pandas as pd 
import os
import random

query_txt = input("1.크롤링 할 영화의 제목을 입력하세요: ")
query_url = 'https://movie.naver.com'
cnt = int(input("2.크롤링 할 리뷰건수는 몇건입니까?: "))
page_cnt = math.ceil(cnt / 10)
f_dir = input("3.파일을 저장할 폴더명만 쓰세요(예:c:\\temp\\):")

now = time.localtime()
s = '%04d-%02d-%02d-%02d-%02d-%02d' % (now.tm_year, now.tm_mon, now.tm_mday, now.tm_hour, now.tm_min, now.tm_sec)
os.makedirs(f_dir+s+'-'+query_txt)
os.chdir(f_dir+s+'-'+query_txt)

ff_name=f_dir+s+'-'+query_txt+'\\'+s+'-'+query_txt+'.txt'
fc_name=f_dir+s+'-'+query_txt+'\\'+s+'-'+query_txt+'.csv'
fx_name=f_dir+s+'-'+query_txt+'\\'+s+'-'+query_txt+'.xls'

s_time = time.time()

path = "c:/temp/chromedriver.exe"
driver = webdriver.Chrome(path)

driver.get(query_url)
time.sleep(2)

element = driver.find_element_by_id("ipt_tx_srch")
element.send_keys(query_txt)
driver.find_element_by_xpath("""//*[@id="jSearchArea"]/div/button""").click()
driver.find_element_by_xpath("""//*[@id="old_content"]/ul[2]/li/dl/dt/a""").click()
driver.find_element_by_xpath("""//*[@id="movieEndTabMenu"]/li[5]/a""").click()

driver.switch_to.frame('pointAfterListIframe')

html = driver.page_source
soup = BeautifulSoup(html, 'html.parser')

result= soup.find('div', class_='score_total').find('strong','total').find_all('em')
result2 = result[0].get_text()

print("=" *80)
result3 = result2.replace(",","")
result4 = re.search("\d+",result3)
search_cnt = int(result4.group())

if cnt > search_cnt :
    cnt = search_cnt

print("전체 검색 결과 건수 :",search_cnt,"건")
print("실제 최종 출력 건수",cnt)

print("\n")
page_cnt = math.ceil(cnt / 10)
print("크롤링 할 총 페이지 번호: ",page_cnt)
print("=" *80)

score2=[]
review2=[]
writer2=[]
wdate2=[]
gogam2=[]
g_gogam2=[]
b_gogam2=[]
dwlist2=[]

count = 0 
click_count = 1 

while ( click_count <= page_cnt )  :
            
            html = driver.page_source
            soup = BeautifulSoup(html, 'html.parser')
            
            score_result = soup.find('div', class_='score_result').find('ul')
            slist = score_result.find_all('li')

            for li in slist:

                count += 1
                
                f = open(ff_name, 'a',encoding='UTF-8')
            
                print("\n")
                print("총 %s 건 중 %s 번째 리뷰 데이터를 수집합니다==============" %(cnt,count))
                score = li.find('div', class_='star_score').find('em').get_text()
                print("1.별점:","*"*int(score),": ", score)
                score2.append(score)
                f.write("\n")
                f.write("총 %s 건 중 %s 번째 리뷰 데이터를 수집합니다==============" %(cnt,count) + "\n")
                f.write("1.별점:"+score + "\n")
            
                review = li.find('div', class_='score_reple').find('p').get_text().strip()
                print("2.리뷰내용:",review)
                f.write("2.리뷰내용:" + review + "\n")
                review2.append(review)
         
                dwlist = li.find('div',class_='score_reple').find_all('em')
                writer = dwlist[0].find('span').get_text()
                print("3.작성자:",writer)
                f.write("3.작성자:" + writer + "\n")
                writer2.append(writer)
                
                wdate = dwlist[1].text
                print('4.작성일자:',wdate)
                f.write("4.작성일자:" + wdate + "\n")
                wdate2.append(wdate)
            
                gogam = li.find('div', class_='btn_area').find_all('strong')
                g_gogam = gogam[0].text
                print('5.공감:',g_gogam)
                f.write("5.공감:" + g_gogam + "\n")
                g_gogam2.append(g_gogam)
                
                b_gogam = gogam[1].text
                print('6.비공감:', b_gogam)
                f.write("6.비공감:" + b_gogam + "\n")
                b_gogam2.append(b_gogam)
                print("\n")

                if count == cnt :
                    break
            
            time.sleep(random.randrange(1,2))  

            click_count += 1
            
            if click_count > page_cnt :
                break
            else :
                driver.find_element_by_link_text("%s" %click_count).click()

            time.sleep(random.randrange(1,2))

naver_movie = pd.DataFrame()
naver_movie['별점(평점)']=score2
naver_movie['리뷰내용']=review2
naver_movie['작성자']=writer2
naver_movie['작성일자']=wdate2
naver_movie['공감횟수']=g_gogam2
naver_movie['비공감횟수']=b_gogam2

naver_movie.to_csv(fc_name,encoding="utf-8-sig",index=True)
naver_movie.to_excel(fx_name ,index=True)

e_time = time.time()
t_time = e_time - s_time

driver.close()
