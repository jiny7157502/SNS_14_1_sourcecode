#Step 1. 필요한 모듈과 라이브러리를 로딩합니다.

from bs4 import BeautifulSoup
from selenium import webdriver

import time
import sys
import re
import math
import numpy 
import pandas as pd
import xlwt 
import random
import os

import urllib.request
import urllib

# Step 2. 크롤링할 건 수와 정보를 입력 받습니다.
print("=" *80)
print("연습문제 9-4. 지마켓 Best Seller 상품 정보 추출하기 ")
print("=" *80)

query_txt = 'G마켓'
query_url='http://corners.gmarket.co.kr/Bestsellers'

cnt = int(input('  1.크롤링 할 건수는 몇건입니까?: '))
real_cnt = math.ceil(cnt / 20)

f_dir=input('  2.파일이 저장될 경로만 쓰세요;(예:c:\\temp\\): ')
print("\n")

if cnt > 30 :
      print("  요청 건수가 많아서 시간이 제법 소요되오니 잠시만 기다려 주세요~~")
else :
      print("  요청하신 데이터를 수집하고 있으니 잠시만 기다려 주세요~~")
    
# Step 3. 저장될 파일위치와 이름을 지정 한 후 크롬 드라이버를 실행하여 페이지를 엽니다
now = time.localtime()
s = '%04d-%02d-%02d-%02d-%02d-%02d' % (now.tm_year, now.tm_mon, now.tm_mday, now.tm_hour, now.tm_min, now.tm_sec)

img_dir = f_dir+s+'-'+query_txt+"\\images"

os.makedirs(img_dir)

os.chdir(f_dir+s+'-'+query_txt)

ff_name=f_dir+s+'-'+query_txt+'\\'+s+'-'+query_txt+'.txt'
fc_name=f_dir+s+'-'+query_txt+'\\'+s+'-'+query_txt+'.csv'
fx_name=f_dir+s+'-'+query_txt+'\\'+s+'-'+query_txt+'.xls'

s_time = time.time( )

path = "c:/temp/chromedriver.exe"
driver = webdriver.Chrome(path)

driver.get(query_url)
time.sleep(1)

#Step 4. 화면을 스크롤해서 아래로 이동한 후 요청된 데이터를 수집합니다.

def scroll_down(driver):
      driver.execute_script("window.scrollTo(0,document.body.scrollHeight);")
      time.sleep(1)

i = 1
while (i <= 3):
         scroll_down(driver) 
         i += 1
            
ranking2=[]
title2=[]
o_price2=[]
s_price2=[]
discount2=[]

html = driver.page_source
soup = BeautifulSoup(html, 'html.parser')
        
count = 1

sale_result = soup.select('div.best-list')
slist = sale_result[1].select('ul > li')

img_file_no = 0

for li in slist:
                
            os.chdir(img_dir)

            try :
                  photo = li.find('div','thumb').find('img')['src']
            except AttributeError :
                continue
            
            img_file_no += 1

            urllib.request.urlretrieve(photo,str(img_file_no)+'.jpg')
            time.sleep(2)

            if img_file_no > cnt :
                break

for li in slist:
            
            f = open(ff_name, 'a',encoding='UTF-8')
            f.write("-----------------------------------------------------"+"\n")

            print("\n")
            
            # 판매순위
            print("-" *70)
            sid = '#no' + str(count)
            try :
             ranking = li.select(sid)[0].get_text()
            except AttributeError :
             ranking = ''
             print('1.판매순위:',ranking.replace("\n",""))
            else :
             print("1.판매순위:",ranking)

             f.write('1.판매순위:'+ ranking + "\n")
             
            #제품 설명 
            try :
             title = li.select('a.itemname')[0].get_text()
            except AttributeError :
             title = ''
             print(title)
             f.write('2.제품소개:'+ title + "\n")
            else :
             print("2.제품소개:", title.replace("\n",""))
             f.write('2.제품소개:'+ title + "\n")
            
             # 원래 가격
             o_price = li.find('div', class_='item_price').find('div','o-price').get_text()
             print("3.원래가격:", o_price.replace("\n",""))
             f.write('3.원래가격:'+ o_price + "\n")

             # 판매 가격
             s_price = li.find('div', class_='item_price').find('div','s-price').find('strong').get_text()
             print("4.판매가격:", s_price.replace("\n",""))
             f.write('4.판매가격:'+ s_price + "\n")

            #할인율 구하기
            try :
               discount = li.find('div', class_='item_price').find('div','s-price').find('em').get_text()
            except AttributeError:
               discount = '0%'
            print("5.할인율:", discount.replace("\n",""))
            f.write('5.할인율:'+ discount + "\n")


            ranking2.append(ranking)
            title2.append(title.replace("\n",""))
            o_price2.append(o_price.replace("\n",""))
            s_price2.append(s_price.replace("\n",""))
            discount2.append(discount.replace("\n",""))

            if count == cnt :
                break

            count += 1

            time.sleep(0.5)

#Step 5. 검색 결과를 다양한 형태로 저장하기
              
g_best_seller = pd.DataFrame()

g_best_seller['판매순위']=ranking2
g_best_seller['제품소개']=pd.Series(title2)
g_best_seller['원래가격']=pd.Series(o_price2)
g_best_seller['판매가격']=pd.Series(s_price2)
g_best_seller['할인율']=pd.Series(discount2)

g_best_seller.to_csv(fc_name,encoding="utf-8-sig",index=True)

g_best_seller.to_excel(fx_name ,index=True)

# Step 6. 요약 정보를 출력하기

e_time = time.time( )
t_time = e_time - s_time

orig_stdout = sys.stdout
f = open(ff_name, 'a',encoding='UTF-8')
sys.stdout = f

print("\n")
print("=" *50)
print("총 소요시간은 %s 초 이며," %t_time)
print("총 저장 건수는 %s 건 입니다 " %count)
print("=" *50)

sys.stdout = orig_stdout
f.close( )

print("\n") 
print("=" *80)
print("1.요청된 총 %s 건의 리뷰 중에서 실제 크롤링 된 건수는 %s 건입니다" %(cnt,count))
print("2.총 소요시간은 %s 초 입니다 " %round(t_time,1))
print("3.파일 저장 완료: txt 파일명 : %s " %ff_name)
print("4.파일 저장 완료: csv 파일명 : %s " %fc_name)
print("5.파일 저장 완료: xls 파일명 : %s " %fx_name)
print("=" *80)
     
import win32com.client as win32   
import win32api  
                
excel = win32.gencache.EnsureDispatch('Excel.Application')
wb = excel.Workbooks.Open(fx_name)
sheet = wb.ActiveSheet
sheet.Columns(3).ColumnWidth = 30
row_cnt = cnt+1
sheet.Rows("2:%s" %row_cnt).RowHeight = 120  

ws = wb.Sheets("Sheet1")
col_name2=[]
file_name2=[]

for a in range(2,cnt+2) :
      col_name='C'+str(a)
      col_name2.append(col_name)

for b in range(1,cnt+1) :
      file_name=img_dir+'\\'+str(b)+'.jpg'
      file_name2.append(file_name)
      
for i in range(0,cnt) :
      rng = ws.Range(col_name2[i])
      image = ws.Shapes.AddPicture(file_name2[i], False, True, rng.Left, rng.Top, 130, 100)
      excel.Visible=True
      excel.ActiveWorkbook.Save()

driver.close()
