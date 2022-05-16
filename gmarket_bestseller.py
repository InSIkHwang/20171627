#!/usr/bin/env python
# coding: utf-8

# In[ ]:


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


print("=" *80)
print("지마켓의 분야별 Best Seller 상품 정보 추출하기")
print("=" *80)

query_txt='지마켓'
query_url='http://corners.gmarket.co.kr/Bestsellers'

cnt = int(input('1.건수를 입력하세요(1-200 건 사이 입력): '))

f_dir = input("2.폴더명을 입력하세요(예:c:\\temp\\):")
print("\n")

      
now = time.localtime()
s = '%04d-%02d-%02d-%02d-%02d-%02d' % (now.tm_year, now.tm_mon, now.tm_mday, now.tm_hour, now.tm_min, now.tm_sec)

os.makedirs(f_dir+s+'-'+query_txt)
os.chdir(f_dir+s+'-'+query_txt)

ff_dir=f_dir+s+'-'+query_txt
ff_name=f_dir+s+'-'+query_txt+'\\'+s+'-'+query_txt+'.txt'
fc_name=f_dir+s+'-'+query_txt+'\\'+s+'-'+query_txt+'.csv'
fx_name=f_dir+s+'-'+query_txt+'\\'+s+'-'+query_txt+'.xls'

s_time = time.time( )

path = "E:/temp/chromedriver_240/chromedriver.exe"
driver = webdriver.Chrome(path)
    
driver.get(query_url)
time.sleep(5)


def scroll_down(driver):
      
      driver.execute_script("window.scrollTo(0,document.body.scrollHeight);")
      time.sleep(1)

scroll_down(driver)

bmp_map = dict.fromkeys(range(0x10000, sys.maxunicode + 1), 0xfffd)

img_src2=[]
file_no = 0


html = driver.page_source
soup = BeautifulSoup(html, 'html.parser')

reple_result = soup.select('.best-list')[1]
slist = reple_result.find_all('li')

      
ranking2=[]
title3=[]
price2=[]
score2=[]
s_price2=[]
store2=[]
       
count = 0
    
img_dir = ff_dir+"\\images"
os.makedirs(img_dir)
os.chdir(img_dir)

for li in slist:
        
        try :
          photo = li.find('div','thumb').find('img')['src']
        except AttributeError :
          continue
        file_no += 1
        
        urllib.request.urlretrieve(photo,str(file_no)+'.jpg')
        time.sleep(1)
        
        if cnt == file_no :
          break
            
        f = open(ff_name, 'a',encoding='UTF-8')
        f.write("-----------------------------------------------------"+"\n")

        print("-" *70)
        try :
         ranking = li.find('p').get_text()
        except AttributeError :
         ranking = ''
         print(ranking.replace("#",""))
        else :
         print("1.판매순위:",ranking)

        f.write('1.판매순위:'+ ranking + "\n")

        try :
         title1 = li.find('a',class_='itemname').get_text().replace("\n","")
        except AttributeError :
         title1 = ''
         print(title1.replace("\n",""))
         f.write('2.제품소개:'+ title1 + "\n")
        else :
         title2=title1.translate(bmp_map).replace("\n","") 
         print("2.제품소개:", title2.replace("\n",""))

         count += 1
             
         f.write('2.제품소개:'+ title2 + "\n")
            
        try :
          price = li.find('div','o-price').get_text().replace("\n","")
        except AttributeError :
          price = ''
               
        print("3.원래가격:", price.replace("\n",""))
        f.write('3.원래가격:'+ price + "\n")
                  
        try :
          s_price = li.find('strong').get_text().replace("\n","")
        except AttributeError :
          s_price = ''
               
        print("4.할인가격:", s_price.replace("\n",""))
        f.write('4.할인가격:'+ s_price + "\n")

        try :
          score = li.find('em').get_text()
        except (IndexError , AttributeError) :
          score='0%'
          print('5.할인율:',score.replace("\n",""))
          f.write('5.할인율:'+ score + "\n")
        else :
          print('5.할인율:',score.replace("\n",""))
          f.write('5.할인율:'+ score + "\n")
               
        

        print("-" *70)
                          
        f.close( )
              
        time.sleep(0.3)
            
        ranking2.append(ranking)
        title3.append(title2.replace("\n",""))
        price2.append(price.replace("\n",""))
        s_price2.append(s_price.replace("\n",""))          
        score2.append(score.replace("\n",""))

        if count == cnt+1 :
            break
                          

amazon_best_seller = pd.DataFrame()
amazon_best_seller['판매순위']=ranking2
amazon_best_seller['제품소개']=pd.Series(title3)
amazon_best_seller['원래가격']=pd.Series(price2)
amazon_best_seller['할인가격']=pd.Series(s_price2)
amazon_best_seller['할인율']=pd.Series(score2)


amazon_best_seller.to_csv(fc_name,encoding="utf-8-sig",index=True)

amazon_best_seller.to_excel(fx_name ,index=True)

e_time = time.time( )
t_time = e_time - s_time

orig_stdout = sys.stdout
f = open(ff_name, 'a',encoding='UTF-8')
sys.stdout = f



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


print("\n")
print("=" *50)
print("총 소요시간은 %s 초 이며," %t_time)
print("총 저장 건수는 %s 건 입니다 " %count)
print("=" *50)

sys.stdout = orig_stdout
f.close( )

print("\n") 
print("=" *80)
print("1.총 %s 건 중에서 실제 검색 건수: %s 건" %(cnt,count))
print("2.총 소요시간은: %s 초" %round(t_time,1))
print("3.파일 저장 완료: txt 파일명 : %s " %ff_name)
print("4.파일 저장 완료: csv 파일명 : %s " %fc_name)
print("5.파일 저장 완료: xls 파일명 : %s " %fx_name)
print("=" *80)    


# In[ ]:




