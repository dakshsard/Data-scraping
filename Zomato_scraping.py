from bs4 import BeautifulSoup
import requests
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import xlsxwriter
import win32com.client as win32


PATH= "C:\Program Files (x86)\chromedriver.exe"


url="https://www.zomato.com/ncr/connaught-place-delhi-restaurants"


workbook = xlsxwriter.Workbook('Zomato_scrap.xlsx')
worksheet = workbook.add_worksheet() 

bold = workbook.add_format({'bold': True})


names=[]
address=[]
phnum=[]
flag=1

driver=webdriver.Chrome(PATH)

driver.get(url)
clk=1
while flag==1:
    try:
        if clk==11:
            print('Finish')
            break
        root=WebDriverWait(driver,5).until(EC.presence_of_element_located((By.ID,"mainframe")))
        pol=root.find_elements_by_class_name('search-result')
        if clk==1:
            lim=len(pol)
        for i in pol:
            try:
                name=i.find_element_by_class_name('result-title')
                names.append(name.text)
                add=i.find_element_by_class_name('search-result-address')
                address.append(add.text)
                elem = driver.find_element_by_xpath("//*")
                html = elem.get_attribute("outerHTML")
                soup = BeautifulSoup(html, 'html.parser')
                for link in soup.find_all('a'):
                    num=link.get('data-phone-no-str')
                    if isinstance(num,str):
                        phnum.append(num)
            except:
                print("No search result")
        
        print('Done page ',clk)
        try:
            driver.find_element_by_xpath('//*[@title="Next Page"]').click()
            clk+=1
        except:
            print('Last page reached')
            break
    except:
        print("No search result")





row=0
i=0
x=0
for ind in names:
    col=0
    if i%lim==0:
        stri='SEARCH RESULTS FOR PAGE '+str(x+1)
        x+=1
        worksheet.write(row,col,stri,bold)
        col=0
        row+=2
        worksheet.write(row,col,'NAME',bold)
        col+=1
        worksheet.write(row,col,'ADDRESS',bold)
        col+=1
        worksheet.write(row,col,'PHONE NUMBER',bold)
        row+=2
        col=0
    worksheet.write(row,col,ind)
    col+=1
    worksheet.write(row,col,address[i])
    col+=1
    lis=phnum[i].split(',')
    trow=row
    jj=0
    for j in lis:
        worksheet.write(trow,col,lis[jj])
        trow+=1
        jj+=1
    row=trow+1
    i+=1



workbook.close()
excel = win32.gencache.EnsureDispatch('Excel.Application')
wb = excel.Workbooks.Open(r'C:\Users\Admin\AppData\Local\Programs\Python\Python38\Zomato_scrap.xlsx')
ws = wb.Worksheets("Sheet1")
ws.Columns.AutoFit()
wb.Save()
excel.Application.Quit()



#time.sleep(5)

driver.quit()
