# A more generic pythonn script to extract data from zomato website
#import all necessary classes
from bs4 import BeautifulSoup # need to pip install bs4 for BeautifulSoup to work which is used to work on html sources
from selenium import webdriver #need to pip install selenium for automated website login otherwise you won't be able to access a website
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
#uptil here all classes required for selenium to work seemlessly
import xlsxwriter # need to pip install this for excel files operations through python
import win32com.client as win32 # an extra extension to edit, beautify, improve use of xlsx files using python


PATH= "C:\Program Files (x86)\chromedriver.exe"   # this is the location of the chromedriver extension where it is installed, so that we may be able to work with selenium

# keep this path same if you also download and keep your chromedriver extension in this location
# or change this path to the location where you keep your chromedriver extension


url="https://www.zomato.com/ncr/connaught-place-delhi-restaurants" # url of the root webpage from where you want to start your data scrapping

# The script works only for zomato website

# BEFORE RUNNING THE SCRIPT, MAKE SURE YOU ARE CONNECTED TO A STABLE INTERNET CONNECTION

# IF YOU HAVE RUN THE SCRIPT AND ARE VIEWING THE EXCEL FILE AND WANT TO RUN THE SCRIPT AGAIN, FIRST CLOSE THE EXCEL FILE OTHERWISE PYTHON WILL NOT GET PERMISSION TO WRITE THE FILE

# IF YOU TERMINATE THE SCRIPT WHILE IT IS RUNNING, YOU WILL NOT GET AN EXCEL FILE FILLED WITH THE DATA, SO BEWARE 

workbook = xlsxwriter.Workbook('Zomato_scrap.xlsx') # opens a workbook with the name 'Zomato_scrap' with extension .xlsx as it is an excel file
worksheet = workbook.add_worksheet() # adds a worksheet to the opened excel file

bold = workbook.add_format({'bold': True}) # defining bold so to write with bold font in excel file


names=[]        # stores the name of the restaurants which we scrap from the website
address=[]      # stores the address of the restaurants which we scrap from the website
phnum=[]        # stores the given phone numbers of the restaurants which we scrap from the website
flag=1          # defined a flag so we know that we are still navigating the website through it's pages and not try to reach the next page from the last page in the website


# every command which is contained in a try block ensures that if and when an error is encountered, it is handled in a manner such that the script does not crash

driver=webdriver.Chrome(PATH)  # this makes an instance from the chrome webdriver 

driver.get(url)                # this send the url link to the instance of the webdriver and opens a new chrome window from where we will scrap our data
clk=1                          # counter to count how many pages we have completed the scraping of
while flag==1:                 # iterate through the next page until we reach the last page
    try:
        if clk==11:            # if you want the data to the first 10 pages of the website, check page numbers with pages+1 i.e. 11th page
            print('Finish')    # this is an explicit way to stop the scraping before the last page is reached
            break
        root=WebDriverWait(driver,5).until(EC.presence_of_element_located((By.ID,"mainframe")))  # locates the webpage using the id 'mainframe'
	
	# every website has a unique id by which it can be found to exist
	
	# the above command copies the webpage in variable root which will contain all html data
	
	# we are waiting for 5 seconds until the webpage is identified by it's id, if it is not present the command will throw an error which will be shown by the except block
	
	
        pol=root.find_elements_by_class_name('search-result')    # all restaurants contain the class 'search-result' which contains every info about the restaurants
        if clk==1:
            lim=len(pol)      # counts how many restaurants are shown to us on a single webpage of the website
        for i in pol:
            try:
                name=i.find_element_by_class_name('result-title') # finding by class to find name of restaurnts
                names.append(name.text)       #appends to the name list
                add=i.find_element_by_class_name('search-result-address') # finding by class to find address of restaurants
                address.append(add.text)      #appends to the name list
                elem = driver.find_element_by_xpath("//*") #finds the path to a class using xpath - explicit address
                html = elem.get_attribute("outerHTML") #gets the html file
                soup = BeautifulSoup(html, 'html.parser')  #contains the parsed version of the html website
                for link in soup.find_all('a'):   #finds all <a> tags
                    num=link.get('data-phone-no-str') #finds phone number from this arttribute
                    if isinstance(num,str): #it is being checked that what we got from website was of data sype str which will mean you know.
                        phnum.append(num)
            except:
                print("No search result")
        
        print('Done page ',clk) # this is printed to the console saying that a particular page is scraped
        try:
            driver.find_element_by_xpath('//*[@title="Next Page"]').click() # finds the button for the 'Next Page' and clicks the button
            clk+=1
        except:
            print('Last page reached') # ll pages are parsed
            break
    except:
        print("No search result")     # absence of a search result part





row=0 # row on where to write in excel sheet
i=0   # counter keeping count of addresses and phone numbers
x=0   # used to print heading for a specific webpage before scraping it's data
for ind in names: # iterating through all the names which have been extracted from the website
    col=0   #  column on where to write in excel sheet
    if i%lim==0:
        stri='SEARCH RESULTS FOR PAGE '+str(x+1)  # segregates scrapped data according to the web pge 
        x+=1
        worksheet.write(row,col,stri,bold)   # writing data to the worksheet
        col=0
        row+=2
        worksheet.write(row,col,'NAME',bold)               ###             
        col+=1                                                    ##
        worksheet.write(row,col,'ADDRESS',bold)            ###           # writing column names for the data 
        col+=1                                                    ##
        worksheet.write(row,col,'PHONE NUMBER',bold)       ###      
        row+=2
        col=0
    worksheet.write(row,col,ind)               # writing restaurant name to worksheet
    col+=1
    worksheet.write(row,col,address[i])        # writing restaurant address to worksheet
    col+=1
    lis=phnum[i].split(',')                    # splitting the string of phone numbers using a delimeter (in this case a comma)
    trow=row
    jj=0
    for j in lis:                              # iterating through the list of phone numbers obtained for a particular retaurant
        worksheet.write(trow,col,lis[jj])      # writing phone number to worksheet
        trow+=1
        jj+=1
    row=trow+1
    i+=1



workbook.close()                        # workbook is saved and closed
excel = win32.gencache.EnsureDispatch('Excel.Application')    # given access to open and modify and save exel files
wb = excel.Workbooks.Open(r'C:\Users\Admin\AppData\Local\Programs\Python\Python38\Zomato_scrap.xlsx')  # providing the full address to the file where it is located
ws = wb.Worksheets("Sheet1")       # explicitly told to work on sheet 1 of workbook
ws.Columns.AutoFit()               # autofits all columns according to the size of data entered
wb.Save()                          # saves the modified excel file
excel.Application.Quit()           # excel application closed

driver.quit()                      # ends the automated test window from where we scrapped our data


# The data scraping process is a real time rendition of the current webpage of zomato
