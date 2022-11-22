# -*- coding: utf-8 -*-
"""
Created on Thu Jul 14 15:50:10 2022

@author: AI01478
"""

# -*- coding: utf-8 -*-
"""
Created on Wed Jul 13 14:00:09 2022

@author: AI01478
"""

# -*- coding: utf-8 -*-
"""
Created on Tue Jul 12 14:01:29 2022

@author: AI01478
"""

# -*- coding: utf-8 -*-
"""
Created on Fri Jul  8 18:28:30 2022

@author: AI01478
"""

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time 
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
import datetime as DT
import pyautogui
import pandas as pd
import re
from pdf2image import convert_from_path
import glob,os
import glob
import os, subprocess
import matplotlib.pyplot as plt 
import os.path
from os import path

import pandas as pd
from datetime import datetime
from threading import Thread

import numpy as np

import timeit


import pyperclip


# datetime object containing current date and time
now = datetime.now()
start = timeit.default_timer()

# =============================================================================
# for error of easyocr
# =============================================================================
os.environ['KMP_DUPLICATE_LIB_OK']='True'
FinalTrueFlase = []
pathofpopper = r'C:\Users\AI01478\Downloads\poppler-0.68.0\bin'
# =============================================================================
#   read the WGS xlsx
#   get all the column of the xlsx into list
# =============================================================================
df = pd.read_excel("WGS DATA Full.xlsx", sheet_name= 'Sheet2')


dcn_noxl = df.KEY_CHK_DCN_NBR+','+'20'+df.KEY_CHK_DCN_NBR

hcidxl = df.ITS_PRFX_CD+df.HC_ID
hc_idxl = []
# =============================================================================
# adding prefic in the hcid
# =============================================================================
for i in hcidxl:
    hc_idxl.append(re.sub('\W+','', i))
    

sdatexl =pd.to_datetime(df.SRVC_FROM_DT).dt.strftime('%m/%d/%Y')
npixl = df.BILLG_NPI
cdatexl = pd.to_datetime(df.ORGNL_ENTRY_DT).dt.strftime('%m/%d/%Y')

patient_name= df.PAT_FRST_NM
patient_dob= pd.to_datetime(df.PAT_BRTH_DT).dt.strftime('%m/%d/%Y')
provider_name= df.PROV_NM



# =============================================================================
# medical Recored Found Function & 
# =============================================================================
#Need to add popper path

 


def getFormType(name):
    img = plt.imread(name)
    
    crop1 = img[200:470 , 0:850]
    crop2 = img[100:400 , 0:600]
    
    t1 = FunctionEasyOcrSTR(crop1)
    t2 = FunctionEasyOcrSTR(crop2)
    if(t1.find('Institutional') != -1):
        #FInalOutPut.append('Institutional')
        return 'Institutional'
    elif(t1.find('Professional') != -1):
        #FInalOutPut.append('Professional')
        return 'Professional'
    else:
        #FInalOutPut.append('Can\'t identify')
        return 'Can\'t identify'
    
    
def FunctionEasyOcrSTR(temp):
    import easyocr
    reader = easyocr.Reader(['en']) 
 
    result = reader.readtext(temp, detail = 0)   
    res = ""
    for x in result:
        res = res + x
    
    return res   
    
def istherepdf(name):
    return path.exists(name)

def medRecord(F_Name,BD,S_Date,P_Name,pdf_path):
    import matplotlib.pyplot as plt 
    from pdf2image import convert_from_path
    import glob,os
    import os, subprocess
    from pdf2image import convert_from_path

    import re
    # converting the pdf to image
    import PyPDF2

    fileisthere = istherepdf(pdf_path)
    if(fileisthere == False):
        FInalOutPut[pdf_path]=False
        return False
    
    file = open(pdf_path, 'rb')
    readpdf = PyPDF2.PdfFileReader(file)
    totalpages = readpdf.numPages
    if(totalpages==1 or totalpages==2):
        FInalOutPut[pdf_path]=False
        return False
    pdfs = pdf_path
    pages = convert_from_path(pdfs, 600,poppler_path= pathofpopper)
    stop_count = min(5,totalpages)
    i = 0
    for page in pages:
        image_name = "Page_" +pdfs[:35]+ str(i) + ".jpg"  
        page.save(image_name, "JPEG")
        if(i==stop_count):
            break
        i = i+1 
    #reading the images with the help of ocr
    import easyocr
    
    formtype = getFormType('Page_' +pdfs[:35]+ str(1) + '.jpg')
    if(formtype == 'Can\'t identify'):
        Pnameflag = False
        Pronameflag = False
        DOBflag = False
        Sdateflag  = False
#         result_fin = []
        
        reader = easyocr.Reader(['en'])
        #getting all the text from the pages
        matched_count = 0
        for i in range(0,stop_count):
            result = reader.readtext('Page_'+pdfs[:35]+str(i)+'.jpg',detail = 0)
        
            #all pages combine list
            #for x in result:
            #    result_fin+=x
            
            final_text = ""
            #final list to string con
            for i in result:
                final_text+=i

 

            #remove spaces
            res = re.sub(' +', '', final_text)

 

            #date format matching
            final_text = str(res)
            S_Date1 = S_Date[:6]+S_Date[8:]
            BD1 = BD[:6]+BD[8:]
            final_text=final_text.lower()
            F_Name=F_Name.lower().split(' ')[0]
            P_Name=P_Name.lower().split(' ')[0]            

 


            if(Pnameflag == False and final_text.find(F_Name)!=-1):
                matched_count+=1
                Pnameflag = True

 

            if(DOBflag == False and final_text.find(BD)!=-1 or final_text.find(BD1)!=-1):
                matched_count+=1
                DOBflag = True

 

            if(Pronameflag == False and final_text.find(P_Name)!=-1):
                matched_count+=1
                Pronameflag = True

 

            if(Sdateflag == False and final_text.find(S_Date)!=-1 or final_text.find(S_Date1)!=-1):
                matched_count+=1
                Sdateflag = True

 
            print(matched_count)
            if(matched_count>=3):
                FInalOutPut[pdf_path]=True
                
                if(Pnameflag):
                    matched_fields[0]=True
                if(DOBflag):
                    matched_fields[1]=True
                if(Pronameflag):
                    matched_fields[2]=True
                if(Sdateflag):
                    matched_fields[3]=True
                return True
                
        
    else:
        FInalOutPut[pdf_path]=False
        return False
    FInalOutPut[pdf_path]=False
    return False


  
def findFiles(dcn):
    folder = os.getcwd()
    files = [filename for filename in os.listdir(folder) if filename.startswith(dcn)]
    
    return files


# =============================================================================
# Get all the Service Date Function
# =============================================================================
def getAlldates(finaltable):
    allDate = []
    
    
    for i in range(1 , len(finaltable)):
        mon = ""
        day = ""
        year = ""
        cdate = finaltable[i][3]
        for i in range(len(cdate)):
            if(i >= 0 and i <= 1):
                mon = mon + cdate[i]
            elif(i >= 3 and i<= 4):
                day = day + cdate[i]
            elif(i >= 6 and i<=9 ):
                year = year + cdate[i]
        
        mon = int(mon)
        day = int(day)
        year = int(year)
        temp =  DT.datetime(year,mon,day)
        allDate.append(temp)
        
    return allDate
 

# =============================================================================
# Automation starts
# =============================================================================

userName = "AI01477"
password = "Intern@123"
#dummyDFRows = [3,2,4]
#ddummyDFRows = [2]
#range(len(df))
oneto10 = range(0,11)     #download time :  1825.1513297999918 Seconds
tento20 = range(11,21)    #download time :  2472.3805816999957 Seconds
twentyto30 = range(21,31) #download time :  1385.7880070999963 Seconds
thityto40 = range(31,41)  #download time :  2564.217776900012 Seconds
fortyto50 = range(41,51)  #download time :  1951.0671120000043 Seconds
fiftyto60 = range(51,61)  #download time :  3318.3408240999997 seconds
sixtyto70 = range(61,71)  #download time :  1727.0894068000052 seconds

    
start=0
end=10
dummyDFRows = [3,2,1]
FinalSearchingRows = dummyDFRows
for indexOfXL in FinalSearchingRows:
    
    FInalOutPut = []
    
    with webdriver.Ie() as driver:
    
     
    # =============================================================================
    #     open browser and chnage the window then log in
    # =============================================================================
        driver.get("https://fnetp8aeprod.us.ad.wellpoint.com/CONTENTONLY/Logon")
        
        wait = WebDriverWait(driver,1)
        original_window = driver.current_window_handle
        
        wait.until(EC.number_of_windows_to_be(2))
        for window_handle in driver.window_handles:
            if window_handle != original_window:
               second_handle = window_handle
               #print(second_handle)
               driver.switch_to.window(second_handle)
               break
        time.sleep(3)
        driver.find_element(By.NAME,"txtUserID").send_keys(userName)
        driver.find_element(By.NAME,"txtPWD").send_keys(password)
        driver.find_element(By.NAME,"btnLogin").send_keys(Keys.ENTER)
        
        
    # =============================================================================
    # selct claims exteneded search and role by down keys
    # =============================================================================
        
        time.sleep(5)
        actions = ActionChains(driver)   
        actions.send_keys(Keys.TAB*7)
        actions.perform()
        
        actions.send_keys(Keys.ARROW_DOWN)
        actions.perform()
        time.sleep(5)
        actions.send_keys(Keys.TAB*8)
        actions.perform()
        actions.send_keys(Keys.ARROW_DOWN)
        actions.perform()         
       
    # =============================================================================
    # enter DCN and 20+DCN and then search
    # =============================================================================
        time.sleep(5)
        driver.find_element(By.NAME,"txtsql00DCN or ClaimNum or PWK#").send_keys(dcn_noxl[indexOfXL])
        time.sleep(1)
        driver.find_element(By.NAME,"btnbrowse").send_keys(Keys.ENTER)
        
        time.sleep(5)
        all_text = driver.find_elements(By.XPATH,"//td[@class='sectionHeading3W']")
       
        tableRows=[]
        for t in all_text:
            tableRows.append((t.text).split(' ')[-1])
        #print(tableRows)
        
        AllEmptyTable = ['[0]','[0]','[0]']
        
        if(tableRows != AllEmptyTable):
            print('Found Using DCN')
            time.sleep(1)
            actions = ActionChains(driver)   
            actions.send_keys(Keys.TAB*37)
            actions.perform() 
            time.sleep(1)
            actions.send_keys(Keys.ENTER)  
            actions.perform()
            time.sleep(30)
            #time.sleep(30)
            
            pyautogui.keyDown('shift')
            time.sleep(2)
            pyautogui.press('P')
            time.sleep(1)
            pyautogui.keyUp('shift')
            time.sleep(2) 
            
            
    # taking only 5 pages
            # taking only 5 pages
            pyautogui.press('tab' , presses = 5)
            pyautogui.keyDown('ctrl')
            pyautogui.press('c')
            pyautogui.keyUp('ctrl')
            
            
            txtx = pyperclip.paste()
            print(txtx)
            txtx  = txtx.split('-')
            if(len(txtx) > 1):
                limitPage = min(5 , int(txtx[1]))
                writetxt = txtx[0] + '-' + str(limitPage)
                pyautogui.press('backspace')
                pyautogui.typewrite(writetxt)
            
            pyautogui.press('enter')
            time.sleep(1)
            
            #print('======================================================================')
            pdfname = dcn_noxl[indexOfXL] +str(int(datetime.today().timestamp()))
            pyautogui.typewrite(pdfname)
            #print('======================================================================')
            time.sleep(1)        
            
            
            pyautogui.press('enter')
            time.sleep(20)
            #time.sleep(59)
            
            # pyautogui.typewrite(pdfname)
            # time.sleep(1)
            # pyautogui.press('enter')
            # time.sleep(15)
            
            pdfname = pdfname + '.pdf'
            #FInalOutPut.append(medRecord(patient_name[indexOfXL] , patient_dob[indexOfXL] ,sdatexl[indexOfXL],provider_name[indexOfXL],pdfname))
            
            
        elif(True):
            print('Not Found Using DCN SO going with the service date and npi')
    # =============================================================================
    #         Searching Using HCPC and service Date gettinf things from table
    # =============================================================================
        driver.find_element(By.XPATH,'//a[contains(@href,"contentHome.jsp")]').send_keys(Keys.ENTER)
        time.sleep(5)
        driver.find_element(By.NAME,"txtsql00DCN or ClaimNum or PWK#").send_keys(Keys.BACKSPACE*25)
        time.sleep(1)
        driver.find_element(By.NAME,"txtsql01Member Cert Num or Customer ID").send_keys(hc_idxl[indexOfXL])    
        time.sleep(1)
        driver.find_element(By.NAME,"btnbrowse").send_keys(Keys.ENTER)   
        
        
        finaltable = []
        time.sleep(5)
        
# =============================================================================
#         #2->Doc ID 
#          #3-> DCN
#          #5-> DocType
#          #7->creation Date
#          #10-> service date
#          #22->NPI
#           get above data from the web 
# =============================================================================
        xstr = '//table[@id = "doclistHD"]/tbody/tr/td'
        neededColumn = [2,3,5,7,10,22]
        for i in neededColumn:
            temp = '['+str(i)+']'  
            finalpath = xstr + temp
            txt = driver.find_elements(By.XPATH , finalpath)
            finaltable.append(txt)
# =============================================================================
#     convert each entery from web element to text
# =============================================================================
        for i in range(len(finaltable)):
            for j in range(len(finaltable[i])):
                finaltable[i][j] = finaltable[i][j].text
      
# =============================================================================
#         transposing the whole table
# =============================================================================
        finaltable = [[row[i] for row in finaltable] for i in range(len(finaltable[0]))]
        sd = []
        npilist = []
# =============================================================================
#         getting first servise date and then npi search
# =============================================================================
        print("........................................................")
        for i in range(len(finaltable)):
            if(finaltable[i][4].find(sdatexl[indexOfXL]) != -1):
                print(finaltable[i])
                sd.append(i)
         
        print("........................................................")
        for i in sd:
            if(finaltable[i][5].find(npixl[indexOfXL]) != -1):
                print(finaltable[i])
                npilist.append(i)

        if(len(npilist) == 0):
            npilist = sd
            
# =============================================================================
#                 creation date logic
# =============================================================================
        if(len(npilist) == 0):
            cdate = cdatexl[indexOfXL]
     
            mon = ""
            day = ""
            year = ""
            
            allDate = getAlldates(finaltable)
            
            for i in range(len(cdate)):
                if(i >= 0 and i <= 1):
                    mon = mon + cdate[i]
                elif(i >= 3 and i<= 4):
                    day = day + cdate[i]
                elif(i >= 6 and i<=9 ):
                    year = year + cdate[i]
            
            mon = int(mon)
            day = int(day)
            year = int(year)
            today =  DT.datetime(year,mon,day)
            allselectedDates = []
            
            
            for i in range(-8 , 9):
                prev = today + DT.timedelta(days=i)
                for j in range(len(allDate)):
                    if(prev == allDate[j]):
                        allselectedDates.append(j)
                        print(allDate[j])
                        break
            
            nplist = allselectedDates
        
        
        # return false cause no data found
        if(len(npilist) == 0):
            print("No data found")
        else:
        
            time.sleep(1)
            actions = ActionChains(driver)   
            actions.send_keys(Keys.TAB*37)
            actions.perform() 
            time.sleep(1)
            last_index=1
            c=0
    
    # =============================================================================
    #     going till the first document using tab and enter the first document
    # =============================================================================
            
            for indexofnpi in npilist:
                actions.send_keys(Keys.TAB*(3*(indexofnpi-last_index)))
                time.sleep(1)
                actions.send_keys(Keys.ENTER)  
                actions.perform()
                time.sleep(30)
                time.sleep(30)
                
                
                pyautogui.keyDown('shift')
                time.sleep(2)
                pyautogui.press('P')
                time.sleep(1)
                pyautogui.keyUp('shift')
                time.sleep(1)
                
                # taking only 5 pages
                pyautogui.press('tab' , presses = 5)
                pyautogui.keyDown('ctrl')
                pyautogui.press('c')
                pyautogui.keyUp('ctrl')
                
                txtx = pyperclip.paste()
                print(txtx)
                txtx  = txtx.split('-')
                if(len(txtx) > 1):
                    
                    limitPage = min(5 , int(txtx[1]))
                    writetxt = txtx[0] + '-' + str(limitPage)
                    pyautogui.press('backspace')
                    pyautogui.typewrite(writetxt)
                
                
                pyautogui.press('enter')
                
                #print('======================================================================')
                pdfname = dcn_noxl[indexOfXL] +str(int(datetime.today().timestamp()))
                pyautogui.typewrite(pdfname)
                #print('======================================================================')
                time.sleep(5)        
                
                
                pyautogui.press('enter')
                time.sleep(2)
                
                pyautogui.typewrite(pdfname)
                time.sleep(1)
   
                pyautogui.press('enter')
                time.sleep(20)
                #time.sleep(59)
                
                pdfname = pdfname + '.pdf'
                #FInalOutPut.append(medRecord(patient_name[indexOfXL] , patient_dob[indexOfXL] ,sdatexl[indexOfXL],provider_name[indexOfXL],pdfname))
                
                
        # =============================================================================
        # enter shift + p to print documrnt then enter , then again enter nd name PDF and enter again
        # =============================================================================
            
                last_index=j
                c+=1
                
                time.sleep(5)
     
 
    
stop = timeit.default_timer()
neededTime = stop - start
print(neededTime)

dict_file={}
FourVariable = {}
for i in FinalSearchingRows:
        dict_file[df.KEY_CHK_DCN_NBR[i]]=False
        FourVariable[df.KEY_CHK_DCN_NBR[i]] = [False]*4
#print(dict_file)

for indexOfXL in range(start,end):
    file=findFiles(df.KEY_CHK_DCN_NBR[indexOfXL])
    FInalOutPut = {}
    matched_fields = [False]*4
    #print(len(file))
    
    threads = [Thread(target=medRecord, args=(patient_name[indexOfXL] , patient_dob[indexOfXL] ,sdatexl[indexOfXL],provider_name[indexOfXL],file[j])) for j in range(len(file))]
    for thread in threads:      
        thread.start()
        
    #print(len(threads))  
    
    [t.join() for t in threads]
     
    
    for i in FInalOutPut.keys():
        for j in dict_file.keys():
            if i.startswith(j):
                dict_file[j]=FInalOutPut[i] or dict_file[j]
                
                if(dict_file[j] == True):
                    FourVariable[j] = matched_fields
                    break
                     
    #print(FInalOutPut)
    
  
print(dict_file)

def transpose(l1):
    l2 = []
    # iterate over list l1 to the length of an item
    for i in range(len(l1[0])):
        # print(i)
        row =[]
        for item in l1:
            # appending to new list with values and index positions
            # i contains index position and item contains values
            row.append(item[i])
        l2.append(row)
    return l2

temp = []

for x in FourVariable.values():
    temp.append(x)
TransposeFourVariable = transpose(temp)

RecordFound = []
DCNNO = []

for x in dict_file.keys():
    DCNNO.append(x)
    
for x in dict_file.values():
    RecordFound.append(x)
    
    
lst = [DCNNO ,RecordFound, TransposeFourVariable[0] , TransposeFourVariable[1] , TransposeFourVariable[2] , TransposeFourVariable[3] ]
import pandas as pd
data = {'DCN Number': lst[0],
        'Record Found':  lst[1],
        'Patiant Name':  lst[2],
        'DOB':  lst[3],
        'Provider Name': lst[4],
        'Service Date':  lst[5],
        }

df = pd.DataFrame(data)

df.to_excel('First & Final 10 Integration 1 to   3.xlsx')

stop = timeit.default_timer()
neededTime = stop - start
print(neededTime)

