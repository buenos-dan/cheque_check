#!/usr/bin/env python 
#-*-coding:utf-8-*-
import os
import xlrd
import xlwt
from xlutils.copy import copy
import time
import selenium
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from settings import *

__VERSION = "1.0.1"        #MODE="ROW" is a stable method

def getIndex(date):
    fdir = "log/"
    files = os.listdir(fdir)
    index = 0
    for f in files:
        if f.startswith(date):
            t_index = int(f[11:-4])
            if t_index > index:
                index = t_index
    return index+1

def openFile():
    date = time.strftime('%Y_%m_%d',time.localtime(time.time()))
    fname = date+"_"+str(getIndex(date))+".log"
    print "Record message in /log/" + fname
    return open("log/"+fname,"w+")

def logInfo(text):
    global f
    f.write(text+"\n")
    print text

def logWarn(text):
    pass

def logError(text):
    pass

f = openFile()

class PieceOfData():
    def __init__(self):
        self.row = None
        self.item = None
        self.year = None
        self.month = None
        self.day = None 
        self.date = None 
        self.usage = None
        self.taoBaoOrderNum = None
        self.money = None
        self.chequeNum = None
        self.buyer = None
        self.payer = None
        self.payMethod = None
        self.isCheck = None
        self.isReimburse = None
        self.note = None
        self.checkDate = None

        self.isSucess = False

class ExcelHandle():
    def __init__(self,filename,sheetname):
        self.filename = filename
        self.sheetname = sheetname
        self.file = xlrd.open_workbook(filename)
        self.file_copy = copy(self.file)
        self.table = self.file.sheet_by_name(self.sheetname)
        self.rows = self.table.nrows
        self.cols = self.table.ncols
        self.ctype = ["empty","string","number","data","boolean","error"]

    def getYear(self,row):
        timestamp = self.getCellValue(row,0)
        return timestamp.split('.')[0]

    def getMonth(self,row):
        timestamp = self.getCellValue(row,0)
        return timestamp.split('.')[1]

    def getDay(self,row):
        timestamp = self.getCellValue(row,0)
        return timestamp.split('.')[2]

    def getCellValue(self,row,col):
        return self.table.cell(row,col).value

    def writeCellValue(self,row,col,value):
        ws = self.file_copy.get_sheet(0)
        ws.write(row,col,value)
        self.file_copy.save(self.filename)
            
        

    def getCellType(self,row,col):
        type_num = self.table.cell(row,col).ctype
        return self.ctype[type_num]

    def printRowxValue(self,row):
        iass
        for i in range(self.cols):
            print self.table.cell(row,i).value
            print self.getCellType(row,i)

    def printColxValue(self,col):
        for i in range(self.rows):
            print self.table.cell(i,col).value

    def _donothing(self):
        #table = f.sheet_by_index(0)
        #table.cell(i,0).value
        pass

class WebHandle():
    def __init__(self):
        self.browser = None
        self.title = None

    def prehandle(self):
        self.openURL(LOGINURL)
        self.process1()
        self.process2()

    def openURL(self,url):
        if OS == "windows":
            if BROWSER == "Firefox":
                self.browser = webdriver.Firefox(executable_path="geckodriver.exe")
            elif BROWSER == "Chrome":
                self.browser = webdriver.Chrome(executable_path="chromedriver.exe")
        elif OS == "linux":
            if BROWSER == "Firefox":
                self.browser = webdriver.Firefox(executable_path="/usr/local/bin/geckodriver")
            elif BROWSER == "Chrome":
                self.browser = webdriver.Chrome(executable_path="/usr/local/bin/chromedriver")
        else:
            logInfo("[warn]Please choose correct browser!")
            return
        #self.browser.maximize_window()
        self.browser.get(url)

    def updateTitle(self):
        self.title = self.browser.title
        print self.title

    def process1(self):
        WebDriverWait(self.browser, 10).until(
            EC.presence_of_element_located((By.ID, "username"))).send_keys(USERNAME)
        #self.browser.find_element_by_id("username").send_keys(u'drizzle')
        self.browser.find_element_by_id("password").send_keys(PASSWORD)
        self.browser.find_element_by_name("submit").click()
        #self.browser.implicitly_wait(30)
        WebDriverWait(self.browser, 5).until(
            EC.presence_of_element_located((By.ID, "LinkButton_wsyy"))).click()
        logInfo("[message]Login success!")
        #print "login successfully!"
        #self.browser.find_element_by_id("LinkButton_wsyy").click()
        #print "change to baozhang_system successfully!"

    def process2(self):
        time.sleep(0.5)
        #print len(self.browser.window_handles)
        window = self.browser.window_handles[-1]
        self.browser.switch_to_window(window)
        self.browser.switch_to_frame("topFrame")
        #print "switch success"
        #WebDriverWait(self.browser, 10).until(
        #    EC.presence_of_element_located((By.ID, "GWK"))).click()
        WebDriverWait(self.browser, 10).until(
            EC.presence_of_element_located((By.ID, "GWK"))).click()
        #self.browser.find_element_by_id("GWK").click()
        self.browser.switch_to_default_content()
        self.browser.switch_to_frame("mainframe")

    def process3(self,flag):
        if flag == True:
            WebDriverWait(self.browser, 10).until(
                EC.presence_of_element_located((By.ID, "ctl00_ContentPlaceHolder1_IBTN_GWKHKBZ"))).click()
        else:
            self.browser.switch_to_frame("mainframe")
            WebDriverWait(self.browser, 10).until(
                EC.presence_of_element_located((By.ID, "ctl00_ContentPlaceHolder1_IBTN_GWKHKBZ"))).click()
        #self.browser.find_element_by_id("ctl00_ContentPlaceHolder1_IBTN_GWKHKBZ").click()
        #print "change to baozhang_system successfully!"
        
    def loophandle(self,date,money,flag):
        self.process3(flag)
        if not self.process4(date,money):
            return False
        isSuccess = self.process5()
        return isSuccess

    def process4(self,date,money):
        #WebDriverWait(self.browser, 10).until(
        #    EC.presence_of_element_located((By.ID, "ctl00_ContentPlaceHolder1_TB_JE"))).send_keys(money)
        WebDriverWait(self.browser, 10).until(
            EC.presence_of_element_located((By.ID, "ctl00_ContentPlaceHolder1_TB_YGBH"))).send_keys(EMPLOYEEID,Keys.ENTER)
        time.sleep(0.5)
        calendar = WebDriverWait(self.browser, 10).until(
            EC.presence_of_element_located((By.ID, "ctl00_ContentPlaceHolder1_TB_XFRQ")))
        #self.browser.execute_script("calendar.setValue('')")
        #js = 'document.getElementById("train_date").removeAttribute("readonly");'
        #self.browser.execute_script(js)
        calendar.readOnly=False
        #calendar.click()
        print date
        calendar.clear()
        calendar.send_keys(date)
        self.browser.execute_script("calendar.fadeOut()")
        #WebDriverWait(self.browser, 10).until(
        #    EC.presence_of_element_located((By.ID, "ctl00_ContentPlaceHolder1_TB_JE"))).send_keys(money)
        #WebDriverWait(self.browser, 10).until(
        #    EC.presence_of_element_located((By.ID, "ctl00_ContentPlaceHolder1_TB_YGBH"))).send_keys("09815")
        m = self.browser.find_element_by_id("ctl00_ContentPlaceHolder1_TB_JE")
        m.send_keys(money)
        #WebDriverWait(self.browser, 10).until(
        #    EC.presence_of_element_located((By.ID, "ctl00_ContentPlaceHolder1_TB_JE"))).send_keys(money)
        #self.browser.execute_script("calendar.setValue('2019-04-20')")
        #self.browser.find_element_by_xpath("//tr/td[@onclick='setValue']")
        #.click();
        #driver.switchTo().defaultContent();
        #self.browser.find_element_by_id("ctl00_ContentPlaceHolder1_TB_JE").send_keys(money)
        #self.browser.find_element_by_id("ctl00_ContentPlaceHolder1_TB_YGBH").send_keys(EMPLOYEEID)
        #calendar = self.browser.find_element_by_id("ctl00_ContentPlaceHolder1_TB_XFRQ") 
        #self.browser.refresh()
        #ctl00_ContentPlaceHolder1_GV_WEBXFJL
        self.browser.find_element_by_id("ctl00_ContentPlaceHolder1_BT_JS").click()
        #selenium.common.exceptions.StaleElementReferenceException
        for i in range(50):
            try:
                self.browser.find_element_by_id("ctl00_ContentPlaceHolder1_GV_WEBXFJL")
            except selenium.common.exceptions.UnexpectedAlertPresentException,e:
                  #selenium.common.exceptions.UnexpectedAlertPresentException
                self.browser.switch_to_alert().accept()
                logInfo("[error]wrong data!")
                self.browser.back()
                return False
            except selenium.common.exceptions.NoSuchElementException,e:
                time.sleep(1)
                print "times:%d"%i
                continue
            logInfo("[+++]Find  the data!")
            break
        else:
            self.browser.back()
            logInfo("[---]Can not find the data!")
            return False
        return True


    def process5(self):
        try:
            WebDriverWait(self.browser, 1).until(
                EC.presence_of_element_located((By.ID, "ctl00_ContentPlaceHolder1_GV_WEBXFJL_ctl02_ck_selectxfjl"))).click()
            #self.browser.find_element_by_id("ctl00_ContentPlaceHolder1_GV_WEBXFJL_ctl02_ck_selectxfjl").click()
        except selenium.common.exceptions.TimeoutException,e:
            #print("\033[31;0m this data has already checked! \033[0m")
            logInfo("\033[1;33;40m [warn:TimeOut]This data has already checked! \033[0m")
            self.browser.back()
            return False

        except selenium.common.exceptions.UnexpectedAlertPresentException,e:
              #selenium.common.exceptions.UnexpectedAlertPresentException
            self.browser.switch_to_alert().accept()
            logInfo("[error]wrong data!")
            self.browser.back()
            return False

        except selenium.common.exceptions.NoSuchElementException,e:
            logInfo("\033[1;33;40m [warn:NoSuchElement]This data has already checked! \033[0m")
            self.browser.back()
            return False

        try:
            self.browser.find_element_by_id("ctl00_ContentPlaceHolder1_BT_QD").click()
        except selenium.common.exceptions.StaleElementReferenceException,e:
            print self.browser.page_source
            return False
        self.browser.implicitly_wait(30)
        self.browser.find_element_by_xpath("//option[@value='6']").click()
        self.browser.find_element_by_id("ctl00_ContentPlaceHolder1_GV_GWKXF_WEB_ctl02_BT_BZWEB").click()
        WebDriverWait(self.browser, 10).until(EC.alert_is_present())
        self.browser.switch_to_alert().accept()
        WebDriverWait(self.browser, 10).until(
            EC.presence_of_element_located((By.ID, "ctl00_ContentPlaceHolder1_BT_BEFOREWEB"))).click()
        #self.browser.find_element_by_id("ctl00_ContentPlaceHolder1_BT_BEFOREWEB").click()
        #print self.browser.page_source
        time.sleep(1)
        return True


    def workFlow(self):
        self.openURL(LOGINURL)
        self.process1()
        self.process2()
        self.process3()
        self.process4()

    def donothing(self):
        #self.browser.get_screenshot_as_file("1.png")
        pass

class ChequeCheckScript:
    def __init__(self):
        if OS == "windows":
            self.encodeMode = "gbk"
        else: 
            self.encodeMode = "utf-8"
        self.excelhandle = ExcelHandle(FILENAME,SHEETNAME)
        self.webhandle = WebHandle()
        self.datas = self.readAllData()
        self.today = time.strftime('%Y-%m-%d',time.localtime(time.time()))
        self.mode = MODE

    def run(self):
        if self.mode == "AI":
            logInfo("[starting program]Handle cheques which has not checked")
            count = 0
            for data in self.datas:
                if (not data.isCheck) and (data.row >= STARTROW and data.row <= ENDROW):
                    logInfo("row %d prepare to check"%data.row)
                    count += 1
            if count == 0:
                logInfo("no data need check!")
                exit()
            logInfo("[message]Total num: %d"%count)
            try:
                self.webhandle.prehandle()
            except selenium.common.exceptions.TimeoutException,e:
                logInfo("\033[1;31;40m[message]Time Out! \033[0m")
                logInfo("please checkout your internet connection.")
                return
            flag = True
            for data in self.datas:
                if (not data.isCheck) and (data.row >= STARTROW and data.row <= ENDROW):
                    logInfo("[checking row:" + str(data.row) + "] " + str(data.date) +\
                            "   " + data.item.encode(self.encodeMode) + "    RMB:" + str(data.money) +\
                            "   " + data.buyer.encode(self.encodeMode))
                    flag = self.webhandle.loophandle(data.date,str(data.money),flag)
                    if flag:
                        self.excelhandle.writeCellValue(data.row-1,9,"是".decode(self.encodeMode))
                        self.excelhandle.writeCellValue(data.row-1,12,data.checkDate)
                        data.isSuccess = True
                        logInfo("[+++]Success!")
            logInfo("[+++]Program End!")
            logInfo("please press ENTER to quit...")

        elif self.mode == "ROW":
            if STARTROW < 0 or ENDROW < 0 or ENDROW - STARTROW < 0:
                logInfo("please write correct row number in settings.py!")
                return
            logInfo("[starting program] Handle cheques from %d to %d"%(STARTROW,ENDROW))
            logInfo("[message]Total num: %d"%(ENDROW-STARTROW+1))
            try:
                self.webhandle.prehandle()
            except selenium.common.exceptions.TimeoutException,e:
                logInfo("\033[1;31;40m[message]Time Out! \033[0m")
                logInfo("please checkout your internet connection.")
                return
            flag = True
            for data in self.datas:
                if data.row >= STARTROW and data.row <= ENDROW:
                    logInfo("[checking row:" + str(data.row) + "] " + str(data.date) +\
                            "   " + data.item.encode(self.encodeMode) + "    RMB:" + str(data.money) +\
                            "   " + data.buyer.encode(self.encodeMode))
                    flag = self.webhandle.loophandle(data.date,str(data.money),flag)
                    if flag:
                        self.excelhandle.writeCellValue(data.row-1,9,"是".decode(self.encodeMode))
                        data.isSuccess = True
                        logInfo("[+++]Success!")
            logInfo("[+++]Program End!")
            logInfo("please press ENTER to quit...")


        elif self.mode == "HACK":
            if not self.isLegal(HACK_DATE):
                logInfo("please write correct HACK_DATE in settings.py!")
                return
            logInfo("[starting program] Hack cheques in %s"%HACK_DATE)
            logInfo("[message]Total num: %s"%(MAXMONEY))
            try:
                self.webhandle.prehandle()
            except selenium.common.exceptions.TimeoutException,e:
                logInfo("\033[1;31;40m[message]Time Out! \033[0m")
                logInfo("please checkout your internet connection.")
                return
            flag = True
            for i in range(int(MAXMONEY)):
                logInfo("[checking row:" + str(HACK_DATE) + "] " + "RMB:%.2f"%float(i))
                flag = self.webhandle.loophandle(HACK_DATE,str(MAXMONEY),flag)
                if flag:
                    #self.excelhandle.writeCellValue(data.row-1,9,"是".decode(self.encodeMode))
                    data.isSuccess = True
                    logInfo("[+++]odd!")
                    break
            logInfo("[+++]Program End!")
            logInfo("please press ENTER to quit...")

    def isLegal(self,date):
        if len(str(date))<10:
            logInfo("wrong length!")
            return False
        if date[4] != "-" and date[7] != "-":
            logInfo("wrong delimitation!")
            return False
        if int(date[:4])< 2000 or int(date[:4])>2050:
            logInfo("wrong year!")
            return False
        if int(date[5:7])<1 or int(date[5:7])>12:
            logInfo("wrong month!")
            return False
        if int(date[8:10])<1 or int (date[8:10])>31:
            logInfo("wrong day!")
            return False
        else:
            return True
        

    def readAllData(self):
        container = []
        for i in range(1,self.excelhandle.rows):
            if self.excelhandle.getCellType(i,0) != "empty":
                container.append(self.readPieceOfData(i))
        return container

    def readPieceOfData(self,row):
        pod = PieceOfData()
        pod.row = row+1
        pod.item = self.excelhandle.getCellValue(row,1)
        pod.year = self.excelhandle.getYear(row).zfill(4)
        pod.month = self.excelhandle.getMonth(row).zfill(2)
        pod.day = self.excelhandle.getDay(row).zfill(2)
        pod.date = pod.year + "-" + pod.month + "-" + pod.day
        pod.usage = self.excelhandle.getCellValue(row,2)
        pod.taoBaoOrderNum = self.excelhandle.getCellValue(row,3)
        pod.money = self.excelhandle.getCellValue(row,4)
        pod.chequeNum = self.excelhandle.getCellValue(row,5)
        pod.buyer = self.excelhandle.getCellValue(row,6)
        pod.payer = self.excelhandle.getCellValue(row,7)
        pod.payMethod = self.excelhandle.getCellValue(row,8)
        if self.excelhandle.getCellValue(row,9).encode(self.encodeMode) == "是":
            pod.isCheck = True
        else:
            pod.isCheck = False
        pod.isReimburse = self.excelhandle.getCellValue(row,10)
        pod.note = self.excelhandle.getCellValue(row,11)
        if self.excelhandle.getCellType(row,12) == "empty":
            pod.checkDate = time.strftime('%Y-%m-%d',time.localtime(time.time()))
        else:
            pod.checkDate = self.excelhandle.getCellType(row,12)

        return pod

    def printData(self,data):
        print type(data.item)
        print data.row
        print data.year
        print data.month
        print data.day
        print data.money
        print data.date
        print len(data.date)
    
    


if __name__=="__main__":
    print "***********Cheque_Check_Script*************"
    print "VERSION:"+__VERSION
    print "MODE: "+MODE
    ccs = ChequeCheckScript()
    #ccs.printData(ccs.datas[-1])
    ccs.run()
    f.close()










