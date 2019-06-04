#!/usr/bin/env python 
#*****************excel_data****************
FILENAME = "data.xlsx"
SHEETNAME = "Sheet1"

#*****************web_data*****************
CWCURL = "https://cwc.buaa.edu.cn/"
LOGINURL = "https://sso.buaa.edu.cn/login?service=https://icw.buaa.edu.cn/checkbuaa.aspx"
BROWSER = "Chrome"
OS = "linux"

#****************cheque_data***************
EMPLOYEEID = "09815"
HACK_DATE = "2019-04-24"
MAXMONEY = "50.00"

#***************config********************
MODE = "ROW"               #AI:automaticly check cheque which has not checked
#MODE = "AI"                #ROW:input start_raw number and end_raw number
#MODE = "HACK"              #TIME: input star_time and end_time

STARTROW = 378
ENDROW = 380
STARTTIME = ""
ENDTIME = ""

USERNAME = "drizzle"
PASSWORD = "dmd3582810"


#employeeID = ctl00_ContentPlaceHolder1_TB_YGBH
#dateID = ctl00_ContentPlaceHolder1_TB_XFRQ
#moneyID = ctl00_ContentPlaceHolder1_TB_JE
#searchID = ctl00_ContentPlaceHolder1_BT_JS

#selectBoxID = ctl00_ContentPlaceHolder1_GV_WEBXFJL_ctl02_ck_selectxfjl
#confirmID = ctl00_ContentPlaceHolder1_BT_QD

#usageID = ctl00_ContentPlaceHolder1_GV_GWKXF_WEB_ctl02_DROP_XFYT
#select material fee
#click box
#authorizedID = ctl00_ContentPlaceHolder1_GV_GWKXF_WEB_ctl02_BT_BZWEB
