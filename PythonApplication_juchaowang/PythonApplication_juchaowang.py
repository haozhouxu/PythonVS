# -*- coding: utf-8 -*-
import urllib
import urllib2
from bs4 import BeautifulSoup
from collections import OrderedDict
import json
from tempfile import TemporaryFile
from xlwt import Workbook
import datetime
import time
import re

#with open('E:\\mywork\\file2.txt','r') as f:
#	soup = BeautifulSoup(f.read())

#main_url = 'http://www.interotc.com.cn/news/detail_pjax.do'
#main_url = 'http://www.interotc.com.cn/news/detail_pjax.do'
main_url = 'http://www.interotc.com.cn/news/index.do'
main_url2 = 'http://www.cninfo.com.cn/cninfo-new/disclosure/szse'
main_url3 = 'http://www.cninfo.com.cn/cninfo-new/announcement/show'
main_url4 = 'http://www.cninfo.com.cn/cninfo-new/disclosure/szse_latest'
main_url5 = 'http://www.cninfo.com.cn/cninfo-new/announcement/query'

sub_url = 'http://www.interotc.com.cn'

postDict = { 
	'stock' : '',
	'searchkey' : '',
	'plate' : '',
	'category' : '',
	'trade' : '',
	'column' : 'szse',
	'columnTitle' : '历史公告查询',
	'pageNum' : '1',
	'pageSize' : '30',
	'tabName' : 'fulltext',
	'sortName' : '',
	'sortType' : '',
	'limit' : '',
	'showTitle' : '',
	'seDate' : '2016-01-25'
}

# print postDict

# del postDict["seDate"]

# print postDict

# postDict["seDate"] = '请选择日期'

# print postDict


# <form id="AnnoucementsQueryForm" method="post" name="AnnoucementsQueryForm" action="/cninfo-new/announcement/show" enctype="multipart/form-data" target="_blank">
# <input type="hidden" id="stock_hidden_input" name="stock" />
# <input type="hidden" id="searchkey_hidden_input" name="searchkey" />
# <input type="hidden" id="plate_hidden_input" name="plate" />
# <input type="hidden" id="category_hidden_input" name="category" />
# <input type="hidden" id="trade_hidden_input" name="trade" />
# <input type="hidden" id="column_hidden_input" name="column" value='szse'/>
# <input type="hidden" id="columnTitle_hidden_input" name="columnTitle" value='深市公告'/>
# <input type="hidden" id="pageNum_hidden_input" name="pageNum" value="1"/>
# <input type="hidden" id="pageSize_hidden_input" name="pageSize" value="30"/>
# <input type="hidden" id="tabName_hidden_input" name="tabName" value="latest"/>
# <input type="hidden" id="sortName_hidden_input" name="sortName" value=""/>
# <input type="hidden" id="sortType_hidden_input" name="sortType" value=""/>
# <input type="hidden" id="limit_hidden_input" name="limit" value=""/>
# <input type="hidden" id="showTitle_hidden_input" name="showTitle" value=""/>

# stock:002117,9900002186;
# searchkey:
# plate:
# category:
# trade:
# column:szse
# columnTitle:历史公告查询
# pageNum:1
# pageSize:30
# tabName:fulltext
# sortName:
# sortType:
# limit:
# showTitle:
# seDate:2006-02-01 ~ 2016-02-03




postData = urllib.urlencode(postDict); 

req = urllib2.Request(main_url5, postData); 

# in most case, for do POST request, the content-type, is application/x-www-form-urlencoded 

req.add_header('Content-Type', "application/x-www-form-urlencoded"); 

resp = urllib2.urlopen(req)

# with open('test.txt', 'wb') as f:
# 	f.write(resp.read())

# print resp.read()

sjson = resp.read()

data = json.loads(sjson,object_pairs_hook=dict)
# print data["classifiedAnnouncements"][0][0]["secName"]

# count = 0
# dacount = 0

# print data

# print count
# print dacount

#新建EXCEL表格
book = Workbook()
sheet1 = book.add_sheet('Sheet1')
datetime = datetime.datetime.now().strftime('%Y-%m-%d')

#输入标题
sheet1.row(0).write(0,u'id') #产品名称
sheet1.row(0).write(1,u'secCode') #产品代码
sheet1.row(0).write(2,u'secName') #公告时间
sheet1.row(0).write(3,u'orgId') #产品名称
sheet1.row(0).write(4,u'announcementId') #产品代码
sheet1.row(0).write(5,u'announcementTitle') #公告时间
sheet1.row(0).write(6,u'announcementTime') #产品名称
sheet1.row(0).write(7,u'adjunctUrl') #产品代码
sheet1.row(0).write(8,u'adjunctSize') #公告时间
sheet1.row(0).write(9,u'adjunctType') #产品名称
sheet1.row(0).write(10,u'storageTime') #产品代码
sheet1.row(0).write(11,u'columnId') #公告时间
sheet1.row(0).write(12,u'pageColumn') #产品名称
sheet1.row(0).write(13,u'announcementType') #产品代码
sheet1.row(0).write(14,u'associateAnnouncement') #公告时间
sheet1.row(0).write(15,u'important') #产品名称
sheet1.row(0).write(16,u'batchNum') #产品代码
sheet1.row(0).write(17,u'announcementContent') #公告时间
sheet1.row(0).write(18,u'announcementTypeName') #公告时间

 # "id": null,
 #        "secCode": "000058",
 #        "secName": "深 赛 格",
 #        "orgId": "gssz0000058",
 #        "announcementId": "1201970437",
 #        "announcementTitle": "发行股份及支付现金购买资产并募集配套资金暨关联交易预案",
 #        "announcementTime": 1454515200000,
 #        "adjunctUrl": "finalpage/2016-02-04/1201970437.PDF",
 #        "adjunctSize": 3687,
 #        "adjunctType": "PDF",
 #        "storageTime": 1454499210000,
        # "columnId": "250101",
        # "pageColumn": "cninfo_announcement_disclosure_sz_fulltext",
        # "announcementType": "010799",
        # "associateAnnouncement": null,
        # "important": true,
        # "batchNum": 1454561640546,
        # "announcementContent": null,
        # "announcementTypeName": "增发"

#e_开头的为excel表格变量,从第二行（1）开始写入数据
e_row = 1
num = 1

while data["announcements"]:
	for y in data["announcements"]:
		print y["announcementTitle"]
		sheet1.row(e_row).write(0,y["id"]) #产品名称
		sheet1.row(e_row).write(1,y["secCode"])
		sheet1.row(e_row).write(2,y["secName"])
		sheet1.row(e_row).write(3,y["orgId"])
		sheet1.row(e_row).write(4,y["announcementId"])
		sheet1.row(e_row).write(5,y["announcementTitle"])
		sheet1.row(e_row).write(6,time.strftime('%Y%m%d %H:%M:%S', time.localtime(y["announcementTime"]/1000)))
		sheet1.row(e_row).write(7,'http://www.cninfo.com.cn/'+y["adjunctUrl"])
		sheet1.row(e_row).write(8,y["adjunctSize"])
		sheet1.row(e_row).write(9,y["adjunctType"])
		sheet1.row(e_row).write(10,time.strftime('%Y%m%d %H:%M:%S', time.localtime(y["storageTime"]/1000)))
		sheet1.row(e_row).write(11,y["columnId"])
		sheet1.row(e_row).write(12,y["pageColumn"])
		sheet1.row(e_row).write(13,y["announcementType"])
		sheet1.row(e_row).write(14,y["associateAnnouncement"])
		sheet1.row(e_row).write(15,y["important"])
		sheet1.row(e_row).write(16,y["batchNum"])
		sheet1.row(e_row).write(17,y["announcementContent"])
		sheet1.row(e_row).write(18,y["announcementTypeName"])
		e_row = e_row + 1

	num = num +1
	del postDict["pageNum"]
	postDict["pageNum"] = num
	postData2 = urllib.urlencode(postDict); 
	req2 = urllib2.Request(main_url5, postData2); 
	req2.add_header('Content-Type', "application/x-www-form-urlencoded"); 
	resp = urllib2.urlopen(req2)
	sjson2 = resp.read()
	data = json.loads(sjson2,object_pairs_hook=dict)

# while data["classifiedAnnouncements"]:
# 	for x in data["classifiedAnnouncements"]:
# 		if x:
# 			for y in x:
# 				print y["announcementTitle"]
# 				sheet1.row(e_row).write(0,y["id"]) #产品名称
# 				sheet1.row(e_row).write(1,y["secCode"])
# 				sheet1.row(e_row).write(2,y["secName"])
# 				sheet1.row(e_row).write(3,y["orgId"])
# 				sheet1.row(e_row).write(4,y["announcementId"])
# 				sheet1.row(e_row).write(5,y["announcementTitle"])
# 				sheet1.row(e_row).write(6,time.strftime('%Y%m%d %H:%M:%S', time.localtime(y["announcementTime"]/1000)))
# 				sheet1.row(e_row).write(7,'http://www.cninfo.com.cn/'+y["adjunctUrl"])
# 				sheet1.row(e_row).write(8,y["adjunctSize"])
# 				sheet1.row(e_row).write(9,y["adjunctType"])
# 				sheet1.row(e_row).write(10,time.strftime('%Y%m%d %H:%M:%S', time.localtime(y["storageTime"]/1000)))
# 				sheet1.row(e_row).write(11,y["columnId"])
# 				sheet1.row(e_row).write(12,y["pageColumn"])
# 				sheet1.row(e_row).write(13,y["announcementType"])
# 				sheet1.row(e_row).write(14,y["associateAnnouncement"])
# 				sheet1.row(e_row).write(15,y["important"])
# 				sheet1.row(e_row).write(16,y["batchNum"])
# 				sheet1.row(e_row).write(17,y["announcementContent"])
# 				sheet1.row(e_row).write(18,y["announcementTypeName"])
# 				e_row = e_row + 1

# 	num = num +1
# 	del postDict["pageNum"]
# 	postDict["pageNum"] = num
# 	postData2 = urllib.urlencode(postDict); 
# 	req2 = urllib2.Request(main_url5, postData2); 
# 	req2.add_header('Content-Type', "application/x-www-form-urlencoded"); 
# 	resp = urllib2.urlopen(req2)
# 	sjson2 = resp.read()
# 	data = json.loads(sjson2,object_pairs_hook=dict)

#保存数据
file_name = u'公告时间'+datetime+'.xls'
book.save(file_name)
book.save(TemporaryFile())

# print data

# soup = BeautifulSoup(resp,"html.parser")

# with open('test.txt', 'w') as f:
#     f.write(resp.read())

# # 定义一个变量，从3开始获取数据
# i = 1

# #新建EXCEL表格
# book = Workbook()
# sheet1 = book.add_sheet('Sheet1')
# datetime = datetime.datetime.now().strftime('%Y-%m-%d')

# #输入标题
# sheet1.row(0).write(0,u'产品名称') #产品名称
# sheet1.row(0).write(1,u'产品代码') #产品代码
# sheet1.row(0).write(2,u'公告时间') #公告时间

# #e_开头的为excel表格变量,从第二行（1）开始写入数据
# e_row = 1


# for tr in soup.find_all('tr'):
# 	if i > 2:

# 		product_name = tr.find_all('td')[0].a.get('title')
# 		product_time = tr.find_all('td')[1].contents[0]

# 		if True:

# 			#建立连接，获取数据
# 			sub_resp = urllib2.urlopen(sub_url+tr.find_all('td')[0].a.get('href'))
# 			sub_str = sub_resp.read()

# 			res = re.findall(r'产品代码([^，]+)',sub_str)

# 			for product_code in res:
# 				sheet1.row(e_row).write(0,product_name) #产品名称
# 				sheet1.row(e_row).write(1,product_code.decode('utf-8').replace(u'【','').replace(u'】','').replace(u'为','')) #产品代码
# 				sheet1.row(e_row).write(2,product_time) #公告时间
# 				e_row = e_row + 1
# 				#print product_name,product_code,product_time

# 	else:
# 		i = i + 1

	
# #保存数据
# #保存数据
# file_name = '..\\'+u'产品公告时间'+datetime+'.xls'
# book.save(file_name)
# book.save(TemporaryFile())