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
	'columnTitle' : '��ʷ�����ѯ',
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

# postDict["seDate"] = '��ѡ������'

# print postDict


# <form id="AnnoucementsQueryForm" method="post" name="AnnoucementsQueryForm" action="/cninfo-new/announcement/show" enctype="multipart/form-data" target="_blank">
# <input type="hidden" id="stock_hidden_input" name="stock" />
# <input type="hidden" id="searchkey_hidden_input" name="searchkey" />
# <input type="hidden" id="plate_hidden_input" name="plate" />
# <input type="hidden" id="category_hidden_input" name="category" />
# <input type="hidden" id="trade_hidden_input" name="trade" />
# <input type="hidden" id="column_hidden_input" name="column" value='szse'/>
# <input type="hidden" id="columnTitle_hidden_input" name="columnTitle" value='���й���'/>
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
# columnTitle:��ʷ�����ѯ
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

#�½�EXCEL���
book = Workbook()
sheet1 = book.add_sheet('Sheet1')
datetime = datetime.datetime.now().strftime('%Y-%m-%d')

#�������
sheet1.row(0).write(0,u'id') #��Ʒ����
sheet1.row(0).write(1,u'secCode') #��Ʒ����
sheet1.row(0).write(2,u'secName') #����ʱ��
sheet1.row(0).write(3,u'orgId') #��Ʒ����
sheet1.row(0).write(4,u'announcementId') #��Ʒ����
sheet1.row(0).write(5,u'announcementTitle') #����ʱ��
sheet1.row(0).write(6,u'announcementTime') #��Ʒ����
sheet1.row(0).write(7,u'adjunctUrl') #��Ʒ����
sheet1.row(0).write(8,u'adjunctSize') #����ʱ��
sheet1.row(0).write(9,u'adjunctType') #��Ʒ����
sheet1.row(0).write(10,u'storageTime') #��Ʒ����
sheet1.row(0).write(11,u'columnId') #����ʱ��
sheet1.row(0).write(12,u'pageColumn') #��Ʒ����
sheet1.row(0).write(13,u'announcementType') #��Ʒ����
sheet1.row(0).write(14,u'associateAnnouncement') #����ʱ��
sheet1.row(0).write(15,u'important') #��Ʒ����
sheet1.row(0).write(16,u'batchNum') #��Ʒ����
sheet1.row(0).write(17,u'announcementContent') #����ʱ��
sheet1.row(0).write(18,u'announcementTypeName') #����ʱ��

 # "id": null,
 #        "secCode": "000058",
 #        "secName": "�� �� ��",
 #        "orgId": "gssz0000058",
 #        "announcementId": "1201970437",
 #        "announcementTitle": "���йɷݼ�֧���ֽ����ʲ���ļ�������ʽ��߹�������Ԥ��",
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
        # "announcementTypeName": "����"

#e_��ͷ��Ϊexcel������,�ӵڶ��У�1����ʼд������
e_row = 1
num = 1

while data["announcements"]:
	for y in data["announcements"]:
		print y["announcementTitle"]
		sheet1.row(e_row).write(0,y["id"]) #��Ʒ����
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
# 				sheet1.row(e_row).write(0,y["id"]) #��Ʒ����
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

#��������
file_name = u'����ʱ��'+datetime+'.xls'
book.save(file_name)
book.save(TemporaryFile())

# print data

# soup = BeautifulSoup(resp,"html.parser")

# with open('test.txt', 'w') as f:
#     f.write(resp.read())

# # ����һ����������3��ʼ��ȡ����
# i = 1

# #�½�EXCEL���
# book = Workbook()
# sheet1 = book.add_sheet('Sheet1')
# datetime = datetime.datetime.now().strftime('%Y-%m-%d')

# #�������
# sheet1.row(0).write(0,u'��Ʒ����') #��Ʒ����
# sheet1.row(0).write(1,u'��Ʒ����') #��Ʒ����
# sheet1.row(0).write(2,u'����ʱ��') #����ʱ��

# #e_��ͷ��Ϊexcel������,�ӵڶ��У�1����ʼд������
# e_row = 1


# for tr in soup.find_all('tr'):
# 	if i > 2:

# 		product_name = tr.find_all('td')[0].a.get('title')
# 		product_time = tr.find_all('td')[1].contents[0]

# 		if True:

# 			#�������ӣ���ȡ����
# 			sub_resp = urllib2.urlopen(sub_url+tr.find_all('td')[0].a.get('href'))
# 			sub_str = sub_resp.read()

# 			res = re.findall(r'��Ʒ����([^��]+)',sub_str)

# 			for product_code in res:
# 				sheet1.row(e_row).write(0,product_name) #��Ʒ����
# 				sheet1.row(e_row).write(1,product_code.decode('utf-8').replace(u'��','').replace(u'��','').replace(u'Ϊ','')) #��Ʒ����
# 				sheet1.row(e_row).write(2,product_time) #����ʱ��
# 				e_row = e_row + 1
# 				#print product_name,product_code,product_time

# 	else:
# 		i = i + 1

	
# #��������
# #��������
# file_name = '..\\'+u'��Ʒ����ʱ��'+datetime+'.xls'
# book.save(file_name)
# book.save(TemporaryFile())