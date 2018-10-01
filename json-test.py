#!/usr/bin/python
# -*- coding: UTF-8 -*-
print ("Status: 200 OK")
print ("Content-Type: text/html;charset=gb2312")  # HTML is following
print ()                                          # blank line, end of headers

import urllib3
import json

http = urllib3.PoolManager()
urllib3.disable_warnings()
#html=http.request('get',url= "https://api.douban.com/v2/book/1220562").data
html=http.request('get',url= "http://www.kuaidi100.com/query?type=yuantong&postid=11111111111").data
hjson=json.loads(html)
#print(hjson)
#print(hjson['data'])
#for index,val in enumerate(hjson['data']):
#    print(index,val)

for i in hjson['data']:
#    print(i)
#    print(type(i))
#    print(i["time"])
#    print(i["context"])
     
     print("%s  %s<br><br>" %(i["time"],i["context"]))
     
     
