#!/usr/bin/python
# -*- coding: UTF-8 -*-
print ("Status: 200 OK")
print ("Content-Type: text/html;charset=gb2312")
print ()

import urllib3
import json
import cgi
import cgitb
cgitb.enable()

form = cgi.FieldStorage()
#print(form)
if 'company' not in form or 'kdnumber' not in form :
    print("<H1>粗错啦！</H1>")
    print("请选择填写快递公司和快递单号")

else:
     print("<p>快递公司:", form["company"].value)
     print("<p>快递单号:", form["kdnumber"].value)
     com=form["company"].value
     kdnu=form["kdnumber"].value

     http = urllib3.PoolManager()
     urllib3.disable_warnings()
     url="http://www.kuaidi100.com/query?type="+com+"&postid="+kdnu
#html=http.request('get',url= "https://api.douban.com/v2/book/1220562").data
#html=http.request('get',url= "http://www.kuaidi100.com/query?type=yuantong&postid=1111111111111").data
#html=http.request('get',url= "http://www.kuaidi100.com/query?type=yuantong&postid=11111111111").data
     html=http.request('get',url= url).data
#print(url)
     hjson=json.loads(html)
#print(hjson)
#print(hjson['data'])
#for index,val in enumerate(hjson['data']):
#    print(index,val)

     if hjson['data']:
         for i in hjson['data']:
#    print(i)
#    print(type(i))
#    print(i["time"])
#    print(i["context"])
     
              print("<p>%s  %s<br>" %(i["time"],i["context"]))

     else:
         print('<p>',hjson['message'])
     
