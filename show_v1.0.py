print ('Status: 200 OK')
print ('Content-type: text/html')
print ()

print ('<HTML><HEAD><TITLE>Python Sample CGI</TITLE></HEAD>')
print ('<BODY>')
print ("<style>.btn{cursor:pointer}\
        select{height:23px}\
        #id_iframe {font-size:10px;}\
       </style>")


import cgi
import cgitb
cgitb.enable()
#print("<form action=json-test.py method=post>快递公司&nbsp<input type=text name=company  />&nbsp")
print("<form action=json-test.py method=post target=nm_iframe>快递公司\
      <select name=company>\
      <option selected=selected>请选择快递公司</option>\
      <option value=yuantong title=圆通速递>圆通</option>\
      <option value=yunda title=韵达速递>韵达</option>\
      <option value=jd title=京东物流>京东</option>\
      <option value=shentong title=申通快递>申通</option>\
      <option value=shunfeng title=顺丰速运>顺丰</option>\
      </select>")
print("快递单号&nbsp<input type=text name=kdnumber >&nbsp")
print("<input type=submit value=查一下 class=btn></form>")
print("<iframe id=id_iframe name=nm_iframe frameborder=0 width=100% height=100% scrolling=yes></iframe>")



print ('<br>')
print ('</BODY>')
