<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>

<%
If IsEmpty(Session("username")) Then
	response.write "<script>window.location.href='index.html'</script>"
else
set conn1=server.createobject("adodb.connection")
conn1str="provider=Microsoft.Jet.OLEDB.4.0;Data Source=c:\inetpub\kuaidi\kdinfo.mdb"
conn1.Open conn1str
Set rs1=server.CreateObject("adodb.recordset")


'连接excel数据源

Dim xlsconn1,strs1ource,xlbook,xlsheet,i
Dim myconn1_Xsl,xlsrs1,sql,sql2,xlsrs2,aa
Set xlsconn1 = server.CreateObject("adodb.connection")
Set xlsrs1 = Server.CreateObject("Adodb.RecordSet")
'Set xlsrs2 = Server.CreateObject("Adodb.RecordSet")

i=0
myconn1_Xsl="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=c:\inetpub\kuaidi\file\book2.xls;Extended Properties=Excel 8.0"
xlsconn1.open myconn1_Xsl

sql = "Select * from [Sheet1$] where 运单号 is not null"
'sql2="Select str(电话) from [Sheet1$] where 收件人 like '杨'"
xlsrs1.open sql,xlsconn1,1,1
'xlsrs2.open sql2,xlsconn1,1,1
a=xlsrs1("运单号")

'Response.write xlsrs1.Fields.Count
'Response.write xlsrs1.Fields.Item(i).Value
If xlsrs1.eof Then

 elseif  not conn1.execute("select * from kdinfo where kd_number like '"&a&"'").eof Then
 danhao=conn1.execute("select * from kdinfo where kd_number like '"&a&"'").Fields("kd_number")
 'Response.write danhao
  	Response.write danhao&"快递单号已经存在相同的记录<br>"

  elseif not xlsrs1.eof Then
 'For X = 0 To xlsrs2.Fields.Count - 1
''			Response.write( xlsrs2.Fields.Item(X).Value)
''			Next
 'Response.Write("<TABLE><TR>")
	do While Not xlsrs1.EOF
		i=i+1
		
	'response.write (f)
if not isnull(xlsrs1("运单号")) Then
	a=xlsrs1("运单号")'excel表中的字段名称
end if
if not isnull(xlsrs1("快递公司"))  Then
	b=xlsrs1("快递公司")
end if
if not isnull(xlsrs1(9))  Then
	c=xlsrs1(9)
end if 
if not isnull(xlsrs1("秋季课程简名"))  Then
	d=xlsrs1("秋季课程简名")
end if
if not isnull(xlsrs1("地址"))  Then
	e=xlsrs1("地址")
end if
if not isnull(xlsrs1(5))  Then
	f=xlsrs1(5)
	'response.write (f)
end if 
if not isnull(xlsrs1("收件人"))   Then
	g=xlsrs1("收件人")
end if 
if not isnull(xlsrs1("学员编码"))	Then
	h=xlsrs1("学员编码")
end if 
'if not isnull(xlsrs1("班级分校"))	Then
	k=xlsrs1("班级分校")
'end if
if not isnull(xlsrs1("邮寄日期"))	Then
	j=xlsrs1("邮寄日期")
end if 
 
 
'     For X = 0 To xlsrs1.Fields.Count - 1
'         Response.Write("<TD>" & xlsrs1.Fields.Item(X).Name & "</TD>")
''      Next
'      Response.Write("</TR>")
'      xlsrs1.MoveFirs1t
'
'      While Not xlsrs1.EOF
'         Response.Write("<TR>")
'         For X = 0 To xlsrs1.Fields.Count - 1
'            Response.write("<TD>" & xlsrs1.Fields.Item(X).Value)
'         Next
		
		Response.write i&":"&a&"<br>"
		For X = 0 To xlsrs1.Fields.Count - 1
			'Response.write xlsrs1.Fields.Item(X).Value&"<br>"
			'if (xlsrs1.Fields.Item(X).Value) Then exit do 
				
			Next
			
           sql2="insert into kdinfo(kd_number,company,sender,lesson_name,address,phone,receivername,stuNumber,school,kd_date) values('"&a&"','"&b&"','"&c&"','"&d&"','"&e&"','"&f&"','"&g&"','"&h&"','"&k&"','"&j&"')"
  			conn1.execute(sql2)
  			
  			
  		xlsrs1.MoveNext
        
'        Response.Write("</TR>")

     loop
      
'    Response.Write("</TABLE>")
 

   End If
xlsrs1.close
conn1.close

Response.CharSet = "GB2312"
Response.write "共导入<font color='red'>" & i & "</font>条记录.<br>" 

set xlsconn1=nothing
end if
%>
<!--<script>alert("单号导入成功！");</script>-->

<!--
http://support.microsoft.com/kb/295646/

https://support.microsoft.com/en-us/help/195951-->
