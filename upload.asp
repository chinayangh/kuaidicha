<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>

<%
If IsEmpty(Session("username")) Then
	response.write "<script>window.location.href='index.html'</script>"
else
set conn1=server.createobject("adodb.connection")
conn1str="provider=Microsoft.Jet.OLEDB.4.0;Data Source=c:\inetpub\kuaidi\kdinfo.mdb"
conn1.Open conn1str
Set rs1=server.CreateObject("adodb.recordset")


'����excel����Դ

Dim xlsconn1,strs1ource,xlbook,xlsheet,i
Dim myconn1_Xsl,xlsrs1,sql,sql2,xlsrs2,aa
Set xlsconn1 = server.CreateObject("adodb.connection")
Set xlsrs1 = Server.CreateObject("Adodb.RecordSet")
'Set xlsrs2 = Server.CreateObject("Adodb.RecordSet")

i=0
myconn1_Xsl="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=c:\inetpub\kuaidi\file\book2.xls;Extended Properties=Excel 8.0"
xlsconn1.open myconn1_Xsl

sql = "Select * from [Sheet1$] where �˵��� is not null"
'sql2="Select str(�绰) from [Sheet1$] where �ռ��� like '����õ'"
xlsrs1.open sql,xlsconn1,1,1
'xlsrs2.open sql2,xlsconn1,1,1
a=xlsrs1("�˵���")

'Response.write xlsrs1.Fields.Count
'Response.write xlsrs1.Fields.Item(i).Value
If xlsrs1.eof Then

 elseif  not conn1.execute("select * from kdinfo where kd_number like '"&a&"'").eof Then
 danhao=conn1.execute("select * from kdinfo where kd_number like '"&a&"'").Fields("kd_number")
 'Response.write danhao
  	Response.write danhao&"��ݵ����Ѿ�������ͬ�ļ�¼<br>"

  elseif not xlsrs1.eof Then
 'For X = 0 To xlsrs2.Fields.Count - 1
''			Response.write( xlsrs2.Fields.Item(X).Value)
''			Next
 'Response.Write("<TABLE><TR>")
	do While Not xlsrs1.EOF
		i=i+1
		
	'response.write (f)
if not isnull(xlsrs1("�˵���")) Then
	a=xlsrs1("�˵���")'excel���е��ֶ�����
end if
if not isnull(xlsrs1("��ݹ�˾"))  Then
	b=xlsrs1("��ݹ�˾")
end if
if not isnull(xlsrs1(9))  Then
	c=xlsrs1(9)
end if 
if not isnull(xlsrs1("�＾�γ̼���"))  Then
	d=xlsrs1("�＾�γ̼���")
end if
if not isnull(xlsrs1("��ַ"))  Then
	e=xlsrs1("��ַ")
end if
if not isnull(xlsrs1(5))  Then
	f=xlsrs1(5)
	'response.write (f)
end if 
if not isnull(xlsrs1("�ռ���"))   Then
	g=xlsrs1("�ռ���")
end if 
if not isnull(xlsrs1("ѧԱ����"))	Then
	h=xlsrs1("ѧԱ����")
end if 
'if not isnull(xlsrs1("�༶��У"))	Then
	k=xlsrs1("�༶��У")
'end if
if not isnull(xlsrs1("�ʼ�����"))	Then
	j=xlsrs1("�ʼ�����")
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
Response.write "������<font color='red'>" & i & "</font>����¼.<br>" 

set xlsconn1=nothing
end if
%>
<!--<script>alert("���ŵ���ɹ���");</script>-->

<!--
http://support.microsoft.com/kb/295646/

https://support.microsoft.com/en-us/help/195951-->