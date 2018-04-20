<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
	<title>快递查询</title>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
</head>
<body>
    <form method="post" onsubmit="return form_check();"  runat="server">
        <p><b>请输入快递单号:</b>&nbsp<input id="Text1" name="Text1" type="text" value="">&nbsp<input name="mybutton" type="submit" value="查一下"  onclick="document.pressed=this.value"/></p>

    </form>

    <script type="text/javascript">
          function form_check() {
              
               if(document.getElementById('Text1').value=="") 
               {
                    alert("请输入快递单号"); 
                   return false; 
               }
              return true; 
          }
    </script>
      <% 
          strFirst = Request.Form("Text1")
	  Set objconn=server.createobject("adodb.connection")
	  objconn.open "provider=microsoft.jet.oledb.4.0; data source=" & server.mappath("example.mdb")
	  set objrecord=server.createobject("adodb.recordset")
	  set objrecord=objconn.execute("select * from kdinfo where kd_number like'" & strFirst & "'")

          if IsEmpty(strFirst) then
	  Response.Write  "&nbsp"

	  elseif objrecord.eof or objrecord.bof then

	  Response.Write  "单号不存在"
	  response.end
	  else
	  Response.Write("<b>找到了!!!</b><hr height:1px;border:none;border-top:1px solid #555555; />公司：" & objrecord.fields("company") & "<hr height:1px;border:none;border-top:1px solid #555555; />" & _

"日期：" & objrecord.fields("kd_date") & "<hr height:1px;border:none;border-top:1px solid #555555; />" & _

"校区：" & objrecord.fields("school") & "<hr height:1px;border:none;border-top:1px solid #555555; />" & _

"学员编号：" & objrecord.fields("stuNumber") & "<hr height:1px;border:none;border-top:1px solid #555555; />" & _

"收件人：" & objrecord.fields("receiverName") & "<hr height:1px;border:none;border-top:1px solid #555555; />" & _

"电话：" & objrecord.fields("phone") & "<hr height:1px;border:none;border-top:1px solid #555555; />" & _

"地址：" & objrecord.fields("address") & "<hr height:1px;border:none;border-top:1px solid #555555; />" & _

"快递单号：" & objrecord.fields("kd_number") & "<hr height:1px;border:none;border-top:1px solid #555555; />" & _

"讲义：" & objrecord.fields("bookName"))

          end if
     %>

</body>
</html>
