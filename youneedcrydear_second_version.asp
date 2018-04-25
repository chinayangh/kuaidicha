<!DOCTYPE html >

<html >
<head>
	<title>快递查询</title>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
</head>
<body>
    <form id="top" method="post" onsubmit="return form_check();" >
        <p><b>快递单号:</b>&nbsp<input id="Text1" name="Text1" type="text" value="" autocomplete="off">&nbsp
        	<b>收件人姓名:</b>&nbsp<input id="Text2" name="Text2" type="text" value="" autocomplete="off">&nbsp
        	<b>学员编号:</b>&nbsp<input id="Text3" name="Text3" type="text" value="" autocomplete="off">&nbsp
        	<input name="mybutton" type="submit" value="查一下"  onclick="document.pressed=this.value"/></p>

    </form>
    <div style="font-size: 11px;text-align: right;">&#9733查询时候的关键字千万不要复制<font color="red">空格</font>和<font color="red">Tab</font>等带格式不可见字符哦&#9786</div>
    <div id="scrollTop" class="displayNone" style="position: fixed; bottom: 94px; right: 0px; z-index: 100; display: block;">
	<a title="返回顶部" href="#top">
		<img src="up.png"/>
	</a>
	</div>


    <script type="text/javascript">
          function form_check() {
              
               if(document.getElementById('Text1').value==""&&document.getElementById('Text2').value==""&&document.getElementById('Text3').value=="") 
               {
                    alert("请至少输入一项"); 
                   return false; 
               }
             
              return true; 
          }
    </script>
      <% 
          strFirst = Request.Form("Text1")
          strSecond = Request.Form("Text2")
          strThird = Request.Form("Text3")
          i = 0
	  Set objconn=server.createobject("adodb.connection")
	  objconn.open "provider=microsoft.jet.oledb.4.0; data source=" & server.mappath("kdinfo.mdb")
	  set objrecord=server.createobject("adodb.recordset")
	  set objrecord=objconn.execute("select * from kdinfo where kd_number like'" & strFirst & "' or receiverName like'" & strSecond & "' or stuNumber like '" & strThird & "'")

          if IsEmpty(strFirst) then
	  Response.Write  "&nbsp"

	  elseif objrecord.eof or objrecord.bof then

	  Response.Write  "没有相关记录"
	  
	  elseif not objrecord.eof then
	  do while not objrecord.eof
	  	i = i + 1
	  	Response.Write i & "<hr height:1px;border:none;border-top:1px solid #555555; /> " 
	  		if IsEmpty(objrecord.fields("company")) then
	  		Response.Write "公司："
	  		else Response.Write "公司：" & objrecord.fields("company") & "&#124" 
	  		end if
	  		if	IsEmpty(objrecord.fields("kd_date"))	then
	  		Response.Write "日期："
	  		else Response.Write "日期：" & objrecord.fields("kd_date") & "&#124" 
	  		end if
	  		if IsEmpty(objrecord.fields("school"))	then
	  		Response.Write "校区："
	  		else Response.Write "校区：" & objrecord.fields("school") & "&#124" 
	    	end if
	  		if IsEmpty(objrecord.fields("stuNumber"))	then
	  		Response.Write "学员编号："
	  		else Response.Write "学员编号：" & objrecord.fields("stuNumber") & "&#124" 
	  		end if
	  		if IsEmpty(objrecord.fields("receiverName")) then
	  		Response.Write "收件人："
	  		else Response.Write "收件人：" & objrecord.fields("receiverName") & "&#124" 
	   		end if
	  		if IsEmpty(objrecord.fields("phone"))	then
	  		Response.Write "电话："
	  		else Response.Write  "电话：" & objrecord.fields("phone") & "&#124" 
	  		end if
	  		if IsEmpty(objrecord.fields("address"))	then
	  		Response.Write "地址："
	  		else Response.Write "地址：" & objrecord.fields("address") & "&#124" 
	  		end if
	  		if IsEmpty(objrecord.fields("kd_number"))	then
	  		Response.Write "快递单号："
	  		else Response.Write "快递单号：" & objrecord.fields("kd_number") & "&#124" 
	  		end if
	  		if IsEmpty(objrecord.fields("bookName"))	then
	  		Response.Write "讲义："
	  		else Response.Write "讲义：" & objrecord.fields("bookName") & "<hr height:1px;border:none;border-top:1px solid #555555; />"
	  		end if
	  	objrecord.movenext
	  	loop
	  	end if
          objrecord.close
     %>

</body>
</html>
