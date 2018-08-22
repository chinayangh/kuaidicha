<!-- #include file="user.asp" -->
<%
Set conn=server.createobject("ADODB.connection")
		if conn.state<>1 then	'判断现在数据库是否正在连接，才打开一次数据库，避免过多的数据库连接'
			conn.open "provider=Microsoft.Jet.OLEDB.4.0;Data Source=c:\inetpub\kuaidi\database.mdb"
		end if
		Set rs=server.CreateObject("adodb.recordset")

	username=Request.Form("username")
	password=Request.Form("password")
	rs.open "select username,password from user_info where username='"&username&"' and password='"&password&"'" ,conn,1,1
    if rs.eof and rs.bof then
    'response.write "账号密码不对"
          'if rs.recordcount<=0 then
          'response.write "账号密码不对"
          'end if
    else
    session("username")=rs("username") '创建session实现用户追踪
    session.timeout=10
    'response.write "<script>location.href='upload.asp'</script>"
    end if
	info="<script charset='utf-8'>"
	if has_this_user(username) then
		if password=get_password(username) then
			info=info&"alert('登录成功！');"
			response.write "<script>window.location.href='fileupload.asp'</script>"
		else
			info=info&"alert('密码错误！');"
		end if
	else
		info=info&"alert('查无此人！');"
	end if
	info=info&"window.location.href='index.html'"
	info=info&"</script>"
	response.Write info	
%>
