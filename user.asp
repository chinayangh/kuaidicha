<!-- #include file="db.asp" -->
<%
	Set db_control=New DB	'初始化用户操作类'
	
	Function has_this_user(username)	'是否存在该用户的判断'
		MyArray=db_control.getbySQL("select * from user_info where username='"&username&"'")
		has_this_user=isarray(MyArray)
	End Function
	
	Function get_password(username)	'根据用户名取得密码'
		MyArray=db_control.getbySQL("select password from user_info where username='"&username&"'")	
		get_password=MyArray(0,0)		
	End Function
 
	Sub insert_user(username,password)	'插入用户'
		db_control.setbySQL("insert into user_info(username,password) values('"&username&"','"&password&"')")
	End Sub
	
	Sub update_user(username,password)	'修改用户密码'
		db_control.setbySQL("update user_info set password='"&password&"' where username='"&username&"'")
	End Sub
	
%>
