<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
'禁止缓存'
Response.CacheControl="no-cache"
Response.Expires=-1
Response.Charset="UTF-8"	'配合第一行设定网页编码'
 
'数据库操作类'
Class DB
	Private CONN	'数据库连接'
	Private RS	'数据库查询结果集'
	Private Sub Open_db_connect()	'类的构造过程，进行数据库连接、初始化数据库查询结果集的操作'
		Set CONN=server.createobject("ADODB.connection")
		if CONN.state<>1 then	'判断现在数据库是否正在连接，才打开一次数据库，避免过多的数据库连接'
			CONN.open "DBQ="&server.MapPath("database.mdb")&";DRIVER={Microsoft Access Driver (*.mdb)}"
		end if
		Set RS=server.CreateObject("adodb.recordset")
	End Sub
	
	Public Function getbySQL(sql)	'根据sql语句，对数据库有返回结果查询'
		Open_db_connect()
		RS.open sql,CONN,1,3
		if not (RS.bof or RS.eof) then	'如果查询结果非空，就将所有查询结果压到数组里面'
			MyArray=RS.GetRows
			getbySQL=MyArray	'返回这个记载查询结果的数组给前台'
		end if
		Close_db_connect()
	End Function
	
	Public Sub setbySQL(sql)	'根据sql语句，对数据库进行无返回结果的操作'
		Open_db_connect()
		CONN.execute sql
		Close_db_connect()
	End Sub
		
	Private Sub Close_db_connect()	'类的析构函数，主要是关闭数据库连接'
		if CONN.state=1 then	'判断现在数据库是否打开，是，才关闭'
			CONN.close
			set Con=nothing  
		end if		
	End Sub
	
End Class
%>
