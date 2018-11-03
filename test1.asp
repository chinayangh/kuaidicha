<html>
<body>
<div align="right">您访问页面的时间是<%=date()%>&nbsp<%=time()%></div>

<font face="宋体,仿宋体,隶书">显示字体的汉字</font><br>
<%
rem 这是一条注释语句
'这是另外一条注释语句

%>
<% for i=1 to 12 %>

<font size="<% =i %>">使用asp语句控制字体大小</font><br>

<%  next  %>

</body>
</html>