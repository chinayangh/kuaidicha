<html>
<head>
<meta http-equiv="refresh" content="0">
</head>
</html>

<% 
session.abandon()
If IsEmpty(Session("username")) Then
	response.write "<script>window.location.href='index.html';</script>"
	response.redirect "index.html"
end if
%>
