<html>
  <head>
    <title>hello from VBScript!</title>
    <script language="VBScript" RunAt="Server">
      sub show()

      Response.Write("Hello from VBScript!")
      End sub

    </script>

  </head>

<body>
	<%
		Call show()
	%>
<body>


</html>
