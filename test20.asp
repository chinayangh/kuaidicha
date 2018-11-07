<!--IE浏览器版本要求11以下-->

<html>
<head>
	<meta charset="utf-8">
	<title></title>

	
</head>
<body>
	<script Language="VBScript">
		
		Dim strColor
		strColor=inputbox("请从red、green、blue中选择一个，并输入作为页面颜色！")
		
		Select case strColor
			case "red" 
				document.bgColor="red"
			case "green" 
				document.bgColor="green"
			case "blue" 
				document.bgColor="blue"
			case else 
				MsgBox "输入有误"

		end Select
		
	</script>
	
	
</body>


</html>
