function check_equal(password1,password2){//判断密码是否相等
	if(password1==password2){
		return true;	
	}
	else{
		return false;	
	}
}
/*function register_check(){//注册表单的判断
	var username=document.getElementById("register_username").value;
	var password1=document.getElementById("register_password1").value;
	var password2=document.getElementById("register_password2").value;
	if(username.length>0&&password1.length>0&&password2.length>0){
		if(check_equal(password1,password2)){
			return true;
		}
		else{
			alert("两次输入的密码不一致！");		
		}
	}
	else{
		alert("用户名、密码任意一项为空！");	
	}
	return false;		
}
function update_check(){//修改密码表单的判断
	var username=document.getElementById("update_username").value;
	var password=document.getElementById("update_password").value;
	var password1=document.getElementById("update_password1").value;
	var password2=document.getElementById("update_password2").value;
	if(username.length>0&&password.length>0&&password1.length>0&&password2.length>0){
		if(check_equal(password1,password2)){
			return true;
		}
		else{
			alert("两次输入的密码不一致！");		
		}
	}
	else{
		alert("用户名、密码任意一项为空！");	
	}
	return false;		
}*/
