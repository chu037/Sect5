<%@ Page Language="VB" ContentType="text/html" ResponseEncoding="utf-8" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>無標題文件</title>
<style type="text/css">
<!--
body {
	background-color: #FF99CC;
}
.style1 {
	font-size: 24px;
	color: #990000;
}
.style3 {font-size: 16px}
.style4 {color: #990000}
-->
</style></head>
<body>
<div align="center">
  <p><strong><span class="style1">請輸入帳號及密碼</span></strong></p>
  <p><strong>帳號及密碼與<span class="style4">公文系統</span>一樣</strong></p>
</div>
<form id="form1" name="form1" method="post" action="s01delcheckuser.aspx">
  <div align="center"><span class="style3">帳號：
    <input name="admin_username" type="text" id="admin_username" /> 
  、密碼：
  <input name="admin_password" type="password" id="admin_password" />
  </span>
  <input type="submit" name="Submit" value="登入" />
  </div>
</form>
</body>
</html>
