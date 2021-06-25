<%@ Page Language="VB" ContentType="text/html" ResponseEncoding="utf-8" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<%@ Register TagPrefix="MM" Namespace="DreamweaverCtrls" Assembly="DreamweaverCtrls,version=1.0.0.0,publicKeyToken=836f606ede05d46a,culture=neutral" %>
<MM:DataSet 
id="checkname"
runat="Server"
IsStoredProcedure="false"
ConnectionString='<%# System.Configuration.ConfigurationSettings.AppSettings("MM_CONNECTION_STRING_s01") %>'
DatabaseType='<%# System.Configuration.ConfigurationSettings.AppSettings("MM_CONNECTION_DATABASETYPE_s01") %>'
CommandText='<%# "SELECT *  FROM s01admintop  WHERE admin_username = ? and admin_password = ?" %>'
Debug="true"
><Parameters>
  <Parameter  Name="@admin_username"  Value='<%# IIf((Request.Form("admin_username") <> Nothing), Request.Form("admin_username"), "") %>'  Type="WChar"   />  
  <Parameter  Name="@admin_password"  Value='<%# IIf((Request.Form("admin_password") <> Nothing), Request.Form("admin_password"), "") %>'  Type="WChar"   />  
</Parameters></MM:DataSet>
<MM:PageBind runat="server" PostBackBind="true" />
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<meta http-equiv="Content-Type" content="text/html />
<head>
<title>無標題文件</title>
</head>
<body>

<script runat="server">
'自定義一個名為“Login”的函數
	Function Login(admin_username,s01name)
		'當動態變數不為空時
		If admin_username<>nothing Then
			'把使用者帳號賦值給session變數
			Session("MM_chief_num") = admin_username
			'把使用者編號賦值給session變數
			Session("MM_chief_name") = s01name
			'登入成功後轉到管理首頁
			Response.Redirect("s01delindex.aspx")
		Else
			'輸入的使用者帳號、密碼不正確時提示使用者重新登入
			Response.Write("使用者帳號/密碼不正確，請點擊")
			Response.Write("<a href=s01dellogin.aspx>重新登入</a>")
		End If
	End Function
</script>
<%# Login(checkname.FieldValue("admin_username", Container),checkname.FieldValue("s01name", Container)) %>

</body>
</html>
