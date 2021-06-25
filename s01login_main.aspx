<%@ Page Language="VB" ContentType="text/html" ResponseEncoding="utf-8" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>登入首頁</title>
<style type="text/css">
<!--
body {
	background-color: #CCFFCC;
}
.style1 {
	font-size: x-large;
	color: #CC0066;
}
-->
</style>
<script language="vb" runat="server">
  Sub page_load(sender As Object, e As EventArgs)
		If Session("MM_username")="" then
			Response.Write("您還沒有登入呢！，請點擊")
			Response.Write("<a href=s01login.aspx>登入</a>")
			Response.End()
		End If
    session("g_result")= nothing
 end sub
 </script>
</head>
<body>

<div align="center">
  <p><strong><span class="style1">請點選要進入之主題</span></strong></p>
  <ul>
    <li>
      <div align="justify"><strong align="justify"><a href="s01case_update_normal.aspx" target="_blank">列管案件</a></strong></div>
    </li>
    <li>
      <div align="justify"><strong align="justify"><a href="s01total_search.aspx" target="_blank">檢查名冊查詢</a></strong></div>
    </li>

  </ul>
</div>
</body>
</html>
