<%@ Page Language="VB" ContentType="text/html" ResponseEncoding="utf-8" %>
<%@ Register TagPrefix="MM" Namespace="DreamweaverCtrls" Assembly="DreamweaverCtrls,version=1.0.0.0,publicKeyToken=836f606ede05d46a,culture=neutral" %>
<MM:DataSet 
id="DataSet2"
runat="Server"
IsStoredProcedure="false"
ConnectionString='<%# System.Configuration.ConfigurationSettings.AppSettings("MM_CONNECTION_STRING_s01") %>'
DatabaseType='<%# System.Configuration.ConfigurationSettings.AppSettings("MM_CONNECTION_DATABASETYPE_s01") %>'
CommandText='<%# "SELECT * FROM s01major WHERE s01id = ?" %>'
Debug="true"
>
  <Parameters>
    <Parameter  Name="@s01id"  Value='<%# IIf((Request.QueryString("s01id") <> Nothing), Request.QueryString("s01id"), "") %>'  Type="Integer"   />
  </Parameters>
</MM:DataSet>
<MM:PageBind runat="server" PostBackBind="true" />
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>訊息內容</title>
<style type="text/css">
<!--
body {
	background-color: #CCFFCC;
}
a:link {
	color: #0033FF;
	text-decoration: none;
}
a:visited {
	text-decoration: none;
	color: #CC3366;
}
a:hover {
	text-decoration: none;
}
a:active {
	text-decoration: none;
	color: #009933;
}
h1,h2,h3,h4,h5,h6 {
	font-weight: bold;
}
h1 {
	font-size: 24px;
	color: #990000;
}
.style1 {font-size: 16px}
body,td,th {
	font-size: 16px;
}
-->
</style></head>
<body><!--顯示分行結果須用下面程式 -->
<script runat="server">
  Function Clean(str)
   str=Replace(str, vbCrLf, "<br>")
   Clean=Replace(str, chr(32), "&nbsp;&nbsp;")
  End Function
 </script> 
<div align="center">
  <h1>訊息內容</h1>
</div>
<table width="100%" border="1" cellspacing="2" cellpadding="0">
  <tr>
    <td width="13%" bgcolor="#66FFCC"><div align="center"><span class="style1">主題</span></div></td>
    <td width="87%" bgcolor="#FFFF99"><%# DataSet2.FieldValue("s01title", Container) %></td>
  </tr>
  <tr>
    <td bgcolor="#66FFCC"><div align="center">內容</div></td>
    <td bgcolor="#FFFF99"><%# Clean(DataSet2.FieldValue("s01con", Container)) %></td>
  </tr>
  <tr>
    <td bgcolor="#66FFCC"><div align="center">發訊者</div></td>
    <td bgcolor="#FFFF99"><%# DataSet2.FieldValue("s01name", Container) %></td>
  </tr>
  <tr>
    <td bgcolor="#66FFCC"><div align="center">訊息屬性</div></td>
    <td bgcolor="#FFFF99"><%# DataSet2.FieldValue("s01group", Container) %></td>
  </tr>
  <tr>
    <td bgcolor="#66FFCC"><div align="center">發訊日期</div></td>
    <td bgcolor="#FFFF99"><%# DateTime.Parse(DataSet2.FieldValue("s01time", Container)).ToString("D") %></td>
  </tr>
</table>
<p>相關檔案：</p>
<form id="form1" name="form1" method="post" action="">
  <input type="hidden" name="hiddenField" />
</form>
<p><a href=s01data_01/<%# DataSet2.FieldValue("s01data01", Container) %> target="_blank"><%# DataSet2.FieldValue("s01data01", Container) %></a></p>
<p><a href=s01data_01/<%# DataSet2.FieldValue("s01data02", Container) %> target="_blank"><%# DataSet2.FieldValue("s01data02", Container) %></a></p>
<p><a href=s01data_01/<%# DataSet2.FieldValue("s01data03", Container) %> target="_blank"><%# DataSet2.FieldValue("s01data03", Container) %></a></p>
<p><a href=s01data_01/<%# DataSet2.FieldValue("s01data04", Container) %> target="_blank"><%# DataSet2.FieldValue("s01data04", Container) %></a></p>
<p><a href=s01data_01/<%# DataSet2.FieldValue("s01data05", Container) %> target="_blank"><%# DataSet2.FieldValue("s01data05", Container) %></a></p>
<p>&nbsp;</p>
</body>
</html>
