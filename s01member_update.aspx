<%@ Page Language="VB" ContentType="text/html" %>
<%@ Register TagPrefix="MM" Namespace="DreamweaverCtrls" Assembly="DreamweaverCtrls,version=1.0.0.0,publicKeyToken=836f606ede05d46a,culture=neutral" %>
<MM:Update
runat="server"
CommandText='<%# "UPDATE s01admin SET admin_yes=?, admin_password=?, admin_cell=?, admin_mail=?, admin_phone=?, s01name=?, admin_username=? WHERE admin_id=?" %>'
ConnectionString='<%# System.Configuration.ConfigurationSettings.AppSettings("MM_CONNECTION_STRING_s01") %>'
DatabaseType='<%# System.Configuration.ConfigurationSettings.AppSettings("MM_CONNECTION_DATABASETYPE_s01") %>'
Expression='<%# Request.Form("MM_update") = "form_update" %>'
CreateDataSet="false"
SuccessURL='<%# "s01member_index.aspx?msg=2" %>'
><Parameters>
  <Parameter Name="@admin_yes" Value='<%# IIf((Request.Form("admin_yes") <> Nothing), Request.Form("admin_yes"), "") %>' Type="Boolean" />  
  <Parameter Name="@admin_password" Value='<%# IIf((Request.Form("admin_password") <> Nothing), Request.Form("admin_password"), "") %>' Type="WChar" />  
  <Parameter Name="@admin_cell" Value='<%# IIf((Request.Form("admin_cell") <> Nothing), Request.Form("admin_cell"), "") %>' Type="WChar" />  
  <Parameter Name="@admin_mail" Value='<%# IIf((Request.Form("admin_mail") <> Nothing), Request.Form("admin_mail"), "") %>' Type="WChar" />  
  <Parameter Name="@admin_phone" Value='<%# IIf((Request.Form("admin_phone") <> Nothing), Request.Form("admin_phone"), "") %>' Type="WChar" />  
  <Parameter Name="@s01name" Value='<%# IIf((Request.Form("s01name") <> Nothing), Request.Form("s01name"), "") %>' Type="WChar" />  
  <Parameter Name="@admin_username" Value='<%# IIf((Request.Form("admin_username") <> Nothing), Request.Form("admin_username"), "") %>' Type="WChar" />  
  <Parameter Name="@admin_id" Value='<%# IIf((Request.Form("admin_id") <> Nothing), Request.Form("admin_id"), "") %>' Type="Integer" />  
</Parameters>
</MM:Update>
<MM:DataSet 
id="DataSet2"
runat="Server"
IsStoredProcedure="false"
ConnectionString='<%# System.Configuration.ConfigurationSettings.AppSettings("MM_CONNECTION_STRING_s01") %>'
DatabaseType='<%# System.Configuration.ConfigurationSettings.AppSettings("MM_CONNECTION_DATABASETYPE_s01") %>'
CommandText='<%# "SELECT * FROM s01admin WHERE admin_id = ?" %>'
Debug="true"
>
  <Parameters>
    <Parameter  Name="@admin_id"  Value='<%# IIf((Request.QueryString("admin_id") <> Nothing), Request.QueryString("admin_id"), "") %>'  Type="Integer"   />  
  </Parameters>
</MM:DataSet>
<MM:PageBind runat="server" PostBackBind="true" />
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=UTF-8" />
<title>更新組員</title>
<script language="VB" runat="server">
 dim url, url_back
  Sub page_load(sender As Object, e As EventArgs)
		If Session("MM_chief_num")="" then
			Response.Write("您還沒有登入呢！，請點擊")
			Response.Write("<a href=s01dellogin.aspx>登入</a>")
			Response.End()
	    else
		   url_back="s01delindex.aspx"
		End If
	End Sub

    </script>

<script language = "JavaScript">
<!--
function Mcheck(){
	if (document.form_update.admin_username.value=="") {
        window.alert("請輸入帳號");
        return false }
    if (document.form_update.s01name.value=="") {
        window.alert("請輸入姓名");
        return false }
	if (document.form_update.admin_password.value=="") {
        window.alert("請輸入密碼");
		return false }
	 return true;
}
//-->
</Script>

<style type="text/css">
<!--
.style1 {
	color: #CC3300;
	font-weight: bold;
	font-size: 16px;
}
-->
</style>
</head>
<body>
<span class="style1">更新組員</span>
<form runat='server' method='POST' name='form_update' id="form_update" onSubmit="return Mcheck()">
<table width="512" border="0" cellspacing="1" cellpadding="0">
  <tr>
    <td width="86" bgcolor="#CCFFCC">姓名：</td>
    <td width="411" bgcolor="#FFCC99">
      <input name="s01name" type="text" id="s01name" value='<%# DataSet2.FieldValue("s01name", Container) %>' />
    </td>
  </tr>
  <tr>
    <td width="86" bgcolor="#CCFFCC">帳號：</td>
    <td bgcolor="#FFCC99"><input name="admin_username" type="text" id="admin_username" value='<%# DataSet2.FieldValue("admin_username", Container) %>' /></td>
  </tr>
  <tr>
    <td width="86" bgcolor="#CCFFCC">密碼：</td>
    <td bgcolor="#FFCC99">
 <input name="admin_password" type="text" id="admin_password" value='<%# DataSet2.FieldValue("admin_password", Container) %>' /></td>
  </tr>
  <tr>
  <tr>
    <td width="86" bgcolor="#CCFFCC">分機：</td>
    <td bgcolor="#FFCC99">
 <input name="admin_phone" type="text" id="admin_phone" value='<%# DataSet2.FieldValue("admin_phone", Container) %>' /></td>
  </tr>
  <tr>
  <tr>
    <td width="86" bgcolor="#CCFFCC">手機：</td>
    <td bgcolor="#FFCC99">
 <input name="admin_cell" type="text" id="admin_cell" value='<%# DataSet2.FieldValue("admin_cell", Container) %>' /></td>
  </tr>
  <tr>
  <tr>
    <td width="86" bgcolor="#CCFFCC">mail：</td>
    <td bgcolor="#FFCC99">
 <input name="admin_mail" type="text" id="admin_mail" value='<%# DataSet2.FieldValue("admin_mail", Container) %>' /></td>
  </tr>
  <tr>

    <td width="86" bgcolor="#CCFFCC">有效：</td>
    <td bgcolor="#FFCC99"><asp:CheckBox Checked='true' ID="admin_yes" runat="server" Text='<%# DataSet2.FieldValue("admin_yes", Container) %>' /></td>
  </tr>
  <tr>
    <td width="86" bgcolor="#CCFFCC">內容：</td>
    <td bgcolor="#FFCC99">&nbsp;</td>
  </tr>
</table>
  <span class="style1">
  <input name="Submit4" type="submit" value="更新" />
  <input type="reset" name="Submit2" value="重新填寫" />
  <input name="Submit" type="button" onClick="history.back()" value="取消" />
  </span>
  
  <p>
<input name="admin_id" type="hidden" id="admin_id" value="<%# DataSet2.FieldValue("admin_id", Container) %>" />
<input type="hidden" name="MM_update" value="form_update">
</form>

</body>
</html>
