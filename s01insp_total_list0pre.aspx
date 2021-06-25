<%@ Page Language="VB" ContentType="text/html" ResponseEncoding="utf-8" %>
<%@ Register TagPrefix="MM" Namespace="DreamweaverCtrls" Assembly="DreamweaverCtrls,version=1.0.0.0,publicKeyToken=836f606ede05d46a,culture=neutral" %>
<MM:DataSet 
id="DataSet1"
runat="Server"
IsStoredProcedure="false"
ConnectionString='<%# System.Configuration.ConfigurationSettings.AppSettings("MM_CONNECTION_STRING_s01") %>'
DatabaseType='<%# System.Configuration.ConfigurationSettings.AppSettings("MM_CONNECTION_DATABASETYPE_s01") %>'
CommandText='<%# "SELECT * FROM t_group WHERE t_g_id = ?" %>'
Debug="true"
>
  <Parameters>
    <Parameter  Name="@t_g_id"  Value='<%# IIf((Request.QueryString("t_g_id") <> Nothing), Request.QueryString("t_g_id"), "") %>'  Type="Integer"   />
  </Parameters>
</MM:DataSet>
<MM:PageBind runat="server" PostBackBind="true" />
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>check</title>
<script runat="server">
		Sub Page_Load(Src As Object, E As EventArgs)
		If Session("MM_username")="" then
			Response.Write("您還沒有登入呢！，請點擊")
			Response.Write("<a href=s01login.aspx>登入</a>")
			Response.End()
		End If
       end sub

	function Page_start(x01,x02,x03,x04)
		dim url0 as string
		url0 = "s01insp_total_list0.aspx?t_g_1=" & x01 & "&t_g_2=" & x02 & "&t_g_3=" & x03 & "&t_g_num=" & x04
		response.Redirect(url0)
		end function
</script>
</head>
<body>
<%# page_start(DataSet1.FieldValue("t_g_1", Container),DataSet1.FieldValue("t_g_2", Container),DataSet1.FieldValue("t_g_3", Container),DataSet1.FieldValue("t_g_num", Container)) %>
</body>
</html>
