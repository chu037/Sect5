<%@ Page Language="VB" ContentType="text/html"%>
<%@ Register TagPrefix="MM" Namespace="DreamweaverCtrls" Assembly="DreamweaverCtrls,version=1.0.0.0,publicKeyToken=836f606ede05d46a,culture=neutral" %>
<MM:DataSet 
id="DataSet1"
runat="Server"
IsStoredProcedure="false"
ConnectionString='<%# System.Configuration.ConfigurationSettings.AppSettings("MM_CONNECTION_STRING_s01") %>'
DatabaseType='<%# System.Configuration.ConfigurationSettings.AppSettings("MM_CONNECTION_DATABASETYPE_s01") %>'
CommandText='<%# "SELECT t_id, t_insu, t_pre  FROM s01total  WHERE t_pre like ?" %>'
Debug="true"
>
  <Parameters>
    <Parameter  Name="t_pre"  Value='<%# "%" +(IIf((Request.form("p_num") <> Nothing), Request.form("p_num"), "")) + "%" %>'  Type="WChar"   />  
  </Parameters>
</MM:DataSet>
<MM:PageBind runat="server" PostBackBind="true" />
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>檢驗統一編號</title>
</head>
<body>

<script runat="server">
'自定義一個名為“Login”的函數
	Function Login(p_num)
     if p_num = "" then
	  Session("MM_p_num") = request.Form("p_num")
	  Response.Redirect("s01insp_t_add.aspx")
     'response.Write("123")
     'response.Write(request.form("p_num"))
	 else 
	  Session("MM_p_num") = p_num
	  Response.Redirect("s01insp_t_list.aspx")
	 'response.Write(p_num)
	 end if
 	End Function
</script>
<%# Login(DataSet1.FieldValue("t_pre", Container)) %>
</body>
</html>
