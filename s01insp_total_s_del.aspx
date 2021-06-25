<%@ Page Language="VB" ContentType="text/html" ResponseEncoding="utf-8" %>
<%@ Register TagPrefix="MM" Namespace="DreamweaverCtrls" Assembly="DreamweaverCtrls,version=1.0.0.0,publicKeyToken=836f606ede05d46a,culture=neutral" %><MM:Delete
runat="server"
CommandText='<%# "DELETE FROM s01insp_section WHERE insp_s_id=?" %>'
ConnectionString='<%# System.Configuration.ConfigurationSettings.AppSettings("MM_CONNECTION_STRING_s01") %>'
DatabaseType='<%# System.Configuration.ConfigurationSettings.AppSettings("MM_CONNECTION_DATABASETYPE_s01") %>'
Expression='<%# (Request.QueryString("case_id") <> Nothing) %>'
CreateDataSet="false"
SuccessURL='<%# "s01insp_total_s_index.aspx" %>'
Debug="true"
><Parameters>
  <Parameter Name="@insp_s_id" Value='<%# IIf((Request.QueryString("case_id") <> Nothing), Request.QueryString("case_id"), "") %>' Type="Integer" /></Parameters>
</MM:Delete>
<MM:PageBind runat="server" PostBackBind="true" />
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>刪除轄區</title>
</head>
<body>
</body>
</html>
