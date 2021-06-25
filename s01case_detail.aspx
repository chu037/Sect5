<%@ Page Language="VB" ContentType="text/html"  %>
<%@ Register TagPrefix="MM" Namespace="DreamweaverCtrls" Assembly="DreamweaverCtrls,version=1.0.0.0,publicKeyToken=836f606ede05d46a,culture=neutral" %>
<MM:DataSet 
id="DataSet1"
runat="Server"
IsStoredProcedure="false"
ConnectionString='<%# System.Configuration.ConfigurationSettings.AppSettings("MM_CONNECTION_STRING_s01") %>'
DatabaseType='<%# System.Configuration.ConfigurationSettings.AppSettings("MM_CONNECTION_DATABASETYPE_s01") %>'
CommandText='<%# "SELECT * FROM s01_case WHERE case_id = ?" %>'
Debug="true"
>
  <Parameters>
    <Parameter  Name="@case_id"  Value='<%# IIf((Request.QueryString("case_id") <> Nothing), Request.QueryString("case_id"), "") %>'  Type="Integer"   />
  </Parameters>
</MM:DataSet>
<MM:DataSet 
id="DataSet2"
runat="Server"
IsStoredProcedure="false"
ConnectionString='<%# System.Configuration.ConfigurationSettings.AppSettings("MM_CONNECTION_STRING_s01") %>'
DatabaseType='<%# System.Configuration.ConfigurationSettings.AppSettings("MM_CONNECTION_DATABASETYPE_s01") %>'
CommandText='<%# "SELECT * FROM s01_case_appen WHERE case_id = ? ORDER BY case_appen_id ASC" %>'
Debug="true" PageSize="20"
>
  <Parameters>
    <Parameter  Name="@case_id"  Value='<%# IIf((Request.QueryString("case_id") <> Nothing), Request.QueryString("case_id"), "") %>'  Type="Integer"   />
  </Parameters>
</MM:DataSet>
<MM:PageBind runat="server" PostBackBind="true" />
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>列管案件辦理明細</title>
<style type="text/css">
<!--
body,td,th {
	font-size: 14px;
	color: #000000;
}
a:link {
	color: #0033CC;
	text-decoration: none;
}
a:visited {
	color: #99CC33;
	text-decoration: none;
}
a:hover {
	color: #33FFCC;
	text-decoration: none;
}
a:active {
	color: #CC0000;
	text-decoration: none;
}
.style3 {color: #006600}
-->
</style>
<script language="VB" runat="server">

 function showdate(s01time)
 if s01time <> ""
 showdate = FormatDateTime(s01time, DateFormat.ShortDate)
 end if
 end function
 function showtime(t)
 if t <> ""
 showtime = FormatDateTime(t, DateFormat.Shorttime)
 end if
 end function
 function showweek(t)
 if t <> ""
 showweek = weekdayname(datepart("w",t))
 end if
 end function

 function showmessage(vt2,vt3)
 if vt2 <> ""
 showmessage = vt2 & "<br/>" & vt3 & "小時"
 end if
 end function
  Function Clean(str)
   str=Replace(str, vbCrLf, "<br>")
   Clean=Replace(str, chr(32), "&nbsp;&nbsp;")
  End Function

</script>
<script language = "JavaScript">
<!--



function w_back()
{
location.href = "calendartest02.aspx"
}
//-->
</Script>
</head>
<body>
<table width="655" border="1" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF" id="t00">
  <tr>
    <td width="13%" bgcolor="#FFFFFF"><div align="center">案件日期</div></td>
    <td width="87%" bgcolor="#FFFFFF"><span class="style3"><%# showdate(DataSet1.FieldValue("case_start", Container)) %></span>，辦理天數:<span class="style3"><%# DataSet1.FieldValue("case_alldays", Container) %></span></td>
  </tr>
  <tr>
    <td width="13%" bgcolor="#FFFFFF"><div align="center">案件日期</div></td>
    <td width="87%" bgcolor="#FFFFFF"><span class="style3"><%# showdate(DataSet1.FieldValue("case_end", Container)) %></span></td>
  </tr>

  <tr>
    <td bgcolor="#FFFFFF"><div align="center">類型</div></td>
    <td bgcolor="#FFFFFF"><span class="style3"><%# DataSet1.FieldValue("case_group", Container) %></span></td>
  </tr>
  <tr>
    <td bgcolor="#FFFFFF"><div align="center">內容</div></td>
    <td bgcolor="#FFFFFF"><span class="style3"><%# DataSet1.FieldValue("case_content", Container) %></span></td>
  </tr>
  <tr>
    <td bgcolor="#FFFFFF"><div align="center">辦理情形</div></td>
    <td bgcolor="#FFFFFF"><span class="style3"><%# DataSet1.FieldValue("case_result", Container) %></span></td>
  </tr>
  <tr bgcolor="#FFFFFF">
    <td colspan="2" align="center"><div align="left">
      <p>相關檔案：</p>
    </div></td>
  </tr>
</table>
<form runat="server">
    <asp:DataGrid 
  AllowCustomPaging="true" 
  AllowPaging="true" 
  AllowSorting="False" 
  AutoGenerateColumns="false" 
  CellPadding="3" 
  CellSpacing="0" 
  DataSource="<%# DataSet2.DefaultView %>" id="DataGrid1" 
  PagerStyle-Mode="NumericPages" 
  PageSize="<%# DataSet2.PageSize %>" 
  runat="server" 
  ShowFooter="false" 
  ShowHeader="true" Width="656" 
  OnPageIndexChanged="DataSet2.OnDataGridPageIndexChanged" 
  VirtualItemCount="<%# DataSet2.RecordCount %>" 
>
      <HeaderStyle HorizontalAlign="center" BackColor="#E8EBFD" ForeColor="#3D3DB6" Font-Name="Verdana, Arial, Helvetica, sans-serif" Font-Bold="true" Font-Size="smaller" />    
      <ItemStyle BackColor="#F2F2F2" Font-Name="Verdana, Arial, Helvetica, sans-serif" Font-Size="smaller" />    
      <AlternatingItemStyle BackColor="#E5E5E5" Font-Name="Verdana, Arial, Helvetica, sans-serif" Font-Size="smaller" />    
      <FooterStyle HorizontalAlign="center" BackColor="#E8EBFD" ForeColor="#3D3DB6" Font-Name="Verdana, Arial, Helvetica, sans-serif" Font-Bold="true" Font-Size="smaller" />    
      <PagerStyle BackColor="white" Font-Name="Verdana, Arial, Helvetica, sans-serif" Font-Size="smaller" />    
      <Columns>
      <asp:BoundColumn DataField="case_appen_id" 
        HeaderText="case_appen_id" 
        ReadOnly="true" 
        Visible="false"/>    
      <asp:HyperLinkColumn
        DataNavigateUrlField="case_appen" DataNavigateUrlFormatString="/web1/sec05/s01data/{0}"
        DataTextField="case_appen" 
        Visible="True" target="case_detail_blank" 
        HeaderText="檔案"/>
      <asp:BoundColumn DataField="case_appen_con" 
        HeaderText="說明" 
        ReadOnly="true" 
        Visible="True"/>    
      <asp:BoundColumn DataField="s01name" 
        HeaderText="承辦人" 
        ReadOnly="true" 
        Visible="True"/>    
  </Columns>
  </asp:DataGrid>
</form>

  <form id="form1" name="form1" method="post" action="">
<p>
  <input name="Submit" type="button" onclick="window.close()" value="關閉視窗" />
</p>
</form>
</body>
</html>
