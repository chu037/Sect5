<%@ Page Language="VB" ContentType="text/html"%>
<%@ Register TagPrefix="MM" Namespace="DreamweaverCtrls" Assembly="DreamweaverCtrls,version=1.0.0.0,publicKeyToken=836f606ede05d46a,culture=neutral" %>
<html xmlns="http://www.w3.org/1999/xhtml">
<MM:DataSet 
id="DataSet1"
runat="Server"
IsStoredProcedure="false"
ConnectionString='<%# System.Configuration.ConfigurationSettings.AppSettings("MM_CONNECTION_STRING_s01") %>'
DatabaseType='<%# System.Configuration.ConfigurationSettings.AppSettings("MM_CONNECTION_DATABASETYPE_s01") %>'
CommandText='<%# "SELECT DISTINCT s01name, count(*) as case_count FROM s01_case where case_end_yes = ? and s01name <> ? GROUP BY s01name ORDER BY count(*) desc" %>'
Debug="true"
>
  <Parameters>
<Parameter  Name="@case_end_yes"  Value='<%# "否" %>'  Type="WChar"/>
<Parameter  Name="@s01name"  Value='<%# "" %>'  Type="WChar"/>

</Parameters></MM:DataSet>
<MM:DataSet 
id="DataSet2"
runat="Server"
IsStoredProcedure="false"
ConnectionString='<%# System.Configuration.ConfigurationSettings.AppSettings("MM_CONNECTION_STRING_s01") %>'
DatabaseType='<%# System.Configuration.ConfigurationSettings.AppSettings("MM_CONNECTION_DATABASETYPE_s01") %>'
CommandText='<%# "SELECT distinct case_id, case_start, case_group, s01name, case_end_yes, case_end_yes, case_yes, case_content, case_audit, admin_id, case_days, case_limit, case_alldays  FROM s01_case  WHERE case_end_yes = ? and s01name like ? ORDER BY case_days DESC" %>'
PageSize="80"
Debug="true"
>
  <Parameters>
    <Parameter  Name="@case_end_yes"  Value='<%# "否" %>'  Type="WChar"   />  
    <Parameter  Name="s01name"  Value='<%# IIf((Request.QueryString("s01name") <> Nothing), Request.QueryString("s01name"), "%") %>'  Type="WChar"   />  
  </Parameters>
</MM:DataSet>
<MM:PageBind runat="server" PostBackBind="true" />
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<meta http-equiv="refresh" content="600" >
<title>管制案件稽催</title>
<script language="VB" runat="server">
   Sub Page_Load(sender As Object, e As EventArgs)
   End Sub

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
 function showyes(y)
 if y = "True" then
 showyes = "是"
 else
 showyes = "否"
 end if
 end function
 function showmessage(vt2,vt3)
 if vt2 <> ""
 showmessage = vt2 & "<br/>" & vt3 & "小時"
 end if
 end function

    </script>
<style type="text/css">
<!--
body {
	background-color: #CCFFCC;
}
.style1 {
	color: #CC3333;
	font-weight: bold;
	font-size: 16px;
}
.style2 {color: #990000}
.style3 {
	color: #993300;
	font-weight: bold;
	font-size: 16px;
}
.style4 {color: #3300CC}
.style5 {
	font-size: 16px;
	font-weight: bold;
	color: #990000;
}
.style6 {
	font-size: 18px;
	font-weight: bold;
}
-->
</style></head>
<body>
<form runat="server">
 <table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
   <td colspan="2">
   <span class="style1">尚未結案統計一覽表
   (請點選<a href="s01login.aspx" target="_blank">登入</a></div>結案)，已辦畢卻未結案之案件請告知科長   </span>
   </td>
  </tr>
  <tr>
   <td width="20%" valign="top">
   <asp:DataGrid AllowPaging="false" 
  AllowSorting="False" 
  AutoGenerateColumns="false" 
  CellPadding="3" 
  CellSpacing="0" 
  DataSource="<%# DataSet1.DefaultView %>" id="DataGrid1" 
  runat="server" 
  ShowFooter="false" 
  ShowHeader="true" 
>
    <HeaderStyle HorizontalAlign="center" BackColor="#E8EBFD" ForeColor="#3D3DB6" Font-Name="Verdana, Arial, Helvetica, sans-serif" Font-Bold="true" Font-Size="smaller" />
    <ItemStyle BackColor="#F2F2F2" Font-Name="Verdana, Arial, Helvetica, sans-serif" Font-Size="smaller" />
    <AlternatingItemStyle BackColor="#E5E5E5" Font-Name="Verdana, Arial, Helvetica, sans-serif" Font-Size="smaller" />
    <FooterStyle HorizontalAlign="center" BackColor="#E8EBFD" ForeColor="#3D3DB6" Font-Name="Verdana, Arial, Helvetica, sans-serif" Font-Bold="true" Font-Size="smaller" />
    <PagerStyle BackColor="white" Font-Name="Verdana, Arial, Helvetica, sans-serif" Font-Size="smaller" />
    <Columns>
      <asp:BoundColumn DataField="s01name" 
        HeaderText="姓名" 
        ReadOnly="true" 
        Visible="True"/>      
<asp:HyperLinkColumn DataNavigateUrlField="s01name" DataNavigateUrlFormatString="s01case_present.aspx?s01name={0}"
        DataTextField="case_count" 
        Visible="True" 
        HeaderText="尚未結案件數"/>      
</Columns>
  </asp:DataGrid>
  </td>
 <td width="80%" valign="top">
 <asp:DataGrid 
  AllowCustomPaging="true" 
  AllowPaging="true" 
  AllowSorting="true" 
  AutoGenerateColumns="false" 
  CellPadding="3" 
  CellSpacing="0" 
  DataSource="<%# DataSet2.DefaultView %>" id="DataGrid2" 
  PagerStyle-Mode="NumericPages" 
  PageSize="<%# DataSet2.PageSize %>" 
  runat="server" 
  ShowFooter="false" 
  ShowHeader="true" Width="100%" 
  OnPageIndexChanged="DataSet2.OnDataGridPageIndexChanged" 
  VirtualItemCount="<%# DataSet2.RecordCount %>" 
>
                  <HeaderStyle HorizontalAlign="center" BackColor="#FFFFCC" ForeColor="#3D3DB6" Font-Name="Verdana, Arial, Helvetica, sans-serif" Font-Bold="true" Font-Size="smaller" />                
                  <ItemStyle BackColor="#FFCCFF" Font-Name="Verdana, Arial, Helvetica, sans-serif" Font-Size="smaller" />                
                  <AlternatingItemStyle BackColor="#CCCCCC" Font-Name="Verdana, Arial, Helvetica, sans-serif" Font-Size="smaller" />                
                  <FooterStyle HorizontalAlign="center" BackColor="#E8EBFD" ForeColor="#3D3DB6" Font-Name="Verdana, Arial, Helvetica, sans-serif" Font-Bold="true" Font-Size="smaller" />                
                  <PagerStyle BackColor="white" Font-Name="Verdana, Arial, Helvetica, sans-serif" Font-Size="smaller" />                
                  <Columns>
                  <asp:BoundColumn DataField="case_id" 
        HeaderText="case_id" 
        ReadOnly="true" 
        Visible="false"/>                  
                  <asp:BoundColumn DataField="case_group" 
        HeaderText="類型" 
        ReadOnly="true" 
        Visible="True" ItemStyle-Width="12%" HeaderStyle-HorizontalAlign="left"/>                  
<asp:TemplateColumn HeaderText="案件日期" 
        Visible="true" ItemStyle-Width="12%" HeaderStyle-HorizontalAlign="left">
  <ItemTemplate>
    <div align="left"><%# showdate(DataSet2.FieldValue("case_start", Container)) %> </div>
  </ItemTemplate>
</asp:TemplateColumn >
                  <asp:BoundColumn DataField="s01name" 
        HeaderText="承辦人" 
        ReadOnly="true" 
        Visible="True" ItemStyle-Width="10%" HeaderStyle-HorizontalAlign="left"/>                  
                  <asp:BoundColumn DataField="case_content" 
        HeaderText="案件內容" 
        ReadOnly="true" 
        Visible="True" ItemStyle-Width="50%" HeaderStyle-HorizontalAlign="left"/>                  
                  <asp:BoundColumn DataField="case_days" 
        HeaderText="已辦天數" 
        ReadOnly="true" 
        Visible="True" ItemStyle-Width="9%" HeaderStyle-HorizontalAlign="left"/>                  
                  <asp:BoundColumn DataField="case_audit" 
        HeaderText="逾辦" 
        ReadOnly="true" 
        Visible="True" ItemStyle-ForeColor="#CC3300" ItemStyle-Width="7%" HeaderStyle-HorizontalAlign="left"/>                  
                  <asp:BoundColumn DataField="admin_id" 
        HeaderText="admin_id" 
        ReadOnly="true" 
        Visible="false"/>                  
</Columns>
  </asp:DataGrid>
  </td>
   </tr>
 </table>
</form>
</body>
</html>
