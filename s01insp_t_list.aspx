<%@ Page Language="VB" ContentType="text/html" %>
<%@ Register TagPrefix="MM" Namespace="DreamweaverCtrls" Assembly="DreamweaverCtrls,version=1.0.0.0,publicKeyToken=836f606ede05d46a,culture=neutral" %>
<MM:DataSet 
id="DataSet1"
runat="Server"
IsStoredProcedure="false"
ConnectionString='<%# System.Configuration.ConfigurationSettings.AppSettings("MM_CONNECTION_STRING_s01") %>'
DatabaseType='<%# System.Configuration.ConfigurationSettings.AppSettings("MM_CONNECTION_DATABASETYPE_s01") %>'
CommandText='<%# "SELECT * FROM s01_total WHERE t_pre like ?" %>'
Debug="true" PageSize="30"
>
  <Parameters>
<Parameter  Name="@t_pre"  Value='<%# IIf((Session("MM_p_num") <> Nothing), Session("MM_p_num"), "") %>'  Type="WChar"   /></Parameters></MM:DataSet>
<MM:PageBind runat="server" PostBackBind="true" />

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<script runat="server">
	Sub Page_Load(Src As Object, E As EventArgs)
		If Session("MM_username")="" then
			Response.Write("您還沒有登入呢！，請點擊")
			Response.Write("<a href=s01insp_list_login.aspx>登入</a>")
			Response.End()
		End If
	End Sub
</script>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>專案查詢結果</title><!--參數應照順序 %為萬用字元 like 前後 + "%"-->
<style type="text/css">
<!--
.style1 {color: #3333FF}
-->
</style>
</head>
<body>
<form runat="server">
  <p>共找到<span class="style1"><%= DataSet1.RecordCount %></span> 筆資料，請點選欲修改之單位。
    <asp:Button ID="Button1" runat="server" Text="回新增事業單位名冊" OnClick="starsearch" />    
<a href="s01insp_logout.aspx" target="_self">登出</a></p>
  <p>
    <asp:DataGrid 
  AllowCustomPaging="true" 
  AllowPaging="true" 
  AllowSorting="False" 
  AutoGenerateColumns="false" 
  CellPadding="3" 
  CellSpacing="0" 
  DataSource="<%# DataSet1.DefaultView %>" id="DataGrid1" 
  PagerStyle-Mode="NumericPages" 
  PageSize="<%# DataSet1.PageSize %>" 
  runat="server" 
  ShowFooter="false" 
  ShowHeader="true" Width="100%" 
  OnPageIndexChanged="DataSet1.OnDataGridPageIndexChanged" 
  VirtualItemCount="<%# DataSet1.RecordCount %>" 
>
      <HeaderStyle HorizontalAlign="center" BackColor="#E8EBFD" ForeColor="#3D3DB6" Font-Name="Verdana, Arial, Helvetica, sans-serif" Font-Bold="true" Font-Size="smaller" />    
      <ItemStyle BackColor="#F2F2F2" Font-Name="Verdana, Arial, Helvetica, sans-serif" Font-Size="smaller" />    
      <AlternatingItemStyle BackColor="#E5E5E5" Font-Name="Verdana, Arial, Helvetica, sans-serif" Font-Size="smaller" />    
      <FooterStyle HorizontalAlign="center" BackColor="#E8EBFD" ForeColor="#3D3DB6" Font-Name="Verdana, Arial, Helvetica, sans-serif" Font-Bold="true" Font-Size="smaller" />    
      <PagerStyle BackColor="white" Font-Name="Verdana, Arial, Helvetica, sans-serif" Font-Size="smaller" />    
      <Columns>
      <asp:BoundColumn DataField="t_insu" 
        HeaderText="勞保證號" 
        ReadOnly="true" 
        Visible="True"/>      
      <asp:BoundColumn DataField="t_pre" 
        HeaderText="統一編號" 
        ReadOnly="true" 
        Visible="True"/>      
<asp:HyperLinkColumn
        DataNavigateUrlField="t_id" DataNavigateUrlFormatString="s01insp_t_update.aspx?t_id={0}"
        DataTextField="t_name" 
        Visible="True" 
        HeaderText="名稱"/>      
<asp:BoundColumn DataField="t_address" 
        HeaderText="地址" 
        ReadOnly="true" 
        Visible="True"/>
<asp:BoundColumn DataField="t_tel" 
        HeaderText="電話" 
        ReadOnly="true" 
        Visible="True"/>
<asp:BoundColumn DataField="t_id" 
        HeaderText="id" 
        ReadOnly="true" 
        Visible="false"/>
</Columns>
    </asp:DataGrid>
</p>
</form>
<%
  session("cancel_insp") = ""
  session("cancel_insp") = request.url.tostring()
%>

</body>
</html>
<script Language="VB" runat="server">
   Sub starsearch(sender As Object, e As EventArgs) 
	  dim url
	  url = "s01total_add_search.aspx"
	  Response.Redirect( url ) ' 使用Server.Transfer亦可
   End Sub

</script>