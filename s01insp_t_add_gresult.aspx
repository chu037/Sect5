<%@ Page Language="VB" ContentType="text/html" %>
<%@ Register TagPrefix="MM" Namespace="DreamweaverCtrls" Assembly="DreamweaverCtrls,version=1.0.0.0,publicKeyToken=836f606ede05d46a,culture=neutral" %>
<MM:DataSet 
id="DataSet1"
runat="Server"
IsStoredProcedure="false"
ConnectionString='<%# System.Configuration.ConfigurationSettings.AppSettings("MM_CONNECTION_STRING_s01") %>'
DatabaseType='<%# System.Configuration.ConfigurationSettings.AppSettings("MM_CONNECTION_DATABASETYPE_s01") %>'
CommandText='<%# "SELECT * FROM t_group where t_g_name like ? and t_g_num <> ? ORDER BY t_g_id ASC" %>'
Debug="true"
><Parameters>
<Parameter  Name="@t_g_name"  Value='<%# "%"+IIf((Request.QueryString("t_g_name") <> Nothing), Request.QueryString("t_g_name"), "")+"%" %>'  Type="WChar"   />
<Parameter  Name="@t_g_num"  Value='<%# "" %>'  Type="WChar"   />

</Parameters></MM:DataSet>
<MM:DataSet 
id="DataSet2"
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
<title>工廠查詢結果</title><!--參數應照順序 %為萬用字元 like 前後 + "%"-->
<style type="text/css">
<!--
.style2 {
	color: #CC6600;
	font-weight: bold;
}
.style3 {
	color: #CC3300;
	font-weight: bold;
	font-size: 14px;
}
-->
</style>
</head>
<body>
<form runat="server">
  <p>共找到 <span class="style2"><%= DataSet1.RecordCount %></span>筆資料<br />
        <span class="style3">請點選行業別</span></p>
  <p>
    <asp:DataGrid 
  AllowPaging="false" 
  AllowSorting="true" 
  AutoGenerateColumns="false" 
  CellPadding="3" 
  CellSpacing="0" 
  DataSource="<%# DataSet1.DefaultView %>" Enabled="true" id="DataGrid1" 
  runat="server" 
  ShowFooter="false" 
  ShowHeader="true" Width="100%" 
>
      <HeaderStyle HorizontalAlign="left" BackColor="#E8EBFD" ForeColor="#3D3DB6" Font-Name="Verdana, Arial, Helvetica, sans-serif" Font-Bold="true" Font-Size="smaller" />      
<ItemStyle BackColor="#F2F2F2" Font-Name="Verdana, Arial, Helvetica, sans-serif" Font-Size="smaller" />      
<AlternatingItemStyle BackColor="#E5E5E5" Font-Name="Verdana, Arial, Helvetica, sans-serif" Font-Size="smaller" />      
<FooterStyle HorizontalAlign="left" BackColor="#E8EBFD" ForeColor="#3D3DB6" Font-Name="Verdana, Arial, Helvetica, sans-serif" Font-Bold="true" Font-Size="smaller" />      
<PagerStyle BackColor="white" Font-Name="Verdana, Arial, Helvetica, sans-serif" Font-Size="smaller" />      
<Columns>
<asp:HyperLinkColumn DataNavigateUrlField="t_g_id" DataNavigateUrlFormatString="s01insp_t_add_gresult.aspx?t_g_id={0}"
	    DataTextField="t_g_name" 
        Visible="True" target="_self" 
        HeaderText="行業別名稱" 
		/>
<asp:BoundColumn DataField="t_g_1" 
        HeaderText="行業別(一)" 
        ReadOnly="true" 
        Visible="True"/>
<asp:BoundColumn DataField="t_g_2" 
        HeaderText="行業別(二)" 
        ReadOnly="true" 
        Visible="True"/>
<asp:BoundColumn DataField="t_g_3" 
        HeaderText="行業別(三)" 
        ReadOnly="true" 
        Visible="True"/>
<asp:BoundColumn DataField="t_g_num" 
        HeaderText="行業別(四)" 
        ReadOnly="true" 
        Visible="True"/>
</Columns>
    </asp:DataGrid>
    <br />
  <asp:TextBox ID="t_g_id" ReadOnly="true" runat="server" text='<%#DataSet2.FieldValue("t_g_id", Container) %>' Visible="true" />  
  <asp:TextBox ID="t_g_1" ReadOnly="true" runat="server" text='<%# DataSet2.FieldValue("t_g_1", Container) %>' Visible="true" />  
  <asp:TextBox ID="t_g_2" ReadOnly="true" runat="server" text='<%# DataSet2.FieldValue("t_g_2", Container) %>' Visible="true" />    
  <asp:TextBox ID="t_g_3" ReadOnly="true" runat="server" text='<%# DataSet2.FieldValue("t_g_3", Container) %>' Visible="true" />    
  <asp:TextBox ID="t_g_num" ReadOnly="true" runat="server" text='<%# DataSet2.FieldValue("t_g_num", Container) %>' Visible="true" />  
  <asp:TextBox ID="t_g_name" ReadOnly="true" runat="server" text='<%# DataSet2.FieldValue("t_g_name", Container) %>' Visible="false" />  

  </p>
  <p>
    <input name="Submit" type="button" onclick="history.back()" value="取消" />
  </p>
</form>
</body>
</html>
<script Language="VB" runat="server">
	Sub Page_Load(Src As Object, E As EventArgs)
End Sub
	function comb1(t01)
  if t01<>"" then
  comb1 = "查詢:" & t01 & "?"
  end if
End function

</script>
<%
if t_g_id.text <> nothing then
%>
<script language = "JavaScript">
<!--

opener.document.form1.t_g_num.value ="<% =t_g_num.text %>";
opener.document.form1.t_g_name.value ="<% =t_g_name.text %>";
<!--window.opener.document.getElementById('Button01').click();


<!--window.opener.document.location.reload();
<!--document.getElementById("DropDownlist1").onchange();觸發.net下拉控制項

window.close();
//-->

</script>
<%
end if
%>

