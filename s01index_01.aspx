<%@ Page Language="VB" ContentType="text/html"%>
<%@ Register TagPrefix="MM" Namespace="DreamweaverCtrls" Assembly="DreamweaverCtrls,version=1.0.0.0,publicKeyToken=836f606ede05d46a,culture=neutral" %>
<MM:DataSet 
id="DataSet1"
runat="Server"
IsStoredProcedure="false"
ConnectionString='<%# System.Configuration.ConfigurationSettings.AppSettings("MM_CONNECTION_STRING_s01") %>'
DatabaseType='<%# System.Configuration.ConfigurationSettings.AppSettings("MM_CONNECTION_DATABASETYPE_s01") %>'
CommandText='<%# "SELECT * FROM s01major WHERE s01group like @s01group and s01title LIKE @s01filter or s01group like @s01group and s01con like @s01filter or s01group like @s01group and s01name like @s01filter or s01group like @s01group and s01time like @s01filter or s01group like @s01group and s01data01 like @s01filter or s01group like @s01group and s01data02 like @s01filter or s01group like @s01group and s01data03 like @s01filter or s01group like @s01group and s01data04 like @s01filter or s01group like @s01group and s01data05 like @s01filter ORDER BY s01time DESC" %>'
PageSize="50"
Debug="true"
><Parameters>

  <Parameter  Name="@s01group"  Value='<%# "%" + s01group.selecteditem.value + "%" %>'  Type="WChar"   />
  <Parameter  Name="@s01filter"  Value='<%# "%" + s01filter.text + "%" %>'  Type="WChar"   />
  <Parameter  Name="@s01oder"  Value='<%= s01oder.selecteditem.value %>'  Type="WChar"   />
  
</Parameters></MM:DataSet>
<MM:PageBind runat="server" PostBackBind="true" />
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <meta http-equiv="Content-Language" content="zh-tw">
  <meta http-equiv="Content-Type" content="text/html; charset=UTF-8">
  <title>全部訊息</title>
  <style type="text/css">
<!--
a {
	font-family: 新細明體;
	font-size: 14px;
	text-decoration: none;
}
a:visited {
	text-decoration: none;
	color: #0000FF;
}
a:active {
	text-decoration: none;
	color: #000000;
}
a:hover {
	font-family: "新細明體";
	font-size: 14px;
	font-style: normal;
	color: #CC0000;
	text-decoration: underline;
}
body {
	font-family: "新細明體";
	font-size: 14px;
	font-style: normal;
	background-color: #FFFFCC;
}
.style3 {
	font-size: 16px;
	font-weight: bold;
	color: #CC6600;
}
.style4 {
	font-size: 24px;
	font-weight: bold;
	color: #FF3300;
}
.style6 {color: #990000}
.style11 {font-size: 18px}
.style12 {font-size: 16px}
br {
	line-height: 20px;
}
ul {
	line-height: 18px;
	list-style-position: outside;
	list-style-type: square;
	text-align: left;
	display: list-item;
	color: #FF3333;
	margin-left: 16px;
}
.ul01 {
	color: #993366;
	display: list-item;
	list-style-position: outside;
	list-style-type: disc;
}

-->
  </style>
</head>
<body>
<script language="VB" runat="server">
   Sub Page_Load(sender As Object, e As EventArgs) 
	end sub

 function shownew(s01time)
 Dim diff = DateDiff("d",s01time,Now)  
 if diff < 7
 shownew = "<img border=0 src='images/new5.gif'>" 
 end if
 end function
 function showdate(s01time)
 if s01time <> ""
 showdate = FormatDateTime(s01time, DateFormat.ShortDate)
 end if
 end function
 function showpoint(s01data)
 if s01data <> ""
 showpoint = "<ul>"
 end if
 end function
  function showbr(s01data)
 if s01data <> ""
 showbr = "</ul>"
 end if
 end function

 Sub groupchange(sender As Object, e As EventArgs) 
	  s01filter.text = ""
 End Sub

</script>
<form name="filter" id="filter" runat="server">
    
  <table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#ECE9D8">
    <tr>
      <td width="20%"><span class="style4"><%= s01group.selecteditem.text %></span>	  </td>
	  <td colspan="3" width="80%">
	  <span class="style3">選擇業務屬性:
        <asp:DropDownList AutoPostBack="true" ID="s01group" runat="server" >
    <asp:ListItem Value=""  Selected="true">全部訊息</asp:ListItem>
    <asp:ListItem Value="業務統計">業務統計</asp:ListItem>
    <asp:ListItem Value="公文傳閱">公文傳閱</asp:ListItem>
    <asp:ListItem Value="組務文件">組務文件</asp:ListItem>
    <asp:ListItem Value="常用函稿">常用函稿</asp:ListItem>
    <asp:ListItem Value="為民服務文件">為民服務文件</asp:ListItem>
    <asp:ListItem Value="留言板">留言板</asp:ListItem>
  </asp:DropDownList>，搜尋字串:
        <asp:TextBox AutoPostBack="true" ID="s01filter" runat="server" />        
<asp:Button ID="Button1" runat="server" Text="查詢" />        </span></td>
    </tr>

    <tr>
      <td width="20%"><div align="left"></div></td>
      <td width="20%"><div align="left"></div></td>
      <td width="20%"><div align="left"><a href="s01login.aspx" class="style12">新增訊息</a></div></td>
      <td width="40%"><div align="left"><a href="s01changelogin.aspx" class="style12">修改訊息</a></div></td>
    </tr>
</table>	
  <span class="style11">於<span class="style6"><%= s01group.selecteditem.text %></span>搜尋<span class="style6"><%= s01filter.text %></span>共有<span class="style6"><%= DataSet1.RecordCount %></span>筆</span>
<asp:DataGrid 
  AllowCustomPaging="true" 
  AllowPaging="true" 
  AllowSorting="true" AlternatingItemStyle-BackColor="#FFCCFF" 
  AutoGenerateColumns="false" 
  CellPadding="3" 
  CellSpacing="0" 
  DataSource="<%# DataSet1.DefaultView %>" EditItemStyle-BackColor="#FFFFFF" HeaderStyle-BackColor="#CCCCCC" ID="DataGrid1" ItemStyle-BackColor="#FFCCCC" PagerStyle-BackColor="#CCCCFF" 
  PagerStyle-Mode="NumericPages" 
  PageSize="<%# DataSet1.PageSize %>" 
  runat="server" SelectedItemStyle-BackColor="#FFFFFF" SelectedItemStyle-ForeColor="#CC3300" 
  ShowFooter="false" 
  ShowHeader="true"
  Width="100%"  
  OnPageIndexChanged="DataSet1.OnDataGridPageIndexChanged" 
  virtualitemcount="<%# DataSet1.RecordCount %>" 
>
    <headerstyle HorizontalAlign="left" BackColor="#FFCCFF" ForeColor="#000000" Font-Size="smaller" />  
    <itemstyle BackColor="#FFFFFF" Font-Name="Verdana, Arial, Helvetica, sans-serif" Font-Size="smaller" />  
    <alternatingitemstyle BackColor="#FFCCFF" Font-Name="Verdana, Arial, Helvetica, sans-serif" Font-Size="smaller" />  
    <footerstyle HorizontalAlign="center" BackColor="#E8EBFD" ForeColor="#3D3DB6" Font-Name="Verdana, Arial, Helvetica, sans-serif" Font-Bold="true" Font-Size="smaller" />  
    <pagerstyle BackColor="white" Font-Name="Verdana, Arial, Helvetica, sans-serif" Font-Size="smaller"/>  
    <columns>
    <asp:TemplateColumn HeaderText="訊息主題" 
        Visible="True" ItemStyle-Width = "30%">
      <itemtemplate><ul class="ul01"><a href="s01detail_01.aspx?s01id=<%# dataset1.FieldValue("s01id", Container) %>" target="win_t"><%# dataset1.FieldValue("s01title", Container) %><%# shownew(dataset1.FieldValue("s01time", Container)) %></a></ul></itemtemplate>
    </asp:TemplateColumn>
    <asp:BoundColumn DataField="s01group" 
        HeaderText="業務屬性" 
        ReadOnly="true" 
        Visible="True" ItemStyle-Width = "11%"/>  
    <asp:TemplateColumn HeaderText="附檔" 
        Visible="True" ItemStyle-Width = "35%">
      <itemtemplate><%# showpoint(DataSet1.FieldValue("s01data01", Container)) %><a href="s01data_01/<%# DataSet1.FieldValue("s01data01", Container) %>" target="win3"><%# DataSet1.FieldValue("s01data01", Container) %></a><%# showbr(DataSet1.FieldValue("s01data01", Container)) %>
  <%# showpoint(DataSet1.FieldValue("s01data02", Container)) %><a href="s01data_01/<%# DataSet1.FieldValue("s01data02", Container) %>" target="win3"><%# DataSet1.FieldValue("s01data02", Container) %></a><%# showbr(DataSet1.FieldValue("s01data02", Container)) %>
   <%# showpoint(DataSet1.FieldValue("s01data03", Container)) %><a href="s01data_01/<%# DataSet1.FieldValue("s01data03", Container) %>" target="win3"><%# DataSet1.FieldValue("s01data03", Container) %></a><%# showbr(DataSet1.FieldValue("s01data03", Container)) %>
    <%# showpoint(DataSet1.FieldValue("s01data04", Container)) %><a href="s01data_01/<%# DataSet1.FieldValue("s01data04", Container) %>" target="win3"><%# DataSet1.FieldValue("s01data04", Container) %></a><%# showbr(DataSet1.FieldValue("s01data04", Container)) %>
	 <%# showpoint(DataSet1.FieldValue("s01data05", Container)) %><a href="s01data_01/<%# DataSet1.FieldValue("s01data05", Container) %>" target="win3"><%# DataSet1.FieldValue("s01data05", Container) %></a></itemtemplate>
    </asp:TemplateColumn>
    <asp:BoundColumn DataField="s01name" 
        HeaderText="發訊者" 
        ReadOnly="true" 
        Visible="True"
		ItemStyle-Width = "10%" HeaderStyle-HorizontalAlign="left"/>  
    <asp:TemplateColumn HeaderText="發訊時間" ItemStyle-Width="14%" HeaderStyle-HorizontalAlign="left" 
        Visible="True">
      <itemtemplate><%# showdate(dataset1.FieldValue("s01time", Container)) %></itemtemplate>
    </asp:TemplateColumn>
    </columns>
  </asp:DataGrid>
</form>
</body>
</html>
