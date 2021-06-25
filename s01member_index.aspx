<%@ Page Language="VB" ContentType="text/html" ResponseEncoding="utf-8" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<%@ Register TagPrefix="MM" Namespace="DreamweaverCtrls" Assembly="DreamweaverCtrls,version=1.0.0.0,publicKeyToken=836f606ede05d46a,culture=neutral" %>
<MM:DataSet 
id="DataSet1"
runat="Server"
IsStoredProcedure="false"
ConnectionString='<%# System.Configuration.ConfigurationSettings.AppSettings("MM_CONNECTION_STRING_s01") %>'
DatabaseType='<%# System.Configuration.ConfigurationSettings.AppSettings("MM_CONNECTION_DATABASETYPE_s01") %>'
CommandText='<%# "SELECT * FROM s01admin ORDER BY admin_id ASC" %>'
Debug="true"
></MM:DataSet>
<MM:PageBind runat="server" PostBackBind="true" />
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=UTF-8" />
<title>組員管理首頁</title>
<script language="vb" runat="server">
 dim url, url_back
  Sub page_load(sender As Object, e As EventArgs)
		If Session("MM_chief_num")="" then
			Response.Write("您還沒有登入呢！，請點擊")
			Response.Write("<a href=s01dellogin.aspx>登入</a>")
			Response.End()
        else
		  dim msg as integer
		  msg=request ("msg")
		   select case msg
		     case 1
			  msg01.text = "新增成功"
			 case 2
			  msg01.text = "修改成功"
			 case else
			  msg01.text = ""
			end select
		End If
	End Sub
</script>
<style type="text/css">
<!--
.style1 {
	color: #993300;
	font-weight: bold;
	font-size: large;
}
-->
</style>
<script type="text/JavaScript">
<!--
function MM_goToURL() { //v3.0
  var i, args=MM_goToURL.arguments; document.MM_returnValue = false;
  for (i=0; i<(args.length-1); i+=2) eval(args[i]+".location='"+args[i+1]+"'");
}
//-->
</script>
</head>
<body>
<p class="style1">組員管理</p>
<form runat="server">
  <p>
    <asp:Label BackColor="#ECE9D8" BorderColor="#99FF00" Font-Bold="true" Font-Size="16" ForeColor="#CC9900" ID="msg01" runat="server" />
</p>
  <p>&nbsp;
    <asp:DataGrid AllowPaging="false" 
  AllowSorting="False" AlternatingItemStyle-BackColor="#FFFF66" AlternatingItemStyle-BorderColor="#ECE9D8" 
  AutoGenerateColumns="false" BackColor="#FFCCFF" 
  CellPadding="3" 
  CellSpacing="0" 
  DataSource="<%# DataSet1.DefaultView %>" EditItemStyle-BackColor="#CCFFFF" EditItemStyle-ForeColor="#ECE9D8" id="DataGrid1" 
  runat="server" 
  ShowFooter="false" 
  ShowHeader="true"
  Width="100%" 
>
      <HeaderStyle HorizontalAlign="left" BackColor="#E8EBFD" ForeColor="#3D3DB6" Font-Name="Verdana, Arial, Helvetica, sans-serif" Font-Bold="true" Font-Size="smaller" />      
<ItemStyle BackColor="#FFCCCC" Font-Name="Verdana, Arial, Helvetica, sans-serif" Font-Size="smaller" />      
<AlternatingItemStyle BackColor="#CCFF99" Font-Name="Verdana, Arial, Helvetica, sans-serif" Font-Size="smaller" />      
<FooterStyle HorizontalAlign="center" BackColor="#E8EBFD" ForeColor="#3D3DB6" Font-Name="Verdana, Arial, Helvetica, sans-serif" Font-Bold="true" Font-Size="smaller" />
<PagerStyle BackColor="white" Font-Name="Verdana, Arial, Helvetica, sans-serif" Font-Size="smaller" />
<Columns>
<asp:BoundColumn DataField="admin_id" 
        HeaderText="id" 
        ReadOnly="true" 
        Visible="True"/>
<asp:HyperLinkColumn 
        HeaderText="姓名" 
        Visible="True"
        DataTextField="s01name"
        DataNavigateUrlField="admin_id"
        DataNavigateUrlFormatString="s01member_update.aspx?admin_id={0}"/>
<asp:BoundColumn DataField="admin_username" 
        HeaderText="帳號" 
        ReadOnly="true" 
        Visible="True"/>
<asp:BoundColumn DataField="admin_password" 
        HeaderText="密碼" 
        ReadOnly="true" 
        Visible="True"/>
<asp:BoundColumn DataField="admin_yes" 
        HeaderText="有效" 
        ReadOnly="true" 
        Visible="True"/>
</Columns>
    </asp:DataGrid>
</p>
</form>
<form id="form1" name="form1" method="post" action="">
  <p>
    <input name="Submit" type="button" onclick="MM_goToURL('parent','s01member_add.aspx');return document.MM_returnValue" value="新增組員" /></p>
  </form>
</body>

</html>
