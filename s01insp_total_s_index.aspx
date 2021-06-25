<%@ Page Language="VB" ContentType="text/html" ResponseEncoding="utf-8" %>
<%@ Register TagPrefix="MM" Namespace="DreamweaverCtrls" Assembly="DreamweaverCtrls,version=1.0.0.0,publicKeyToken=836f606ede05d46a,culture=neutral" %>
<MM:DataSet 
id="DataSet1"
runat="Server"
IsStoredProcedure="false"
ConnectionString='<%# System.Configuration.ConfigurationSettings.AppSettings("MM_CONNECTION_STRING_s01") %>'
DatabaseType='<%# System.Configuration.ConfigurationSettings.AppSettings("MM_CONNECTION_DATABASETYPE_s01") %>'
CommandText='<%# "SELECT * FROM s01_insp_s ORDER BY s01name ASC" %>'
Debug="true"
></MM:DataSet>
<MM:PageBind runat="server" PostBackBind="true" />
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>綜合行業科轄區一覽表</title>
<script language="vb" runat="server">
  Sub page_load(sender As Object, e As EventArgs)
		If Session("MM_chief_num")="" then
			Response.Write("您還沒有登入呢！，請點擊")
			Response.Write("<a href=s01dellogin.aspx>登入</a>")
			Response.End()
		End If
  end sub

  Sub s01_total_add(sender As Object, e As EventArgs)
   response.Redirect("s01insp_total_s_add.aspx")
  end sub 
</script>  
<script type="text/JavaScript">
<!--
function MM_openBrWindow(theURL,winName,features) { //v2.0
  window.open(theURL,winName,features);
}

function MM_goToURL() { //v3.0
  var i, args=MM_goToURL.arguments; document.MM_returnValue = false;
  for (i=0; i<(args.length-1); i+=2) eval(args[i]+".location='"+args[i+1]+"'");
}
function DelTitle(s01id) {//確認是否要刪除該記錄
	if (confirm('您確實要刪除該主題嗎？' + '\r\r' +
		'注意：如果您刪除該主題，' + '\r' +
		'哪麼該主題下所有的資料' + '\r' +
		'就將全部被刪除！')){
	window.location ='s01insp_total_s_del.aspx?case_id=' + s01id;
    }
}
</script>
</head>
<body>
<form runat="server">
  <asp:DataGrid id="DataGrid1" 
  runat="server" 
  AllowSorting="False" 
  AutoGenerateColumns="false" 
  CellPadding="3" 
  CellSpacing="0" 
  ShowFooter="false" 
  ShowHeader="true" 
  DataSource="<%# DataSet1.DefaultView %>" 
  PagerStyle-Mode="NextPrev" 
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
      <asp:BoundColumn DataField="t_g_name" 
        HeaderText="負責行業" 
        ReadOnly="true" 
        Visible="True"/>
<asp:TemplateColumn HeaderText="修改" 
        Visible="True" ItemStyle-Width="16%">
  <ItemTemplate>
    <input name="Submit2" type="button" onclick="MM_goToURL('parent','s01insp_total_s_update.aspx?t_g_id=<%# DataSet1.FieldValue("t_g_id", Container) %>&insp_s_id=<%# DataSet1.FieldValue("insp_s_id", Container) %>');return document.MM_returnValue" value="修改" />
    <input name="Submit3" type="button" onClick="DelTitle('<%# DataSet1.FieldValue("insp_s_id", Container) %>')" value="刪除" />
  </ItemTemplate>
</asp:TemplateColumn>

    </Columns>
  </asp:DataGrid>
  <p>
    <asp:Button ID="insp_total_add" runat="server" Text="新增轄區" OnClick="s01_total_add" /></p>
</form>
</body>
</html>
