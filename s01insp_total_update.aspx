<%@ Page Language="VB" ContentType="text/html" ResponseEncoding="UTF-8" %>

<%@ Import Namespace="System.Data.OleDb" %>
<%@ Register TagPrefix="MM" Namespace="DreamweaverCtrls" Assembly="DreamweaverCtrls,version=1.0.0.0,publicKeyToken=836f606ede05d46a,culture=neutral" %>
<MM:DataSet 
id="DataSet1"
runat="Server"
IsStoredProcedure="false"
ConnectionString='<%# System.Configuration.ConfigurationSettings.AppSettings("MM_CONNECTION_STRING_s01") %>'
DatabaseType='<%# System.Configuration.ConfigurationSettings.AppSettings("MM_CONNECTION_DATABASETYPE_s01") %>'
CommandText='<%# "SELECT * FROM s01_total WHERE t_id = ?" %>'
Debug="true"
>
  <Parameters>
    <Parameter  Name="@t_id"  Value='<%# IIf((Request.QueryString("t_id") <> Nothing), Request.QueryString("t_id"), "") %>'  Type="Integer"   />
  </Parameters>
</MM:DataSet>
<MM:PageBind runat="server" PostBackBind="true" />
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<script Language="VB" runat="server">
   Sub page_load(sender As Object, e As EventArgs) 
		If Session("MM_username")="" then
			Response.Write("您還沒有登入呢！，請點擊")
			Response.Write("<a href=s01insp_list_login.aspx>登入</a>")
			Response.End()
		End If
if not ispostback then
p_sh_g.selecteditem.text = DataSet1.FieldValue("p_sh_g", nothing)
p_sh_org.selecteditem.text = DataSet1.FieldValue("p_sh_org", nothing)
p_sh_man.selecteditem.text = DataSet1.FieldValue("p_sh_man", nothing)
p_sh_top.selecteditem.text = DataSet1.FieldValue("p_sh_org", nothing)
p_sh_sec.selecteditem.text = DataSet1.FieldValue("p_sh_sec", nothing)
else
exit sub
end if
p_keyin.text = session("MM_s01name") 
   End Sub

   Sub UpdateData(sender As Object, e As EventArgs) 
    if IsValid Then

      Dim Conn As OleDbConnection
      Dim Cmd  As OleDbCommand

      Dim Provider = "Provider=Microsoft.Jet.OLEDB.4.0"
      Dim Database = "Data Source=" & Server.MapPath( "result/result.mdb" )
      Conn = New OleDbConnection( Provider & ";" & DataBase )
      Conn.Open()

      Dim SQL As String
      '以下是為了無結束日期及時間時.要區分的更新碼
	  if case_end.text <> "" then
	  SQL = "UPDATE s01case SET t_name=@t_name, t_address=@t_address, t_g_num=@t_g_num, t_email=@t_email, t_boss=@t_boss, t_per=@t_per, t_sh_g=@t_sh_g, t_sh_org=@t_sh_org, t_sh_man=@t_sh_man, t_sh_top=@t_sh_top, t_sh_sec=@t_sh_sec, t_place=@t_place, t_sh=@t_sh, t_shcell=@t_shcell, t_hazard=@t_hazard, t_ps=@t_ps WHERE t_id=@t_id"
      Cmd = New OleDbCommand( SQL, Conn )

      Cmd.Parameters.Add( New OleDbParameter("@t_name", OleDbType.char,20))
      Cmd.Parameters.Add( New OleDbParameter("@t_address", OleDbType.char,80))
      Cmd.Parameters.Add( New OleDbParameter("@t_g_num", OleDbType.char,12))
      Cmd.Parameters.Add( New OleDbParameter("@t_email", OleDbType.char,80))
      Cmd.Parameters.Add( New OleDbParameter("@t_boss", OleDbType.char,20))
      Cmd.Parameters.Add( New OleDbParameter("@t_per", OleDbType.integer))
      Cmd.Parameters.Add( New OleDbParameter("@t_sh_g", OleDbType.char,20))
      Cmd.Parameters.Add( New OleDbParameter("@t_sh_org", OleDbType.char,20))
      Cmd.Parameters.Add( New OleDbParameter("@t_sh_man", OleDbType.char,20))
      Cmd.Parameters.Add( New OleDbParameter("@t_sh_top", OleDbType.char,80))
      Cmd.Parameters.Add( New OleDbParameter("@t_address", OleDbType.char,80))
      Cmd.Parameters.Add( New OleDbParameter("@t_address", OleDbType.char,80))
      Cmd.Parameters.Add( New OleDbParameter("@t_address", OleDbType.char,80))
      Cmd.Parameters.Add( New OleDbParameter("@t_address", OleDbType.char,80))
      Cmd.Parameters.Add( New OleDbParameter("@t_address", OleDbType.char,80))

      Cmd.Parameters.Add( New OleDbParameter("@t_address", OleDbType.char,80))

      Cmd.Parameters.Add( New OleDbParameter("@case_start", OleDbType.dbdate))
      Cmd.Parameters.Add( New OleDbParameter("@case_content", OleDbType.varChar))
      Cmd.Parameters.Add( New OleDbParameter("@case_result", OleDbType.varChar))
      Cmd.Parameters.Add( New OleDbParameter("@case_yes", OleDbType.char,6))
      Cmd.Parameters.Add( New OleDbParameter("@case_limit", OleDbType.single))
      Cmd.Parameters.Add( New OleDbParameter("@case_group", OleDbType.char,20))
      Cmd.Parameters.Add( New OleDbParameter("@case_end", OleDbType.dbdate))
      Cmd.Parameters.Add( New OleDbParameter("@admin_id", OleDbType.integer))
      Cmd.Parameters.Add( New OleDbParameter("@case_id", OleDbType.integer))

      Cmd.Parameters("@case_start").value = case_start.text
      Cmd.Parameters("@case_content").value = case_content.text
      Cmd.Parameters("@case_result").value = case_result.text
      Cmd.Parameters("@case_yes").value = case_yes.selecteditem.text
      Cmd.Parameters("@case_limit").value = val(case_limit.text)
      Cmd.Parameters("@case_group").value = case_group.selecteditem.text
      Cmd.Parameters("@case_end").value = case_end.text
      Cmd.Parameters("@admin_id").value = case_man.selecteditem.value
      Cmd.Parameters("@case_id").value = val(case_id.text)

      Cmd.ExecuteNonQuery()
      'If Err.Number <> 0 Then
         'Msg.Text = Err.Description
      'Else
		  url = back.text
		  response.Redirect(url)
      'End If

      Conn.Close()
  end sub
 function showdate()
 showdate = FormatDateTime(now)
 end function
</script>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=UTF-8" />
<title>工廠資料</title>
<style type="text/css">
<!--
.style2 {font-size: 12px}
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

<form name='form1' method='POST' runat='server'>
<p>修改事業單位資料：
  <asp:Button ID="total_update" runat="server" Text="更新" OnClick="updatedata" />  
<input type="reset" name="Submit5" value="重新填寫" />
  <input name="Submit6" type="button" onclick="MM_goToURL('parent','s01insp_detail_index.aspx?p_num=<%# DataSet1.FieldValue("p_num", Container) %>');return document.MM_returnValue" value="取消更新" /></br>
<table width="100%" border="1" cellspacing="0" cellpadding="0">
  <tr>
    <td width="10%" bgcolor="#66FFCC">單位名稱：</td>
    <td width="47%" bgcolor="#CCFF99"><asp:TextBox Height="" ID="t_name" runat="server" Text='<%# DataSet1.FieldValue("t_name", Container) %>' Columns="50" /></td>
    <td width="12%" bgcolor="#66FFCC">負責人：</td>
    <td width="31%" bgcolor="#CCFF99"><asp:TextBox ID="t_boss" runat="server" Text='<%# DataSet1.FieldValue("t_boss", Container) %>' Columns="20" /></td>
  </tr>
  <tr>
    <td bgcolor="#66FFCC">地址：</td>
    <td bgcolor="#CCFF99"><asp:TextBox ID="t_address" runat="server" Text='<%# DataSet1.FieldValue("t_address", Container) %>' Columns="50" /></td>
    <td bgcolor="#66FFCC">電話：</td>
    <td bgcolor="#CCFF99"><asp:TextBox ID="t_tel" runat="server" Text='<%# DataSet1.FieldValue("t_tel", Container) %>' Columns="20" /></td>
  </tr>
  <tr>
    <td bgcolor="#66FFCC">行業別：</td>
    <td bgcolor="#CCFF99"><asp:TextBox ID="t_g_name" runat="server" Text='<%# DataSet1.FieldValue("t_group.t_g_name", Container) %>' Columns="20" />
	<asp:TextBox Columns="10" ID="p_keyin" ReadOnly="true" runat="server" Visible="true" />      
      <asp:TextBox Columns="20" ID="p_update_time" ReadOnly="true" runat="server" Text='<%# showdate() %>' Visible="true" />      
    </span></td>
    <td bgcolor="#66FFCC">傳真：</td>
    <td bgcolor="#CCFF99"></td>
  </tr>
  <tr>
    <td bgcolor="#66FFCC">e-mail：</td>
    <td bgcolor="#CCFF99"><asp:TextBox ID="t_email" runat="server" Text='<%# DataSet1.FieldValue("t_email", Container) %>' Columns="50" /></td>
    <td bgcolor="#66FFCC">勞工人數：</td>
    <td bgcolor="#CCFF99"><asp:TextBox ID="t_per" runat="server" Text='<%# DataSet1.FieldValue("t_per", Container) %>' Columns="10" /></td>
  </tr>
    <td bgcolor="#FFCC99">勞安單位人員設置：</td>
      <td colspan="3" bgcolor="#FFCC99">分類：
        <asp:DropDownList ID="t_sh_g" runat="server">
	  <asp:ListItem></asp:ListItem>
	  <asp:ListItem>第一類</asp:ListItem>
	  <asp:ListItem>第二類</asp:ListItem>
	  <asp:ListItem>第三類</asp:ListItem>
	  </asp:DropDownList>
	  勞安單位：<asp:DropDownList ID="t_sh_org" runat="server">
	  <asp:ListItem></asp:ListItem>
	  <asp:ListItem>符合規定</asp:ListItem>
	  <asp:ListItem>不符規定</asp:ListItem>
	  <asp:ListItem>不適用</asp:ListItem>
	  </asp:DropDownList>
	  勞安人員：<asp:DropDownList ID="t_sh_man" runat="server">
	  <asp:ListItem></asp:ListItem>
	  <asp:ListItem>符合規定</asp:ListItem>
	  <asp:ListItem>不符規定</asp:ListItem>
	  </asp:DropDownList>
	 總機構：<asp:DropDownList ID="t_sh_top" runat="server">
	  <asp:ListItem></asp:ListItem>
	  <asp:ListItem>符合規定</asp:ListItem>
	  <asp:ListItem>不符規定</asp:ListItem>
	  <asp:ListItem>不適用</asp:ListItem>
	  </asp:DropDownList>
	  製造一級單位：
	  <asp:DropDownList ID="t_sh_sec" runat="server">
	  <asp:ListItem></asp:ListItem>
	  <asp:ListItem>符合規定</asp:ListItem>
	  <asp:ListItem>不符規定</asp:ListItem>
	  <asp:ListItem>不適用</asp:ListItem>
	  </asp:DropDownList>  	  </td>
  </tr>
  <tr>
    <td height="52" bgcolor="#66FFCC">原物料：</td>
    <td bgcolor="#CCFF99"></td>
    <td bgcolor="#66FFCC">連絡人</td>
    <td bgcolor="#CCFF99">姓名：
      <asp:TextBox ID="p_sh" runat="server" Text='<%# DataSet1.FieldValue("t_sh", Container) %>' Columns="20" />
      <br />
      手機：
      <asp:TextBox ID="p_shcell" runat="server" Text='<%# DataSet1.FieldValue("t_shcell", Container) %>' Columns="20" /></td>
  </tr>
  <tr>
    <td bgcolor="#66FFCC">產品：</td>
    <td bgcolor="#CCFF99"></td>
    <td bgcolor="#66FFCC">危害物：</td>
    <td bgcolor="#CCFF99"><textarea name="t_hazard" cols="20" rows="2" id="textarea2"><%# DataSet1.FieldValue("t_hazard", Container) %></textarea></td>
  </tr>
  <tr>
    <td bgcolor="#66FFCC">受檢場所：</td>
    <td bgcolor="#CCFF99"><textarea name="t_place" cols="30" rows="2" id="t_place"><%# DataSet1.FieldValue("t_place", Container) %></textarea></td>
    <td bgcolor="#66FFCC">備註：</td>
    <td bgcolor="#CCFF99"><textarea name="textarea" cols="20" rows="2" id="t_ps"><%# DataSet1.FieldValue("t_ps", Container) %></textarea></td>
  </tr>
</table>
  <input type="submit" name="Submit" value="更新" />
  <input type="reset" name="Submit2" value="重新填寫" />
  <input name="Submit3" type="button" onclick="MM_goToURL('parent','s01insp_total_detail.aspx?t_id=<%# DataSet1.FieldValue("t_id", Container) %>');return document.MM_returnValue" value="取消更新" />
  <asp:TextBox ID="t_num" ReadOnly="true" runat="server" Text='<%# DataSet1.FieldValue("t_id", Container) %>' />
<input type="hidden" name="MM_update" value="form1">
  
</form>
<p>&nbsp;</p>
</body>
</html>
