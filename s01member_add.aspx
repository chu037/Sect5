<%@ Page Language="VB" ContentType="text/html" ResponseEncoding="utf-8" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Data.OleDb" %>
<head>
<meta http-equiv="Content-Type" content="text/html; charset="utf-8"/>
<title>新增行程</title>
<script language="vb" runat="server">
 dim url, url_back
  Sub page_load(sender As Object, e As EventArgs)
		If Session("MM_chief_num")="" then
			Response.Write("您還沒有登入呢！，請點擊")
			Response.Write("<a href=s01dellogin.aspx>登入</a>")
			Response.End()
	    else
		   url_back="s01delindex.aspx"
		End If
	End Sub
  Sub InsertData(sender As Object, e As EventArgs) 
    if IsValid Then

      Dim Conn As OleDbConnection
      Dim Cmd  As OleDbCommand

      Dim Provider = "Provider=Microsoft.Jet.OLEDB.4.0"
      Dim Database = "Data Source=" & Server.MapPath( "result/result.mdb" )
      Conn = New OleDbConnection( Provider & ";" & DataBase )
      Conn.Open()

      Dim SQL As String
      SQL = "Insert Into s01admin (s01name, admin_username, admin_password, admin_phone, admin_cell, admin_mail, admin_yes) Values(@s01name, @admin_username, @admin_password, @admin_phone, @admin_cell, @admin_mail, @admin_yes)"
      Cmd = New OleDbCommand( SQL, Conn )

      Cmd.Parameters.Add( New OleDbParameter("@s01name", OleDbType.Char, 16))
      Cmd.Parameters.Add( New OleDbParameter("@admin_username", OleDbType.Char, 16))
      Cmd.Parameters.Add( New OleDbParameter("@admin_password", OleDbType.Char, 16))
      Cmd.Parameters.Add( New OleDbParameter("@admin_phone", OleDbType.Char, 20))
      Cmd.Parameters.Add( New OleDbParameter("@admin_cell", OleDbType.Char, 20))
      Cmd.Parameters.Add( New OleDbParameter("@admin_mail", OleDbType.Char, 50))
      Cmd.Parameters.Add( New OleDbParameter("@admin_yes", OleDbType.boolean))

      Cmd.Parameters("@s01name").value = cstr(s01name.text)
      Cmd.Parameters("@admin_username").value = cstr(admin_username.text)
      Cmd.Parameters("@admin_password").value = cstr(admin_password.text)
      Cmd.Parameters("@admin_phone").value = cstr(admin_phone.text)
      Cmd.Parameters("@admin_cell").value = cstr(admin_cell.text)
      Cmd.Parameters("@admin_mail").value = cstr(admin_mail.text)
      Cmd.Parameters("@admin_yes").value = cbool(admin_yes.checked)

      Cmd.ExecuteNonQuery()
      If Err.Number <> 0 Then
         Msg.Text = Err.Description
      Else
		 url = "s01member_index.aspx?msg=1"
      End If

      Conn.Close()
      End If
   End Sub

</script>
<script language = "JavaScript">
<!--
function w_back(uback)
{
parent.location.href =  uback; 
} 
//-->
</Script>

<style type="text/css">
<!--
.style1 {
	font-size: 16px;
	font-weight: bold;
}
-->
</style>
</head>
<body>
<p class="style1">新增組員：</p>
<form runat="server">
<table width="512" border="0" cellspacing="1" cellpadding="0">
  <tr>
    <td width="86" bgcolor="#99FF99">姓名：</td>
    <td width="411" bgcolor="#99FFFF"><asp:TextBox ID="s01name" runat="server" />
      <asp:RequiredFieldValidator ControlToValidate="s01name" Display="Dynamic" ErrorMessage="尚未填寫姓名" runat="server" /></td>
  </tr>
  <tr>
    <td width="86" bgcolor="#99FF99">帳號：</td>
    <td bgcolor="#99FFFF"><asp:TextBox ID="admin_username" runat="server" />
      <asp:RequiredFieldValidator ControlToValidate="admin_username" Display="Dynamic" ErrorMessage="尚未填寫帳號" runat="server" /></td>
  </tr>
  <tr>
    <td width="86" bgcolor="#99FF99">密碼：</td>
    <td bgcolor="#99FFFF">
      <asp:TextBox ID="admin_password" runat="server" />
      <asp:RequiredFieldValidator ControlToValidate="admin_password" Display="Dynamic" ErrorMessage="尚未填寫密碼" runat="server" /></td>
  </tr>
  <tr>
    <td width="86" bgcolor="#99FF99">分機：</td>
    <td bgcolor="#99FFFF">
      <asp:TextBox ID="admin_phone" runat="server" />
      <asp:RequiredFieldValidator ControlToValidate="admin_phone" Display="Dynamic" ErrorMessage="尚未填寫分機" runat="server" /></td>
  </tr>
  <tr>
    <td width="86" bgcolor="#99FF99">手機：</td>
    <td bgcolor="#99FFFF">
      <asp:TextBox ID="admin_cell" runat="server" />
      <asp:RequiredFieldValidator ControlToValidate="admin_cell" Display="Dynamic" ErrorMessage="尚未填寫手機" runat="server" /></td>
  </tr>
  <tr>
    <td width="86" bgcolor="#99FF99">電郵：</td>
    <td bgcolor="#99FFFF">
      <asp:TextBox ID="admin_mail" runat="server" />
      <asp:RequiredFieldValidator ControlToValidate="admin_mail" Display="Dynamic" ErrorMessage="尚未填寫mail" runat="server" /></td>
  </tr>

  <tr>
    <td width="86" bgcolor="#99FF99">有效：</td>
    <td bgcolor="#99FFFF"><asp:CheckBox ID="admin_yes" runat="server" Checked="true" Text='yes' /></td>
  </tr>
  <tr>
    <td width="86" bgcolor="#99FF99">內容：</td>
    <td bgcolor="#99FFFF">&nbsp;</td>
  </tr>
</table>
  <asp:Button ID="Button1" runat="server" Text="新增" OnClick="InsertData" />
  <span class="style1">
  <input type="reset" name="Submit2" value="重新填寫" />
  <input type="button" onclick="w_back('s01member_index.aspx')" value="取消" />
  </span>
  <p>
  <HR><asp:ValidationSummary runat="server"
     HeaderText="必須輸入的欄位還有:"
     DisplayMode="BulletList" />
<asp:Label runat="server" id="Msg" ForeColor="Red" />
</form>
<%
if url <> "" then
%>
<script language = "JavaScript">
<!--
//將已經上傳的檔案名稱傳回到新增照片頁面的Photo_Picture
//隱藏欄位，用於交檔案名稱儲存到資料庫中。
location.href="<% =url %>";
//-->
</Script>
<%
End If
%>

</body>
</html>
