<%@ Page Language="VB" ContentType="text/html" ResponseEncoding="utf-8" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<%@ Register TagPrefix="MM" Namespace="DreamweaverCtrls" Assembly="DreamweaverCtrls,version=1.0.0.0,publicKeyToken=836f606ede05d46a,culture=neutral" %>
<MM:DataSet 
id="DataSet1"
runat="Server"
IsStoredProcedure="false"
ConnectionString='<%# System.Configuration.ConfigurationSettings.AppSettings("MM_CONNECTION_STRING_s01") %>'
DatabaseType='<%# System.Configuration.ConfigurationSettings.AppSettings("MM_CONNECTION_DATABASETYPE_s01") %>'
CommandText='<%# "SELECT * FROM s01admin WHERE admin_yes = yes ORDER BY admin_username ASC" %>'
Debug="true"
>
</MM:DataSet>
<MM:DataSet 
id="DataSet2"
runat="Server"
IsStoredProcedure="false"
ConnectionString='<%# System.Configuration.ConfigurationSettings.AppSettings("MM_CONNECTION_STRING_s01") %>'
DatabaseType='<%# System.Configuration.ConfigurationSettings.AppSettings("MM_CONNECTION_DATABASETYPE_s01") %>'
CommandText='<%# "SELECT case_content FROM s01_case WHERE case_id = ?" %>'
Debug="true"
>
  <Parameters>
    <Parameter  Name="@case_id"  Value='<%# IIf((Request.QueryString("case_id") <> Nothing), Request.QueryString("case_id"), "") %>'  Type="Integer"   />
  </Parameters>
</MM:DataSet>
<MM:PageBind runat="server" PostBackBind="false" />
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Data.OleDb" %>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=UTF-8" />
<title>新增案件</title>
<script language="vb" runat="server">
 dim url, url_back
  Sub page_load(sender As Object, e As EventArgs)
		If Session("MM_chief_num")="" then
			Response.Write("您還沒有登入呢！，請點擊")
			Response.Write("<a href=s01dellogin.aspx>登入</a>")
			Response.End()
		End If

  if not Ispostback Then
    case_start.text = request("case_start")
	if case_start.text = "" then
      response.Redirect("s01case_admin.aspx")
    else
	back.text = session("cancel_case_admin")
	session("cancel_case_admin") = ""
	 end if
	end if
  end sub

 		Sub Date_Selecteden(sender As Object, e As EventArgs)
		  dim diff
		  diff = datediff("d" , showdate(case_start.text) , showdate(Calendaren.SelectedDate.ToShortDateString))
		  if diff < 0 then
		   msgen.text = "結束日期小於啟始日期.請重新點選"
		  else 
		  case_limit.Text = diff
		  msgen.text = ""
          end if
		End Sub
  Sub InsertData(sender As Object, e As EventArgs) 
   ch01.enabled = true
   ch02.enabled = true
   ch03.enabled = true
   ch04.enabled = true

    if IsValid Then

      Dim Conn As OleDbConnection
      Dim Cmd  As OleDbCommand

      Dim Provider = "Provider=Microsoft.Jet.OLEDB.4.0"
      Dim Database = "Data Source=" & Server.MapPath( "result/result.mdb" )
      Conn = New OleDbConnection( Provider & ";" & DataBase )
      Conn.Open()

      Dim SQL As String
      SQL = "Insert Into s01case (case_start, admin_id, case_group, case_content, case_result, case_yes, case_limit) Values(@case_start, @admin_id, @case_group, @case_content, @case_result, @case_yes, @case_limit)"
      Cmd = New OleDbCommand( SQL, Conn )

      Cmd.Parameters.Add( New OleDbParameter("@case_start", OleDbType.Char, 16))
      Cmd.Parameters.Add( New OleDbParameter("@admin_id", OleDbType.single))
      Cmd.Parameters.Add( New OleDbParameter("@case_group", OleDbType.Char, 20))
      Cmd.Parameters.Add( New OleDbParameter("@case_content", OleDbType.VarChar))
      Cmd.Parameters.Add( New OleDbParameter("@case_result", OleDbType.VarChar))
      Cmd.Parameters.Add( New OleDbParameter("@case_yes", OleDbType.Char, 6))
      Cmd.Parameters.Add( New OleDbParameter("@case_limit", OleDbType.single))

      Cmd.Parameters("@case_start").value = case_start.text
      Cmd.Parameters("@admin_id").value = case_man.selecteditem.value
      Cmd.Parameters("@case_group").value = case_group.text
      Cmd.Parameters("@case_content").value = case_content.text
      Cmd.Parameters("@case_result").value = case_result.text
      Cmd.Parameters("@case_yes").value = case_yes.text
      Cmd.Parameters("@case_limit").value = val(case_limit.text)

      Cmd.ExecuteNonQuery()
      If Err.Number <> 0 Then
         Msg.Text = Err.Description
      Else
		 url = "s01case_admin.aspx?case_add=1"
      End If

      Conn.Close()
      response.Redirect("s01case_admin.aspx?case_add=1")
	  End If
   End Sub

function stra1(t01)
 if t01 <> ""  
  dim i as integer
  dim str2 as string
  dim v
   v = split(t01,",")
   str2 = "" 
	for i = 0 to ubound(v)
	str2 &= chr(v(i))
	next
   stra1 = str2
  end if
end function   	

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

Sub startcancel(sender As Object, e As EventArgs) 
 ch01.Enabled="false"
 ch02.Enabled="false"

 response.Redirect(back.text)
end sub

 function showcontent(s01content)
  if s01content <> "" then
   showcontent = s01content
  else
  dim content_chk as string
  content_chk = stra1(request("case_content"))
  showcontent = content_chk
  end if
 end function
</script>
<script language = "JavaScript">
<!--

function w_back()
{
location.href = "<%= url_back %>"
}
//-->
</Script>

<style type="text/css">
<!--
.style1 {
	font-size: 16px;
	font-weight: bold;
}
.style2 {
	color: #FF0000;
	font-weight: bold;
}
.style3 {	color: #CC3300;
	font-weight: bold;
	font-size: 16px;
}
-->
</style>
</head>
<body>
<p class="style1">新增案件：</p>
<form runat="server" name='form_add' id="form_add">
<table width="529" border="0" cellspacing="1" cellpadding="0">
  <tr>
    <td width="80" bgcolor="#CCFFFF">案件日期：</td>
    <td width="446" bgcolor="#CCCC66"><asp:TextBox ID="case_start" ReadOnly="true" runat="server" Wrap="false" /></td>
  </tr>
  <tr>
    <td width="80" bgcolor="#CCFFFF">承辦人：</td>
    <td bgcolor="#CCCC66"><asp:DropDownList DataSource="<%# DataSet1.DefaultView %>" DataTextField="s01name" DataValueField="admin_id" ID="case_man" runat="server">
    </asp:DropDownList></td>
  </tr>
  <tr>
    <td width="80" bgcolor="#CCFFFF">案件種類：</td>
    <td bgcolor="#CCCC66"><asp:DropDownList ID="case_group" runat="server">
      <asp:ListItem>申訴</asp:ListItem>
      <asp:ListItem>重大職災</asp:ListItem>
      <asp:ListItem>重大職災(其他)</asp:ListItem>
      <asp:ListItem>重大職災(勞安法以外)</asp:ListItem>
      <asp:ListItem>重大職災(認定中)</asp:ListItem>
      <asp:ListItem>非重大職災</asp:ListItem>
      <asp:ListItem>專案</asp:ListItem>
      <asp:ListItem>其他</asp:ListItem>
    </asp:DropDownList></td>
  </tr>
    <tr>
    <td width="80" bgcolor="#CCFFFF">限辦天數：</td>
    <td bgcolor="#CCCC66">
            
      限辦天數或點選期限：
        <asp:TextBox Columns="8" ID="case_limit" runat="server" AutoPostBack="true" />
      
  <asp:RequiredFieldValidator ControlToValidate="case_limit" Display="Dynamic" Enabled="true" ErrorMessage="尚未填寫限辦天數" ID="ch01" runat="server" />
  <asp:RangeValidator ControlToValidate="case_limit" Display="Dynamic" Enabled="true" ErrorMessage="" ID="ch03" MaximumValue="999" MinimumValue="0" runat="server" Text="應填數字" type="double" />
  <br />
  <asp:Label Font-Bold="true" forecolor="red" ID="msgen" runat="server" />
  <br />
  <asp:Calendar
            DayHeaderStyle-BackColor="#ffcccc" Enabled=""
            Font-Name="Arial" Font-Size="12px"
            Height="180px" id=Calendaren
            OtherMonthDayStyle-ForeColor="#ffffcc" runat="server"
            SelectedDayStyle-BackColor="Navy"
            SelectedDayStyle-Font-Bold="True"
            SelectorStyle-BackColor="gainsboro"
            TitleStyle-BackColor="#cccc66"
            TitleStyle-Font-Bold="True"
            TitleStyle-Font-Size="12px"
            TodayDayStyle-BackColor="#ffcc33" Visible="true" Width="200px"
            onselectionchanged="Date_Selecteden"
            /></td></tr>
  <tr>
    <td width="80" bgcolor="#CCFFFF">內容：</td>
    <td bgcolor="#CCCC66"><asp:TextBox Columns="30" ID="case_content" Rows="5" runat="server" text='<%# showcontent(DataSet2.FieldValue("case_content", Container)) %>' TextMode="MultiLine" Width="300" />
    <asp:RequiredFieldValidator ControlToValidate="case_content" Enabled="true" ErrorMessage="尚未填寫內容" ID="ch02" runat="server" /></td>
    </tr>
  <tr>
    <td width="80" bgcolor="#CCFFFF">結果及說明：</td>
    <td bgcolor="#CCCC66"><asp:TextBox Columns="30" ID="case_result" Rows="5" runat="server" TextMode="MultiLine" Width="300" />
   </td>
    </tr>

  <tr>
    <td bgcolor="#CCFFFF">已登錄：</td>
    <td bgcolor="#CCCC66"><asp:DropDownList ID="case_yes" runat="server">
      <asp:listitem>是</asp:listitem>
      <asp:listitem Selected="true">否</asp:listitem>
      <asp:listitem>不需</asp:listitem>
	</asp:DropDownList></td>
  </tr>
</table>
  <asp:Button ID="Button1" runat="server" Text="新增" OnClick="InsertData" />
  <span class="style1">
  <input type="reset" name="Submit2" value="重新填寫" />
  </span>
  <span class="style3">
  <asp:Button ID="Buttonc" runat="server" Text="取消" OnClick='startcancel' />  </span>
  <asp:TextBox ID="back" runat="server" Visible="false" Columns="50" />
  
  <p>
  <HR><asp:ValidationSummary
     DisplayMode="BulletList" EnableClientScript="true" Enabled="false"
     HeaderText="必須輸入的欄位還有:" ID="ch04" runat="server" />
<asp:Label runat="server" id="Msg" ForeColor="Red" />
</form>
</body>
</html>