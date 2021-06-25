<%@ Page Language="VB" ContentType="text/html" ResponseEncoding="utf-8" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Data.OleDb" %>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>新增行程</title>
<script language="vb" runat="server">
 dim url, url_back
  Sub page_load(sender As Object, e As EventArgs)
  if not Ispostback Then
    m_date.text = request("m_date")
	else 
	m_date.text = m_date.text
	end if
	if m_date.text = "" then
      response.Redirect("kc_cal.aspx")
    else
	 url_back = "kc_cal_selday.aspx?m_date=" & m_date.text
   end if
  end sub
  Sub change_text1(sender As Object, e As EventArgs) 
    m_time.text = m_time_h.selecteditem.text & ":" & m_time_n.selecteditem.text 
   End Sub
   Sub change_text2(sender As Object, e As EventArgs) 
    if m_date_en.text <> nothing
	m_time_en.text = m_time_enh.selecteditem.text & ":" & m_time_enn.selecteditem.text
	msg02.text = ""
	else 
	msg02.text = "尚未點選結束日期"
	end if 
   End Sub
  Sub del_text2(sender As Object, e As EventArgs) 
    m_date_en.text = "" 
   End Sub
  Sub call_cal(sender As Object, e As EventArgs) 
   Calendar1.visible = true
   End Sub

  Sub InsertData(sender As Object, e As EventArgs) 
    if IsValid Then
      Dim Conn As OleDbConnection
      Dim Cmd  As OleDbCommand

      Dim Provider = "Provider=Microsoft.Jet.OLEDB.4.0"
            Dim Database = "Data Source=" & Server.MapPath("\result\result.mdb")
      Conn = New OleDbConnection( Provider & ";" & DataBase )
      Conn.Open()
  if m_date_en.text <> nothing 
   if m_time_en.text = nothing then
    m_time_en.text = "17:30"
    end if
      Dim SQL As String
      SQL = "Insert Into kc_calen (m_date, m_time, m_hours, m_note, m_messeng, m_date_en, m_time_en) Values(@m_date, @m_time, @m_hours, @m_note, @m_messeng, @m_date_en,@m_time_en)"
      Cmd = New OleDbCommand( SQL, Conn )

      Cmd.Parameters.Add( New OleDbParameter("@m_date", OleDbType.Char, 16))
      Cmd.Parameters.Add( New OleDbParameter("@m_time", OleDbType.Char, 10))
      Cmd.Parameters.Add( New OleDbParameter("@m_hours", OleDbType.single))
      Cmd.Parameters.Add( New OleDbParameter("@m_note", OleDbType.VarChar))
      Cmd.Parameters.Add( New OleDbParameter("@m_messeng", OleDbType.Char, 6))
      Cmd.Parameters.Add( New OleDbParameter("@m_date_en", OleDbType.Char, 16))
      Cmd.Parameters.Add( New OleDbParameter("@m_time_en", OleDbType.Char, 10))

      Cmd.Parameters("@m_date").value = m_date.text
      Cmd.Parameters("@m_time").value = m_time.text
      Cmd.Parameters("@m_hours").value = val(m_hours.text)
      Cmd.Parameters("@m_note").value = m_note.text
      Cmd.Parameters("@m_messeng").value = m_messeng.text
      Cmd.Parameters("@m_date_en").value = cstr(m_date_en.text)
      Cmd.Parameters("@m_time_en").value = cstr(m_time_en.text)

      Cmd.ExecuteNonQuery()
      If Err.Number <> 0 Then
         Msg.Text = Err.Description
      Else
		 url = "kc_cal_selday.aspx?m_date=" & m_date.text & "&m_add=1"
      End If

      Conn.Close()
else
      Dim SQL As String
      SQL = "Insert Into kc_calen (m_date, m_time, m_hours, m_note, m_messeng, m_detail) Values(@m_date, @m_time, @m_hours, @m_note, @m_messeng, @m_detail)"
      Cmd = New OleDbCommand( SQL, Conn )

      Cmd.Parameters.Add( New OleDbParameter("@m_date", OleDbType.Char, 16))
      Cmd.Parameters.Add( New OleDbParameter("@m_time", OleDbType.Char, 10))
      Cmd.Parameters.Add( New OleDbParameter("@m_hours", OleDbType.single))
      Cmd.Parameters.Add( New OleDbParameter("@m_note", OleDbType.Char, 80))
      Cmd.Parameters.Add( New OleDbParameter("@m_messeng", OleDbType.Char, 6))
      Cmd.Parameters.Add( New OleDbParameter("@m_detail", OleDbType.varchar))

      Cmd.Parameters("@m_date").value = m_date.text
      Cmd.Parameters("@m_time").value = m_time.text
      Cmd.Parameters("@m_hours").value = val(m_hours.text)
      Cmd.Parameters("@m_note").value = m_note.text
      Cmd.Parameters("@m_messeng").value = m_messeng.text
      Cmd.Parameters("@m_detail").value = m_detail.text

      Cmd.ExecuteNonQuery()
      If Err.Number <> 0 Then
         Msg.Text = Err.Description
      Else
		 url = "kc_cal_selday.aspx?m_date=" & m_date.text & "&m_add=1"
      End If

      Conn.Close()
   end if
      End If
   End Sub
		Sub Date_Selected(sender As Object, e As EventArgs)
          if datediff("d" , showdate(m_date.text) , showdate(Calendar1.SelectedDate.ToShortDateString)) < 0 then
		  m_date_en.Text = ""
		   msg01.text = "結束日期小於啟始日期"
		  else 
		  m_date_en.Text = Calendar1.SelectedDate.ToShortDateString
		  m_time_en.text = m_time_enh.selecteditem.text & ":" & m_time_enn.selecteditem.text
		  msg01.text = ""
          msg02.text = ""

          end if
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
</script>
<script language = "JavaScript">
<!--

function w_back()
{
location.href = "<% =url_back %>";
}
function MM_goToURL() { //v3.0
  var i, args=MM_goToURL.arguments; document.MM_returnValue = false;
  for (i=0; i<(args.length-1); i+=2) eval(args[i]+".location='"+args[i+1]+"'");
}

//-->
</Script>

<style type="text/css">
<!--
.style1 {
	font-size: 16px;
	font-weight: bold;
}
.style3 {color: #FF0000}
-->
</style>
</head>
<body>
<p class="style1">新增行程：</p>
<form runat="server" name='form_add' id="form_add">
<table width="529" border="0" cellspacing="1" cellpadding="0">
  <tr>
    <td width="80" bgcolor="#CC9966">啟始日期：</td>
    <td width="446" bgcolor="#CCCC66"><asp:TextBox ID="m_date" ReadOnly="true" runat="server" Wrap="false" /></td>
  </tr>
  <tr>
    <td width="80" bgcolor="#CC9966">啟始時間：</td>
    <td bgcolor="#CCCC66">
      <asp:DropDownList AutoPostBack="true" ID="m_time_h" runat="server" OnSelectedIndexChanged="change_text1">
    <asp:listitem>08</asp:listitem>
    <asp:listitem>09</asp:listitem>
    <asp:listitem>10</asp:listitem>
    <asp:listitem>11</asp:listitem>
    <asp:listitem>12</asp:listitem>
    <asp:listitem>13</asp:listitem>
    <asp:listitem>14</asp:listitem>
    <asp:listitem>15</asp:listitem>
    <asp:listitem>16</asp:listitem>
    <asp:listitem>17</asp:listitem>
    <asp:listitem>18</asp:listitem>
    <asp:listitem>19</asp:listitem>
    <asp:listitem>20</asp:listitem>
	</asp:DropDownList>
      時
      <asp:DropDownList ID="m_time_n" runat="server" OnSelectedIndexChanged="change_text1" AutoPostBack="true">
        <asp:listitem>00</asp:listitem>
        <asp:listitem>10</asp:listitem>
        <asp:listitem>20</asp:listitem>
        <asp:listitem>30</asp:listitem>
        <asp:listitem>40</asp:listitem>
        <asp:listitem>50</asp:listitem>
      </asp:DropDownList>
    分   
    <asp:TextBox ID="m_time" ReadOnly="true" runat="server" Text='08:00' /></td>
  </tr>
  <tr>
    <td width="80" bgcolor="#CC9966">預定時數：</td>
    <td bgcolor="#CCCC66">
      <asp:TextBox ID="m_hours" MaxLength="5" Rows="5" runat="server" Width="50" />    
      小時    <asp:RequiredFieldValidator ControlToValidate="m_hours" Display="Dynamic" ErrorMessage="尚未填寫預定時數" runat="server" /><asp:RangeValidator ControlToValidate="m_hours" Display="Dynamic" ErrorMessage="" MaximumValue="999" MinimumValue="0" runat="server" Text="預定時數0~999間" type="double" /></td>
  </tr>
  <tr>
    <td width="80" bgcolor="#CC9966">行程：</td>
    <td bgcolor="#CCCC66"><asp:DropDownList ID="m_messeng" runat="server">
      <asp:listitem>開會</asp:listitem>
      <asp:listitem>宣導</asp:listitem>
      <asp:listitem>上課</asp:listitem>
      <asp:listitem>請假</asp:listitem>
      <asp:listitem>換班</asp:listitem>
      <asp:listitem>其他</asp:listitem>
    </asp:DropDownList></td>
  </tr>
  <tr>
    <td width="80" bgcolor="#CC9966">主題：</td>
    <td bgcolor="#CCCC66"><asp:TextBox Columns="30" ID="m_note" Rows="3" runat="server" TextMode="MultiLine" Width="300" />
      <asp:RequiredFieldValidator ControlToValidate="m_note" ErrorMessage="尚未填寫主題" runat="server" /></td>
  </tr>
  <tr>
    <td width="80" bgcolor="#CC9966">內容：</td>
    <td bgcolor="#CCCC66"><asp:TextBox Columns="30" ID="m_detail" Rows="5" runat="server" TextMode="MultiLine" Width="300" />
  </tr>

  <tr>
    <td width="80" bgcolor="#CC9966">結束日期：</td>
    <td width="446" bgcolor="#CCCC66">
      <span class="style3">跨日行程請點選:</span>
        <asp:Button ID="Button2" runat="server" Text="結束日期" OnClick="call_cal" />
      <br />
      <asp:Calendar
            DayHeaderStyle-BackColor="#ffcccc" DayNameFormat="shortest"
            Font-Name="Arial" Font-Size="12px"
            Height="180px" id=Calendar1
            OtherMonthDayStyle-ForeColor="#ffffcc" runat="server"
            SelectedDayStyle-BackColor="Navy"
            SelectedDayStyle-Font-Bold="True"
            SelectorStyle-BackColor="gainsboro"
            TitleStyle-BackColor="#cccc66"
            TitleStyle-Font-Bold="True"
            TitleStyle-Font-Size="12px"
            TodayDayStyle-BackColor="#ffcc33" Visible="false" Width="200px"
            onselectionchanged="Date_Selected"
            />      
    <br/>
      <asp:TextBox ID="m_date_en" ReadOnly="true" runat="server" Wrap="false" />
      <asp:Label Font-Bold="true" forecolor="red" ID="msg01" runat="server" /></td>
  </tr>
  <tr>
    <td width="80" bgcolor="#CC9966">結束時間：</td>
    <td bgcolor="#CCCC66">
      <asp:DropDownList AutoPostBack="true" ID="m_time_enh" runat="server" OnSelectedIndexChanged="change_text2">
    <asp:listitem>08</asp:listitem>
    <asp:listitem>09</asp:listitem>
    <asp:listitem>10</asp:listitem>
    <asp:listitem>11</asp:listitem>
    <asp:listitem>12</asp:listitem>
    <asp:listitem>13</asp:listitem>
    <asp:listitem>14</asp:listitem>
    <asp:listitem>15</asp:listitem>
    <asp:listitem>16</asp:listitem>
    <asp:listitem>17</asp:listitem>
    <asp:listitem>18</asp:listitem>
    <asp:listitem>19</asp:listitem>
    <asp:listitem>20</asp:listitem>

	</asp:DropDownList>
      時
      <asp:DropDownList ID="m_time_enn" runat="server" OnSelectedIndexChanged="change_text2" AutoPostBack="true">
        <asp:listitem>00</asp:listitem>
        <asp:listitem>10</asp:listitem>
        <asp:listitem>20</asp:listitem>
        <asp:listitem>30</asp:listitem>
        <asp:listitem>40</asp:listitem>
        <asp:listitem>50</asp:listitem>
      </asp:DropDownList>
    分   
    <asp:TextBox ID="m_time_en" ReadOnly="true" runat="server" Text="" />
    <asp:Label Font-Bold="true" ForeColor="#FF0000" ID="msg02" runat="server" font-color="red" /></td>
  </tr>
</table>
  <asp:Button ID="Button1" runat="server" Text="新增" OnClick="InsertData" />
  <span class="style1">
  <input type="reset" name="Submit2" value="重新填寫" />
  <input name="Submit" type="button" onclick="w_back()" value="取消" />
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
