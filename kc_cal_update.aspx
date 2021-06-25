<%@ Page Language="VB" ContentType="text/html" EnableEventValidation="false"%>
<%@ Register TagPrefix="MM" Namespace="DreamweaverCtrls" Assembly="DreamweaverCtrls,version=1.0.0.0,publicKeyToken=836f606ede05d46a,culture=neutral" %>
<MM:DataSet 
id="DataSet1"
runat="Server"
IsStoredProcedure="false"
ConnectionString='<%# System.Configuration.ConfigurationSettings.AppSettings("MM_CONNECTION_STRING_s01") %>'
DatabaseType='<%# System.Configuration.ConfigurationSettings.AppSettings("MM_CONNECTION_DATABASETYPE_s01") %>'
CommandText='<%# "SELECT m_date, m_date_en, m_hours, m_messeng, m_note, m_num, m_time, m_time_en, m_detail FROM s01_kc_calen WHERE m_num = ?" %>'
Debug="true"
><Parameters>
<Parameter  Name="@m_num"  Value='<%# IIf((Request.QueryString("m_num") <> Nothing), Request.QueryString("m_num"), "") %>'  Type="Integer"   /></Parameters></MM:DataSet>
<MM:DataSet 
id="DataSet2"
runat="Server"
IsStoredProcedure="false"
ConnectionString='<%# System.Configuration.ConfigurationSettings.AppSettings("MM_CONNECTION_STRING_s01") %>'
DatabaseType='<%# System.Configuration.ConfigurationSettings.AppSettings("MM_CONNECTION_DATABASETYPE_s01") %>'
CommandText='<%# "SELECT * FROM s01_kc_calen_appen WHERE m_num = ? ORDER BY m_appen_id DESC" %>'
Debug="true"
>
  <Parameters>
    <Parameter  Name="@m_num"  Value='<%# IIf((Request.QueryString("m_num") <> Nothing), Request.QueryString("m_num"), "") %>'  Type="Integer"   />  
  </Parameters>
</MM:DataSet>
<MM:PageBind runat="server" PostBackBind="false" />
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Data.OleDb" %>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=UTF-8" />
<title>修改行事曆</title>
<script language="VB" runat="server">
 dim url, url_back
  Sub page_load(sender As Object, e As EventArgs)
	dim m_numcheck, t_now, url_now
	m_numcheck = request("m_num")
	t_now = request("t_now")
	if m_numcheck = "" then
      response.Redirect("kc_cal.aspx")
	 end if
	 if t_now = 1 then
	  url_now = "kc_cal_upfile.aspx?m_num=" & m_numcheck
	  response.Redirect(url_now)
	 end if
   if not ispostback
   m_messeng.SelectedIndex = m_messeng.Items.IndexOf(m_messeng.Items.FindByValue(DataSet1.FieldValue("m_messeng", Nothing) ))
   end if
  end sub
  Sub cancel_url(sender As Object, e As EventArgs)
      url_back = session("cancel")
	  response.Redirect(url_back)
	  end sub
  Sub change_textst(sender As Object, e As EventArgs) 
    m_time.text = m_time_h.selecteditem.text & ":" & m_time_n.selecteditem.text 
   End Sub
   Sub change_texten(sender As Object, e As EventArgs) 
    if m_date_en.text <> nothing
	m_time_en.text = m_time_enh.selecteditem.text & ":" & m_time_enn.selecteditem.text
	msgent.text = ""
	else 
	msgent.text = "請先點選結束日期"
	end if 
   End Sub
  Sub del_text2(sender As Object, e As EventArgs) 
    m_date_en.text = "" 
   End Sub
  Sub call_calst(sender As Object, e As EventArgs) 
   Calendarst.visible = true
   Calendaren.visible = false
	msgst.text = ""
	msgen.text = ""
   End Sub
  Sub call_calen(sender As Object, e As EventArgs) 
   Calendaren.visible = true
   Calendarst.visible = false
	msgst.text = ""
	msgen.text = ""
   End Sub
 Sub showgrid(sender As Object, e As EventArgs)
 if DataSet2.FieldValue("m_appen",nothing) = nothing then
 datagrid1.visible = false
 end if 
 end sub

 		Sub Date_Selecteden(sender As Object, e As EventArgs)
          if  m_date_en.Text <> "" then
		  if datediff("d" , showdate(m_date.text) , showdate(Calendaren.SelectedDate.ToShortDateString)) < 0 then
		   msgen.text = "結束日期小於啟始日期.請重新點選"
		  else 
		  m_date_en.Text = Calendaren.SelectedDate.ToShortDateString
		  msgen.text = ""
          end if
		  else
		  if datediff("d" , showdate(m_date.text) , showdate(Calendaren.SelectedDate.ToShortDateString)) < 0 then
		   msgen.text = "結束日期小於啟始日期.請重新點選"
		  else
		  m_date_en.Text = Calendaren.SelectedDate.ToShortDateString
		  m_time_en.text = m_time_enh.selecteditem.text & ":" & m_time_enn.selecteditem.text
		  msgen.text = ""
          end if
         end if
		End Sub  
 		Sub Date_Selectedst(sender As Object, e As EventArgs)
          if m_date_en.text <> "" then
		  if datediff("d" , showdate(Calendarst.SelectedDate.ToShortDateString) , showdate(m_date_en.text)) < 0 then
		   msgst.text = "結束日期小於啟始日期.請重新點選"
		  else 
		  m_date.Text = Calendarst.SelectedDate.ToShortDateString
		   msgst.text = ""
          end if
		   else
		  m_date.Text = Calendarst.SelectedDate.ToShortDateString
           msgst.text = ""
         end if	  
		End Sub 
  Sub UpdateData(sender As Object, e As EventArgs) 
    if IsValid Then

      Dim Conn As OleDbConnection
      Dim Cmd  As OleDbCommand

      Dim Provider = "Provider=Microsoft.Jet.OLEDB.4.0"
            Dim Database = "Data Source=" & Server.MapPath("result/result.mdb")
      Conn = New OleDbConnection( Provider & ";" & DataBase )
      Conn.Open()

      Dim SQL As String
      '以下是為了無結束日期及時間時.要區分的更新碼
	  if m_date_en.text <> "" then
	  SQL = "UPDATE kc_calen SET m_time=@m_time, m_time_en=@m_time_en, m_hours=@m_hours, m_messeng=@m_messeng, m_date=@m_date, m_date_en=@m_date_en, m_detail=@m_detail, m_note=@m_note WHERE m_num=@m_num"
      Cmd = New OleDbCommand( SQL, Conn )

      Cmd.Parameters.Add( New OleDbParameter("@m_time", OleDbType.Char,16))
      Cmd.Parameters.Add( New OleDbParameter("@m_time_en", OleDbType.Char,16))
      Cmd.Parameters.Add( New OleDbParameter("@m_hours", OleDbType.single))
      Cmd.Parameters.Add( New OleDbParameter("@m_messeng", OleDbType.Char,6))
      Cmd.Parameters.Add( New OleDbParameter("@m_date", OleDbType.dbdate))
      Cmd.Parameters.Add( New OleDbParameter("@m_date_en", OleDbType.dbdate))
      Cmd.Parameters.Add( New OleDbParameter("@m_detail", OleDbType.varchar))
      Cmd.Parameters.Add( New OleDbParameter("@m_note", OleDbType.Char,80))
      Cmd.Parameters.Add( New OleDbParameter("@m_num", OleDbType.integer))

      Cmd.Parameters("@m_time").value = m_time.text
      Cmd.Parameters("@m_time_en").value = m_time_en.text
      Cmd.Parameters("@m_hours").value = val(m_hours.text)
      Cmd.Parameters("@m_messeng").value = m_messeng.selecteditem.text
      Cmd.Parameters("@m_date").value = m_date.text
      Cmd.Parameters("@m_date_en").value = m_date_en.text
      Cmd.Parameters("@m_detail").value = m_detail.text
      Cmd.Parameters("@m_note").value = m_note.text
      Cmd.Parameters("@m_num").value = val(m_num.text)

      Cmd.ExecuteNonQuery()
      If Err.Number <> 0 Then
         Msg.Text = Err.Description
      Else
		 url = session("cancel") & "&m_add=2"
      End If

      Conn.Close()
 	   else
	  SQL = "UPDATE kc_calen SET m_time=@m_time, m_hours=@m_hours, m_messeng=@m_messeng, m_date=@m_date, m_note=@m_note, m_detail=@m_detail WHERE m_num=@m_num"
      Cmd = New OleDbCommand( SQL, Conn )

      Cmd.Parameters.Add( New OleDbParameter("@m_time", OleDbType.Char,16))
      'Cmd.Parameters.Add( New OleDbParameter("@m_time_en", OleDbType.Char,16))
      Cmd.Parameters.Add( New OleDbParameter("@m_hours", OleDbType.single))
      Cmd.Parameters.Add( New OleDbParameter("@m_messeng", OleDbType.Char,6))
      Cmd.Parameters.Add( New OleDbParameter("@m_date", OleDbType.dbdate))
      'Cmd.Parameters.Add( New OleDbParameter("@m_date_en", OleDbType.dbdate))
      Cmd.Parameters.Add( New OleDbParameter("@m_note", OleDbType.Char,80))
      Cmd.Parameters.Add( New OleDbParameter("@m_detail", OleDbType.varchar))
      Cmd.Parameters.Add( New OleDbParameter("@m_num", OleDbType.integer))

      Cmd.Parameters("@m_time").value = m_time.text
      'Cmd.Parameters("@m_time_en").value = m_time_en.text
      Cmd.Parameters("@m_hours").value = val(m_hours.text)
      Cmd.Parameters("@m_messeng").value = m_messeng.selecteditem.text
      Cmd.Parameters("@m_date").value = m_date.text
      'Cmd.Parameters("@m_date_en").value = m_date_en.text
      Cmd.Parameters("@m_note").value = m_note.text
      Cmd.Parameters("@m_detail").value = m_detail.text
      Cmd.Parameters("@m_num").value = val(m_num.text)

      Cmd.ExecuteNonQuery()
      If Err.Number <> 0 Then
         Msg.Text = Err.Description
      Else
		 url = session("cancel") & "&m_add=2"
      End If

      Conn.Close()
	   end if
      End If
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
 function showmessage(vt2,vt3)
 if vt2 <> ""
 showmessage = vt2 & "<br/>" & vt3 & "小時"
 end if
 end function
 sub trans_go(sender As Object, e As EventArgs)
 dim trans_text = "kc_cal_upfile.aspx?m_num=" & m_num.text
 if trans_text <> "" then
 response.Redirect(trans_text)
 end if 
 end sub

</script>
<script language = "JavaScript">
<!--

function w_back()
{
var x1;
x1 == document.form_update.m_date.value;
location.href = "kc_cal_selday.aspx?m_date=" & x1;
}

function Mcheck(){
	if (document.form_update.m_note.value=="") {
        window.alert("請輸入內容");
        return false }
    if (document.form_update.m_hours.value=="") {
        window.alert("請輸入預定時數");
        return false }
	if (isNaN(document.form_update.m_hours.value)) {
        window.alert("時數請輸入數值");
        return false }
	 return true;
}

function MM_callJS(jsStr) { //v2.0
  return eval(jsStr)
}
//-->
</Script>

<style type="text/css">
<!--
.style1 {
	color: #CC3300;
	font-weight: bold;
	font-size: 16px;
}
.style2 {color: #FF0000}
-->
</style>
</head>
<body>
<span class="style1">修改行程</span>
<form name='form_update' id="form_update" runat="server" >
<table width="551" border="0" cellpadding="0" cellspacing="1" bordercolor="#FFFFFF">
  <tr>
    <td width="86" bgcolor="#CCFFCC">啟始日期：</td>
    <td width="462" bgcolor="#FFCC99"><asp:Button ID="Button2" runat="server" Text="啟始日期" OnClick="call_calst" />
      <asp:TextBox AutoPostBack="true" Columns="9" ID="m_date" ReadOnly="true" runat="server" Text='<%# showdate(DataSet1.FieldValue("m_date", Container)) %>' /><br/>
      <asp:Label Font-Bold="true" forecolor="red" ID="msgst" runat="server" />
      <asp:TextBox AutoPostBack="true" Columns="6" ID="m_num" ReadOnly="true" runat="server" text='<%# DataSet1.FieldValue("m_num", Container) %>' Visible="false" />
      <br />
<asp:Calendar
            DayHeaderStyle-BackColor="#ffcccc" DayNameFormat="shortest"
            Font-Name="Arial" Font-Size="12px"
            Height="180px" id=Calendarst
            OtherMonthDayStyle-ForeColor="#ffffcc" runat="server"
            SelectedDayStyle-BackColor="Navy"
            SelectedDayStyle-Font-Bold="True"
            SelectorStyle-BackColor="gainsboro"
            TitleStyle-BackColor="#cccc66"
            TitleStyle-Font-Bold="True"
            TitleStyle-Font-Size="12px"
            TodayDayStyle-BackColor="#ffcc33" Visible="false" Width="200px"
            onselectionchanged="Date_Selectedst"
            /></td>
  </tr>
  <tr>
    <td width="80" bgcolor="#CCFFCC">啟始時間：</td>
    <td bgcolor="#FFCC99">
      <asp:DropDownList AutoPostBack="true" ID="m_time_h" runat="server" OnSelectedIndexChanged="change_textst">
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
      <asp:DropDownList ID="m_time_n" runat="server" OnSelectedIndexChanged="change_textst" AutoPostBack="true">
        <asp:listitem>00</asp:listitem>
        <asp:listitem>10</asp:listitem>
        <asp:listitem>20</asp:listitem>
        <asp:listitem>30</asp:listitem>
        <asp:listitem>40</asp:listitem>
        <asp:listitem>50</asp:listitem>
      </asp:DropDownList>
    分   
    <asp:TextBox AutoPostBack="true" Columns="6" ID="m_time" ReadOnly="true" runat="server" Text='<%# showtime(DataSet1.FieldValue("m_time", Container)) %>' /></td>
  </tr>

  <tr>
    <td width="86" bgcolor="#CCFFCC">預定時數：</td>
    <td bgcolor="#FFCC99">
 <asp:TextBox ID="m_hours" MaxLength="5" Rows="5" runat="server" text='<%# DataSet1.FieldValue("m_hours", Container) %>' Width="50"/>    
      小時    <asp:RequiredFieldValidator ControlToValidate="m_hours" Display="Dynamic" ErrorMessage="尚未填寫預定時數" runat="server" /><asp:RangeValidator ControlToValidate="m_hours" Display="Dynamic" ErrorMessage="" MaximumValue="999" MinimumValue="0" runat="server" Text="預定時數0~999間" type="double" />   </td>
  </tr>
  <tr>
    <td width="86" bgcolor="#CCFFCC">行程：</td>
    <td bgcolor="#FFCC99">
	<asp:DropDownList ID="m_messeng" runat="server">
        <asp:listitem value="開會">開會</asp:listitem>
        <asp:listitem value="宣導">宣導</asp:listitem>
        <asp:listitem value="上課">上課</asp:listitem>
        <asp:listitem value="請假">請假</asp:listitem>
        <asp:listitem value="換班">換班</asp:listitem>
        <asp:listitem value="其他">其他</asp:listitem>

	</asp:DropDownList> </td>
  </tr>
  <tr>
    <td bgcolor="#CCFFCC">結束日期：</td>
    <td bgcolor="#FFCC99">
<asp:Button ID="Button3" runat="server" Text="結束日期" OnClick="call_calen" />      
(<span class="style2">跨日行程請點選</span>)
<asp:TextBox AutoPostBack="true" Columns="9" ID="m_date_en" ReadOnly="true" runat="server" Text='<%# showdate(DataSet1.FieldValue("m_date_en", Container)) %>' Wrap="false" />
<br/>
<asp:Label Font-Bold="true" forecolor="red" ID="msgen" runat="server" />
<br/>
	  <asp:Calendar
            DayHeaderStyle-BackColor="#ffcccc" DayNameFormat="shortest"
            Font-Name="Arial" Font-Size="12px"
            Height="180px" id=Calendaren
            OtherMonthDayStyle-ForeColor="#ffffcc" runat="server"
            SelectedDayStyle-BackColor="Navy"
            SelectedDayStyle-Font-Bold="True"
            SelectorStyle-BackColor="gainsboro"
            TitleStyle-BackColor="#cccc66"
            TitleStyle-Font-Bold="True"
            TitleStyle-Font-Size="12px"
            TodayDayStyle-BackColor="#ffcc33" Visible="false" Width="200px"
            onselectionchanged="Date_Selecteden"
            />	  </td>
  </tr>
  <tr>
    <td width="80" bgcolor="#CCFFCC">結束時間：</td>
    <td bgcolor="#FFCC99">
      <asp:DropDownList AutoPostBack="true" ID="m_time_enh" runat="server" OnSelectedIndexChanged="change_texten">
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
      <asp:DropDownList ID="m_time_enn" runat="server" OnSelectedIndexChanged="change_texten" AutoPostBack="true">
        <asp:listitem>00</asp:listitem>
        <asp:listitem>10</asp:listitem>
        <asp:listitem>20</asp:listitem>
        <asp:listitem>30</asp:listitem>
        <asp:listitem>40</asp:listitem>
        <asp:listitem>50</asp:listitem>
      </asp:DropDownList>
    分   
    <asp:TextBox AutoPostBack="true" Columns="6" ID="m_time_en" ReadOnly="true" runat="server" Text='<%# showtime(DataSet1.FieldValue("m_time_en", Container)) %>' />
    <asp:Label Font-Bold="true" forecolor="red" ID="msgent" runat="server" /></td>
  </tr>
  <tr>
    <td width="86" bgcolor="#CCFFCC">主題：</td>
    <td bgcolor="#FFCC99"><asp:TextBox Columns="80" ID="m_note" Rows="3" runat="server" text='<%# DataSet1.FieldValue("m_note", Container) %>' Width="300" TextMode="MultiLine" />
    <asp:RequiredFieldValidator ControlToValidate="m_note" ErrorMessage="尚未填寫主題" runat="server" /></td>
  </tr>
  <tr>
    <td bgcolor="#CCFFCC">詳細內容：</td>
    <td bgcolor="#FFCC99"><asp:TextBox Columns="80" ID="m_detail" Rows="5" runat="server" text='<%# DataSet1.FieldValue("m_detail", Container) %>' Width="300" TextMode="MultiLine" /></td>
  </tr>
</table>

<table width="551" border="0" cellpadding="0" cellspacing="1">
    <tr>
	<td bgcolor="#FF9999">
	 <span class="style1">
  <asp:Button ID="Button1" runat="server" Text="更新" OnClick="UpdateData" />
  <input type="reset" name="Submit2" value="重新填寫" /></span>
	 <asp:Button ID="Button4" runat="server" Text="取消" OnClick="cancel_url" /> 
	     
相關圖照：
<asp:Button ID="Button5" runat="server" Text="增修檔案" OnClick= "trans_go" /></td>
  </tr>
</table>
<asp:DataGrid 
  AllowPaging="false" 
  AllowSorting="False" 
  AutoGenerateColumns="false" 
  CellPadding="3" 
  CellSpacing="0" 
  DataSource="<%# DataSet2.DefaultView %>" id="DataGrid1" 
  runat="server" 
  ShowFooter="false" 
  ShowHeader="true" Visible="" Width="551" 
>
      <HeaderStyle HorizontalAlign="center" BackColor="#E8EBFD" ForeColor="#3D3DB6" Font-Name="Verdana, Arial, Helvetica, sans-serif" Font-Bold="true" Font-Size="smaller" />    
      <ItemStyle BackColor="#F2F2F2" Font-Name="Verdana, Arial, Helvetica, sans-serif" Font-Size="smaller" />    
      <AlternatingItemStyle BackColor="#E5E5E5" Font-Name="Verdana, Arial, Helvetica, sans-serif" Font-Size="smaller" />    
      <FooterStyle HorizontalAlign="center" BackColor="#E8EBFD" ForeColor="#3D3DB6" Font-Name="Verdana, Arial, Helvetica, sans-serif" Font-Bold="true" Font-Size="smaller" />    
      <PagerStyle BackColor="white" Font-Name="Verdana, Arial, Helvetica, sans-serif" Font-Size="smaller" />    
      <Columns>
      <asp:HyperLinkColumn
        DataNavigateUrlField="m_appen"
        DataNavigateUrlFormatString="/web1/sec05/s01data/{0}"
        DataTextField="m_appen" 
        Visible="True" target="show_photo" 
        HeaderText="檔名"/>      
<asp:BoundColumn DataField="m_appen_con" 
        HeaderText="檔案說明" 
        ReadOnly="true" 
        Visible="True"/>      

      </Columns>
  </asp:DataGrid>

 
  
  <p>
  <input name="m_num" type="hidden" id="m_num" value="<%# DataSet1.FieldValue("m_num", Container) %>" />
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
