<%@ Page Language="VB" ContentType="text/html" ResponseEncoding="utf-8" validateRequest="false" EnableEventValidation="false"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<%@ Register TagPrefix="MM" Namespace="DreamweaverCtrls" Assembly="DreamweaverCtrls,version=1.0.0.0,publicKeyToken=836f606ede05d46a,culture=neutral" %>
<MM:DataSet 
id="DataSet1"
runat="Server"
IsStoredProcedure="false"
ConnectionString='<%# System.Configuration.ConfigurationSettings.AppSettings("MM_CONNECTION_STRING_s01") %>'
DatabaseType='<%# System.Configuration.ConfigurationSettings.AppSettings("MM_CONNECTION_DATABASETYPE_s01") %>'
CommandText='<%# "SELECT *  FROM s01_case  WHERE case_id = ?" %>'
Debug="true"
>
  <Parameters>
    <Parameter  Name="@case_id"  Value='<%# IIf((Request.QueryString("case_id") <> Nothing), Request.QueryString("case_id"), "") %>'  Type="Integer"   />  
  </Parameters>
</MM:DataSet>
<MM:DataSet 
id="DataSet2"
runat="Server"
IsStoredProcedure="false"
ConnectionString='<%# System.Configuration.ConfigurationSettings.AppSettings("MM_CONNECTION_STRING_s01") %>'
DatabaseType='<%# System.Configuration.ConfigurationSettings.AppSettings("MM_CONNECTION_DATABASETYPE_s01") %>'
CommandText='<%# "SELECT * FROM s01admin WHERE admin_yes = yes ORDER BY admin_username ASC" %>'
Debug="true"
>
</MM:DataSet>
<MM:DataSet 
id="DataSet3"
runat="Server"
IsStoredProcedure="false"
ConnectionString='<%# System.Configuration.ConfigurationSettings.AppSettings("MM_CONNECTION_STRING_s01") %>'
DatabaseType='<%# System.Configuration.ConfigurationSettings.AppSettings("MM_CONNECTION_DATABASETYPE_s01") %>'
CommandText='<%# "SELECT *  FROM s01_case_appen  WHERE case_id = ?" %>'
Debug="true" PageSize="30"
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
<meta http-equiv="Content-Type" content="text/html/>
<title>案件修改</title>
<script language="VB" runat="server">
 dim url, url_back
dim str1 as string
  Sub page_load(sender As Object, e As EventArgs)
		If Session("MM_chief_num")="" then
			Response.Write("您還沒有登入呢！，請點擊")
			Response.Write("<a href=s01dellogin.aspx>登入</a>")
			Response.End()
		End If
		  if not ispostback then
	dim m_numcheck
	m_numcheck = request("case_id")
	if m_numcheck = "" then
      response.Redirect("s01case_admin.aspx")
	else
	back.text = session("cancel_case_admin")
	 end if
	 
	 
str1 = DataSet1.FieldValue("admin_id", Nothing)
t1.text = str1

case_group.SelectedIndex =case_group.Items.IndexOf(case_group.Items.FindByValue(DataSet1.FieldValue("case_group", Nothing) ))
case_yes.SelectedIndex =case_yes.Items.IndexOf(case_yes.Items.FindByValue(DataSet1.FieldValue("case_yes", Nothing) ))
   end if
  end sub
  Sub call_calst(sender As Object, e As EventArgs) 
   Calendaren_1.visible = false
   Calendaren.visible = false
   Calendarst.visible = true
	msgst.text = ""
	msgen.text = ""
	msgen_1.text = ""
   End Sub

  Sub mansel(sender As Object, e As EventArgs)
if not ispostback then
case_man.SelectedIndex =case_man.Items.IndexOf(case_man.Items.FindByValue(DataSet1.FieldValue("admin_id", Nothing)))
end if
 end sub

  Sub call_calen(sender As Object, e As EventArgs) 
   Calendaren_1.visible = false
   Calendaren.visible = true
   Calendarst.visible = false
	msgst.text = ""
	msgen.text = ""
	msgen_1.text = ""
   End Sub
  Sub call_calen_1(sender As Object, e As EventArgs) 
   Calendaren_1.visible = true
   Calendaren.visible = false
   Calendarst.visible = false
	msgst.text = ""
	msgen.text = ""
	msgen_1.text = ""
   End Sub
  Sub tchange(sender As Object, e As EventArgs)
   dim d03
    d03 = case_limit.text
	diffchange(d03)
	end sub 
  Sub t1change(sender As Object, e As EventArgs)
   if case_man.selecteditem.value <> "" then
    t1.text = case_man.selecteditem.value
   end if
  end sub 
	 
 		function diffchange(diff)
         dim d01, d02
		  d01 = showdate(case_start.text)
		   d02 = dateadd("d", diff, d01)
		  Calendaren.SelectedDate = d02
		  end function  
 		Sub Date_Selecteden(sender As Object, e As EventArgs)
		  dim diff01
		  diff01 = datediff("d" , showdate(case_start.text) , showdate(Calendaren.SelectedDate.ToShortDateString))
		  if diff01 < 0 then
		   msgen.text = "結束日期小於啟始日期.請重新點選"
		  else 
		  case_limit.Text = diff01
		  diffchange(diff01)
		  msgen.text = ""
          end if
		End Sub 
 		Sub Date_Selecteden_1(sender As Object, e As EventArgs)
          if  case_end.Text <> "" then
		  if datediff("d" , showdate(case_start.text) , showdate(Calendaren_1.SelectedDate.ToShortDateString)) < 0 then
		   msgen_1.text = "結案日期小於啟始日期.請重新點選"
		  else 
		  case_end.Text = Calendaren_1.SelectedDate.ToShortDateString
		  msgen_1.text = ""
          end if
		  else
		  if datediff("d" , showdate(case_start.text) , showdate(Calendaren_1.SelectedDate.ToShortDateString)) < 0 then
		   msgen_1.text = "結案日期小於啟始日期.請重新點選"
		  else
		  case_end.Text = Calendaren_1.SelectedDate.ToShortDateString
		  msgen_1.text = ""
          end if
         end if
		End Sub  
 		Sub Date_Selectedst(sender As Object, e As EventArgs)
		  case_start.Text = showdate(Calendarst.SelectedDate.ToShortDateString)
		  dim dst01
		   dst01 = val(case_limit.text)
		   diffchange(dst01)
		  msgst.text = "案件日期已修改.限辦日期將順延"
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
	  SQL = "UPDATE s01case SET case_start=@case_start, case_content=@case_content, case_result=@case_result, case_yes=@case_yes, case_limit=@case_limit, case_group=@case_group, case_end=@case_end, admin_id=@admin_id WHERE case_id=@case_id"
      Cmd = New OleDbCommand( SQL, Conn )

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
      Cmd.Parameters("@admin_id").value = val(case_man.selecteditem.value)
      Cmd.Parameters("@case_id").value = val(case_id.text)

      Cmd.ExecuteNonQuery()
      'If Err.Number <> 0 Then
         'Msg.Text = Err.Description
      'Else
		  url = back.text
      'End If

      Conn.Close()
		  response.Redirect(url)
       else
	  SQL = "UPDATE s01case SET case_start=@case_start, case_content=@case_content, case_result=@case_result, case_yes=@case_yes, case_limit=@case_limit, case_group=@case_group, case_end=null, admin_id=@admin_id WHERE case_id=@case_id"
      Cmd = New OleDbCommand( SQL, Conn )

      Cmd.Parameters.Add( New OleDbParameter("@case_start", OleDbType.dbdate))
      Cmd.Parameters.Add( New OleDbParameter("@case_content", OleDbType.varChar))
      Cmd.Parameters.Add( New OleDbParameter("@case_result", OleDbType.varChar))
      Cmd.Parameters.Add( New OleDbParameter("@case_yes", OleDbType.char,6))
      Cmd.Parameters.Add( New OleDbParameter("@case_limit", OleDbType.single))
      Cmd.Parameters.Add( New OleDbParameter("@case_group", OleDbType.char, 20))
      'Cmd.Parameters.Add( New OleDbParameter("@case_end", OleDbType.varChar))
      Cmd.Parameters.Add( New OleDbParameter("@admin_id", OleDbType.integer))
      Cmd.Parameters.Add( New OleDbParameter("@case_id", OleDbType.integer))

      Cmd.Parameters("@case_start").value = case_start.text
      Cmd.Parameters("@case_content").value = case_content.text
      Cmd.Parameters("@case_result").value = case_result.text
      Cmd.Parameters("@case_yes").value = case_yes.selecteditem.text
      Cmd.Parameters("@case_limit").value = val(case_limit.text)
      Cmd.Parameters("@case_group").value = case_group.selecteditem.text
      'Cmd.Parameters("@case_end").value = cstr(case_end.text)
      Cmd.Parameters("@admin_id").value = val(case_man.selecteditem.value)
      Cmd.Parameters("@case_id").value = val(case_id.text)

      Cmd.ExecuteNonQuery()
      'If Err.Number <> 0 Then
         'Msg.Text = Err.Description
      'Else
		  url = back.text
      'End If

      Conn.Close()	   
		  response.Redirect(url)
	  end if 
	  end if 
	  end sub
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

Sub startcancel(sender As Object, e As EventArgs) 
 response.Redirect(back.text)
end sub

sub trans_go(sender As Object, e As EventArgs)
 dim trans_text = "s01case_upfile.aspx?case_id=" & case_id.text
 if trans_text <> "" then
 session("cancel_case01") = ""
 session("cancel_case01") = request.url.tostring()
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
location.href = "s01cal_selday.aspx?m_date=" & x1;
}

function MM_openBrWindow(theURL,winName,features) { //v2.0
  window.open(theURL,winName,features);
}
function Mcheck(){
	if (document.form_update.case_content.value=="") {
        window.alert("請輸入內容");
        return false }
	if (document.form_update.case_limit.value=="") {
        window.alert("請輸入限辦天數");
        return false }
	if (isNaN(document.form_update.case_limit.value)) {
        window.alert("天數請輸入數值");
        return false }

	 return true;
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
	color: #CC3300;
	font-weight: bold;
	font-size: 16px;
}
.style2 {color: #FF0000}
-->
</style>
</head>
<body>
<span class="style1">修改案件</span>
<form name='form_update' id="form_update" runat="server">
<table width="551" border="0" cellspacing="1" cellpadding="0">
  <tr>
    <td width="86" bgcolor="#CCFFCC">案件日期：</td>
    <td width="462" bgcolor="#FFCC99"><a href="s01cal_cal.aspx" target="_blank"></a>
      <asp:Button Enabled="true" ID="Button2" runat="server" Text="案件日期" OnClick="call_calst" />
      <asp:TextBox AutoPostBack="true" Columns="9" ID="case_start" ReadOnly="true" runat="server" Text='<%# showdate(DataSet1.FieldValue("case_start", Container)) %>' />
      <asp:TextBox ID="t1" ReadOnly="true" runat="server" />
<br/>
      <asp:Label Font-Bold="true" forecolor="red" ID="msgst" runat="server" />
      <asp:TextBox AutoPostBack="true" Columns="6" ID="case_id" ReadOnly="true" runat="server" text='<%# DataSet1.FieldValue("case_id", Container) %>' Visible="false" />
      <br />
<asp:Calendar
            DayHeaderStyle-BackColor="#ffcccc"
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
    <td width="86" bgcolor="#CCFFCC">類型：</td>
    <td bgcolor="#FFCC99">
	<asp:DropDownList AutoPostBack="true" Enabled="true" EnableViewState="true" ID="case_group" runat="server">
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
    <td width="86" bgcolor="#CCFFCC">承辦人：</td>
    <td bgcolor="#FFCC99">
	<asp:DropDownList DataSource="<%# dataset2.defaultview %>" DataTextField="s01name" DataValueField="admin_id" ID="case_man" runat="server" OnSelectedIndexChanged="t1change" OnLoad="mansel" AutoPostBack="true" ></asp:DropDownList></td>
  </tr>
  <tr>
    <td bgcolor="#CCFFCC">內容：</td>
    <td bgcolor="#FFCC99"><asp:TextBox Columns="30" ID="case_content" ReadOnly="false" Rows="5" runat="server" text='<%# DataSet1.FieldValue("case_content", Container) %>' TextMode="MultiLine" Width="300" /></td>
  </tr>
  <tr>
    <td bgcolor="#CCFFCC">辦理情形：</td>
    <td bgcolor="#FFCC99"><asp:TextBox Columns="30" ID="case_result" ReadOnly="false" Rows="6" runat="server" text='<%# DataSet1.FieldValue("case_result", Container) %>' TextMode="MultiLine" Width="300" /></td>
  </tr>

  <tr>
    <td bgcolor="#CCFFCC">限辦天數</td>
    <td bgcolor="#FFCC99">
	  <asp:Button Enabled="true" ID="Button3" runat="server" Text="限辦日期" OnClick="call_calen" />
	  <asp:TextBox AutoPostBack="true" Columns="8" ID="case_limit" ReadOnly="true" runat="server" text='<%# DataSet1.FieldValue("case_limit", Container) %>' OnTextChanged="tchange" />
<br/>
<asp:Label Font-Bold="true" forecolor="red" ID="msgen" runat="server" />
<br/>
	  <asp:Calendar
            DayHeaderStyle-BackColor="#ffcccc"
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
            />
	  </td>
  </tr>
  <tr>
    <td bgcolor="#CCFFCC">是否登錄：</td>
    <td bgcolor="#FFCC99">
	<asp:DropDownList AutoPostBack="true" Enabled="true" ID="case_yes" runat="server">
      <asp:listitem>是</asp:listitem>
      <asp:listitem Selected="true">否</asp:listitem>
      <asp:listitem>不需</asp:listitem>
    </asp:DropDownList></td>
  </tr>
  <tr>
    <td bgcolor="#CCFFCC">結案日期：</td>
    <td bgcolor="#FFCC99"><asp:Button Enabled="true" ID="Button5" runat="server" Text="結案日期" OnClick="call_calen_1" />
    
    <asp:TextBox ID="case_end" ReadOnly="true" runat="server" Text='<%# showdate(DataSet1.FieldValue("case_end", Container)) %>' Width="60" Wrap="false" />
    <br/>
<asp:Label Font-Bold="true" forecolor="red" ID="msgen_1" runat="server" />
<br/>
	  <asp:Calendar
            DayHeaderStyle-BackColor="#ffcccc"
            Font-Name="Arial" Font-Size="12px"
            Height="180px" id=Calendaren_1
            OtherMonthDayStyle-ForeColor="#ffffcc" runat="server"
            SelectedDayStyle-BackColor="Navy"
            SelectedDayStyle-Font-Bold="True"
            SelectorStyle-BackColor="gainsboro"
            TitleStyle-BackColor="#cccc66"
            TitleStyle-Font-Bold="True"
            TitleStyle-Font-Size="12px"
            TodayDayStyle-BackColor="#ffcc33" Visible="false" Width="200px"
            onselectionchanged="Date_Selecteden_1"
            /> </td>
  </tr>
</table>
  <p><span class="style1">
    <asp:Button ID="Button1" runat="server" Text="更新" OnClick="UpdateData" />    
<input type="reset" name="Submit2" value="重新填寫" />
    </span>
    
      <span class="style1">
      <asp:Button ID="Buttonc" runat="server" Text="取消" OnClick='startcancel' />      
相關檔案：
<asp:Button ID="add_appen" runat="server" Text="增修檔案" OnClick= "trans_go" /></span>  
      <asp:TextBox ID="back" runat="server" Visible="false" Columns="50" />
</p>
  <p>&nbsp;
    <asp:DataGrid 
  AllowCustomPaging="true" 
  AllowPaging="true" 
  AllowSorting="False" 
  AutoGenerateColumns="false" 
  CellPadding="3" 
  CellSpacing="0" 
  DataSource="<%# DataSet3.DefaultView %>" id="DataGrid1" 
  PagerStyle-Mode="NumericPages" 
  PageSize="<%# DataSet3.PageSize %>" 
  runat="server" 
  ShowFooter="false" 
  ShowHeader="true" Width="80%" 
  OnPageIndexChanged="DataSet3.OnDataGridPageIndexChanged" 
  VirtualItemCount="<%# DataSet3.RecordCount %>" 
>
      <HeaderStyle HorizontalAlign="center" BackColor="#E8EBFD" ForeColor="#3D3DB6" Font-Name="Verdana, Arial, Helvetica, sans-serif" Font-Bold="true" Font-Size="smaller" />      
<ItemStyle BackColor="#F2F2F2" Font-Name="Verdana, Arial, Helvetica, sans-serif" Font-Size="smaller" />      
<AlternatingItemStyle BackColor="#E5E5E5" Font-Name="Verdana, Arial, Helvetica, sans-serif" Font-Size="smaller" />      
<FooterStyle HorizontalAlign="center" BackColor="#E8EBFD" ForeColor="#3D3DB6" Font-Name="Verdana, Arial, Helvetica, sans-serif" Font-Bold="true" Font-Size="smaller" />
<PagerStyle BackColor="white" Font-Name="Verdana, Arial, Helvetica, sans-serif" Font-Size="smaller" />
<Columns>
      <asp:BoundColumn DataField="case_id" 
        HeaderText="case_id" 
        ReadOnly="true" 
        Visible="false"/>      
      <asp:BoundColumn DataField="case_appen_id" 
        HeaderText="case_appen_id" 
        ReadOnly="true" 
        Visible="false"/>      
<asp:HyperLinkColumn
        DataNavigateUrlField="case_appen" DataNavigateUrlFormatString="/web1/sec05/s01data/{0}"
        DataTextField="case_appen" 
        Visible="True" target="case_appen_admin" 
        HeaderText="檔案"/>      
<asp:BoundColumn DataField="case_appen_con" 
        HeaderText="說明" 
        ReadOnly="true" 
        Visible="True"/>
</Columns>
    </asp:DataGrid>
</p>
  <p>
  <input name="case_id" type="hidden" id="case_id" value="<%# DataSet1.FieldValue("case_id", Container) %>" />
<HR>
</form>
</body>
</html>