<%@ Page Language="VB" ContentType="text/html" ResponseEncoding="utf-8" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<%@ Register TagPrefix="MM" Namespace="DreamweaverCtrls" Assembly="DreamweaverCtrls,version=1.0.0.0,publicKeyToken=836f606ede05d46a,culture=neutral" %>
<MM:DataSet 
id="DataSet1"
runat="Server"
IsStoredProcedure="false"
ConnectionString='<%# System.Configuration.ConfigurationSettings.AppSettings("MM_CONNECTION_STRING_s01") %>'
DatabaseType='<%# System.Configuration.ConfigurationSettings.AppSettings("MM_CONNECTION_DATABASETYPE_s01") %>'
CommandText='<%# "SELECT distinct t_g_1, t_g_name, t_g_2 FROM t_group where t_g_2 like ? or t_g_2 is null ORDER BY t_g_1 ASC" %>'
Debug="true"
>
  <Parameters>
<Parameter  Name="@t_g_2"  Value='<%# "" %>'  Type="WChar"   />
</Parameters>
</MM:DataSet>
<MM:DataSet 
id="DataSet2"
runat="Server"
IsStoredProcedure="false"
ConnectionString='<%# System.Configuration.ConfigurationSettings.AppSettings("MM_CONNECTION_STRING_s01") %>'
DatabaseType='<%# System.Configuration.ConfigurationSettings.AppSettings("MM_CONNECTION_DATABASETYPE_s01") %>'
CommandText='<%# "SELECT distinct t_g_2, t_g_name, t_g_1, t_g_3 FROM t_group where t_g_2 <> ? and t_g_1 like ? and(t_g_3 like ? or t_g_3 is null) ORDER BY t_g_2 ASC" %>'
Debug="true"
>
  <Parameters>
<Parameter  Name="@t_g_2"  Value='<%# "" %>'  Type="WChar"   />
<Parameter  Name="@t_g_1"  Value='<%# IIf(cstr(t_g_1_t.text) = "", "", cstr(t_g_1_t.text)) %>'  Type="WChar"   />
<Parameter  Name="@t_g_3"  Value='<%# "" %>'  Type="WChar"   />

</Parameters>
</MM:DataSet>
<MM:DataSet 
id="DataSet3"
runat="Server"
IsStoredProcedure="false"
ConnectionString='<%# System.Configuration.ConfigurationSettings.AppSettings("MM_CONNECTION_STRING_s01") %>'
DatabaseType='<%# System.Configuration.ConfigurationSettings.AppSettings("MM_CONNECTION_DATABASETYPE_s01") %>'
CommandText='<%# "SELECT t_g_2, t_g_name, t_g_3, t_g_num FROM t_group where t_g_3 <> ? and t_g_2 like ? and( t_g_num like ? or t_g_num is null) ORDER BY t_g_3 ASC" %>'
Debug="true"
>
  <Parameters>
<Parameter  Name="@t_g_3"  Value='<%# "" %>'  Type="WChar"   />
<Parameter  Name="@t_g_2"  Value='<%# IIf(cstr(t_g_2_t.text) = "", "", cstr(t_g_2_t.text)) %>'  Type="WChar"   />
<Parameter  Name="@t_g_num"  Value='<%# "" %>'  Type="WChar"   />
</Parameters>
</MM:DataSet>
<MM:DataSet 
id="DataSet5"
runat="Server"
IsStoredProcedure="false"
ConnectionString='<%# System.Configuration.ConfigurationSettings.AppSettings("MM_CONNECTION_STRING_s01") %>'
DatabaseType='<%# System.Configuration.ConfigurationSettings.AppSettings("MM_CONNECTION_DATABASETYPE_s01") %>'
CommandText='<%# "SELECT t_g_2, t_g_name, t_g_3, t_g_num FROM t_group where t_g_num <> ? and t_g_3 like ? ORDER BY t_g_num ASC" %>'
Debug="true"
>
  <Parameters>
<Parameter  Name="@t_g_num"  Value='<%# "" %>'  Type="WChar"   />
<Parameter  Name="@t_g_3"  Value='<%# IIf(cstr(t_g_3_t.text) = "", "", cstr(t_g_3_t.text)) %>'  Type="WChar"   />

</Parameters>
</MM:DataSet>
<MM:DataSet 
id="DataSet7"
runat="Server"
IsStoredProcedure="false"
ConnectionString='<%# System.Configuration.ConfigurationSettings.AppSettings("MM_CONNECTION_STRING_s01") %>'
DatabaseType='<%# System.Configuration.ConfigurationSettings.AppSettings("MM_CONNECTION_DATABASETYPE_s01") %>'
CommandText='<%# "SELECT * FROM s01admin WHERE admin_yes = yes ORDER BY admin_username ASC" %>'
Debug="true"
>
</MM:DataSet>
<MM:DataSet 
id="DataSet6"
runat="Server"
IsStoredProcedure="false"
ConnectionString='<%# System.Configuration.ConfigurationSettings.AppSettings("MM_CONNECTION_STRING_s01") %>'
DatabaseType='<%# System.Configuration.ConfigurationSettings.AppSettings("MM_CONNECTION_DATABASETYPE_s01") %>'
CommandText='<%# "SELECT t_g_id FROM t_group where t_g_1 like ? and t_g_2 like ? and t_g_3 like ? and t_g_num like ? ORDER BY t_g_num ASC" %>'
Debug="true"
>
  <Parameters>
<Parameter  Name="@t_g_1"  Value='<%# IIf(cstr(shownum(t_g_1_t.text)) = "", "", cstr(shownum(t_g_1_t.text))) %>'  Type="WChar"   />
<Parameter  Name="@t_g_2"  Value='<%# IIf(cstr(shownum(t_g_2_t.text)) = "", "", cstr(shownum(t_g_2_t.text))) %>'  Type="WChar"   />
<Parameter  Name="@t_g_3"  Value='<%# IIf(cstr(shownum(t_g_3_t.text)) = "", "", cstr(shownum(t_g_3_t.text))) %>'  Type="WChar"   />
<Parameter  Name="@t_g_num"  Value='<%# IIf(cstr(shownum(t_g_num_t.text)) = "", "", cstr(shownum(t_g_num_t.text))) %>'  Type="WChar"   />
</Parameters>
</MM:DataSet>
<MM:PageBind runat="server" PostBackBind="true" />
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
    'case_start.text = request("case_start")
	'if case_start.text = "" then
     ' response.Redirect("s01case_admin.aspx")
    'else
	'back.text = session("cancel_case")
	'session("cancel_case") = ""
	 'end if
	'end if
	   't_g_1.SelectedIndex = t_g_1.Items.IndexOf(t_g_1.Items.FindByValue(t_g_1_t.text))
	
	't_g_1.items.add("全部")
	t_g_1.datasource = DataSet1.DefaultView
	case_man.datasource = DataSet7.DefaultView
	't_address.datasource = DataSet4.DefaultView
	't_g_2.datasource = DataSet2.DefaultView
	t_g_1_t.text = "%"
	t_g_2_t.text = "%"
	t_g_3_t.text = "%"
	t_g_num_t.text = "%"
	't_g_1.DataTextField = t_g_name
	't_g_1.DataValueField = DataSet1.FieldValue("t_g_1")
	t_g_2.enabled = false
	t_g_3.enabled = false
	t_g_num.enabled = false
	
	else
	t_g_1_t.text = cstr(t_g_1.selecteditem.value) '放在page load因為此會先執行
	't_g_2.SelectedIndex = t_g_2.Items.IndexOf(t_g_2.Items.FindByValue(t_g_2_t.text))
	'if t_g_2.selectedindex < (DataSet2.RecordCount - 1) then
	'else
	 't_g_2_t.text = t_g_1_t.text 
    'end if
    if t_g_1.selectedindex = 0 then
	t_g_2.enabled = false
	  t_g_2_t.text = "%"
	t_g_3.enabled = false
	  t_g_3_t.text = "%"
	t_g_num.enabled = false
	  t_g_num_t.text = "%"
	  exit sub
	else 
	t_g_2.enabled = true
	t_g_2.datasource = DataSet2.DefaultView
	 t_g_2_t.text = t_g_2.selecteditem.value
	t_g_3.enabled = false
	  t_g_3_t.text = "%"
	t_g_num.enabled = false
	  t_g_num_t.text = "%"
    if t_g_2.selectedindex = 0 then
	t_g_3.enabled = false
	  t_g_3_t.text = "%"
	t_g_num.enabled = false
	  t_g_num_t.text = "%"
	  exit sub
    else
	 t_g_3.enabled = true
	 t_g_3.datasource = DataSet3.DefaultView
	 t_g_3_t.text = t_g_3.selecteditem.value
     t_g_num.enabled = false
	  t_g_num_t.text = "%"
    if t_g_3.selectedindex = 0 then
	t_g_num.enabled = false
	  t_g_num_t.text = "%"
 	  exit sub
    else
	 t_g_num.enabled = true
	 t_g_num.datasource = DataSet5.DefaultView
	 t_g_num_t.text = t_g_num.selecteditem.value
	  exit sub
    end if
	end if
	 end if
    't_g_2.Items.Insert(0, New ListItem("全部", "%"))
	't_g_2.SelectedIndex = t_g_2.Items.IndexOf(t_g_2.Items.FindByValue(t_g_2_t.text))
	't_g_2_t.text = t_g_1_t.text '一改變時.二便要全部
	end if
  end sub
 
    Sub t_g_1_pre(sender As Object, e As EventArgs)
     if not ispostback then
	 t_g_1.Items.Insert(0, New ListItem("全部", "%"))
	 t_g_1.selectedindex = 0
	 't_g_1_t.text = cstr(t_g_1.selecteditem.value)
	 end if
	 end sub

   Sub t_g_1_cha(sender As Object, e As EventArgs)
	if ispostback then
	't_g_1_t.text = cstr(t_g_1.selecteditem.value)
	'if t_g_1.selectedindex > 0 then
	't_g_2_t.text = cstr(t_g_1.selecteditem.value)  '一改變時.二便要全部
	't_g_2.datasource = DataSet2.DefaultView
	'else
    t_g_2_t.text = "%"
	'end if
	t_g_3.enabled = false
	  t_g_3_t.text = "%"
	t_g_num.enabled = false
	  t_g_num_t.text = "%"

	end if
	 end sub

   Sub t_g_2_pre(sender As Object, e As EventArgs)
	 'if t_g_1_t.text = t_g_2_t.text then
	 t_g_2.Items.Insert(0, New ListItem("全部", "%"))
	 't_g_2_t.text = cstr(t_g_2.selecteditem.value)
	 't_g_2.SelectedIndex = 2 
	't_g_2.SelectedIndex = 2
    if t_g_1.selectedindex > 0 then 
	t_g_2.SelectedIndex = t_g_2.Items.IndexOf(t_g_2.Items.FindByValue(t_g_2_t.text))
	else
	t_g_2.SelectedIndex = 0
	end if
	't_g_3_t.text = t_g_2.Selectedindex
	 'end if
	 end sub

   Sub t_g_2_on(sender As Object, e As EventArgs)
     if ispostback then
	t_g_2_t.text = t_g_2.selecteditem.value
	't_g_3.datasource = DataSet2.DefaultView
	end if
	 end sub

   Sub t_g_2_cha(sender As Object, e As EventArgs)
	if ispostback then
	't_g_1_t.text = cstr(t_g_1.selecteditem.value)
	'if t_g_1.selectedindex > 0 then
	't_g_2_t.text = cstr(t_g_1.selecteditem.value)  '一改變時.二便要全部
	't_g_2.datasource = DataSet2.DefaultView
	'else
    t_g_3_t.text = "%"
	'end if
	t_g_3.SelectedIndex = 0
	t_g_num.enabled = false
	  t_g_num_t.text = "%"

	end if
	 end sub

   Sub t_g_2_bin(sender As Object, e As EventArgs)
	 if t_g_1_t.text = t_g_2_t.text then
	 t_g_2.datasource = DataSet2.DefaultView
	 't_g_2.SelectedIndex =t_g_2.Items.IndexOf(t_g_2.Items.FindByValue(cstr(t_g_2_t.text)))
	 end if
	 end sub

   Sub t_g_2_un(sender As Object, e As EventArgs)
	 t_g_2.datasource = ""
	 't_g_2.SelectedIndex =t_g_2.Items.IndexOf(t_g_2.Items.FindByValue(cstr(t_g_2_t.text)))
	 end sub

   Sub t_g_3_pre(sender As Object, e As EventArgs)
	 'if t_g_1_t.text = t_g_2_t.text then
	 t_g_3.Items.Insert(0, New ListItem("全部", "%"))
	 't_g_2_t.text = cstr(t_g_2.selecteditem.value)
	 't_g_2.SelectedIndex = 2 
	't_g_2.SelectedIndex = 2
    if t_g_2.selectedindex > 0 then 
	t_g_3.SelectedIndex = t_g_3.Items.IndexOf(t_g_3.Items.FindByValue(t_g_3_t.text))
	else
	t_g_3.SelectedIndex = 0
	end if
	't_g_3_t.text = t_g_2.Selectedindex
	 'end if
	 end sub

   Sub t_g_3_cha(sender As Object, e As EventArgs)
	if ispostback then
	't_g_1_t.text = cstr(t_g_1.selecteditem.value)
	'if t_g_1.selectedindex > 0 then
	't_g_2_t.text = cstr(t_g_1.selecteditem.value)  '一改變時.二便要全部
	't_g_2.datasource = DataSet2.DefaultView
	'else
    t_g_num_t.text = "%"
	'end if
	t_g_num.SelectedIndex = 0
	end if
	 end sub

   Sub t_g_num_pre(sender As Object, e As EventArgs)
	 'if t_g_1_t.text = t_g_2_t.text then
	 t_g_num.Items.Insert(0, New ListItem("全部", "%"))
	 't_g_2_t.text = cstr(t_g_2.selecteditem.value)
	 't_g_2.SelectedIndex = 2 
	't_g_2.SelectedIndex = 2
    if t_g_3.selectedindex > 0 then 
	t_g_num.SelectedIndex = t_g_num.Items.IndexOf(t_g_num.Items.FindByValue(t_g_num_t.text))
	else
	t_g_num.SelectedIndex = 0
	end if
	't_g_3_t.text = t_g_2.Selectedindex
	 'end if
	 end sub 


  Sub InsertData(sender As Object, e As EventArgs) 
    if IsValid Then

      Dim Conn As OleDbConnection
      Dim Cmd  As OleDbCommand

      Dim Provider = "Provider=Microsoft.Jet.OLEDB.4.0"
      Dim Database = "Data Source=" & Server.MapPath( "result/result.mdb" )
      Conn = New OleDbConnection( Provider & ";" & DataBase )
      Conn.Open()

      Dim SQL As String
      SQL = "Insert Into s01insp_section (t_g_id, admin_id) Values(@t_g_id, @admin_id)"
      Cmd = New OleDbCommand( SQL, Conn )

      Cmd.Parameters.Add( New OleDbParameter("@t_g_id", OleDbType.integer))
      Cmd.Parameters.Add( New OleDbParameter("@admin_id", OleDbType.integer))

      Cmd.Parameters("@t_g_id").value = cint(t_g_id_t.text)
      Cmd.Parameters("@admin_id").value = cint(case_man.selecteditem.value)

      Cmd.ExecuteNonQuery()
      If Err.Number <> 0 Then
         Msg.Text = Err.Description
      Else
		 url = "s01insp_total_s_index.aspx"
      End If

      Conn.Close()
      response.Redirect(url)
	  End If
   End Sub

 function shownum(t_num)
 if t_num = "%"
 shownum = ""
 else
 shownum = t_num
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
<p class="style1">新增轄區：</p>
<form runat="server" name='form_add' id="form_add">
<table width="529" border="0" cellspacing="1" cellpadding="0">
  <tr>
    <td width="22%">行業別(一)</td>
    <td width="78%"><asp:DropDownList AutoPostBack="true" DataTextField="t_g_name" DataValueField="t_g_1" ID="t_g_1" runat="server" OnSelectedIndexChanged="t_g_1_cha" OnPreRender="t_g_1_pre">

	</asp:DropDownList>
      <asp:TextBox ID="t_g_1_t" ReadOnly="true" runat="server" Visible="false" />
      <asp:TextBox ID="t_g_id_t" ReadOnly="true" runat="server" text='<%# DataSet6.FieldValue("t_g_id", Container) %>' Visible="true" /></td>
  </tr>
  <tr>
    <td>行業別(二)</td>
    <td><asp:DropDownList AutoPostBack="true" DataTextField="t_g_name" DataValueField="t_g_2" ID="t_g_2" runat="server" OnSelectedIndexChanged="t_g_2_cha" OnPreRender="t_g_2_pre"></asp:DropDownList>
      <asp:TextBox ID="t_g_2_t" ReadOnly="true" runat="server" Visible="false" /></td>
  </tr>
  <tr>
    <td>行業別(三)</td>
    <td><asp:DropDownList AutoPostBack="true" DataTextField="t_g_name" DataValueField="t_g_3" ID="t_g_3" runat="server" OnSelectedIndexChanged="t_g_3_cha" OnPreRender="t_g_3_pre"></asp:DropDownList>
      <asp:TextBox ID="t_g_3_t" ReadOnly="true" runat="server" Visible="false" /></td>
  </tr>
  <tr>
    <td>行業別(四)</td>
    <td><asp:DropDownList AutoPostBack="true" DataTextField="t_g_name" DataValueField="t_g_num" ID="t_g_num" runat="server" OnPreRender="t_g_num_pre"></asp:DropDownList>
      <asp:TextBox ID="t_g_num_t" ReadOnly="true" runat="server" Visible="false" /></td>
  </tr>
  <tr>
    <td width="80" bgcolor="#CCFFFF">承辦人：</td>
    <td bgcolor="#CCCC66"><asp:DropDownList DataTextField="s01name" DataValueField="admin_id" ID="case_man" runat="server">
    </asp:DropDownList></td>
  </tr>

</table>
  <asp:Button ID="Button1" runat="server" Text="新增" OnClick="InsertData" />
  <span class="style1">
  <input type="reset" name="Submit2" value="重新填寫" />
  </span>
  <span class="style3">
  <asp:Button ID="Buttonc" runat="server" Text="取消" />  </span>
  <asp:TextBox ID="back" runat="server" Visible="false" Columns="50" />
  
  <p>
  <HR><asp:ValidationSummary
     DisplayMode="BulletList" EnableClientScript="true" Enabled="false"
     HeaderText="必須輸入的欄位還有:" ID="ch04" runat="server" />
<asp:Label runat="server" id="Msg" ForeColor="Red" />
</form>
</body>
</html>