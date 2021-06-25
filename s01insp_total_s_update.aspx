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
<Parameter  Name="@t_g_1"  Value='<%# IIf(cstr(t_g_1_t.value) = "", "", cstr(t_g_1_t.value)) %>'  Type="WChar"   />
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
<Parameter  Name="@t_g_2"  Value='<%# IIf(cstr(t_g_2_t.value) = "", "", cstr(t_g_2_t.value)) %>'  Type="WChar"   />
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
<Parameter  Name="@t_g_3"  Value='<%# IIf(cstr(t_g_3_t.value) = "", "", cstr(t_g_3_t.value)) %>'  Type="WChar"   />

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
<Parameter  Name="@t_g_1"  Value='<%# IIf(cstr(shownum(t_g_1_t.value)) = "", "", cstr(shownum(t_g_1_t.value))) %>'  Type="WChar"   />
<Parameter  Name="@t_g_2"  Value='<%# IIf(cstr(shownum(t_g_2_t.value)) = "", "", cstr(shownum(t_g_2_t.value))) %>'  Type="WChar"   />
<Parameter  Name="@t_g_3"  Value='<%# IIf(cstr(shownum(t_g_3_t.value)) = "", "", cstr(shownum(t_g_3_t.value))) %>'  Type="WChar"   />
<Parameter  Name="@t_g_num"  Value='<%# IIf(cstr(shownum(t_g_num_t.value)) = "", "", cstr(shownum(t_g_num_t.value))) %>'  Type="WChar"   />
</Parameters>
</MM:DataSet>
<MM:DataSet 
id="DataSet8"
runat="Server"
IsStoredProcedure="false"
ConnectionString='<%# System.Configuration.ConfigurationSettings.AppSettings("MM_CONNECTION_STRING_s01") %>'
DatabaseType='<%# System.Configuration.ConfigurationSettings.AppSettings("MM_CONNECTION_DATABASETYPE_s01") %>'
CommandText='<%# "SELECT * FROM s01_insp_s where insp_s_id = ? " %>'
Debug="true"
>
  <Parameters>
    <Parameter  Name="@insp_s_id"  Value='<%# IIf((Request.QueryString("insp_s_id") <> Nothing), Request.QueryString("insp_s_id"), "") %>'  Type="Integer"   />
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
	t_g_id_t.value = DataSet6.FieldValue("t_g_id", nothing) 
  if not Ispostback Then
    'case_start.value = request("case_start")
	'if case_start.value = "" then
     ' response.Redirect("s01case_admin.aspx")
    'else
	'back.value = session("cancel_case")
	'session("cancel_case") = ""
	 'end if
	'end if
	   't_g_1.SelectedIndex = t_g_1.Items.IndexOf(t_g_1.Items.FindByValue(t_g_1_t.value))
	
	't_g_1.items.add("全部")
	't_address.datasource = DataSet4.DefaultView
	't_g_2.datasource = DataSet2.DefaultView
	't_g_1.DataTextField = t_g_name
	't_g_1.DataValueField = DataSet8.FieldValue("t_g_1")

	't_g_1.datasource = DataSet7.DefaultView
	't_address.datasource = DataSet4.DefaultView
	't_g_2.enabled = true
	't_g_3.enabled = true
	't_g_2.datasource = DataSet2.DefaultView
	't_g_3.datasource = DataSet3.DefaultView
	't_g_num.datasource = DataSet7.DefaultView

	t_g_1.datasource = DataSet1.DefaultView
	't_address.datasource = DataSet4.DefaultView
	t_g_2.datasource = DataSet2.DefaultView
	t_g_3.datasource = DataSet3.DefaultView
	t_g_num.datasource = DataSet5.DefaultView
	case_man.datasource = DataSet7.DefaultView
	
	t_g_1_t.value = DataSet8.FieldValue("t_g_1", nothing)
	t_g_2_t.value = DataSet8.FieldValue("t_g_2", nothing)
	t_g_3_t.value = DataSet8.FieldValue("t_g_3", nothing)
	t_g_num_t.value = DataSet8.FieldValue("t_g_num", nothing)

	t_g_1.SelectedIndex = t_g_1.Items.IndexOf(t_g_1.Items.FindByValue(t_g_1_t.value))
	case_man.SelectedIndex = case_man.Items.IndexOf(case_man.Items.FindByValue(DataSet8.FieldValue("s01name", nothing)))
	't_g_3.SelectedIndex = t_g_3.Items.IndexOf(t_g_3.Items.FindByValue(t_g_3_t.value))
	't_g_num.SelectedIndex = t_g_num.Items.IndexOf(t_g_num.Items.FindByValue(t_g_num_t.value))
	't_g_2.enabled = false
	't_g_3.enabled = false
	't_g_num.enabled = false
	'button01.enabled = false
	't_g_id_t.value = ""


	't_g_1.datasource = DataSet1.DefaultView
	't_g_1_t.value = "%"
	't_g_2_t.value = "%"
	't_g_3_t.value = "%"
	't_g_num_t.value = "%"
	't_g_2.enabled = false
	't_g_3.enabled = false
	't_g_num.enabled = false
	
	else
	t_g_1_t.value = cstr(t_g_1.selecteditem.value) '放在page load因為此會先執行
	't_g_2.SelectedIndex = t_g_2.Items.IndexOf(t_g_2.Items.FindByValue(t_g_2_t.value))
	'if t_g_2.selectedindex < (DataSet2.RecordCount - 1) then
	'else
	 't_g_2_t.value = t_g_1_t.value 
    'end if
    if t_g_1.selectedindex = 0 then
	t_g_2.enabled = false
	  t_g_2_t.value = "%"
	t_g_3.enabled = false
	  t_g_3_t.value = "%"
	t_g_num.enabled = false
	  t_g_num_t.value = "%"
	  exit sub
	else 
	t_g_2.enabled = true
	t_g_2_t.value = t_g_2.selecteditem.value
    t_g_2.datasource = DataSet2.DefaultView
	t_g_3.enabled = false
	  t_g_3_t.value = "%"
	t_g_num.enabled = false
	  t_g_num_t.value = "%"
    if t_g_2.selectedindex = 0 then
	t_g_3.enabled = false
	  t_g_3_t.value = "%"
	t_g_num.enabled = false
	  t_g_num_t.value = "%"
	  exit sub
    else
	 t_g_3.enabled = true
	 t_g_3_t.value = t_g_3.selecteditem.value
	 t_g_3.datasource = DataSet3.DefaultView
     t_g_num.enabled = false
	  t_g_num_t.value = "%"
    if t_g_3.selectedindex = 0 then
	t_g_num.enabled = false
	  t_g_num_t.value = "%"
 	  exit sub
    else
	 t_g_num.enabled = true
	 t_g_num_t.value = t_g_num.selecteditem.value
	 t_g_num.datasource = DataSet5.DefaultView
	  exit sub
    end if
	end if
	 end if
    't_g_2.Items.Insert(0, New ListItem("全部", "%"))
	't_g_2.SelectedIndex = t_g_2.Items.IndexOf(t_g_2.Items.FindByValue(t_g_2_t.value))
	't_g_2_t.value = t_g_1_t.value '一改變時.二便要全部
	end if
  end sub
 
    Sub t_g_1_pre(sender As Object, e As EventArgs)
     if not ispostback then
	 t_g_1.Items.Insert(0, New ListItem("全部", "%"))
	 't_g_1.selectedindex = 0
	't_g_1_t.value = DataSet8.FieldValue("t_g_1", nothing)
	t_g_1.SelectedIndex = t_g_1.Items.IndexOf(t_g_1.Items.FindByValue(t_g_1_t.value))
	 end if
	 end sub

   Sub t_g_1_cha(sender As Object, e As EventArgs)
	if ispostback then
	't_g_1_t.value = cstr(t_g_1.selecteditem.value)
	'if t_g_1.selectedindex > 0 then
	't_g_2_t.value = cstr(t_g_1.selecteditem.value)  '一改變時.二便要全部
	't_g_2.datasource = DataSet2.DefaultView
	'else
    t_g_2_t.value = "%"
	'end if
	t_g_3.enabled = false
	  t_g_3_t.value = "%"
	t_g_num.enabled = false
	  t_g_num_t.value = "%"

	end if
	 end sub

   Sub t_g_1_load(sender As Object, e As EventArgs)
    if not ispostback then
    t_g_1.SelectedIndex =t_g_1.Items.IndexOf(t_g_1.Items.FindByValue(DataSet8.FieldValue("t_g_1", Nothing)))
    t_g_2.SelectedIndex =t_g_2.Items.IndexOf(t_g_2.Items.FindByValue(DataSet8.FieldValue("t_g_2", Nothing)))
    t_g_3.SelectedIndex =t_g_3.Items.IndexOf(t_g_3.Items.FindByValue(DataSet8.FieldValue("t_g_3", Nothing)))
    t_g_num.SelectedIndex =t_g_num.Items.IndexOf(t_g_num.Items.FindByValue(DataSet8.FieldValue("t_g_num", Nothing)))
    case_man.SelectedIndex =case_man.Items.IndexOf(case_man.Items.FindByValue(DataSet8.FieldValue("admin_id", Nothing)))
	t_g_id_t.value= cint(DataSet8.FieldValue("t_g_id", Nothing))
	end if
   end sub		

   Sub t_g_2_pre(sender As Object, e As EventArgs)
	 'if t_g_1_t.value = t_g_2_t.value then
	 t_g_2.Items.Insert(0, New ListItem("全部", "%"))
	 't_g_2_t.value = cstr(t_g_2.selecteditem.value)
	 't_g_2.SelectedIndex = 2 
	't_g_2.SelectedIndex = 2
    if t_g_1.selectedindex > 0 then 
	t_g_2.SelectedIndex = t_g_2.Items.IndexOf(t_g_2.Items.FindByValue(t_g_2_t.value))
	else
	t_g_2.SelectedIndex = 0
	end if
	't_g_3_t.value = t_g_2.Selectedindex
	 'end if
	 end sub

   Sub t_g_2_on(sender As Object, e As EventArgs)
     if ispostback then
	t_g_2_t.value = t_g_2.selecteditem.value
	't_g_3.datasource = DataSet2.DefaultView
	end if
	 end sub

   Sub t_g_2_cha(sender As Object, e As EventArgs)
	if ispostback then
	't_g_1_t.value = cstr(t_g_1.selecteditem.value)
	'if t_g_1.selectedindex > 0 then
	't_g_2_t.value = cstr(t_g_1.selecteditem.value)  '一改變時.二便要全部
	't_g_2.datasource = DataSet2.DefaultView
	'else
	t_g_3_t.value = "%"
	'end if
	t_g_3.SelectedIndex = 0
	t_g_num.enabled = false
	  t_g_num_t.value = "%"
	end if
	 end sub

   Sub t_g_2_bin(sender As Object, e As EventArgs)
	 if t_g_1_t.value = t_g_2_t.value then
	 t_g_2.datasource = DataSet2.DefaultView
	 't_g_2.SelectedIndex =t_g_2.Items.IndexOf(t_g_2.Items.FindByValue(cstr(t_g_2_t.value)))
	 end if
	 end sub

   Sub t_g_2_un(sender As Object, e As EventArgs)
	 t_g_2.datasource = ""
	 't_g_2.SelectedIndex =t_g_2.Items.IndexOf(t_g_2.Items.FindByValue(cstr(t_g_2_t.value)))
	 end sub

   Sub t_g_3_pre(sender As Object, e As EventArgs)
	 'if t_g_1_t.value = t_g_2_t.value then
	 t_g_3.Items.Insert(0, New ListItem("全部", "%"))
	 't_g_2_t.value = cstr(t_g_2.selecteditem.value)
	 't_g_2.SelectedIndex = 2 
	't_g_2.SelectedIndex = 2
   if not ispostback then
	t_g_3.SelectedIndex = t_g_3.Items.IndexOf(t_g_3.Items.FindByValue(t_g_3_t.value))
	't_g_3.enabled = false
   else 

    if t_g_2.selectedindex > 0 then 
	t_g_3.SelectedIndex = t_g_3.Items.IndexOf(t_g_3.Items.FindByValue(t_g_3_t.value))
	else
	t_g_3.SelectedIndex = 0
	end if
	't_g_3_t.value = t_g_2.Selectedindex
	end if
	 end sub

   Sub t_g_3_cha(sender As Object, e As EventArgs)
	if ispostback then
	't_g_1_t.value = cstr(t_g_1.selecteditem.value)
	'if t_g_1.selectedindex > 0 then
	't_g_3_t.value = cstr(t_g_3.selecteditem.value)  '一改變時.二便要全部
	't_g_2.datasource = DataSet2.DefaultView
	'else
    t_g_num_t.value = "%"
	'end if
	t_g_num.SelectedIndex = 0
	end if
	 end sub

   Sub t_g_num_pre(sender As Object, e As EventArgs)
	 'if t_g_1_t.value = t_g_2_t.value then
	 t_g_num.Items.Insert(0, New ListItem("全部", "%"))
	 't_g_2_t.value = cstr(t_g_2.selecteditem.value)
	 't_g_2.SelectedIndex = 2 
	't_g_2.SelectedIndex = 2
   if not ispostback then
	t_g_num.SelectedIndex = t_g_num.Items.IndexOf(t_g_num.Items.FindByValue(t_g_num_t.value))
	't_g_num.enabled = false
   else 
    if t_g_3.selectedindex > 0 then 
	t_g_num.SelectedIndex = t_g_num.Items.IndexOf(t_g_num.Items.FindByValue(t_g_num_t.value))
	else
	t_g_num.SelectedIndex = 0
	end if
	't_g_3_t.value = t_g_2.Selectedindex
	end if
	 end sub 


  Sub updateData(sender As Object, e As EventArgs) 
    if IsValid Then

      Dim Conn As OleDbConnection
      Dim Cmd  As OleDbCommand

      Dim Provider = "Provider=Microsoft.Jet.OLEDB.4.0"
      Dim Database = "Data Source=" & Server.MapPath( "result/result.mdb" )
      Conn = New OleDbConnection( Provider & ";" & DataBase )
      Conn.Open()

      Dim SQL As String 
      SQL = "update s01insp_section set t_g_id=@t_g_id, admin_id=@admin_id where insp_s_id=@insp_s_id"
      Cmd = New OleDbCommand( SQL, Conn )

      Cmd.Parameters.Add( New OleDbParameter("@t_g_id", OleDbType.integer))
      Cmd.Parameters.Add( New OleDbParameter("@admin_id", OleDbType.integer))
      Cmd.Parameters.Add( New OleDbParameter("@insp_s_id", OleDbType.integer))      
      
	  Cmd.Parameters("@t_g_id").value = t_g_id_t.value
      Cmd.Parameters("@admin_id").value = cint(case_man.selecteditem.value)
      Cmd.Parameters("@insp_s_id").value = cint(insp_s_id.text)

      Cmd.ExecuteNonQuery()
      If Err.Number <> 0 Then
         Msg.text = Err.Description
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
<p class="style1">修改轄區：</p>
<form runat="server" name='form_add' id="form_add">
<table width="529" border="0" cellspacing="1" cellpadding="0">
  <tr>
    <td width="22%">行業別(一)</td>
    <td width="78%"><asp:DropDownList AutoPostBack="true" DataTextField="t_g_name" DataValueField="t_g_1" ID="t_g_1" runat="server" OnSelectedIndexChanged="t_g_1_cha" OnPreRender="t_g_1_pre">

	</asp:DropDownList>
      <input type="hidden" name="hiddenField" ID="t_g_1_t" runat="server"/>
      
      <input name="Submit" type="button" id="in01" onclick="MM_openBrWindow('s01total_g_search.aspx','gsearch','scrollbars=yes,width=500,height=500')" value="查詢行業名稱" runat="server"/>
      <input name="hiddenField" type="hidden" ID="t_g_id_t" runat="server"/>
      <asp:Button EnableViewState="false" ID="Button01" runat="server" Text="確定" /></td>
  </tr>
  <tr>
    <td>行業別(二)</td>
    <td><asp:DropDownList AutoPostBack="true" DataTextField="t_g_name" DataValueField="t_g_2" ID="t_g_2" runat="server" OnSelectedIndexChanged="t_g_2_cha" OnPreRender="t_g_2_pre"></asp:DropDownList>
      <input type="hidden" name="hiddenField" ID="t_g_2_t" runat="server"/></td>
  </tr>
  <tr>
    <td>行業別(三)</td>
    <td><asp:DropDownList AutoPostBack="true" DataTextField="t_g_name" DataValueField="t_g_3" ID="t_g_3" runat="server" OnSelectedIndexChanged="t_g_3_cha" OnPreRender="t_g_3_pre"></asp:DropDownList>
	        <input type="hidden" name="hiddenField" ID="t_g_3_t" runat="server"/>	  </td>
  </tr>
  <tr>
    <td>行業別(四)</td>
    <td><asp:DropDownList AutoPostBack="true" DataTextField="t_g_name" DataValueField="t_g_num" ID="t_g_num" runat="server" OnPreRender="t_g_num_pre"></asp:DropDownList>
	        <input type="hidden" name="hiddenField" ID="t_g_num_t" runat="server"/>	  </td>
  </tr>
  <tr>
    <td width="80" bgcolor="#CCFFFF">承辦人：</td>
    <td bgcolor="#CCCC66"><asp:DropDownList DataTextField="s01name" DataValueField="admin_id" ID="case_man" runat="server">
    </asp:DropDownList></td>
  </tr>

</table>
  <asp:Button ID="Button1" runat="server" Text="修改" OnClick="updateData" />
  <span class="style1">
  <input type="reset" name="Submit2" value="重新填寫" />
  </span>
  <span class="style3">
  <asp:Button ID="Buttonc" runat="server" Text="取消" />  </span>
  <asp:TextBox ID="back" runat="server" Visible="false" Columns="50" />
  
  <asp:TextBox Columns="10" ID="insp_s_id" runat="server" text='<%# DataSet8.FieldValue("insp_s_id", Container) %>' />
  
  <p>
  <HR><asp:ValidationSummary
     DisplayMode="BulletList" EnableClientScript="true" Enabled="false"
     HeaderText="必須輸入的欄位還有:" ID="ch04" runat="server" />
<asp:Label runat="server" id="Msg" ForeColor="Red" />
</form>
</body>
</html>