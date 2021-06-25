<%@ Page Language="VB" ContentType="text/html"%>
<%@ Register TagPrefix="MM" Namespace="DreamweaverCtrls" Assembly="DreamweaverCtrls,version=1.0.0.0,publicKeyToken=836f606ede05d46a,culture=neutral" %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Data.OleDb" %>
<MM:DataSet 
id="DataSet2"
runat="Server"
IsStoredProcedure="false"
ConnectionString='<%# System.Configuration.ConfigurationSettings.AppSettings("MM_CONNECTION_STRING_s01") %>'
DatabaseType='<%# System.Configuration.ConfigurationSettings.AppSettings("MM_CONNECTION_DATABASETYPE_s01") %>'
CommandText='<%# "SELECT * FROM s01cs ORDER BY cs_id ASC" %>'
Debug="true"
></MM:DataSet>
<MM:DataSet 
id="DataSet1"
runat="Server"
IsStoredProcedure="false"
ConnectionString='<%# System.Configuration.ConfigurationSettings.AppSettings("MM_CONNECTION_STRING_s01") %>'
DatabaseType='<%# System.Configuration.ConfigurationSettings.AppSettings("MM_CONNECTION_DATABASETYPE_s01") %>'
CommandText='<%# "SELECT * FROM s01_total WHERE t_id = ?" %>'
Debug="true"
><Parameters>
<Parameter  Name="@t_id"  Value='<%# IIf((Request.QueryString("t_id") <> Nothing), Request.QueryString("t_id"), "") %>'  Type="Integer"   /></Parameters></MM:DataSet>
<MM:PageBind runat="server" PostBackBind="false" />
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=UTF-8" />
<title>修改事業單位資料</title>
<script type="text/JavaScript">
<!--

function MM_openBrWindow(theURL,winName,features) { //v2.0
  window.open(theURL,winName,features);
}
//-->
</script>
<style type="text/css">
<!--
.style1 {color: #FF0000}
-->
</style>
</head>
<body>
修改事業單位資料：
    
  <form method='POST' name='form1' id="form1" runat="server">
  <table width="100%" border="1" cellspacing="0" cellpadding="0">
  <tr>
    <td width="10%" bgcolor="#99FF99">勞保證號：</td>
    <td width="47%" bgcolor="#FFFF99"><asp:TextBox ID="t_insu" runat="server" text='<%# DataSet1.FieldValue("t_insu", Container) %>' />
    <td width="12%" bgcolor="#99FF99">統一編號：</td>
    <td width="31%" bgcolor="#FFFF99"><asp:TextBox ID="t_pre" ReadOnly="true" runat="server" text='<%# DataSet1.FieldValue("t_pre", Container) %>' />
      <asp:TextBox ID="t_id" ReadOnly="true" runat="server" text='<%# DataSet1.FieldValue("t_id", Container) %>' Columns="8" /></td>
  </tr>

  <tr>
    <td width="10%" bgcolor="#99FF99">單位名稱：</td>
    <td width="47%" bgcolor="#FFFF99"><asp:TextBox Columns="50" ID="t_name" runat="server" text='<%# DataSet1.FieldValue("t_name", Container) %>'/>
      <span class="style1">必填</span></td>
    <td width="12%" bgcolor="#99FF99">負責人：</td>
    <td width="31%" bgcolor="#FFFF99"><asp:TextBox ID="t_boss" runat="server" text='<%# DataSet1.FieldValue("t_boss", Container) %>' /></td>
  </tr>
  <tr>
    <td bgcolor="#99FF99">地址：</td>
    <td bgcolor="#FFFF99">高雄市
       <asp:DropDownList AutoPostBack="true" DataSource="<%# DataSet2.DefaultView %>" DataTextField="cs_name" DataValueField="cs_id" ID="p_address1" runat="server" OnSelectedIndexChanged="change_text1"></asp:DropDownList>

	<asp:TextBox Columns="40" ID="t_address" runat="server" text='<%# DataSet1.FieldValue("t_address", Container) %>'/>
	<span class="style1">必填</span></td>
    <td bgcolor="#99FF99">電話：</td>
    <td bgcolor="#FFFF99"><asp:TextBox ID="t_tel" runat="server" text='<%# DataSet1.FieldValue("t_tel", Container) %>' /></td>
  </tr>
  <tr>
    <td bgcolor="#99FF99">行業別：</td>
    <td bgcolor="#FFFF99" colspan="3">    
      <input name="Submit2" type="button" onclick="MM_openBrWindow('s01insp_t_add_gsearch.aspx','addgsearch','scrollbars=yes,width=500,height=500')" value="查詢行業別" />
<asp:TextBox Columns="16" ID="t_g_name" runat="server" text='<%# DataSet1.FieldValue("t_group.t_g_name", Container) %>'/>
      <span class="style1">必選</span>
      <asp:TextBox BackColor="#FFFF99" BorderColor="#FFFF99" BorderStyle="None" Columns="5" ForeColor="#FFFF99" ID="t_g_num" runat="server" text='<%# DataSet1.FieldValue("t_group.t_g_num", Container) %>' /></td>
  </tr>
  <tr>
    <td bgcolor="#99FF99">e-mail：</td>
    <td bgcolor="#FFFF99"><asp:TextBox Columns="50" ID="t_email" runat="server" text='<%# DataSet1.FieldValue("t_email", Container) %>'/></td>
    <td bgcolor="#99FF99">勞工人數：</td>
    <td bgcolor="#FFFF99"><asp:TextBox ID="t_per" runat="server" text='<%# DataSet1.FieldValue("t_per", Container) %>' />
      <span class="style1">必填</span></td>
  </tr>
  <tr>
    <td bgcolor="#99FF99">勞安人員：</td>
    <td bgcolor="#FFFF99"><asp:TextBox ID="t_sh" runat="server" text='<%# DataSet1.FieldValue("t_sh", Container) %>'/></td>
    <td bgcolor="#99FF99">手機：</td>
    <td bgcolor="#FFFF99"><asp:TextBox ID="t_shcell" runat="server" text='<%# DataSet1.FieldValue("t_shcell", Container) %>' /></td>
  </tr>
  <tr>
    <td bgcolor="#99FF99">受檢場所：</td>
    <td bgcolor="#FFFF99" colspan="3"><asp:TextBox ID="t_place" runat="server" text='<%# DataSet1.FieldValue("t_place", Container) %>'/></td>
  </tr>
  <tr>
    <td bgcolor="#99FF99">相關危害：</td>
    <td bgcolor="#FFFF99" colspan="3"><asp:TextBox Columns="120" ID="t_hazard" Rows="6" runat="server" text='<%# DataSet1.FieldValue("t_hazard", Container) %>' TextMode="MultiLine" /></td>
  </tr>
  <tr>
    <td width="10%" bgcolor="#99FF99">輸入日期</td>
    <td width="47%" bgcolor="#FFFF99"><asp:TextBox ID="t_update_time" runat="server" Text='<%# showdate %>' Visible="false" /></td>
    <td width="12%" bgcolor="#99FF99">登錄者：</td>
    <td width="31%" bgcolor="#FFFF99"><asp:TextBox ID="t_keyin" ReadOnly="true" runat="server" /></td>
  </tr>

</table>
  <asp:Button ID="Button1" runat="server" Text="送出" OnClick="updateData" />
  <asp:Button ID="Button2" runat="server" Text="取消" OnClick="cancel_insp" />
  
<HR>
  <asp:Label ForeColor="Red" id="Msg" runat="server" />
</form>

<p>&nbsp;</p>
</body>
</html>
<script Language="VB" runat="server">
   Sub page_load(sender As Object, e As EventArgs) 
		If Session("MM_username")="" then
			Response.Write("您還沒有登入呢！，請點擊")
			Response.Write("<a href=s01insp_list_login.aspx>登入</a>")
			Response.End()
		End If
if ispostback then
exit sub
end if
t_pre.text = session("MM_p_num")
t_keyin.text = session("MM_s01name") 
   End Sub

   Sub change_text1(sender As Object, e As EventArgs) 
t_address.text = "高雄市" & p_address1.selecteditem.text 

   End Sub

   Sub cancel_insp(sender As Object, e As EventArgs) 
   response.Redirect(session("cancel_insp"))
   End Sub


 function showdate()
 showdate = FormatDateTime(now, DateFormat.ShortDate)
 end function

   Sub updateData(sender As Object, e As EventArgs) 
    dim x01
    x01 = ""
	msg.text = x01
	dim a() = {t_pre.text, t_address.text, t_name.text, t_g_num.text, t_per.text}
	dim b() = {"統一編號", "地址", "名稱", "行業別", "勞工人數"}
	dim i, j
	j = array.indexof(a, "")
	if j >= 0 then
	x01 = "尚未輸入:" & "<br>"
	for i = 0 to ubound(a)
	if a(i) = "" then
	x01 &= b(i)
	end if
	next
    if isnumeric(t_per.text) = false then
	x01 = x01 & "<br>" & "勞工人數請輸入正確數字"
     else
    if val(t_per.text) < 0 then 
	x01 = x01 & "<br>" & "勞工人數請輸入正確數字"
	 end if
	end if
	msg.text = x01
    exit sub
	end if
    if IsValid Then

      Dim Conn As OleDbConnection
      Dim Cmd  As OleDbCommand

      Dim Provider = "Provider=Microsoft.Jet.OLEDB.4.0"
      Dim Database = "Data Source=" & Server.MapPath( "result/result.mdb" )
      Conn = New OleDbConnection( Provider & ";" & DataBase )
      Conn.Open()

      Dim SQL As String
      SQL = "update s01total set t_address=@t_address, t_email=@t_email, t_g_num=@t_g_num, t_hazard=@t_hazard, t_boss=@t_boss, t_name=@t_name, t_insu=@t_insu, t_place=@t_place, t_per=@t_per, t_sh=@t_sh, t_shcell=@t_shcell, t_tel=@t_tel, t_keyin=@t_keyin, t_update_time=@t_update_time, t_pre=@t_pre where t_id=@tid "
      Cmd = New OleDbCommand( SQL, Conn )

      Cmd.Parameters.Add( New OleDbParameter("@t_address", OleDbType.Char, 200))
      Cmd.Parameters.Add( New OleDbParameter("@t_email", OleDbType.Char, 80))
      Cmd.Parameters.Add( New OleDbParameter("@t_g_num", OleDbType.Char, 10))
      Cmd.Parameters.Add( New OleDbParameter("@t_hazard", OleDbType.VarChar))
      Cmd.Parameters.Add( New OleDbParameter("@t_boss", OleDbType.Char, 20))
      Cmd.Parameters.Add( New OleDbParameter("@t_name", OleDbType.Char, 80))
      Cmd.Parameters.Add( New OleDbParameter("@t_insu", OleDbType.Char, 20))
      Cmd.Parameters.Add( New OleDbParameter("@t_place", OleDbType.Char, 200))
      Cmd.Parameters.Add( New OleDbParameter("@t_per", OleDbType.smallint))
      Cmd.Parameters.Add( New OleDbParameter("@t_sh", OleDbType.Char, 50))
      Cmd.Parameters.Add( New OleDbParameter("@t_shcell", OleDbType.Char, 50))
      Cmd.Parameters.Add( New OleDbParameter("@t_tel", OleDbType.Char, 30))
      Cmd.Parameters.Add( New OleDbParameter("@t_keyin", OleDbType.Char, 50))
      Cmd.Parameters.Add( New OleDbParameter("@t_update_time", OleDbType.Char, 50))
      Cmd.Parameters.Add( New OleDbParameter("@t_pre", OleDbType.Char, 10))
      Cmd.Parameters.Add( New OleDbParameter("@t_id", OleDbType.smallint))

      Cmd.Parameters("@t_address").value = t_address.text
      Cmd.Parameters("@t_email").value = t_email.text
      Cmd.Parameters("@t_g_num").value = t_g_num.text
      Cmd.Parameters("@t_hazard").value = t_hazard.text
      Cmd.Parameters("@t_boss").value = t_boss.text
      Cmd.Parameters("@t_name").value = t_name.text
      Cmd.Parameters("@t_insu").value = t_insu.text
      Cmd.Parameters("@t_place").value = t_place.text
      Cmd.Parameters("@t_per").value = cint(t_per.text)
      Cmd.Parameters("@t_sh").value = t_sh.text
      Cmd.Parameters("@t_shcell").value = t_shcell.text
      Cmd.Parameters("@t_tel").value = t_tel.text
      Cmd.Parameters("@t_keyin").value = t_keyin.text
      Cmd.Parameters("@t_update_time").value = t_update_time.text
      Cmd.Parameters("@t_pre").value = t_pre.text
      Cmd.Parameters("@t_id").value = cint(t_id.text)

      Cmd.ExecuteNonQuery()
      If Err.Number <> 0 Then
         Msg.Text = Err.Description
      Else
         response.Redirect(session("cancel_insp"))
      End If

      Conn.Close()
      End If
   End Sub

</script>

<script type="text/JavaScript">
function Mcheck(){
    if (document.form1.t_name.value=="") {
        window.alert("請輸入單位名稱");
        return false }
    if (document.form1.t_address.value=="") {
        window.alert("請輸入地址");
        return false }
    if (document.form1.t_g_num.value=="") {
        window.alert("請點選行業別");
        return false }
     return true;
}
//-->
</script>
