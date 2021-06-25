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
<MM:PageBind runat="server" PostBackBind="false" />
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=UTF-8" />
<title>新增事業單位資料</title>
<script type="text/JavaScript">
<!--
function MM_goToURL() { //v3.0
  var i, args=MM_goToURL.arguments; document.MM_returnValue = false;
  for (i=0; i<(args.length-1); i+=2) eval(args[i]+".location='"+args[i+1]+"'");
}

function MM_openBrWindow(theURL,winName,features) { //v2.0
  window.open(theURL,winName,features);
}
//-->
</script>
<style type="text/css">
<!--
.style1 {font-size: 12px}
-->
</style>
</head>
<body>
<form method='POST' name='form1' id="form1" runat="server">
  <p>新增事業單位資料：
    <input name="Submit" type="button" onclick="MM_goToURL('parent','s01total_add_search.aspx');return document.MM_returnValue" value="回新增事業單位" runat="server"/>
  </p>
<table width="100%" border="1" cellspacing="0" cellpadding="0">
  <tr>
    <td width="10%" bgcolor="#FFCCCC">勞保證號：</td>
    <td width="47%" bgcolor="#CCFFFF"><asp:TextBox ID="t_insu" runat="server" />
    <td width="12%" bgcolor="#FFCCCC">統一編號：</td>
    <td width="31%" bgcolor="#CCFFFF"><asp:TextBox ID="t_pre" ReadOnly="true" runat="server" /></td>
  </tr>

  <tr>
    <td width="10%" bgcolor="#FFCCCC">單位名稱：</td>
    <td width="47%" bgcolor="#CCFFFF"><asp:TextBox ID="t_name" runat="server" Columns="50"/>
    必填</td>
    <td width="12%" bgcolor="#FFCCCC">負責人：</td>
    <td width="31%" bgcolor="#CCFFFF"><asp:TextBox ID="t_boss" runat="server" /></td>
  </tr>
  <tr>
    <td bgcolor="#FFCCCC">地址：</td>
    <td bgcolor="#CCFFFF"><span class="style1">高雄市</span>
       <asp:DropDownList AutoPostBack="true" DataSource="<%# DataSet2.DefaultView %>" DataTextField="cs_name" DataValueField="cs_id" ID="p_address1" runat="server" OnSelectedIndexChanged="change_text1"></asp:DropDownList>

	<asp:TextBox ID="t_address" runat="server" Columns="40"/>
	必填</td>
    <td bgcolor="#FFCCCC">電話：</td>
    <td bgcolor="#CCFFFF"><asp:TextBox ID="t_tel" runat="server" /></td>
  </tr>
  <tr>
    <td bgcolor="#FFCCCC">行業別：</td>
    <td bgcolor="#CCFFFF" colspan="3">    
      <input name="Submit2" type="button" onclick="MM_openBrWindow('s01insp_t_add_gsearch.aspx','addgsearch','scrollbars=yes,width=500,height=500')" value="查詢行業別" />
<asp:TextBox Columns="16" ID="t_g_name" runat="server" />
      必選
      <asp:TextBox BackColor="#CCFFFF" BorderColor="#CCFFFF" BorderStyle="None" Columns="5" ForeColor="#CCFFFF" ID="t_g_num" runat="server" /></td>
  </tr>
  <tr>
    <td bgcolor="#FFCCCC">e-mail：</td>
    <td bgcolor="#CCFFFF"><asp:TextBox ID="t_email" runat="server" Columns="50"/></td>
    <td bgcolor="#FFCCCC">勞工人數：</td>
    <td bgcolor="#CCFFFF"><asp:TextBox ID="t_per" runat="server" />
    必填</td>
  </tr>
  <tr>
    <td bgcolor="#FFCCCC">勞安人員：</td>
    <td bgcolor="#CCFFFF"><asp:TextBox ID="t_sh" runat="server"/></td>
    <td bgcolor="#FFCCCC">手機：</td>
    <td bgcolor="#CCFFFF"><asp:TextBox ID="t_shcell" runat="server" /></td>
  </tr>
  <tr>
    <td bgcolor="#FFCCCC">受檢場所：</td>
    <td bgcolor="#CCFFFF" colspan="3"><asp:TextBox ID="t_place" runat="server"/></td>
  </tr>
  <tr>
    <td bgcolor="#FFCCCC">相關危害：</td>
    <td bgcolor="#CCFFFF" colspan="3"><asp:TextBox ID="t_hazard" runat="server" TextMode="MultiLine" Columns="120" Rows="6" /></td>
  </tr>
  <tr>
    <td width="10%" bgcolor="#FFCCCC">輸入日期</td>
    <td width="47%" bgcolor="#CCFFFF"><asp:TextBox ID="t_update_time" runat="server" Text='<%# showdate %>' Visible="false" /></td>
    <td width="12%" bgcolor="#FFCCCC">登錄者：</td>
    <td width="31%" bgcolor="#CCFFFF"><asp:TextBox ID="t_keyin" ReadOnly="true" runat="server" /></td>
  </tr>

</table>
  <asp:Button ID="Button1" runat="server" Text="送出" OnClick="InsertData" />
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

 function showdate()
 showdate = FormatDateTime(now, DateFormat.ShortDate)
 end function

   Sub InsertData(sender As Object, e As EventArgs) 
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
      SQL = "Insert Into s01total (t_address, t_email, t_g_num, t_hazard, t_boss, t_name, t_insu, t_place, t_per, t_sh, t_shcell, t_tel, t_keyin, t_update_time, t_pre) Values(@t_address, @t_email, @t_g_num, @t_hazard, @t_boss, @t_name, @t_insu, @t_place, @t_per, @t_sh, @t_shcell, @t_tel, @t_keyin, @t_update_time, @t_pre)"
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

      Cmd.ExecuteNonQuery()
      If Err.Number <> 0 Then
         Msg.Text = Err.Description
      Else
         response.Redirect("s01total_add_search.aspx")
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
