<%@ Page Language="VB" ContentType="text/html" ResponseEncoding="utf-8" %>
<%@ Register TagPrefix="MM" Namespace="DreamweaverCtrls" Assembly="DreamweaverCtrls,version=1.0.0.0,publicKeyToken=836f606ede05d46a,culture=neutral" %>

<MM:DataSet 
id="DataSet1"
runat="Server"
IsStoredProcedure="false"
ConnectionString='<%# System.Configuration.ConfigurationSettings.AppSettings("MM_CONNECTION_STRING_s01") %>'
DatabaseType='<%# System.Configuration.ConfigurationSettings.AppSettings("MM_CONNECTION_DATABASETYPE_s01") %>'
CommandText='<%# "SELECT distinct t_g_id, t_g_1, t_g_name, t_g_2 FROM t_group where t_g_2 like ? or t_g_2 is null ORDER BY t_g_1 ASC" %>'
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
<Parameter  Name="@t_g_1"  Value='<%# IIf(cstr(t_g_1_t.value) = "", "%", cstr(t_g_1_t.value)) %>'  Type="WChar"   />
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
<Parameter  Name="@t_g_2"  Value='<%# IIf(cstr(t_g_2_t.value) = "", "%", cstr(t_g_2_t.value)) %>'  Type="WChar"   />
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
<Parameter  Name="@t_g_3"  Value='<%# IIf(cstr(t_g_3_t.value) = "", "%", cstr(t_g_3_t.value)) %>'  Type="WChar"   />
</Parameters>
</MM:DataSet>

<MM:DataSet 
id="DataSet4"
runat="Server"
IsStoredProcedure="false"
ConnectionString='<%# System.Configuration.ConfigurationSettings.AppSettings("MM_CONNECTION_STRING_s01") %>'
DatabaseType='<%# System.Configuration.ConfigurationSettings.AppSettings("MM_CONNECTION_DATABASETYPE_s01") %>'
CommandText='<%# "SELECT * FROM s01cs ORDER BY cs_num ASC" %>'
Debug="true"
></MM:DataSet>
<MM:DataSet 
id="DataSet6"
runat="Server"
IsStoredProcedure="false"
ConnectionString='<%# System.Configuration.ConfigurationSettings.AppSettings("MM_CONNECTION_STRING_s01") %>'
DatabaseType='<%# System.Configuration.ConfigurationSettings.AppSettings("MM_CONNECTION_DATABASETYPE_s01") %>'
CommandText='<%# "SELECT * FROM s01_insp_s ORDER BY s01name ASC" %>'
Debug="true"
></MM:DataSet>
<MM:PageBind runat="server" PostBackBind="true" />
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>綜合行業科檢查名冊查詢</title>
<script language="VB" runat="server">
   Sub Page_Load(sender As Object, e As EventArgs)
		If Session("MM_username")="" then
			Response.Write("您還沒有登入呢！，請點擊")
			Response.Write("<a href=s01login.aspx>登入</a>")
			Response.End()
		End If
    'call t_g_1_load()
    'call t_g_2_load()
    if not ispostback then
    'session("g_result")= nothing
	   't_g_1.SelectedIndex = t_g_1.Items.IndexOf(t_g_1.Items.FindByValue(t_g_1_t.value))
	 session("total_cancel") = request.url.tostring()
	't_g_1.items.add("全部")
	t_g_id_t.value = ""
	t_g_1.datasource = DataSet1.DefaultView
	t_address.datasource = DataSet4.DefaultView
	t_g_2.datasource = DataSet2.DefaultView
	t_g_3.datasource = DataSet3.DefaultView
	t_g_num.datasource = DataSet5.DefaultView
	't_g_2.datasource = DataSet2.DefaultView
	t_g_1_t.value = "%"
	t_g_2_t.value = "%"
	t_g_3_t.value = "%"
	t_g_num_t.value = "%"
	't_g_1.DataTextField = t_g_name
	't_g_1.DataValueField = DataSet1.FieldValue("t_g_1")
	t_g_2.enabled = false
	t_g_3.enabled = false
	t_g_num.enabled = false
	else
	'if session("g_result") > 0 then '若用行業查詢則顯示的內容
	't_g_1.datasource = DataSet1.DefaultView
	't_address.datasource = DataSet4.DefaultView
	't_g_2.datasource = DataSet2.DefaultView
	't_g_3.datasource = DataSet3.DefaultView
	't_g_num.datasource = DataSet5.DefaultView
	't_g_id_t.value = session("g_result")
	't_g_2.enabled = false
	't_g_3.enabled = false
	't_g_num.enabled = false
	't_g_2_t.value = cstr(t_g_2.selecteditem.value) '放在page load因為此會先執行
	't_g_3_t.value = cstr(t_g_3.selecteditem.value) '放在page load因為此會先執行
	't_g_num_t.value = cstr(t_g_num.selecteditem.value) '放在page load因為此會先執行
	't_g_2.enabled = false
	't_g_3.enabled = false
	't_g_num.enabled = false
	'exit sub
	'end if

	if t_g_id_t.value <> "" then
	't_g_1.datasource = DataSet7.DefaultView
	't_address.datasource = DataSet4.DefaultView
	't_g_2.enabled = true
	't_g_3.enabled = true
	't_g_2.datasource = DataSet2.DefaultView
	't_g_3.datasource = DataSet3.DefaultView
	't_g_num.datasource = DataSet7.DefaultView

	't_g_1.datasource = DataSet1.DefaultView
	t_address.datasource = DataSet4.DefaultView
	t_g_2.datasource = DataSet2.DefaultView
	t_g_3.datasource = DataSet3.DefaultView
	t_g_num.datasource = DataSet5.DefaultView

	t_g_1.SelectedIndex = t_g_1.Items.IndexOf(t_g_1.Items.FindByValue(t_g_1_t.value))
	't_g_2.SelectedIndex = t_g_2.Items.IndexOf(t_g_2.Items.FindByValue(t_g_2_t.value))
	't_g_3.SelectedIndex = t_g_3.Items.IndexOf(t_g_3.Items.FindByValue(t_g_3_t.value))
	't_g_num.SelectedIndex = t_g_num.Items.IndexOf(t_g_num.Items.FindByValue(t_g_num_t.value))
	t_g_2.enabled = false
	t_g_3.enabled = false
	t_g_num.enabled = false
	'button01.enabled = false

	t_g_id_t.value = ""
	't_g_id_t.value = session("g_result")
	't_g_2.enabled = false
	't_g_3.enabled = false
	't_g_num.enabled = false
	exit sub

	else
	t_g_1_t.value = cstr(t_g_1.selecteditem.value) '放在page load因為此會先執行
	end if
	't_g_2_t.value = cstr(t_g_2.selecteditem.value) '放在page load因為此會先執行
	't_g_3_t.value = cstr(t_g_3.selecteditem.value) '放在page load因為此會先執行
	't_g_num_t.value = cstr(t_g_num.selecteditem.value) '放在page load因為此會先執行
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
        End Sub

   Sub t_g_1_pre(sender As Object, e As EventArgs)
     't_g_id_t.value = ""
	 if not ispostback then
	 t_g_1.Items.Insert(0, New ListItem("全部", "%"))
	 t_g_1.selectedindex = 0
	 't_g_1_t.value = cstr(t_g_1.selecteditem.value)
	 end if
	 end sub

   Sub t_g_1_cha(sender As Object, e As EventArgs)
	if ispostback then
	 t_g_id_t.value = ""
     't_g_name.value = ""
	t_g_1_t.value = cstr(t_g_1.selecteditem.value)
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
    
	 'end if
	end if
	 end sub

   Sub t_g_1_t_cha(sender As Object, e As EventArgs)
	if t_g_id_t.value <> "" then
	t_g_1.SelectedIndex = t_g_1.Items.IndexOf(t_g_1.Items.FindByValue(t_g_1_t.value))
	t_g_2.SelectedIndex = t_g_2.Items.IndexOf(t_g_2.Items.FindByValue(t_g_2_t.value))
	t_g_3.SelectedIndex = t_g_3.Items.IndexOf(t_g_3.Items.FindByValue(t_g_3_t.value))

    t_g_2.enabled = false
    t_g_3.enabled = false
    t_g_id_t.value = ""
	exit sub
	else
	t_g_1_t.value = cstr(t_g_1.selecteditem.value) '放在page load因為此會先執行
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

   Sub t_g_3_pre(sender As Object, e As EventArgs)
	 'if t_g_1_t.value = t_g_2_t.value then
	 t_g_3.Items.Insert(0, New ListItem("全部", "%"))
	 't_g_2_t.value = cstr(t_g_2.selecteditem.value)
	 't_g_2.SelectedIndex = 2 
	't_g_2.SelectedIndex = 2
	if t_g_2.selectedindex > 0 then 
	t_g_3.SelectedIndex = t_g_3.Items.IndexOf(t_g_3.Items.FindByValue(t_g_3_t.value))
	else
	t_g_3.SelectedIndex = 0
	end if
	't_g_3_t.value = t_g_2.Selectedindex
	 'end if
	 end sub

   Sub t_g_3_cha(sender As Object, e As EventArgs)
	if ispostback then
	't_g_1_t.value = cstr(t_g_1.selecteditem.value)
	'if t_g_1.selectedindex > 0 then
	't_g_2_t.value = cstr(t_g_1.selecteditem.value)  '一改變時.二便要全部
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
	if t_g_3.selectedindex > 0 then 
	t_g_num.SelectedIndex = t_g_num.Items.IndexOf(t_g_num.Items.FindByValue(t_g_num_t.value))
	else
	t_g_num.SelectedIndex = 0
	end if
	't_g_3_t.value = t_g_2.Selectedindex
	 end sub

   Sub starsearch(sender As Object, e As EventArgs) 
	 session("total_cancel") = request.url.tostring()
	  dim url
	  url = "s01insp_total_list0.aspx?t_g_1=" & t_g_1_t.value & "&t_g_2=" & t_g_2_t.value & "&t_g_3=" & t_g_3_t.value & "&t_g_num=" & t_g_num_t.value & "&t_address=" & t_address.selecteditem.text & t_address0.text & "&t_name=" & t_name.text & "&t_pre=" & t_pre.text & "&t_person1=" & t_person1.text & "&t_person2=" & t_person2.text 
    session("g_result")= nothing
	  response.Redirect( url ) ' 使用Server.Transfer亦可
   End Sub

   Sub star_g_search(sender As Object, e As EventArgs) 
	t_g_1_t.value = "%"
	t_g_2_t.value = "%"
	t_g_3_t.value = "%"
	t_g_num_t.value = "%"
	t_g_2.enabled = false
	t_g_3.enabled = false
	t_g_num.enabled = false

	't_g_1.datasource = DataSet2.DefaultView
	't_g_2.datasource = DataSet2.DefaultView
	't_g_3.datasource = DataSet3.DefaultView
	't_g_num.datasource = DataSet5.DefaultView
   End Sub

   Sub star_total_add(sender As Object, e As EventArgs) 
   dim url01
   url01= "s01total_add_search.aspx"
   response.Redirect(url01)
   End Sub


</script>
<script type="text/JavaScript">
<!--

function MM_goToURL() { //v3.0
  var i, args=MM_goToURL.arguments; document.MM_returnValue = false;
  for (i=0; i<(args.length-1); i+=2) eval(args[i]+".location='"+args[i+1]+"'");
}
function DelTitle(s01id) {//確認是否要刪除該記錄
	if (confirm('您確實要刪除該主題嗎？' + '\r\r' +
		'注意：如果您刪除該主題，' + '\r' +
		'哪麼該主題下所有的資料' + '\r' +
		'就將全部被刪除！')){
	window.location ='s01case_del.aspx?case_id=' + s01id;
    }
}

function MM_openBrWindow(theURL,winName,features) { //v2.0
  window.open(theURL,winName,features);
}
//-->
</script>
</head>
<body>
<form name='formts' id="formts" runat="server">
<table width="80%" border="1">
  <tr>
    <td width="22%">行業別(一)</td>
    <td width="78%"><asp:DropDownList AutoPostBack="true" DataTextField="t_g_name" DataValueField="t_g_1" ID="t_g_1" runat="server" OnSelectedIndexChanged="t_g_1_cha" OnPreRender="t_g_1_pre">

	</asp:DropDownList>
      <input name="hiddenField" type="hidden" ID="t_g_1_t" runat="server"/>
      
      <input name="Submit" type="button" id="in01" onclick="MM_openBrWindow('s01total_g_search.aspx','gsearch','scrollbars=yes,width=500,height=500')" value="查詢行業名稱" runat="server"/>
      <input type="hidden" name="hiddenField" ID="t_g_id_t" runat="server"/>
      <asp:Button Enabled="true" EnableViewState="true" ID="Button01" runat="server" Text="" Visible="true" /></td>
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
<td width="14%">行政區：</td>
      <td width="86%">      <asp:DropDownList DataTextField="cs_name" DataValueField="cs_id" ID="t_address" runat="server" ></asp:DropDownList>
        道路或地址：
        <asp:TextBox ID="t_address0" Columns="30" runat="server" /></td>
    </tr>
    <tr>
      <td>事業單位名稱：</td>
      <td><asp:TextBox ID="t_name" runat="server" Columns="30" />
        (可輸入事業單位部分名稱)</td>
    </tr>
    <tr>
      <td>統一編號：</td>
      <td><asp:TextBox ID="t_pre" runat="server" Columns="20" />
        (<span class="style2">8碼</span>)</td>
    </tr>
    <tr>
      <td>人數：</td>
      <td><asp:TextBox ID="t_person1" runat="server" Width="40" />
        ~
        <asp:TextBox ID="t_person2" runat="server" Width="40" /></td>
    </tr>
</table>
<p><br />
    <asp:Button ID="insp_search" runat="server" Text="查詢" OnClick="starsearch" />    
    <asp:Button ID="total_add_search" runat="server" Text="新增事業單位" OnClick="star_total_add" />    
</p>
<p>轄區檢查名冊</p>
<asp:DataGrid 
  AllowSorting="False" 
  AutoGenerateColumns="false" 
  CellPadding="3" 
  CellSpacing="0" 
  DataSource="<%# DataSet6.DefaultView %>" id="DataGrid6" 
  PagerStyle-Mode="NextPrev" 
  runat="server" 
  ShowFooter="false" 
  ShowHeader="true" Width="80%" 
>
    <HeaderStyle HorizontalAlign="center" BackColor="#E8EBFD" ForeColor="#3D3DB6" Font-Name="Verdana, Arial, Helvetica, sans-serif" Font-Bold="true" Font-Size="smaller" />
    <ItemStyle BackColor="#F2F2F2" Font-Name="Verdana, Arial, Helvetica, sans-serif" Font-Size="smaller" />
    <AlternatingItemStyle BackColor="#E5E5E5" Font-Name="Verdana, Arial, Helvetica, sans-serif" Font-Size="smaller" />
    <FooterStyle HorizontalAlign="center" BackColor="#E8EBFD" ForeColor="#3D3DB6" Font-Name="Verdana, Arial, Helvetica, sans-serif" Font-Bold="true" Font-Size="smaller" />
    <PagerStyle BackColor="white" Font-Name="Verdana, Arial, Helvetica, sans-serif" Font-Size="smaller" />
    <Columns>
      <asp:BoundColumn DataField="s01name" 
        HeaderText="姓名" 
        ReadOnly="true" 
        Visible="True"/>
      <asp:HyperLinkColumn DataNavigateUrlField="t_g_id" DataNavigateUrlFormatString="s01insp_total_list0pre.aspx?t_g_id={0}"
        DataTextField="t_g_name" 
        Visible="True" 
        HeaderText="負責行業" 
		/>      
<asp:BoundColumn DataField="t_g_1" 
        HeaderText="行業別(一)" 
        ReadOnly="true" 
        Visible="True"/>
<asp:BoundColumn DataField="t_g_2" 
        HeaderText="行業別(二)" 
        ReadOnly="true" 
        Visible="True"/>
<asp:BoundColumn DataField="t_g_3" 
        HeaderText="行業別(三)" 
        ReadOnly="true" 
        Visible="True"/>
<asp:BoundColumn DataField="t_g_num" 
        HeaderText="行業別(四)" 
        ReadOnly="true" 
        Visible="True"/>
    </Columns>
  </asp:DataGrid>
</form>
</body>
</html>
