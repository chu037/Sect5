<%@ Page Language="VB" ContentType="text/html" ResponseEncoding="utf-8" %>
<%@ Register TagPrefix="MM" Namespace="DreamweaverCtrls" Assembly="DreamweaverCtrls,version=1.0.0.0,publicKeyToken=836f606ede05d46a,culture=neutral" %>
<MM:DataSet 
id="DataSet1"
runat="Server"
IsStoredProcedure="false"
ConnectionString='<%# System.Configuration.ConfigurationSettings.AppSettings("MM_CONNECTION_STRING_s01") %>'
DatabaseType='<%# System.Configuration.ConfigurationSettings.AppSettings("MM_CONNECTION_DATABASETYPE_s01") %>'
CommandText='<%# "SELECT distinct case_id, case_content, case_end, case_end_yes, case_group, case_result, case_start, s01name FROM s01_case WHERE case_content like ? order by case_start desc, case_id desc" %>'
Debug="true"
><Parameters>
<Parameter  Name="@case_content"  Value='<%#  "%" +  (IIf((Request.QueryString("case_content") <> Nothing), stra1(Request.QueryString("case_content")), "")) + "%" %>'  Type="WChar" /></Parameters></MM:DataSet>
<MM:PageBind runat="server" PostBackBind="true" />
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>案件查詢</title>
<script language="VB" runat="server">
  Sub page_load(sender As Object, e As EventArgs)
		If Session("MM_chief_num")="" then
			Response.Write("您還沒有登入呢！，請點擊")
			Response.Write("<a href=s01dellogin.aspx>登入</a>")
			Response.End()
		End If
  if not Ispostback Then
    case_start.text = request("case_start")
    case_content.text = stra1(request("case_content"))
	if case_start.text = "" then
      response.Redirect("s01case_admin.aspx")
    else
	back.text = session("cancel_case_admin")
	 end if
	end if
  end sub

Sub startcancel(sender As Object, e As EventArgs) 
 response.Redirect(back.text)
end sub
  
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

function trans(str01)
 if str01 <> ""
  dim i as integer
  dim stra
  stra = ""
   for i = 1 to len(str01)-1
	stra &= asc(mid(str01,i,1)) & ","
   next
    if i = len(str01)
	stra = stra & asc(mid(str01,i,1))
	end if
   trans = stra
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

</script>
<script type="text/JavaScript">
<!--
function MM_openBrWindow(theURL,winName,features) { //v2.0
  window.open(theURL,winName,features);
}

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
</script>

</head>
<body>
<form runat="server" name='form_addchk' id="form_addchk">
<table width="529" border="0" cellspacing="1" cellpadding="0">
  <tr>
    <td width="80" bgcolor="#CCFFFF">案件日期：</td>
    <td width="446" bgcolor="#CCCC66"><asp:TextBox ID="case_start" ReadOnly="true" runat="server" Wrap="false" /></td>
  </tr>
  <tr>
    <td width="80" bgcolor="#CCFFFF">案件內容：</td>
    <td width="446" bgcolor="#CCCC66">
	<asp:TextBox ID="case_content" ReadOnly="true" runat="server" /></td>
  </tr>
</table>
<span class="style3">
<asp:Button ID="Buttonc" runat="server" Text="取消" OnClick='startcancel' />  </span>
  <asp:TextBox ID="back" runat="server" Visible="false" Columns="50" /><p>
  請點選欲新增之案件:</p>
<asp:DataGrid 
  AllowPaging="false" 
  AllowSorting="true" 
  AutoGenerateColumns="false" 
  CellPadding="3" 
  CellSpacing="0" 
  DataSource="<%# DataSet1.DefaultView %>" id="DataGrid1" 
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
                  <asp:TemplateColumn HeaderText="案件日期" 
        Visible="True" ItemStyle-Width="10%" HeaderStyle-HorizontalAlign="left">
                    <ItemTemplate><%# showdate(DataSet1.FieldValue("case_start", Container)) %> </ItemTemplate>
                  </asp:TemplateColumn >
                  <asp:BoundColumn DataField="case_group" 
        HeaderText="類型" 
        ReadOnly="true" 
        Visible="True"/>                  
                  <asp:BoundColumn DataField="s01name" 
        HeaderText="承辦人" 
        ReadOnly="true" 
        Visible="True"/>                  
                  <asp:BoundColumn DataField="case_content" 
        HeaderText="案件內容" 
        ReadOnly="true" 
        Visible="True"/>                  
                  <asp:BoundColumn DataField="case_result" 
        HeaderText="結果" 
        ReadOnly="true" 
        Visible="True"/>                  
<asp:TemplateColumn HeaderText="結案日期" 
        Visible="True" ItemStyle-Width="10%" HeaderStyle-HorizontalAlign="left">
  <ItemTemplate><%# showdate(DataSet1.FieldValue("case_end", Container)) %> </ItemTemplate>
</asp:TemplateColumn>
<asp:TemplateColumn HeaderText="新增" 
        Visible="true">
  <ItemTemplate>
    <input name="Submit2" type="button" onclick="MM_goToURL('parent','s01case_add.aspx?case_id=<%# DataSet1.FieldValue("case_id", Container) %>&case_start=<%# showdate(case_start.text) %>');return document.MM_returnValue" value="新增" />
  </ItemTemplate>
</asp:TemplateColumn>

</Columns>
</asp:DataGrid>  
  <p>
  <HR>
</p>
</form>
</body>
<% 
if dataset1.recordcount = 0 then
dim url_add as string
url_add = "s01case_add.aspx?case_start=" & showdate(case_start.text) & "&case_content=" & trans(case_content.text)
response.redirect( url_add)
end if
%>
</html>
