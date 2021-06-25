<%@ Page Language="VB" ContentType="text/html" ResponseEncoding="utf-8" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<%@ Register TagPrefix="MM" Namespace="DreamweaverCtrls" Assembly="DreamweaverCtrls,version=1.0.0.0,publicKeyToken=836f606ede05d46a,culture=neutral" %>
<MM:DataSet 
id="DataSet1"
runat="Server"
IsStoredProcedure="false"
ConnectionString='<%# System.Configuration.ConfigurationSettings.AppSettings("MM_CONNECTION_STRING_s01") %>'
DatabaseType='<%# System.Configuration.ConfigurationSettings.AppSettings("MM_CONNECTION_DATABASETYPE_s01") %>'
CommandText='<%# "SELECT distinct case_id, case_start, case_end, case_group, s01name, case_end_yes, case_end_yes, case_yes, case_content, case_audit, admin_id, case_days, case_limit, case_alldays  FROM s01_case  WHERE case_start between ? and ? AND case_group like ? AND s01name like ? AND case_end_yes like ? AND case_yes like ? AND case_content like ? AND case_audit like ?  ORDER BY case_start DESC ,case_id DESC" %>'
PageSize="50"
Debug="true"
><Parameters>
  <Parameter  Name="@case_start"  Value='<%# IIf((Request.QueryString("case_start") <> Nothing), Request.QueryString("case_start"), "#1/1/2008#") %>'  Type="Date"   />  
  <Parameter  Name="@case_start1"  Value='<%# IIf((Request.QueryString("case_start1") <> Nothing), Request.QueryString("case_start1"), "#12/1/2099#") %>'  Type="Date"   />  
  <Parameter  Name="@case_group"  Value='<%# IIf((Request.QueryString("case_group") <> Nothing), stra1(Request.QueryString("case_group")), "%") %>'  Type="WChar"   />  
  <Parameter  Name="@s01name"  Value='<%# IIf((Request.QueryString("s01name") <> Nothing), stra1(Request.QueryString("s01name")), "%") %>'  Type="WChar"   />  
  <Parameter  Name="@case_end_yes"  Value='<%# IIf((Request.QueryString("case_end_yes") <> Nothing), stra1(Request.QueryString("case_end_yes")), "%") %>'  Type="WChar"   />  
  <Parameter  Name="@case_yes"  Value='<%# IIf((Request.QueryString("case_yes") <> Nothing), stra1(Request.QueryString("case_yes")), "%") %>'  Type="WChar"   />  
  <Parameter  Name="@case_content"  Value='<%# "%" + (IIf((Request.QueryString("case_content") <> Nothing), stra1(Request.QueryString("case_content")), "%")) + "%" %>'  Type="WChar"   />  
  <Parameter  Name="@case_audit"  Value='<%# IIf((Request.QueryString("case_audit") <> Nothing), stra1(Request.QueryString("case_audit")), "%") %>'  Type="WChar"   />
</Parameters></MM:DataSet>
<MM:DataSet 
id="DataSet2"
runat="Server"
IsStoredProcedure="false"
ConnectionString='<%# System.Configuration.ConfigurationSettings.AppSettings("MM_CONNECTION_STRING_s01") %>'
DatabaseType='<%# System.Configuration.ConfigurationSettings.AppSettings("MM_CONNECTION_DATABASETYPE_s01") %>'
CommandText='<%# "SELECT * FROM s01admin ORDER BY admin_username ASC" %>'
Debug="true"
></MM:DataSet>
<MM:PageBind runat="server" PostBackBind="false" />
<%@ Import Namespace="System.Data.OleDb" %>
<%@ Import Namespace="System.Globalization.Calendar" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Data.OleDb" %>
<head>
<meta http-equiv="Content-Type"; charset="UTF-8" />
<meta http-equiv="refresh" content="600"/>
<title>綜合行業科案件管理</title>

<script language="VB" runat="server">
        Sub Page_Load(sender As Object, e As EventArgs)
         If Session("MM_chief_num")="" then
			Response.Write("您還沒有登入呢！，請點擊")
			Response.Write("<a href=s01dellogin.aspx>登入</a>")
			Response.End()
		End If

		 dim add
		 add = request ("case_add")
		  select case add
		   case 1 
		   label2.text = "新增成功"
		   case 2 
		   label2.text = "修改成功"
		   case 3 
		   label2.text = "刪除成功"

		  end select 
     if not ispostback then
	 dim url_back
	 session("cancel_case_admin") = ""
	  url_back = request.url.tostring()
	  session("cancel_case_admin") = url_back
	  end if
	    End Sub

 sub s01name_di(sender As Object, e As EventArgs)
 if not ispostback then
s01name.DataTextField = dataset2.FieldValue("s01name", nothing)
s01name.DatavalueField = dataset2.FieldValue("admin_id", nothing)
end if
 end sub

    Sub Date_Selected(sender As Object, e As EventArgs)
	  dim url
	  url = "s01case_add0.aspx?case_start=" & Calendar1.SelectedDate.ToShortDateString
	  response.Redirect( url ) ' 使用Server.Transfer亦可
   End Sub
   
   Sub starsearch(sender As Object, e As EventArgs) 
	  dim url0
	  url0 = "s01case_admin.aspx?case_start=" & case_start.text & "&case_start1=" & case_start1.text & "&case_group=" & trans(case_group.selecteditem.text) & "&s01name=" & trans(s01name.selecteditem.text) & "&case_end_yes=" & trans(case_end_yes.selecteditem.text) & "&case_yes=" & trans(case_yes.selecteditem.text) & "&case_content=" & trans(case_content.text) & "&case_audit=" & trans(case_audit.selecteditem.text) 
	  response.Redirect( url0 ) ' 使用Server.Transfer亦可
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
 function showyes(y)
 if y = "True" then
 showyes = "是"
 else
 showyes = "否"
 end if
 end function
 function showmessage(vt2,vt3)
 if vt2 <> ""
 showmessage = vt2 & "<br/>" & vt3 & "小時"
 end if
 end function

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

<body bgcolor="#FFCCFF">

    <h3><strong><font color="#CC3300" face="Verdana, 新細明體">綜合行業科案件管制：</font><font color="#339900" face="Verdana, 新細明體">新增管制案件請點選日期</font></strong></h3>

    <form runat=server>
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td width="30%" valign="top"><p>
              <asp:Label Font-Bold="true" Font-Size="16" ForeColor="#990000" ID="Label2" runat="server" />              
            </p>
              <p>
                <asp:Calendar BorderColor="#CC9933"
            BorderWidth="1" DayHeaderStyle-BorderColor="#CC9933" DayHeaderStyle-Font-Size="8" DayNameFormat="SHORT" DayStyle-BorderColor="#CC9933"
            DayStyle-Height=
            DayStyle-VerticalAlign="Top"
            DayStyle-Width="12%"
            Font-Name="Verdana"
            Font-Size="9" ID=Calendar1
            NextMonthText = "下一月" NextPrevStyle-BorderColor="#990033" NextPrevStyle-Font-Underline="false" NextPrevStyle-ForeColor="#0000FF" NextPrevStyle-Wrap="false"
            PrevMonthText = "上一月" runat="server"
            SelectedDayStyle-BackColor="#FFCC66" SelectedDayStyle-BorderColor="#FF9933" SelectedDayStyle-ForeColor="#000000"
            ShowGridLines="true"
            TitleStyle-BackColor="Gainsboro" TitleStyle-BorderColor="#FF9966"
            TitleStyle-Font-Bold="true"
            TitleStyle-Font-Size="12px"
            TodayDayStyle-BackColor="#CCFF33" TodayDayStyle-BorderColor="#FF9966" TodayDayStyle-Font-Bold="false"
            TodayDayStyle-ForeColor="#993333" WeekendDayStyle-BackColor="#FF99CC" WeekendDayStyle-BorderColor="#FF9900"
            Width="250px"
            OnSelectionChanged="Date_Selected"
            />                
</p>
              <p>案件日期：
                <asp:TextBox ID="case_start" width="50" runat="server" />
                ~<asp:TextBox ID="case_start1" width="50" runat="server" />
                <br />
                案件類型：
<asp:DropDownList ID="case_group" runat="server">
	  <asp:ListItem></asp:ListItem>
      <asp:ListItem>申訴</asp:ListItem>
      <asp:ListItem>重大職災</asp:ListItem>
      <asp:ListItem>重大職災(其他)</asp:ListItem>
      <asp:ListItem>重大職災(勞安法以外)</asp:ListItem>
      <asp:ListItem>重大職災(認定中)</asp:ListItem>
      <asp:ListItem>非重大職災</asp:ListItem>
      <asp:ListItem>專案</asp:ListItem>
      <asp:ListItem>其他</asp:ListItem>
</asp:DropDownList>
              <br />
 承辦人：
 <asp:DropDownList DataSource="<%# dataset2.defaultview %>" DataTextField="s01name" DataValueField="admin_id" ID="s01name" runat="server" ></asp:DropDownList>
              <br />
已結案：<asp:DropDownList ID="case_end_yes" runat="server">
	  <asp:ListItem></asp:ListItem>
	  <asp:ListItem>是</asp:ListItem>
	  <asp:ListItem>否</asp:ListItem>
</asp:DropDownList>
              <br />
已登錄：<asp:DropDownList ID="case_yes" runat="server">
	  <asp:ListItem></asp:ListItem>
	  <asp:ListItem>是</asp:ListItem>
	  <asp:ListItem>否</asp:ListItem>
	  <asp:ListItem>不需</asp:ListItem>
</asp:DropDownList>
<br />
內容：
<asp:TextBox ID="case_content" runat="server" />              
<br />
              逾期：
              <asp:DropDownList ID="case_audit" runat="server">
                <asp:ListItem></asp:ListItem>
                <asp:ListItem>逾期</asp:ListItem>
              </asp:DropDownList>
              </p>
              <p>
            <asp:Button ID="case_search" runat="server" Text="查詢" OnClick="starsearch" />                              </p></td>
            <td width="70%" colspan="" valign="top">
            
              <p>
                </br>
                共<strong><font color="#CC3300"><%= DataSet1.RecordCount %></font></strong>筆
                <asp:DataGrid 
  AllowCustomPaging="true" 
  AllowPaging="true" 
  AllowSorting="False" 
  AutoGenerateColumns="false" 
  CellPadding="3" 
  CellSpacing="0" 
  DataSource="<%# DataSet1.DefaultView %>" id="DataGrid1" 
  PagerStyle-Mode="NumericPages" 
  PageSize="<%# DataSet1.PageSize %>" 
  runat="server" 
  ShowFooter="false" 
  ShowHeader="true" Width="100%" 
  OnPageIndexChanged="DataSet1.OnDataGridPageIndexChanged" 
  VirtualItemCount="<%# DataSet1.RecordCount %>" 
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
                  <asp:BoundColumn DataField="case_group" 
        HeaderText="類型" 
        ReadOnly="true" 
        Visible="True" ItemStyle-Width="10%"/>                  
<asp:TemplateColumn HeaderText="案件日期" 
        Visible="True" ItemStyle-Width="10%">
              <ItemTemplate><%# showdate(DataSet1.FieldValue("case_start", Container)) %> </ItemTemplate>
                  </asp:TemplateColumn>
                  <asp:BoundColumn DataField="s01name" 
        HeaderText="承辦人" 
        ReadOnly="true" 
        Visible="True" ItemStyle-Width="8%"/>                  
                  <asp:BoundColumn DataField="case_content" 
        HeaderText="案件內容" 
        ReadOnly="true" 
        Visible="True" ItemStyle-Width="18%"/>                  
                  <asp:BoundColumn DataField="case_days" 
        HeaderText="已辦天數" 
        ReadOnly="true" 
        Visible="True" ItemStyle-Width="6%"/>                  
                  <asp:BoundColumn DataField="case_limit" 
        HeaderText="限辦天數" 
        ReadOnly="true" 
        Visible="True" ItemStyle-Width="6%"/>                  
                  <asp:BoundColumn DataField="case_audit" 
        HeaderText="逾辦" 
        ReadOnly="true" 
        Visible="True" ItemStyle-ForeColor="#CC3300" ItemStyle-Width="8%"/>                  
                  <asp:BoundColumn DataField="case_end_yes" 
        HeaderText="已結案" 
        ReadOnly="true" 
        Visible="True" ItemStyle-Width="8%"/>                  
                  <asp:TemplateColumn HeaderText="結案日期" 
        Visible="True" ItemStyle-Width="10%">
                    <ItemTemplate><%# showdate(DataSet1.FieldValue("case_end", Container)) %></ItemTemplate>
                  </asp:TemplateColumn>
                  <asp:BoundColumn DataField="case_alldays" 
        HeaderText="完成天數" 
        ReadOnly="true" 
        Visible="false"/>                  
<asp:TemplateColumn HeaderText="修改" 
        Visible="True" ItemStyle-Width="16%">
  <ItemTemplate>
    <input name="Submit2" type="button" onclick="MM_goToURL('parent','s01case_update.aspx?case_id=<%# DataSet1.FieldValue("case_id", Container) %>');return document.MM_returnValue" value="修改" />
    <input name="Submit3" type="button" onClick="DelTitle('<%# DataSet1.FieldValue("case_id", Container) %>')" value="刪除" />
  </ItemTemplate>
</asp:TemplateColumn>
                  </Columns>
                </asp:DataGrid>
              </p>
      </table>
        <p>
        <asp:Label id=Label1 runat="server" />
        
</form>
</body>
</html>