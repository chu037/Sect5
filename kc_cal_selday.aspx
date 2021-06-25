<%@ Page Language="VB" ContentType="text/html" ResponseEncoding="utf-8" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<%@ Register TagPrefix="MM" Namespace="DreamweaverCtrls" Assembly="DreamweaverCtrls,version=1.0.0.0,publicKeyToken=836f606ede05d46a,culture=neutral" %>
<MM:DataSet 
id="DataSet1"
runat="Server"
IsStoredProcedure="false"
ConnectionString='<%# System.Configuration.ConfigurationSettings.AppSettings("MM_CONNECTION_STRING_s01") %>'
DatabaseType='<%# System.Configuration.ConfigurationSettings.AppSettings("MM_CONNECTION_DATABASETYPE_s01") %>'
CommandText='<%# "SELECT m_date, m_hours, m_maker, m_messeng, m_note, m_num, m_time, m_date_en, m_time_en  FROM s01_kc_calen  WHERE m_date = ?  ORDER BY m_time ASC" %>'
Debug="true"
>
  <Parameters>
<Parameter  Name="@m_date"  Value='<%# IIf((Request.QueryString("m_date") <> Nothing), Request.QueryString("m_date"), "") %>'    /></Parameters></MM:DataSet>
<MM:PageBind runat="server" PostBackBind="false" />

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>當日行程</title>
<script language="VB" runat="server">
    Sub page_load(sender As Object, e As EventArgs)
        If Request("m_date") = Nothing Then
            Response.Redirect("kc_cal.aspx")
        End If
        Dim t1 As DateTime = Request("m_date")
        If t1 = "" Then
            Response.Redirect("kc_cal.aspx")
        End If
        Dim add_yes = Request("m_add")
        Select Case add_yes
            Case 1
                msg1.Text = "新增成功"
            Case 2
                msg1.Text = "修改成功"
        End Select
        Dim weekstr() = {"日", "一", "二", "三", "四", "五", "六"}
        Dim m_week As String = DatePart(DateInterval.Weekday, t1) - 1
        Dim week_match = weekstr(m_week)
        Dim selected_day = FormatDateTime(t1, DateFormat.ShortDate)
        Dim show_day = selected_day & "，星期" & week_match
        msg02.Text = show_day
        m_date.Text = Request("m_date")
        If Not IsPostBack Then
            Calendar1.SelectedDate = showdate(Request("m_date"))
            Call Populate_YearList()
            Call Populate_MonthList()
            Dim d1 = DatePart("m", CDate(drpCalMonth.SelectedItem.Value & " ,1, " & drpCalYear.SelectedItem.Value))
            Dim d0 = DatePart("m", Now().ToString())
            If d1 <> d0 Then
                Calendar1.TodaysDate = CDate(drpCalMonth.SelectedItem.Value & " ,1, " & drpCalYear.SelectedItem.Value)
                'Calendar1.TodayDayStyle.BackColor="FFFF99"
            Else
                Calendar1.TodaysDate = CDate(Now().ToString())
                'Calendar1.TodaysDate = showdate(now)
            End If
        End If
    End Sub
Sub Set_Calendar(Sender As Object, E As EventArgs)
    
        'Whenever month or year selection changes display the calendar for that month/year        
    dim d1 = datepart("m",CDate(drpCalMonth.SelectedItem.Value & " ,1, " & drpCalYear.SelectedItem.Value))
	dim d0 = datepart("m",now().tostring())
	if d1 <> d0 then
    Calendar1.TodaysDate = CDate(drpCalMonth.SelectedItem.Value & " ,1, " & drpCalYear.SelectedItem.Value)
    'Calendar1.TodayDayStyle.BackColor="FFFF99"
	else
	Calendar1.TodaysDate = cdate(now().tostring())
	'Calendar1.TodaysDate = showdate(now)
    end if
	End Sub

Sub Populate_MonthList()
dim sel_month 
sel_month = showmonth(request("m_date"))

    	 drpCalMonth.Items.Add("一月")   
         drpCalMonth.Items.Add("二月")   
         drpCalMonth.Items.Add("三月")   
         drpCalMonth.Items.Add("四月")   
         drpCalMonth.Items.Add("五月")   
         drpCalMonth.Items.Add("六月")   
         drpCalMonth.Items.Add("七月")   
         drpCalMonth.Items.Add("八月")   
         drpCalMonth.Items.Add("九月")   
         drpCalMonth.Items.Add("十月")   
         drpCalMonth.Items.Add("十一月")   
         drpCalMonth.Items.Add("十二月") 
    
        '把這行註解起來好像就不會出錯了..(因為我們的月份是國字的..這個範例是外國的時間格式...)
		drpCalMonth.Items.FindByValue(sel_month).selected = true
        
        '看這行就知道為什麼了.
        'Response.Write(MonthName(DateTime.Now.Month))
    
    End Sub   
    Sub Populate_YearList()
dim sel_year as object
sel_year = showyear(request("m_date"))
        'Year list can be extended
        Dim intYear As Integer
    
        For intYear = sel_year - 2 to sel_year + 2
    
             drpCalYear.Items.Add(intYear)
        Next
    
        drpCalYear.Items.FindByValue(sel_year).selected = true
    
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
  function showmonth(t)
 if t <> ""
 showmonth = monthname(month(t))
 end if
 end function
  function showyear(t)
 if t <> ""
 showyear = year(t)
 end if
 end function
 function showmessage(vt2,vt3)
 if vt2 <> ""
 showmessage = vt2 & "<br/>" & vt3 & "小時"
 end if
 end function
 Sub change_page(sender As Object, e As EventArgs)
	  dim url
	  url = "kc_cal_add.aspx?m_date=" & m_date.text
	  response.Redirect( url ) ' 使用Server.Transfer亦可
 End Sub
    Sub Date_Selected(sender As Object, e As EventArgs)
	  dim url
	  url = "kc_cal_selday.aspx?m_date=" & Calendar1.SelectedDate.ToShortDateString
	  response.Redirect( url ) ' 使用Server.Transfer亦可
   End Sub

    </script>
<script language = "JavaScript">
<!--

function w_back()
{
location.href = "kc_cal.aspx"
}
function MM_goToURL() { //v3.0
  var i, args=MM_goToURL.arguments; document.MM_returnValue = false;
  for (i=0; i<(args.length-1); i+=2) eval(args[i]+".location='"+args[i+1]+"'");
}

//-->
</Script>

<style type="text/css">
<!--
a:link {
	color: #0000FF;
	text-decoration: none;
}
a:visited {
	color: #0000CC;
	text-decoration: none;
}
a:hover {
	color: #00CC00;
	text-decoration: none;
}
a:active {
	color: #33FFFF;
	text-decoration: none;
}
.style1 {
	color: #CC3300;
	font-weight: bold;
}
-->
</style>
</head>
<body>

<form action="" runat="server">
<table width="100%">
 <tr width="100%">
  <td width="30%" align="center" valign="bottom">
西元
  <asp:DropDownList id="drpCalYear" Runat="Server" AutoPostBack="True" OnSelectedIndexChanged="Set_Calendar" cssClass="calTitle"></asp:DropDownList> 年<asp:DropDownList id="drpCalMonth" Runat="Server" AutoPostBack="True" OnSelectedIndexChanged="Set_Calendar" cssClass="calTitle"></asp:DropDownList>
  </td>
  <td width="70%">共<span class="style1"><%= DataSet1.RecordCount %></span>筆行程
  <asp:Label Font-Bold="true" ForeColor="#CC0033" ID="msg1" runat="server" />  
  </td></tr>
<tr>
 <td valign="top">
<asp:Calendar BackColor="#FFFF99"
            DayHeaderStyle-BackColor="#ffcccc"
            DayNameFormat="Shortest" EnableViewState=""
            Font-Name="Arial" Font-Size="11px" 
            ID=Calendar1  OtherMonthDayStyle-BackColor="#FFFFFF" OtherMonthDayStyle-BorderColor="#FFFFFF" OtherMonthDayStyle-Font-Size="0" OtherMonthDayStyle-Height="0" runat="server"  
            SelectedDayStyle-BackColor="Navy"
            SelectedDayStyle-Font-Bold="True" SelectionMode="Day"
            SelectorStyle-BackColor="gainsboro" ShowDayHeader="true" ShowGridLines="true" ShowNextPrevMonth="false" ShowTitle="false" TitleFormat="MonthYear"
            TitleStyle-BackColor="#cccc66"
            TitleStyle-Font-Bold="True"
            TitleStyle-Font-Size="12px" 
            TodayDayStyle-BackColor="#FFFF00" Width="100%"
			OnSelectionChanged="Date_Selected"
            /> 
 <br />
 <asp:Label ID="msg02" runat="server" />
 <br />
<asp:Button ID="Button3" runat="server" Text="新增行程" OnClick="change_page" />

<input type="button" name="Submit3" value="回行事曆" onclick="return w_back()" /></td> 
 <td valign="top">
<asp:DataGrid AllowPaging="false" 
  AllowSorting="False" 
  AutoGenerateColumns="false" 
  CellPadding="3" 
  CellSpacing="0" 
  DataSource="<%# DataSet1.DefaultView %>" ID="DataGrid1" 
  runat="server" 
  ShowFooter="false" 
  ShowHeader="true" Width="100%" 
>
  <headerstyle HorizontalAlign="left" BackColor="#E8EBFD" ForeColor="#3D3DB6" Font-Name="Verdana, Arial, Helvetica, sans-serif" Font-Bold="true" Font-Size="smaller" />
  <itemstyle BackColor="#F2F2F2" Font-Name="Verdana, Arial, Helvetica, sans-serif" Font-Size="smaller" />
  <alternatingitemstyle BackColor="#E5E5E5" Font-Name="Verdana, Arial, Helvetica, sans-serif" Font-Size="smaller" />
  <footerstyle HorizontalAlign="left" BackColor="#E8EBFD" ForeColor="#3D3DB6" Font-Name="Verdana, Arial, Helvetica, sans-serif" Font-Bold="true" Font-Size="smaller" />
  <pagerstyle BackColor="white" Font-Name="Verdana, Arial, Helvetica, sans-serif" Font-Size="smaller" />
  <columns>
  <asp:TemplateColumn
	    HeaderText="啟始" ItemStyle-Width="20%" 
        Visible="True">
    <itemtemplate><%# showdate(DataSet1.FieldValue("m_date", Container)) %> <br />
        <%# showtime(DataSet1.FieldValue("m_time", Container)) %> </itemtemplate>
  </asp:TemplateColumn>
  <asp:TemplateColumn
	    HeaderText="結束" ItemStyle-Width="20%" 
        Visible="True">
    <itemtemplate><%# showdate(DataSet1.FieldValue("m_date_en", Container)) %> <br />
        <%# showtime(DataSet1.FieldValue("m_time_en", Container)) %> </itemtemplate>
  </asp:TemplateColumn>

  <asp:TemplateColumn
	    HeaderText="行程" ItemStyle-Width="10%"
        Visible="True">
    <itemtemplate><%# showmessage(DataSet1.FieldValue("m_messeng", Container),DataSet1.FieldValue("m_hours", Container)) %> </itemtemplate>
  </asp:TemplateColumn>
                <asp:HyperLinkColumn
        DataNavigateUrlField="m_num"
        DataNavigateUrlFormatString="kc_cal_detail.aspx?m_num={0}"
        DataTextField="m_note" 
        Visible="True" target="cal_detail" 
        ItemStyle-Width="50%"
		HeaderText="主題"/> 
<asp:TemplateColumn HeaderText="修改" 
        Visible="True">
  <ItemTemplate>
    <input type="button" name="Submit3" value="修改" onclick= "MM_goToURL('self','kc_cal_update.aspx?m_num=<%# DataSet1.FieldValue("m_num", Container) %>');return document.MM_returnValue" />
  </ItemTemplate>
</asp:TemplateColumn>
  </columns>
</asp:DataGrid>
<p>
  <asp:TextBox ID="m_date" MaxLength="20" ReadOnly="true" Rows="10" runat="server" />
  <asp:Button ID="Button1" runat="server" Text="新增行程" OnClick="change_page" />
  
  <input type="button" name="Submit2" value="回行事曆" onclick="return w_back()" />
</td>
</tr>
</table>

</form>
<%
  session("cancel") = ""
  session("cancel") = request.url.tostring()
%>

</body>
</html>
