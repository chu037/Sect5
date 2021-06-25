<%@ Page Language="VB" ContentType="text/html"%>
<%@ Register TagPrefix="MM" Namespace="DreamweaverCtrls" Assembly="DreamweaverCtrls,version=1.0.0.0,publicKeyToken=836f606ede05d46a,culture=neutral" %>
<MM:DataSet 
id="DataSet1"
runat="Server"
IsStoredProcedure="false"
ConnectionString='<%# System.Configuration.ConfigurationSettings.AppSettings("MM_CONNECTION_STRING_s01") %>'
DatabaseType='<%# System.Configuration.ConfigurationSettings.AppSettings("MM_CONNECTION_DATABASETYPE_s01") %>'
CommandText='<%# "SELECT diff, m_date, m_hours, m_messeng, m_note, m_num, m_time, diffen, m_date_en  FROM s01_calen  WHERE diff = 0 or diff < 0 and diffen > -1  ORDER BY m_date asc, m_time ASC" %>'
Debug="true"
></MM:DataSet>
<MM:DataSet 
id="DataSet2"
runat="Server"
IsStoredProcedure="false"
ConnectionString='<%# System.Configuration.ConfigurationSettings.AppSettings("MM_CONNECTION_STRING_s01") %>'
DatabaseType='<%# System.Configuration.ConfigurationSettings.AppSettings("MM_CONNECTION_DATABASETYPE_s01") %>'
CommandText='<%# "SELECT diff, m_date, m_hours, m_messeng, m_note, m_time, diffen, m_date_en, m_num  FROM s01_calen  WHERE diff = 1   ORDER BY m_date asc, m_time ASC" %>'
Debug="true"
></MM:DataSet>
<MM:DataSet 
id="DataSet3"
runat="Server"
IsStoredProcedure="false"
ConnectionString='<%# System.Configuration.ConfigurationSettings.AppSettings("MM_CONNECTION_STRING_s01") %>'
DatabaseType='<%# System.Configuration.ConfigurationSettings.AppSettings("MM_CONNECTION_DATABASETYPE_s01") %>'
CommandText='<%# "SELECT diff, m_date, m_hours, m_messeng, m_note, m_time, m_num FROM s01_calen  WHERE diff < ? and diff > 1  ORDER BY m_date asc, m_time ASC" %>'
PageSize="30"
Debug="true"
>
<parameters>
<Parameter  Name="@diff"  Value='<%# diff_3.selecteditem.value %>'  Type="Integer"   /></Parameters></MM:DataSet>
</MM:DataSet>
<MM:PageBind runat="server" PostBackBind="true" />
<%@ Import Namespace="System.Data.OleDb" %>
<%@ Import Namespace="System.Globalization.Calendar" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<meta http-equiv="refresh" content="600"/>
<title>綜合行業科行事曆</title>

<script language="VB" runat="server">
        Sub Page_Load(sender As Object, e As EventArgs)
        if not ispostback then
        session("cancel") = ""
		end if 
		End Sub
    Function duty_d3(st_d As Integer, v_d As Integer) As Integer '不論人數是否為3的倍數皆通用
        Dim diff_w As Integer = DatePart("w", v_d) 'v1是一週內的第幾天
        Dim w_n As Integer = Int(DateDiff("d", st_d, v_d) / 7) '第幾週.0~
        Select Case diff_w
            Case 1, 6, 7
                duty_d3 = 3 + w_n
            Case 2, 3
                duty_d3 = 3 + w_n + 1
            Case 4, 5
                duty_d3 = 3 + w_n + 2
        End Select
        Return duty_d3
    End Function
    Function duty_d(st_d As Integer, v_d As Integer) As Integer
        Dim diff_w As Integer = DatePart("w", v_d) 'v1是一週內的第幾天
        Dim w_n As Integer = Int(DateDiff("d", st_d, v_d) / 7) '第幾週.0~
        Select Case diff_w
            Case 1, 6, 7
                duty_d = 3 * w_n
            Case 2, 3
                duty_d = 3 * w_n + 1
            Case 4, 5
                duty_d = 3 * w_n + 2
        End Select
        Return duty_d
    End Function
    Function duty_d2(st_d As Integer, v_d As Integer) As Integer
        Dim diff_w As Integer = DatePart("w", v_d) 'v1是一週內的第幾天
        Dim w_n As Integer = Int(DateDiff("d", st_d, v_d) / 7) '第幾週.0~
        Select Case diff_w
            Case 2, 3, 4
                duty_d2 = 2 * w_n
            Case 5, 6, 7, 1
                duty_d2 = 2 * w_n + 1
        End Select
        Return duty_d2
    End Function

    Sub Calendar1_DayRender(sender As Object, e As DayRenderEventArgs)
        Dim Conn As OleDbConnection
        Dim Cmd As OleDbCommand
        Dim Rd As OleDbDataReader
        Dim I As Integer
        Dim j As Integer
        Dim I1 As Integer = 0

        Dim Provider = "Provider=Microsoft.Jet.OLEDB.4.0"
        Dim Database = "Data Source=" & Server.MapPath("result/result.mdb")
        Dim v As CalendarDay
        Dim c As TableCell
        Conn = New OleDbConnection(Provider & ";" & Database)
        Conn.Open()
        Dim SQL = "Select * From s01_calen where m_date = @m_date"
        Cmd = New OleDbCommand(SQL, Conn)
        Cmd.Parameters.Clear()
        Cmd.Parameters.AddWithValue("m_date", e.Day.Date)
        Rd = Cmd.ExecuteReader()
        v = e.Day
        c = e.Cell
        Dim v1 As Date
        v1 = v.Date
        Dim start_d() As Date = {#3/31/2014#} '要從星期一開始
        Dim dx As Integer
        For dx = 0 To start_d.Length - 1
            If dx < start_d.Length - 1 Then
                Dim diff_d0 = DateDiff("d", start_d(dx), v1)
                Dim diff_d = DateDiff("d", start_d(dx + 1), v1)
                If diff_d0 >= 0 And diff_d < 0 Then Exit For
            Else
                Dim diff_d0 = DateDiff("d", start_d(dx), v1)
                If diff_d0 >= 0 Then Exit For
            End If
        Next

        Select Case dx
            Case 0
                Dim ad_d() = {"1", "2", "3", "4", "5", "6", "7", "8"}
                Dim author_w() = {"顏廷諭", "姜智敏", "楊尚淳", "楊勝安", "蘇銘源", "吳俐節", "朱志杰"}
                'dim i01 as integer
                'if author_w.length mod 2 = 0 then '2的倍數時要+第幾輪參數
                Dim w_n As Integer = Int(DateDiff("d", start_d(dx), v1) / 7) Mod (author_w.Length) '第幾輪
                'i01 = (duty_d2(start_d(dx),v1)+w_n) mod author_w.length 
                'else
                'i = (duty_d2(start_d(dx),v1)) mod author_w.length
                'end if 
                If author_w.Length < ad_d.Length Then '輪的人小於8人
                    'dim k as integer = ad_d.length -1
                    j = 0
                    I = w_n
                    While j < ad_d.Length
                        'for i = 0 to author_w.length-1-(ad_d.length-author_w.length)
                        Dim author As String = author_w(I)
                        Dim ad As String = ad_d(j)
                        c.Controls.Add(New LiteralControl("<br>" + ad + author))
                        j = j + 1
                        I = I + 1
                        If I > author_w.Length - 1 Then
                            I = 0
                        End If

                    End While
                End If

        End Select

 

        'dim start_d3 = #8/16/2010# '輪值組程式
        'dim diff_star_w3 = int((datediff("d",start_d3,v1)/7)) '第幾星期 
        'dim diff_star_s3 = (int((datediff("d",start_d,v1)/7))) mod 7 '假日輪勤第幾個星期 
        'dim author_w3() = {"一組","二組","三組","四組"}
        'if datediff("d",start_d3,v1) >= 0 then
        'i = (datediff("d",start_d0,v1) + diff_star_w0) mod 7
        'i = diff_star_w3 mod 4
        'dim author3 as string = author_w3(i)
        'c.Controls.Add(new LiteralControl("<br>" + author3))
        'end if

        While Rd.Read()
            Dim ltrCr As New LiteralControl("<br>")
            Dim link As New HyperLink()
            link.NavigateUrl = "s01cal_detail.aspx?m_num=" & Rd.Item(3)
            link.Text = Rd.GetString(0)
            link.Target = "cal_detail"
            c.Controls.Add(ltrCr)
            c.Controls.Add(link)
        End While
        Conn.Close()
        If v.IsOtherMonth Then
            c.Controls.Clear()
        End If

    End Sub

    Sub Date_Selected(sender As Object, e As EventArgs)
	  dim url
	  url = "s01cal_selday.aspx?m_date=" & Calendar1.SelectedDate.ToShortDateString
	  response.Redirect( url ) ' 使用Server.Transfer亦可
   End Sub

    Function showdate(s01time As String) As String
        showdate = ""
        If s01time <> "" Then
            showdate = FormatDateTime(s01time, DateFormat.ShortDate)
        End If
        Return showdate
    End Function

    Function showtime(t As String) As String
        showtime = ""
        If t <> "" Then
            showtime = FormatDateTime(t, DateFormat.ShortTime)
        End If
        Return showtime
    End Function

    Function showweek(t As Date) As String
        showweek = ""
        If t <> "" Then
            showweek = WeekdayName(DatePart("w", t))
        End If
        Return showweek
    End Function

    Function showmessage(vt2 As String, vt3 As String) As String
        showmessage = ""
        If vt2 <> "" Then
            showmessage = vt2 & "<br/>" & vt3 & "小時"
        End If
        Return showmessage
    End Function

    Function showdetail(vm1 As String) As String
        showdetail = ""
        If vm1 <> "" Then
            If Len(vm1) > 30 Then
                showdetail = Left(vm1, 30) & "..."
            Else
                showdetail = vm1
            End If
        End If
        Return showdetail
    End Function

 
   Sub starsearch(sender As Object, e As EventArgs) 
	  dim url0
	  dim asc_m, asc_n as string
	  if m_messeng.selecteditem.text <> "" then
	   asc_m = m_messeng.selecteditem.text
	  else
	   asc_m = "" 
	  end if
	  if m_note.text <> "" then
	   asc_n = m_note.text
	  else
	   asc_n = ""
	  end if   
	  url0 = "s01cal_search.aspx?m_date=" & m_date.text & "&m_date1=" & m_date1.text & "&m_messeng=" & trans(m_messeng.selecteditem.text) & "&m_note=" & trans(m_note.text)
	  response.Redirect( url0 ) ' 使用Server.Transfer亦可
   End Sub

    Function trans(str01 As String) As String
        trans = ""
        If str01 <> "" Then
            Dim i As Integer
            Dim stra As String
            stra = ""
            For i = 1 To Len(str01) - 1
                stra &= Asc(Mid(str01, i, 1)) & ","
            Next
            If i = Len(str01) Then
                stra = stra & Asc(Mid(str01, i, 1))
            End If
            trans = stra
        End If
        Return trans
    End Function
    
    </script>

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
.calWeekendDay
{
      background-color: #FF6600 !important;
}
body {
	background-color: #FFCCCC;
}


-->
</style>
</head>

<body>

    <h3 align="center"><strong><font color="#CC3300" size="+2" face="Verdana, 新細明體">綜合行業科行事曆</font><font color="#339900" face="Verdana, 新細明體">新增或更改行程請點選日期</font></strong></h3>

    <form runat=server>
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td width="39%" valign="top"><asp:Calendar BorderColor="#CC9933"
            BorderWidth="1" DayHeaderStyle-BorderColor="#CC9933" DayHeaderStyle-Font-Size="8" DayNameFormat="Shortest" DayStyle-BorderColor="#CC9933"
            DayStyle-Height=
            DayStyle-VerticalAlign="Top"
            DayStyle-Width="14%"
            Font-Name="Verdana"
            Font-Size="9" ID=Calendar1
            NextMonthText = "下一月" NextPrevStyle-BorderColor="#990033" NextPrevStyle-Font-Underline="false" NextPrevStyle-ForeColor="#0000FF" NextPrevStyle-Wrap="false"
            PrevMonthText = "上一月" runat="server"
            SelectedDayStyle-BackColor="#FFCC66" SelectedDayStyle-BorderColor="#FF9933" SelectedDayStyle-ForeColor="#000000" ShowDayHeader="true"
            ShowGridLines="true"
            TitleStyle-BackColor="Gainsboro" TitleStyle-BorderColor="#FF9966"
            TitleStyle-Font-Bold="true"
            TitleStyle-Font-Size="12px"
            TodayDayStyle-BackColor="#CCFF33" TodayDayStyle-BorderColor="#FF9966" TodayDayStyle-Font-Bold="false"
            TodayDayStyle-ForeColor="#993333" WeekendDayStyle-CssClass="calWeekendDay"

            Width="300px"
            OnDayRender="Calendar1_DayRender"
            OnSelectionChanged="Date_Selected"
            />
<br>
<table width="81%" valign="top" border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td><strong><font color="#CC3300" size="+1">行程查詢：</font></strong></td>
			  <td>日期格式:<br/>YYYY/M/D              </tr>
              <tr>
                <td width="50%">行程日期：</td>
                <td width="50%"><asp:TextBox ID="m_date" runat="server" Width="60" />~<asp:TextBox ID="m_date1" runat="server" Width="60" /></td>
              </tr>
              <tr>
                <td>行程類型：</td>
                <td><asp:DropDownList ID="m_messeng" runat="server">
                  <asp:ListItem></asp:ListItem>
                  <asp:ListItem>開會</asp:ListItem>
                  <asp:ListItem>宣導</asp:ListItem>
                  <asp:ListItem>上課</asp:ListItem>
                  <asp:ListItem>請假</asp:ListItem>
                  <asp:ListItem>換班</asp:ListItem>
                  <asp:ListItem>0239</asp:ListItem>
                  <asp:ListItem>0257</asp:ListItem>
                  <asp:ListItem>其他</asp:ListItem>

 
                </asp:DropDownList></td>
              </tr>
              <tr>
                <td>內容：</td>
                <td><asp:TextBox ID="m_note" runat="server" Width="100" /></td>
              </tr>
        </table>
              <p>
              <asp:Button ID="case_search" runat="server" Text="查詢" OnClick="starsearch" />                              </p>

            </td>
            <td width="61%" rowspan="3" valign="top"><strong><font color="#FF0000">今日<font color="#0000CC"><%= FormatDateTime(now(), DateFormat.ShortDate) %>(<%= weekdayname(datepart("w",now())) %>)</font>行程:共</font></strong>              <font color="#000000"><strong><%= DataSet1.RecordCount %></strong></font><strong><font color="#FF0000">筆</font></strong>
              <asp:DataGrid AllowPaging="false" 
  AllowSorting="False" AlternatingItemStyle-Wrap="false" 
  AutoGenerateColumns="false" 
  CellPadding="3" 
  CellSpacing="0" 
  DataSource="<%# DataSet1.DefaultView %>" EditItemStyle-Wrap="false" FooterStyle-Wrap="false" HeaderStyle-Wrap="false" ID="DataGrid1" ItemStyle-Wrap="true" PagerStyle-Wrap="false" 
  runat="server" SelectedItemStyle-Wrap="false" 
  ShowFooter="false" 
  ShowHeader="true" Width="100%" 
>
                <headerstyle HorizontalAlign="left" BackColor="#E8EBFD" ForeColor="#3D3DB6" Font-Name="Verdana, Arial, Helvetica, sans-serif" Font-Bold="true" Font-Size="smaller" />              
                <itemstyle BackColor="#F2F2F2" Font-Name="Verdana, Arial, Helvetica, sans-serif" Font-Size="smaller" />              
                <alternatingitemstyle BackColor="#E5E5E5" Font-Name="Verdana, Arial, Helvetica, sans-serif" Font-Size="smaller" Wrap="true" />              
                <footerstyle HorizontalAlign="left" BackColor="#E8EBFD" ForeColor="#3D3DB6" Font-Name="Verdana, Arial, Helvetica, sans-serif" Font-Bold="true" Font-Size="smaller" />              
                <pagerstyle BackColor="white" Font-Name="Verdana, Arial, Helvetica, sans-serif" Font-Size="smaller" />              
                <columns>
                <asp:TemplateColumn
	    HeaderText="日期" ItemStyle-Width="25%" HeaderStyle-Width="25%" 
        Visible="True">
                  <itemtemplate><%# showdate(DataSet1.FieldValue("m_date", Container)) %> <br />
                    <%# showtime(DataSet1.FieldValue("m_time", Container)) %><font color="#3333CC"><%# showweek(DataSet1.FieldValue("m_date", Container)) %></font> </itemtemplate>
                </asp:TemplateColumn>
                <asp:TemplateColumn
	    HeaderText="行程" ItemStyle-Width="15%"
        Visible="True">
                  <itemtemplate><%# showmessage(DataSet1.FieldValue("m_messeng", Container),DataSet1.FieldValue("m_hours", Container)) %> </itemtemplate>
                </asp:TemplateColumn>
                <asp:HyperLinkColumn
        DataNavigateUrlField="m_num"
        DataNavigateUrlFormatString="s01cal_detail.aspx?m_num={0}"
        DataTextField="m_note" 
        Visible="True" target="cal_detail"
		HeaderText="主題" 
        ItemStyle-Width="60%"/> 
		             
                <asp:BoundColumn DataField="m_time" 
        HeaderText="時間" 
        ReadOnly="true" 
        Visible="false"/>                
</columns>
              </asp:DataGrid>
              <strong><font color="#CC6600">明日行程:共<font color="#000000"><%= DataSet2.RecordCount %></font>筆              </font></strong>
              <asp:DataGrid 
  AllowPaging="false" 
  AllowSorting="False" AlternatingItemStyle-Wrap="false" 
  AutoGenerateColumns="false" 
  CellPadding="3" 
  CellSpacing="0" 
  DataSource="<%# DataSet2.DefaultView %>" EditItemStyle-Wrap="false" FooterStyle-Wrap="false" HeaderStyle-Wrap="false" id="DataGrid2" ItemStyle-Wrap="true" PagerStyle-Wrap="false" 
  runat="server" SelectedItemStyle-Wrap="false" 
  ShowFooter="false" 
  ShowHeader="true" Width="100%" 
>
        <HeaderStyle HorizontalAlign="left" BackColor="#FFFFCC" ForeColor="#3D3DB6" Font-Name="Verdana, Arial, Helvetica, sans-serif" Font-Bold="true" Font-Size="smaller" />        
        <ItemStyle BackColor="#F2F2F2" Font-Name="Verdana, Arial, Helvetica, sans-serif" Font-Size="smaller" Wrap="true" />        
        <AlternatingItemStyle BackColor="#FFFFCC" Font-Name="Verdana, Arial, Helvetica, sans-serif" Font-Size="smaller" />        
        <FooterStyle HorizontalAlign="left" BackColor="#E8EBFD" ForeColor="#3D3DB6" Font-Name="Verdana, Arial, Helvetica, sans-serif" Font-Bold="true" Font-Size="smaller" />        
        <PagerStyle BackColor="white" Font-Name="Verdana, Arial, Helvetica, sans-serif" Font-Size="smaller" />        
      <Columns>
      <asp:TemplateColumn
	    HeaderText="日期" ItemStyle-Width="25%" 
        Visible="True">
        <ItemTemplate><%# showdate(DataSet2.FieldValue("m_date", Container)) %> <br />
            <%# showtime(DataSet2.FieldValue("m_time", Container)) %><font color="#3333CC"><%# showweek(DataSet2.FieldValue("m_date", Container)) %></font> </ItemTemplate>
      </asp:TemplateColumn>
      <asp:TemplateColumn
	    HeaderText="行程" ItemStyle-Width="15%"
        Visible="True">
        <ItemTemplate><%# showmessage(DataSet2.FieldValue("m_messeng", Container),DataSet2.FieldValue("m_hours", Container)) %> </ItemTemplate>
      </asp:TemplateColumn>
                <asp:HyperLinkColumn
        DataNavigateUrlField="m_num"
        DataNavigateUrlFormatString="s01cal_detail.aspx?m_num={0}"
        DataTextField="m_note" 
        Visible="True" target="cal_detail" 
        ItemStyle-Width="50%"
		HeaderText="主題"/>       <asp:BoundColumn DataField="m_time" 
        HeaderText="時間" 
        ReadOnly="true" 
        Visible="false"/>      
</Columns>
      </asp:DataGrid>
              <strong><font color="#FF9933">
              <asp:DropDownList ID="diff_3" runat="server" AutoPostBack="true">
			  <asp:ListItem Value="8" Selected="true">一週</asp:ListItem>
			  <asp:ListItem Value="15">二週</asp:ListItem>
			  <asp:ListItem Value="22">三週</asp:ListItem>
			  <asp:ListItem Value="31">一個月</asp:ListItem>
			  <asp:ListItem Value="61">二個月</asp:ListItem>
			  <asp:ListItem Value="91">三個月</asp:ListItem>
			  <asp:ListItem Value="9999">所有預定</asp:ListItem>
			  </asp:DropDownList>
              行程:共<font color="#000000"><%= DataSet3.RecordCount %></font> 筆              </font></strong>
              <asp:DataGrid id="DataGrid3" 
  runat="server" 
  AllowSorting="False"
  Width="100%" 
  AutoGenerateColumns="false" 
  CellPadding="3" 
  CellSpacing="0" 
  ShowFooter="false" 
  ShowHeader="true" 
  DataSource="<%# DataSet3.DefaultView %>" 
  PagerStyle-Mode="NumericPages" 
  AllowPaging="true" 
  AllowCustomPaging="true" 
  PageSize="<%# DataSet3.PageSize %>" 
  VirtualItemCount="<%# DataSet3.RecordCount %>" 
  OnPageIndexChanged="DataSet3.OnDataGridPageIndexChanged" 
>
  <HeaderStyle HorizontalAlign="left" BackColor="#FFCCFF" ForeColor="#3D3DB6" Font-Name="Verdana, Arial, Helvetica, sans-serif" Font-Bold="true" Font-Size="smaller" />
  <ItemStyle BackColor="#F2F2F2" Font-Name="Verdana, Arial, Helvetica, sans-serif" Font-Size="smaller" Wrap="true" />
  <AlternatingItemStyle BackColor="#FFCCFF" Font-Name="Verdana, Arial, Helvetica, sans-serif" Font-Size="smaller" />
  <FooterStyle HorizontalAlign="left" BackColor="#E8EBFD" ForeColor="#3D3DB6" Font-Name="Verdana, Arial, Helvetica, sans-serif" Font-Bold="true" Font-Size="smaller" />
  <PagerStyle BackColor="white" Font-Name="Verdana, Arial, Helvetica, sans-serif" Font-Size="smaller" />
      <Columns>
      <asp:TemplateColumn
	    HeaderText="日期" ItemStyle-Width="25%" 
        Visible="True">
        <ItemTemplate><%# showdate(DataSet3.FieldValue("m_date", Container)) %> <br /> <%# showtime(DataSet3.FieldValue("m_time", Container)) %><font color="#3333CC"><%# showweek(DataSet3.FieldValue("m_date", Container)) %></font> </ItemTemplate>
      </asp:TemplateColumn>
      <asp:TemplateColumn
	    HeaderText="行程" ItemStyle-Width="15%"
        Visible="True">
        <ItemTemplate><%# showmessage(DataSet3.FieldValue("m_messeng", Container),DataSet3.FieldValue("m_hours", Container)) %> </ItemTemplate>
      </asp:TemplateColumn>
                <asp:HyperLinkColumn
        DataNavigateUrlField="m_num"
        DataNavigateUrlFormatString="s01cal_detail.aspx?m_num={0}"
        DataTextField="m_note" 
        Visible="True" target="cal_detail" 
        ItemStyle-Width="50%"
		HeaderText="主題"/> </Columns>
</asp:DataGrid>
</td>
</tr>

</table>
        <p>
        <asp:Label id=Label1 runat="server" />
    </form>


</body>
</html>
