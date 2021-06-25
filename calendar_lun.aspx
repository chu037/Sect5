<%@ Page Language="VB" ContentType="text/html"  %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<%@ Import Namespace="System.Data.OleDb" %>
<%@ Import Namespace="System.Globalization.Calendar" %>
<%@ Import Namespace="System.Globalization.EastAsianLunisolarCalendar" %>
<%@ Import Namespace="System.Globalization.TaiwanLunisolarCalendar" %>
<%@ Import Namespace="System.Globalization" %>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5" />
<title>�L���D���</title>
<script language="VB" runat="server">
    Sub Calendar1_DayRender(sender As Object, e As DayRenderEventArgs)
        Dim v As CalendarDay
        Dim c As TableCell
        Dim v1 As DateTime
        Dim instance As New TaiwanLunisolarCalendar()
        Dim rev As Integer
        Dim rev_m As Integer
	  
        Dim day_arry() = {"��@", "��G", "��T", "��|", "�줭", "�줻", "��C", "��K", "��E", "�Q", "�Q�@", "�Q�G", "�Q�T", "�Q�|", "�Q��", "�Q��", "�Q�C", "�Q�K", "�Q�E", "��", "�ܤ@", "�ܤG", "�ܤT", "�ܥ|", "�ܤ�", "�ܤ�", "�ܤC", "�ܤK", "�ܤE", "��"}
        Dim month_arry() = {"�@��", "�G��", "�T��", "�|��", "����", "����", "�C��", "�K��", "�E��", "�Q��", "�Q�@��", "�Q�G��"}
        Dim fe_arry() = {"�߬K", "�B��", "���h", "�K��", "�M��", "�\�B", "�߮L", "�p��", "�~��", "�L��", "�p��", "�j��", "�߬�", "�B��", "���S", "���", "�H�S", "����", "�ߥV", "�p��", "�j��", "�V��", "�p�H", "�j�H"}
        v = e.Day
        v1 = v.Date
        c = e.Cell
			
        rev = instance.GetDayOfMonth(v1) - 1
        rev_m = instance.GetMonth(v1) - 1
        Dim ltrCr As New LiteralControl("<br>")
        If rev = 0 Then
            c.Controls.Add(New LiteralControl("<br>" + "<font style=""color:#ff9966;font-size:10pt"">" + month_arry(rev_m)))
        Else
            c.Controls.Add(New LiteralControl("<br>" + "<font style=""color:#ffcc33;font-size:10pt"">" + day_arry(rev)))
        End If
        'c.Controls.Add(ltrcr)
       
        If e.Day.IsOtherMonth Then
            c.Controls.Clear()
        End If

    End Sub

    Sub page_load(sender As Object, e As EventArgs)
        Dim tlc As New TaiwanLunisolarCalendar()
        tlc.MaxSupportedDateTime.ToShortDateString()
        txtContent.Text = tlc.GetDayOfMonth(#1/1/2015#).ToString()
    End Sub

    Function showdate(s01time As DateTime) As Date
        If s01time <> "" Then
            showdate = FormatDateTime(s01time, DateFormat.ShortDate)
        End If
        Return showdate
    End Function
    Function GetDayOfMonth(time1 As DateTime) As Integer
        If time1 <> "" Then
            GetDayOfMonth = GetDayOfMonth(time1)
        End If
        Return GetDayOfMonth
    End Function

    </script>

</head>
<body>
    <form runat=server>

      <p>
        <asp:Calendar BorderColor="#CC9933"
            BorderWidth="1" DayHeaderStyle-BorderColor="#CC9933" DayHeaderStyle-Font-Size="16" DayHeaderStyle-ForeColor="#0000CC" DayNameFormat="Shortest" DayStyle-BorderColor="#CC9933" DayStyle-Font-Bold="true"
            DayStyle-Height="30" DayStyle-HorizontalAlign="center" DayStyle-VerticalAlign="top"
            DayStyle-Width="50"
            Font-Name="Verdana"
            Font-Size="16" Font-Underline="false" ID=Calendar1
            NextMonthText = "�U�@��" NextPrevStyle-BorderColor="#990033" NextPrevStyle-Font-Underline="false" NextPrevStyle-ForeColor="#0000FF" NextPrevStyle-Wrap="false"
            PrevMonthText = "�W�@��" runat="server"
            SelectedDayStyle-BackColor="#FFCC66" SelectedDayStyle-BorderColor="#FF9933" SelectedDayStyle-ForeColor="#000000" 
            ShowGridLines="true"
            TitleStyle-BackColor="Gainsboro" TitleStyle-BorderColor="#FF9966"
            TitleStyle-Font-Bold="true"
            TitleStyle-Font-Size="12px"
            TodayDayStyle-BackColor="#CCFF33" TodayDayStyle-BorderColor="#FF9966" TodayDayStyle-Font-Bold="false"
            TodayDayStyle-ForeColor="#993333" WeekendDayStyle-BackColor="#66FF99" WeekendDayStyle-CssClass="calWeekendDay" WeekendDayStyle-Font-Bold="true" WeekendDayStyle-ForeColor="#CC0000"

            Width="500px"
            OnDayRender="Calendar1_DayRender" 
            />      
      </p>
      <p>
        <asp:TextBox ID="txtContent" runat="server" />              </p>
    </form>
</body>
</html>
