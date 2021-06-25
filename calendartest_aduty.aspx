<%@ Page Language="VB" ContentType="text/html"%></MM:DataSet>
<%@ Import Namespace="System.Globalization.Calendar" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<meta http-equiv="refresh" content="600"/>
<title>綜合行業科軍區輪值</title>

<script language="VB" runat="server">
    Sub Page_Load(sender As Object, e As EventArgs)
        If Not IsPostBack Then
            Session("cancel") = ""
        End If
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
        'Dim Conn As OleDbConnection
        'Dim Cmd  As OleDbCommand
        'Dim Rd   As OleDbDataReader
        Dim I As Integer
        Dim j As Integer
        Dim I1 As Integer = 0

        'Dim Provider = "Provider=Microsoft.Jet.OLEDB.4.0"
        'Dim Database = "Data Source=" & Server.MapPath( "result/result.mdb" )
        Dim v As CalendarDay
        Dim c As TableCell
        'Conn = New OleDbConnection( Provider & ";" & DataBase )
        'Conn.Open()
        'Dim SQL = "Select * From s01_calen where m_date = @m_date" 
        'Cmd = New OleDbCommand( SQL, Conn )
        'cmd.Parameters.Clear()
        'cmd.Parameters.Add("m_date", e.day.Date)
        'Rd = Cmd.ExecuteReader()
        v = e.Day
        c = e.Cell
        Dim v1
        v1 = v.Date
        Dim start_d() As Date = {#6/30/2014#, #8/31/2015#, #7/18/2016#} '要從星期一開始
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
                Dim ad_d() = {"鳥", "永", "鳳", "路", "園", "鎮", "美", "橋"}
                Dim author_w() = {"顏廷諭", "姜智敏", "蘇銘源", "楊勝安", "林紫蓁", "吳俐節", "朱志杰"}
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
                        c.Controls.Add(New LiteralControl("<br>" + ad + "-" + author))
                        j = j + 1
                        I = I + 1
                        If I > author_w.Length - 1 Then
                            I = 0
                        End If

                    End While
                End If

            Case 1
                Dim ad_d() = {"鳥", "永", "鳳", "路", "園", "鎮", "美", "橋"}
                Dim author_w() = {"吳俐節", "朱志杰", "顏廷諭", "姜智敏", "蘇銘源", "楊勝安", "林紫蓁", "楊尚淳"}
                Dim w_n As Integer = Int(DateDiff("d", start_d(dx), v1) / 7) Mod (author_w.Length) '第幾輪
                'if author_w.length < ad_d.length '輪的人小於8人
                j = 0
                I = w_n
                While j < ad_d.Length
                    Dim author As String = author_w(I)
                    Dim ad As String = ad_d(j)
                    c.Controls.Add(New LiteralControl("<br>" + ad + "-" + author))
                    j = j + 1
                    I = I + 1
                    If I > author_w.Length - 1 Then
                        I = 0
                    End If

                End While
 
            Case 2
                Dim ad_d() = {"鳥", "永", "鳳", "路", "園", "鎮", "美", "橋"}
                Dim author_w() = {"林紫蓁", "楊尚淳", "吳俐節", "朱志杰", "顏廷諭", "姜智敏", "蘇銘源", "楊勝安", "王孝中"}
                Dim w_n As Integer = Int(DateDiff("d", start_d(dx), v1) / 7) Mod (author_w.Length) '第幾輪
                'if author_w.length < ad_d.length '輪的人小於8人
                j = 0
                I = w_n
                While j < ad_d.Length
                    Dim author As String = author_w(I)
                    Dim ad As String = ad_d(j)
                    c.Controls.Add(New LiteralControl("<br>" + ad + "-" + author))
                    j = j + 1
                    I = I + 1
                    If I > author_w.Length - 1 Then
                        I = 0
                    End If

                End While
                'end if
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

        'While rd.Read()
        'Dim ltrCr As New LiteralControl("<br>")
        'Dim link As New HyperLink()
        'link.NavigateUrl = "s01cal_detail.aspx?m_num=" & rd.item(3)
        'link.Text = rd.Getstring(0)
        'link.target = "cal_detail"
        'c.Controls.Add(ltrCr)
        'c.Controls.Add(link)
        'end while
        'conn.close()
        If v.IsOtherMonth Then
            c.Controls.Clear()
        End If

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
.style1 {color: #0000FF}


-->
</style>
</head>

<body>

    <h3 align="center"><strong><font color="#CC3300" size="+2" face="Verdana, 新細明體">綜合行業科動態稽查輪值</font><font color="#339900" face="Verdana, 新細明體"></font></strong></h3>

    <form runat=server>
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td width="60%" valign="top"><asp:Calendar BorderColor="#CC9933"
            BorderWidth="1" DayHeaderStyle-BorderColor="#CC9933" DayHeaderStyle-Font-Size="12" DayNameFormat="Shortest" DayStyle-BorderColor="#CC9933"
            DayStyle-Height=
            DayStyle-VerticalAlign="Top"
            DayStyle-Width="14%"
            Font-Name="Verdana"
            Font-Size="12" ID=Calendar1
            NextMonthText = "下一月" NextPrevStyle-BorderColor="#990033" NextPrevStyle-Font-Underline="false" NextPrevStyle-ForeColor="#0000FF" NextPrevStyle-Wrap="false"
            PrevMonthText = "上一月" runat="server"
            SelectedDayStyle-BackColor="#FFCC66" SelectedDayStyle-BorderColor="#FF9933" SelectedDayStyle-ForeColor="#000000" ShowDayHeader="true"
            ShowGridLines="true"
            TitleStyle-BackColor="Gainsboro" TitleStyle-BorderColor="#FF9966"
            TitleStyle-Font-Bold="true"
            TitleStyle-Font-Size="12px"
            TodayDayStyle-BackColor="#CCFF33" TodayDayStyle-BorderColor="#FF9966" TodayDayStyle-Font-Bold="false"
            TodayDayStyle-ForeColor="#993333" WeekendDayStyle-CssClass="calWeekendDay"

            Width="600px"
            OnDayRender="Calendar1_DayRender"
            />
</td>
            <td width="40%" valign="top"><table width="100%" border="2" cellspacing="0" cellpadding="1">
			<tr>
			<td width="30%" align="center" valign="top"><span class="style1">軍區</span></td>
			<td width="70%" valign="top"><span class="style1">範圍</span></td>
			</tr>
			<tr>
			<td width="30%" align="center" valign="top">鳥</td>
			<td width="70%" valign="top">鳥松、仁武、大樹、大社</td>
			<tr>
			<td width="30%" align="center" valign="top">永</td><td width="70%" valign="top">永安、彌陀、茄萣</td>
			<tr>
			<td width="30%" align="center" valign="top">鳳</td><td width="70%" valign="top">鳳山、三民、新興、前金、鹽埕、苓雅、鼓山</td>
			<tr>
			<td width="30%" align="center" valign="top">路</td><td width="70%" valign="top">路竹、岡山、湖內、阿蓮</td></tr>
			<tr>
			<td width="30%" align="center" valign="top">園</td><td width="70%" valign="top">林園、大寮</td>
			</tr>
			<tr>
			<td width="30%" align="center" valign="top">鎮</td><td width="70%" valign="top">前鎮、小港、旗津</td>
			</tr>
			<tr>
			<td width="30%" align="center" valign="top">美</td><td width="70%" valign="top">美濃、燕巢、田寮、旗山、內門、杉林...</td>
			</tr>
			<tr>
			<td width="30%" align="center" valign="top">橋</td><td width="70%" valign="top">橋頭、楠梓、左營、梓官</td>
			<tr>
			  <td align="center" valign="top">高屏道路挖掘</td>
			  <td valign="top"><a href="http://kproad.kcg.gov.tw/kproad/gmap/indexkp.asp" target="new_round">http://kproad.kcg.gov.tw/kproad/gmap/indexkp.asp</a></td>
			  <tr>
			    <td align="center" valign="top">公共管線管理平台</td>
			    <td valign="top"><a href="http://pipegis.kcg.gov.tw/default.aspx" target="new_round">http://pipegis.kcg.gov.tw/default.aspx</a></td>
		      </table>
			</tr>
		</table>


    </form>


</body>
</html>
