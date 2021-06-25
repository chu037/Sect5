<%@ Page Language="VB" ContentType="text/html"%>
<%@ Register TagPrefix="MM" Namespace="DreamweaverCtrls" Assembly="DreamweaverCtrls,version=1.0.0.0,publicKeyToken=836f606ede05d46a,culture=neutral" %>
<MM:DataSet 
id="DataSet1"
runat="Server"
IsStoredProcedure="false"
ConnectionString='<%# System.Configuration.ConfigurationSettings.AppSettings("MM_CONNECTION_STRING_s01") %>'
DatabaseType='<%# System.Configuration.ConfigurationSettings.AppSettings("MM_CONNECTION_DATABASETYPE_s01") %>'
CommandText='<%# "SELECT diff, m_date, m_hours, m_messeng, m_note, m_num, m_time, diffen, m_date_en  FROM s01_calen  WHERE diff = 0 or diff < 0 and diffen > -1  ORDER BY m_time ASC" %>'
Debug="true"
></MM:DataSet>
<MM:DataSet 
id="DataSet2"
runat="Server"
IsStoredProcedure="false"
ConnectionString='<%# System.Configuration.ConfigurationSettings.AppSettings("MM_CONNECTION_STRING_s01") %>'
DatabaseType='<%# System.Configuration.ConfigurationSettings.AppSettings("MM_CONNECTION_DATABASETYPE_s01") %>'
CommandText='<%# "SELECT diff, m_date, m_hours, m_messeng, m_note, m_time, diffen, m_date_en, m_num  FROM s01_calen  WHERE diff = 1   ORDER BY m_time ASC" %>'
Debug="true"
></MM:DataSet>
<MM:DataSet 
id="DataSet3"
runat="Server"
IsStoredProcedure="false"
ConnectionString='<%# System.Configuration.ConfigurationSettings.AppSettings("MM_CONNECTION_STRING_s01") %>'
DatabaseType='<%# System.Configuration.ConfigurationSettings.AppSettings("MM_CONNECTION_DATABASETYPE_s01") %>'
CommandText='<%# "SELECT diff, m_date, m_hours, m_messeng, m_note, m_time, m_num FROM s01_calen  WHERE diff < ? and diff > 1  ORDER BY m_date ASC" %>'
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
<meta http-equiv="Content-Type" content="text/html; charset=UTF-8" />
<meta http-equiv="refresh" content="600"/>
<title>一組行事曆</title>

<script language="VB" runat="server">
        Sub Page_Load(sender As Object, e As EventArgs)
        if not ispostback then
        session("cancel") = ""
		end if 
		End Sub
  function duty_d3(st_d,v_d) '不論人數是否為3的倍數皆通用
   	  dim diff_w as integer = datepart("w",v_d) 'v1是一週內的第幾天
	  dim w_n as integer = int(datediff("d",st_d,v_d)/7) '第幾週.0~
	   select case diff_w
	    case 1, 6, 7
		duty_d3 = 3 + w_n
		case 2, 3
		duty_d3 = 3 + w_n + 1
		case 4, 5
		duty_d3 = 3 + w_n + 2
		end select  
     end function
  function duty_d(st_d,v_d) 
   	  dim diff_w as integer = datepart("w",v_d) 'v1是一週內的第幾天
	  dim w_n as integer = int(datediff("d",st_d,v_d)/7) '第幾週.0~
	   select case diff_w
	    case 1, 6, 7
		duty_d = 3 * w_n
		case 2, 3
		duty_d = 3 * w_n + 1
		case 4, 5
		duty_d = 3 * w_n + 2
		end select  
     end function
  function duty_d2(st_d,v_d) 
   	  dim diff_w as integer = datepart("w",v_d) 'v1是一週內的第幾天
	  dim w_n as integer = int(datediff("d",st_d,v_d)/7) '第幾週.0~
	   select case diff_w
	    case 2, 3, 4
		duty_d2 = 2 * w_n
		case 5, 6, 7, 1
		duty_d2 = 2 * w_n + 1
		end select  
     end function

        Sub Calendar1_DayRender(sender As Object, e As DayRenderEventArgs)
      Dim Conn As OleDbConnection
      Dim Cmd  As OleDbCommand
      Dim Rd   As OleDbDataReader
      Dim I    As Integer
      Dim I1   As Integer

      Dim Provider = "Provider=Microsoft.Jet.OLEDB.4.0"
      Dim Database = "Data Source=" & Server.MapPath( "result/result.mdb" )
      Dim v as CalendarDay
      Dim c as TableCell
      Conn = New OleDbConnection( Provider & ";" & DataBase )
      Conn.Open()
      Dim SQL = "Select * From s01_calen where m_date = @m_date" 
      Cmd = New OleDbCommand( SQL, Conn )
       cmd.Parameters.Clear()
       cmd.Parameters.Add("m_date", e.day.Date)
      Rd = Cmd.ExecuteReader()
            v = e.Day
            c = e.Cell
      dim v1
			v1 = v.date 
	  dim start_d() as date = {#4/1/2010#,#7/12/2010#,#10/22/2010#,#11/8/2010#,#12/24/2010#,#3/26/2011#,#8/1/2011#,#9/2/2011#,#11/23/2011#,#4/30/2012#}
	  dim dx as integer
	  for dx = 0 to start_d.length-1
	   if dx < start_d.length-1
	   dim diff_d0 = datediff("d",start_d(dx),v1)
	   dim diff_d = datediff("d",start_d(dx+1),v1)
	   if diff_d0 >= 0 and diff_d < 0 then exit for
	   else
	   dim diff_d0 = datediff("d",start_d(dx),v1)
	   if diff_d0 >= 0 then exit for
	  end if
	   next

	  select case dx
	   case 0
		dim author_w() = {"段濬豪","嚴瑞宏","洪文哲","郭進寶","陳柏良"}
	    dim author_h() = {"林天成","郭進寶","林炳賦","嚴瑞宏","陳柏良","段濬豪","洪文哲"}
			  dim diff_w as integer = datepart("w",start_d(dx))
	  dim diff_star_d = datediff("d",start_d(dx),v1) + diff_w '+3 星期六距開始日幾天 +之後除以7會=0

	  dim au_w as integer = author_w.length
	  dim au_h as integer = author_h.length
	  dim diff_star_h = (int((datediff("d",start_d(dx),v1)/7))) mod au_h '假日輪勤第幾個星期 

    Try
		 if diff_star_d mod 7 >1 then
		  'i = (datediff("d",start_d0,v1) + diff_star_w0) mod 7
           i = (datediff("d",start_d(dx),v1)) mod au_w
			dim author as string = author_w(i)
			c.Controls.Add(new LiteralControl("<br>" + author))
		 else
		try 
		  i1 = diff_star_h
            dim author_hol as string = author_h(i1)
			c.Controls.Add(new LiteralControl("<br>" + author_hol))
         Catch exc as Exception
         End Try
		 end if
     Catch exc as Exception
	end try	
       case 1
	   	dim author_w() = {"洪文哲","郭進寶","陳柏良","陳泰安","嚴瑞宏"}
	    dim author_h() = {"郭進寶","林炳賦","嚴瑞宏","陳柏良","陳泰安","洪文哲","林天成"}
			  dim diff_w as integer = datepart("w",start_d(dx))
	  dim diff_star_d = datediff("d",start_d(dx),v1) + diff_w '+3 星期六距開始日幾天 +之後除以7會=0

	  dim au_w as integer = author_w.length
	  dim au_h as integer = author_h.length
	  dim diff_star_h = (int((datediff("d",start_d(dx),v1)/7))) mod au_h '假日輪勤第幾個星期 

    Try
		 if diff_star_d mod 7 >1 then
		  'i = (datediff("d",start_d0,v1) + diff_star_w0) mod 7
           i = (datediff("d",start_d(dx),v1)) mod au_w
			dim author as string = author_w(i)
			c.Controls.Add(new LiteralControl("<br>" + author))
		 else
		try 
		  i1 = diff_star_h
            dim author_hol as string = author_h(i1)
			c.Controls.Add(new LiteralControl("<br>" + author_hol))
         Catch exc as Exception
         End Try
		 end if
     Catch exc as Exception
	end try	
       case 2
	   	dim author_w() = {"陳泰安","嚴瑞宏","洪文哲","郭進寶"}
	    dim author_h() = {"郭進寶","林炳賦","嚴瑞宏"}
			  dim diff_w as integer = datepart("w",start_d(dx))
	  dim diff_star_d = datediff("d",start_d(dx),v1) + diff_w '+3 星期六距開始日幾天 +之後除以7會=0

	  dim au_w as integer = author_w.length
	  dim au_h as integer = author_h.length
	  dim diff_star_h = (int((datediff("d",start_d(dx),v1)/7))) mod au_h '假日輪勤第幾個星期 

    Try
		 if diff_star_d mod 7 >1 then
		  'i = (datediff("d",start_d0,v1) + diff_star_w0) mod 7
           i = (datediff("d",start_d(dx),v1)) mod au_w
			dim author as string = author_w(i)
			c.Controls.Add(new LiteralControl("<br>" + author))
		 else
		try 
		  i1 = diff_star_h
            dim author_hol as string = author_h(i1)
			c.Controls.Add(new LiteralControl("<br>" + author_hol))
         Catch exc as Exception
         End Try
		 end if
     Catch exc as Exception
	end try	
       case 3
	   	dim author_w() = {"嚴瑞宏","洪文哲","郭進寶","陳泰安","詹兆熙"}
	    dim author_h() = {"陳泰安","洪文哲","郭進寶","林炳賦","嚴瑞宏","詹兆熙"}
			  dim diff_w as integer = datepart("w",start_d(dx))
	  dim diff_star_d = datediff("d",start_d(dx),v1) + diff_w '+3 星期六距開始日幾天 +之後除以7會=0

	  dim au_w as integer = author_w.length
	  dim au_h as integer = author_h.length
	  dim diff_star_h = (int((datediff("d",start_d(dx),v1)/7))) mod au_h '假日輪勤第幾個星期 

    Try
		 if diff_star_d mod 7 >1 then
		  'i = (datediff("d",start_d0,v1) + diff_star_w0) mod 7
           i = (datediff("d",start_d(dx),v1)) mod au_w
			dim author as string = author_w(i)
			c.Controls.Add(new LiteralControl("<br>" + author))
		 else
		try 
		  i1 = diff_star_h
            dim author_hol as string = author_h(i1)
			c.Controls.Add(new LiteralControl("<br>" + author_hol))
         Catch exc as Exception
         End Try
		 end if
     Catch exc as Exception
	end try	
       case 4
	   	dim author_w() = {"郭進寶","陳泰安","詹兆熙","嚴瑞宏"}
	    dim author_h() = {"陳泰安","林天成","郭進寶","林炳賦","嚴瑞宏","詹兆熙"}
			  dim diff_w as integer = datepart("w",start_d(dx))
	  dim diff_star_d = datediff("d",start_d(dx),v1) + diff_w '+3 星期六距開始日幾天 +之後除以7會=0

	  dim au_w as integer = author_w.length
	  dim au_h as integer = author_h.length
	  dim diff_star_h = (int((datediff("d",start_d(dx),v1)/7))) mod au_h '假日輪勤第幾個星期 

    Try
		 if diff_star_d mod 7 >1 then
		  'i = (datediff("d",start_d0,v1) + diff_star_w0) mod 7
           i = (datediff("d",start_d(dx),v1)) mod au_w
			dim author as string = author_w(i)
			c.Controls.Add(new LiteralControl("<br>" + author))
		 else
		try 
		  i1 = diff_star_h
            dim author_hol as string = author_h(i1)
			c.Controls.Add(new LiteralControl("<br>" + author_hol))
         Catch exc as Exception
         End Try
		 end if
     Catch exc as Exception
	end try	
       case 5
	   	dim author_w() = {"郭進寶","陳泰安","詹兆熙","嚴瑞宏"}
	    dim author_h() = {"郭進寶","林炳賦","嚴瑞宏","詹兆熙","陳泰安"}
			  dim diff_w as integer = datepart("w",start_d(dx))
	  dim diff_star_d = datediff("d",start_d(dx),v1) + diff_w '+3 星期六距開始日幾天 +之後除以7會=0

	  dim au_w as integer = author_w.length
	  dim au_h as integer = author_h.length
	  dim diff_star_h = (int((datediff("d",start_d(dx),v1)/7))) mod au_h '假日輪勤第幾個星期 

    Try
		 if diff_star_d mod 7 >1 then
		  'i = (datediff("d",start_d0,v1) + diff_star_w0) mod 7
           i = (datediff("d",start_d(dx),v1)) mod au_w
			dim author as string = author_w(i)
			c.Controls.Add(new LiteralControl("<br>" + author))
		 else
		try 
		  i1 = diff_star_h
            dim author_hol as string = author_h(i1)
			c.Controls.Add(new LiteralControl("<br>" + author_hol))
         Catch exc as Exception
         End Try
		 end if
     Catch exc as Exception
	end try	
       case 6
	   	dim author_w() = {"郭進寶","陳泰安","詹兆熙","嚴瑞宏","陳怡臻"}
	    dim author_h() = {"陳泰安","郭進寶","林炳賦","嚴瑞宏","詹兆熙","陳怡臻"}
			  dim diff_w as integer = datepart("w",start_d(dx))
	  dim diff_star_d = datediff("d",start_d(dx),v1) + diff_w '+3 星期六距開始日幾天 +之後除以7會=0

	  dim au_w as integer = author_w.length
	  dim au_h as integer = author_h.length
	  dim diff_star_h = (int((datediff("d",start_d(dx),v1)/7))) mod au_h '假日輪勤第幾個星期 

    Try
		 if diff_star_d mod 7 >1 then
		  'i = (datediff("d",start_d0,v1) + diff_star_w0) mod 7
           i = (datediff("d",start_d(dx),v1)) mod au_w
			dim author as string = author_w(i)
			c.Controls.Add(new LiteralControl("<br>" + author))
		 else
		try 
		  i1 = diff_star_h
            dim author_hol as string = author_h(i1)
			c.Controls.Add(new LiteralControl("<br>" + author_hol))
         Catch exc as Exception
         End Try
		 end if
     Catch exc as Exception
	end try	
 
       case 7
	   	dim author_w() = {"嚴瑞宏","郭進寶","陳泰安","詹兆熙","潘玉峰"}
		 i = duty_d(start_d(dx),v1) mod author_w.length
 		dim author as string = author_w(i)
 		c.Controls.Add(new LiteralControl("<br>" + author))

       case 8
	   	dim author_w() = {"陳泰安","洪文哲","詹兆熙","郭進寶","潘玉峰"}
		 i = duty_d3(start_d(dx),v1) mod author_w.length
 		dim author as string = author_w(i)
 		c.Controls.Add(new LiteralControl("<br>" + author))

       case 9
	   	dim author_w() = {"洪文哲","郭進寶","潘玉峰","陳泰安","莊永寧","詹兆熙"}
	    dim w_n as integer = int(datediff("d",start_d(9),v1)/21) '第幾週.0~
		 i = (duty_d2(start_d(dx),v1) + w_n) mod author_w.length
 		dim author as string = author_w(i)
 		c.Controls.Add(new LiteralControl("<br>" + author))

	  end select

 

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

        While rd.Read()
            Dim ltrCr As New LiteralControl("<br>")
            Dim link As New HyperLink()
            link.NavigateUrl = "s01cal_detail.aspx?m_num=" & rd.item(3)
            link.Text = rd.Getstring(0)
            link.target = "cal_detail"
			c.Controls.Add(ltrCr)
            c.Controls.Add(link)
        end while
		conn.close()
       If v.IsOtherMonth Then
           c.Controls.Clear
        end if

		End Sub

    Sub Date_Selected(sender As Object, e As EventArgs)
	  dim url
	  url = "s01cal_selday.aspx?m_date=" & Calendar1.SelectedDate.ToShortDateString
	  response.Redirect( url ) ' 使用Server.Transfer亦可
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
 function showmessage(vt2,vt3)
 if vt2 <> ""
 showmessage = vt2 & "<br/>" & vt3 & "小時"
 end if
 end function
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


-->
</style>
</head>

<body bgcolor="#CCCC99">

    <h3 align="center"><strong><font color="#CC3300" size="+2" face="Verdana, 新細明體">製造業科行事曆</font><font color="#339900" face="Verdana, 新細明體">新增或更改行程請點選日期</font></strong></h3>

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
            <td width="61%" rowspan="3" valign="top"><strong><font color="#FF0000">今日行程:共</font></strong>              <font color="#000000"><strong><%= DataSet1.RecordCount %></strong></font><strong><font color="#FF0000">筆</font></strong>
              <asp:DataGrid AllowPaging="false" 
  AllowSorting="False" AlternatingItemStyle-Wrap="false" 
  AutoGenerateColumns="false" 
  CellPadding="3" 
  CellSpacing="0" 
  DataSource="<%# DataSet1.DefaultView %>" EditItemStyle-Wrap="false" FooterStyle-Wrap="false" HeaderStyle-Wrap="false" ID="DataGrid1" ItemStyle-Wrap="false" PagerStyle-Wrap="false" 
  runat="server" SelectedItemStyle-Wrap="false" 
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
	    HeaderText="日期" ItemStyle-Width="25%" 
        Visible="True">
                  <itemtemplate><%# showdate(DataSet1.FieldValue("m_date", Container)) %> <br />
                    <%# showtime(DataSet1.FieldValue("m_time", Container)) %> </itemtemplate>
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
        ItemStyle-Width="50%"
		HeaderText="主題"/>                
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
  DataSource="<%# DataSet2.DefaultView %>" EditItemStyle-Wrap="false" FooterStyle-Wrap="false" HeaderStyle-Wrap="false" id="DataGrid2" ItemStyle-Wrap="false" PagerStyle-Wrap="false" 
  runat="server" SelectedItemStyle-Wrap="false" 
  ShowFooter="false" 
  ShowHeader="true" Width="100%" 
>
        <HeaderStyle HorizontalAlign="left" BackColor="#FFFFCC" ForeColor="#3D3DB6" Font-Name="Verdana, Arial, Helvetica, sans-serif" Font-Bold="true" Font-Size="smaller" />        
        <ItemStyle BackColor="#F2F2F2" Font-Name="Verdana, Arial, Helvetica, sans-serif" Font-Size="smaller" />        
        <AlternatingItemStyle BackColor="#FFFFCC" Font-Name="Verdana, Arial, Helvetica, sans-serif" Font-Size="smaller" />        
        <FooterStyle HorizontalAlign="left" BackColor="#E8EBFD" ForeColor="#3D3DB6" Font-Name="Verdana, Arial, Helvetica, sans-serif" Font-Bold="true" Font-Size="smaller" />        
        <PagerStyle BackColor="white" Font-Name="Verdana, Arial, Helvetica, sans-serif" Font-Size="smaller" />        
      <Columns>
      <asp:TemplateColumn
	    HeaderText="日期" ItemStyle-Width="25%" 
        Visible="True">
        <ItemTemplate><%# showdate(DataSet2.FieldValue("m_date", Container)) %> <br />
            <%# showtime(DataSet2.FieldValue("m_time", Container)) %> </ItemTemplate>
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
  <ItemStyle BackColor="#F2F2F2" Font-Name="Verdana, Arial, Helvetica, sans-serif" Font-Size="smaller" />
  <AlternatingItemStyle BackColor="#FFCCFF" Font-Name="Verdana, Arial, Helvetica, sans-serif" Font-Size="smaller" />
  <FooterStyle HorizontalAlign="left" BackColor="#E8EBFD" ForeColor="#3D3DB6" Font-Name="Verdana, Arial, Helvetica, sans-serif" Font-Bold="true" Font-Size="smaller" />
  <PagerStyle BackColor="white" Font-Name="Verdana, Arial, Helvetica, sans-serif" Font-Size="smaller" />
      <Columns>
      <asp:TemplateColumn
	    HeaderText="日期" ItemStyle-Width="25%" 
        Visible="True">
        <ItemTemplate><%# showdate(DataSet3.FieldValue("m_date", Container)) %> <br /> <%# showtime(DataSet3.FieldValue("m_time", Container)) %> </ItemTemplate>
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
