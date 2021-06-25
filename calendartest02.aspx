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
    End Function
    Function duty_d2(st_d, v_d)
        Dim diff_w As Integer = datepart("w", v_d) 'v1是一週內的第幾天
        Dim w_n As Integer = int(datediff("d", st_d, v_d) / 7) '第幾週.0~
        Select Case diff_w
            Case 2, 3, 4
                duty_d2 = 2 * w_n
            Case 5, 6, 7, 1
                duty_d2 = 2 * w_n + 1
        End Select
    End Function
    Function duty_d4(st_d, v_d)
        Dim diff_w As Integer = DatePart("w", v_d) 'v1是一週內的第幾天
        Dim w_n As Integer = Int(DateDiff("d", st_d, v_d) / 7) '第幾週.0~
        Select Case diff_w
            Case 1, 6, 7
                duty_d4 = 5 * w_n + 1
            Case 2
                duty_d4 = 5 * w_n + 3
            Case 3
                duty_d4 = 5 * w_n + 5
            Case 4
                duty_d4 = 5 * w_n + 7
            Case 5
                duty_d4 = 5 * w_n + 9
        End Select
    End Function
    
    Function duty_d5(st_d, v_d)
        Dim diff_w As Integer = DatePart("w", v_d) 'v1是一週內的第幾天
        Dim w_n As Integer = Int(DateDiff("d", st_d, v_d) / 7) '第幾週.0~
        Select Case diff_w
            Case 1, 6, 7
                duty_d5 = 5 * w_n + 1
            Case 2
                duty_d5 = 5 * w_n + 2
            Case 3
                duty_d5 = 5 * w_n + 3
            Case 4
                duty_d5 = 5 * w_n + 4
            Case 5
                duty_d5 = 5 * w_n + 5
        End Select
    End Function
    Function duty_d6(st_d, v_d)
        Dim diff_w As Integer = DatePart("w", v_d) 'v1是一週內的第幾天
        Dim w_n As Integer = Int(DateDiff("d", st_d, v_d) / 7) '第幾週.0~
        Select Case diff_w
            Case 1, 6, 7
                duty_d6 = 5
            Case 2
                duty_d6 = 5 * w_n + 1
            Case 3
                duty_d6 = 5 * w_n + 2
            Case 4
                duty_d6 = 5 * w_n + 3
            Case 5
                duty_d6 = 5 * w_n + 4
        End Select
    End Function
    Function duty_d7(st_d, v_d)
        Dim diff_w As Integer = DatePart("w", v_d) 'v1是一週內的第幾天
        Dim w_n As Integer = Int(DateDiff("d", st_d, v_d) / 7) '第幾週.0~
        Select Case diff_w
            Case 1, 6, 7
                duty_d7 = 5 * w_n
            Case 2
                duty_d7 = 5 * w_n + 1
            Case 3
                duty_d7 = 5 * w_n + 2
            Case 4
                duty_d7 = 5 * w_n + 3
            Case 5
                duty_d7 = 5 * w_n + 4
        End Select
    End Function
 Function duty_d8(st_d, v_d)
        Dim diff_w As Integer = DatePart("w", v_d) 'v1是一週內的第幾天
        Dim w_n As Integer = Int(DateDiff("d", st_d, v_d) / 7) '第幾週.0~
        Select Case diff_w
            Case 1, 6, 7
                duty_d8 = 5 * w_n + 1
            Case 2
                duty_d8 = 5 * w_n + 2
            Case 3
                duty_d8 = 5 * w_n + 3
            Case 4
                duty_d8 = 5 * w_n + 4
            Case 5
                duty_d8 = 5 * w_n + 5
        End Select
    End Function
        Sub Calendar1_DayRender(sender As Object, e As DayRenderEventArgs)
      Dim Conn As OleDbConnection
      Dim Cmd  As OleDbCommand
      Dim Rd   As OleDbDataReader
        Dim i As Integer
        Dim ii As Integer
        Dim iii As Integer
      Dim i1   As Integer
      Dim i2   As Integer


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
        Dim v1 As Date
			v1 = v.date 
        Dim start_d() As Date = {#12/2/2013#, #2/3/2014#, #7/7/2014#, #10/20/2014#, #11/24/2014#, #1/5/2015#, #8/31/2015#, #3/7/2016#, #5/16/2016#, #8/1/2016#, #12/12/2016#, #2/13/2017#, #11/6/2017#, #2/5/2018#, #3/5/2018#, #4/23/2018#, #5/7/2018#, #5/21/2018#, #5/28/2018#, #7/23/2018#, #8/6/2018#, #12/3/2018#, #1/7/2019#, #2/11/2019#, #4/1/2019#, #5/6/2019#, #6/24/2019#, #9/2/2019#, #10/21/2019#, #12/2/2019#, #2/3/2020#, #8/3/2020#, #1/4/2021#} '要從星期一開始
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
	   	dim author_w() = {"楊勝安", "蘇銘源","吳俐節","顏廷諭","姜智敏","楊尚淳"}
	    if author_w.length mod 2 = 0 then '2的倍數時要+第幾輪參數
		dim w_n as integer = int(datediff("d",start_d(dx),v1)/(author_w.length*7/2)) '第幾輪
		 i = (duty_d2(start_d(dx),v1)+w_n) mod author_w.length 
		else
		 i = (duty_d2(start_d(dx),v1)) mod author_w.length
		end if 
 		dim author as string = author_w(i)
 		c.Controls.Add(new LiteralControl("<br>" + author))

	   case 1

	   	dim author_w() = {"顏廷諭","姜智敏","楊尚淳","楊勝安", "蘇銘源","吳俐節","朱志杰"}
	    if author_w.length mod 2 = 0 then '2的倍數時要+第幾輪參數
		dim w_n as integer = int(datediff("d",start_d(dx),v1)/(author_w.length*7/2)) '第幾輪
		 i = (duty_d2(start_d(dx),v1)+w_n) mod author_w.length 
		else
		 i = (duty_d2(start_d(dx),v1)) mod author_w.length
		end if 
 		dim author as string = author_w(i)
 		c.Controls.Add(new LiteralControl("<br>" + author))

	   case 2

	   	dim author_w() = {"楊尚淳","楊勝安", "陳柏鈞","吳俐節","朱志杰","姜智敏"}
	    if author_w.length mod 2 = 0 then '2的倍數時要+第幾輪參數
		dim w_n as integer = int(datediff("d",start_d(dx),v1)/(author_w.length*7/2)) '第幾輪
		 i = (duty_d2(start_d(dx),v1)+w_n) mod author_w.length 
		else
		 i = (duty_d2(start_d(dx),v1)) mod author_w.length
		end if 
 		dim author as string = author_w(i)
 		c.Controls.Add(new LiteralControl("<br>" + author))

	   case 3

	   	dim author_w() = {"林紫蓁","楊尚淳","顏廷諭", "陳柏鈞","吳俐節","朱志杰"}
	    if author_w.length mod 2 = 0 then '2的倍數時要+第幾輪參數
		dim w_n as integer = int(datediff("d",start_d(dx),v1)/(author_w.length*7/2)) '第幾輪
		 i = (duty_d2(start_d(dx),v1)+w_n) mod author_w.length 
		else
		 i = (duty_d2(start_d(dx),v1)) mod author_w.length
		end if 
 		dim author as string = author_w(i)
 		c.Controls.Add(new LiteralControl("<br>" + author))

	   case 4

	   	dim author_w() = {"朱志杰","姜智敏","林紫蓁","楊勝安","顏廷諭","吳俐節","楊尚淳"}
	    if author_w.length mod 2 = 0 then '2的倍數時要+第幾輪參數
		dim w_n as integer = int(datediff("d",start_d(dx),v1)/(author_w.length*7/2)) '第幾輪
		 i = (duty_d2(start_d(dx),v1)+w_n) mod author_w.length 
		else
		 i = (duty_d2(start_d(dx),v1)) mod author_w.length
		end if 
 		dim author as string = author_w(i)
 		c.Controls.Add(new LiteralControl("<br>" + author))

	   case 5

	   	dim author_w() = {"吳俐節","蘇銘源","朱志杰","姜智敏","林紫蓁","楊勝安","顏廷諭"}
	    if author_w.length mod 2 = 0 then '2的倍數時要+第幾輪參數
		dim w_n as integer = int(datediff("d",start_d(dx),v1)/(author_w.length*7/2)) '第幾輪
		 i = (duty_d2(start_d(dx),v1)+w_n) mod author_w.length 
		else
		 i = (duty_d2(start_d(dx),v1)) mod author_w.length
		end if 
 		dim author as string = author_w(i)
 		c.Controls.Add(new LiteralControl("<br>" + author))

	   case 6

	   	dim author_w() = {"楊勝安","顏廷諭","吳俐節","蘇銘源","朱志杰","姜智敏","林紫蓁","楊尚淳"}
	    if author_w.length mod 2 = 0 then '2的倍數時要+第幾輪參數
		dim w_n as integer = int(datediff("d",start_d(dx),v1)/(author_w.length*7/2)) '第幾輪
		 i = (duty_d2(start_d(dx),v1)+w_n) mod author_w.length 
		else
		 i = (duty_d2(start_d(dx),v1)) mod author_w.length
		end if 
 		dim author as string = author_w(i)
 		c.Controls.Add(new LiteralControl("<br>" + author))

	   case 7

	   	dim author_w() = {"朱志杰","姜智敏","林紫蓁","楊尚淳","楊勝安","顏廷諭","吳俐節","蘇銘源","王孝中"}
	    if author_w.length mod 2 = 0 then '2的倍數時要+第幾輪參數
		dim w_n as integer = int(datediff("d",start_d(dx),v1)/(author_w.length*7/2)) '第幾輪
		 i = (duty_d2(start_d(dx),v1)+w_n) mod author_w.length 
		else
		 i = (duty_d2(start_d(dx),v1)) mod author_w.length
		end if 
 		dim author as string = author_w(i)
 		c.Controls.Add(new LiteralControl("<br>" + author))

	   case 8

	   	dim author_w() = {"楊尚淳","楊勝安","顏廷諭","吳俐節","蘇銘源","王孝中","朱志杰","姜智敏"}
	    if author_w.length mod 2 = 0 then '2的倍數時要+第幾輪參數
		dim w_n as integer = int(datediff("d",start_d(dx),v1)/(author_w.length*7/2)) '第幾輪
		 i = (duty_d2(start_d(dx),v1)+w_n) mod author_w.length 
		else
		 i = (duty_d2(start_d(dx),v1)) mod author_w.length
		end if 
 		dim author as string = author_w(i)
 		c.Controls.Add(new LiteralControl("<br>" + author))

	   case 9

	   	dim author_w() = {"楊尚淳","楊勝安","顏廷諭","吳俐節","蘇銘源","林紫蓁","王孝中","朱志杰","姜智敏"}
	    if author_w.length mod 2 = 0 then '2的倍數時要+第幾輪參數
		dim w_n as integer = int(datediff("d",start_d(dx),v1)/(author_w.length*7/2)) '第幾輪
		 i = (duty_d2(start_d(dx),v1)+w_n) mod author_w.length
		else
		 i = (duty_d2(start_d(dx),v1)) mod author_w.length
		end if 
		  if i = author_w.length-1
		  i1 = 0
		  else
		  i1 = i+1
		  end if 
 		dim author as string = author_w(i)
 		dim author1 as string = author_w(i1)
 		c.Controls.Add(new LiteralControl("<br>" + "內-" + author + "<br>" + "外-" + author1))

	   case 10

	   	dim author_w() = {"顏廷諭","蘇銘源","林紫蓁","王孝中","朱志杰","姜智敏","楊尚淳","楊勝安"}
	    if author_w.length mod 2 = 0 then '2的倍數時要+第幾輪參數
		dim w_n as integer = int(datediff("d",start_d(dx),v1)/(author_w.length*7/2)) '第幾輪
		 i = (duty_d2(start_d(dx),v1)+w_n) mod author_w.length
		else
		 i = (duty_d2(start_d(dx),v1)) mod author_w.length
		end if 
		  if i = author_w.length-1
		  i1 = 0
		  else
		  i1 = i+1
		  end if 
 		dim author as string = author_w(i)
 		dim author1 as string = author_w(i1)
 		c.Controls.Add(new LiteralControl("<br>" + "內-" + author + "<br>" + "外-" + author1))

	   case 11

	   	dim author_w() = {"朱志杰","姜智敏","楊尚淳","楊勝安","顏廷諭","蘇銘源","林紫蓁","王孝中"}
	    if author_w.length mod 2 = 0 then '2的倍數時要+第幾輪參數
		dim w_n as integer = int(datediff("d",start_d(dx),v1)/(author_w.length*7/2)) '第幾輪
		 i = (duty_d2(start_d(dx),v1)+w_n) mod author_w.length
		else
		 i = (duty_d2(start_d(dx),v1)) mod author_w.length
        end if
		 i1 = i+1
		 i2 = i1+1
		  if i = author_w.length-1
		  i1 = 0
		  i2 = 1
		  else
   	       if i1 = author_w.length-1
           i2 = 0
		   end if
		  end if

 		dim author as string = author_w(i)
 		dim author1 as string = author_w(i1)
 		dim author2 as string = author_w(i2)
 		c.Controls.Add(new LiteralControl("<br>" + "內-" + author +  "<br>" + "外-" + author1 + "<br>" +  "外-" + author2 ))

	   case 12

	   	dim author_w() = {"蘇銘源","林紫蓁","王孝中","朱志杰","姜智敏","楊尚淳","顏廷諭"}
	    if author_w.length mod 2 = 0 then '2的倍數時要+第幾輪參數
		dim w_n as integer = int(datediff("d",start_d(dx),v1)/(author_w.length*7/2)) '第幾輪
		 i = (duty_d2(start_d(dx),v1)+w_n) mod author_w.length
		else
		 i = (duty_d2(start_d(dx),v1)) mod author_w.length
        end if
		 i1 = i+1
		 i2 = i1+1
		  if i = author_w.length-1
		  i1 = 0
		  i2 = 1
		  else
   	       if i1 = author_w.length-1
           i2 = 0
		   end if
		  end if

 		dim author as string = author_w(i)
 		dim author1 as string = author_w(i1)
 		dim author2 as string = author_w(i2)
                c.Controls.Add(New LiteralControl("<br>" + "內-" + author + "<br>" + "外-" + author1 + "<br>" + "外-" + author2))
                
            Case 13

                Dim author_w() = {"賴易昌", "賴曉蓉", "蘇銘源", "林紫蓁", "王孝中", "朱志杰", "姜智敏", "楊尚淳", "顏廷諭"}
                If author_w.Length Mod 2 = 0 Then '2的倍數時要+第幾輪參數
                    Dim w_n As Integer = Int(DateDiff("d", start_d(dx), v1) / (author_w.Length * 7 / 2)) '第幾輪
                    i = (duty_d2(start_d(dx), v1) + w_n) Mod author_w.Length
                Else
                    i = (duty_d2(start_d(dx), v1)) Mod author_w.Length
                End If
                i1 = i + 1
                i2 = i1 + 1
                If i = author_w.Length - 1 Then
                    i1 = 0
                    i2 = 1
                Else
                    If i1 = author_w.Length - 1 Then
                        i2 = 0
                    End If
                End If

                Dim author As String = author_w(i)
                Dim author1 As String = author_w(i1)
                Dim author2 As String = author_w(i2)
                c.Controls.Add(New LiteralControl("<br>" + "內-" + author + "<br>" + "外-" + author1 + "<br>" + "外-" + author2))
            Case 14

                Dim author_w() = {"陳進富", "賴易昌", "賴曉蓉", "蘇銘源", "林紫蓁", "王孝中", "朱志杰", "姜智敏", "楊尚淳", "顏廷諭"}
                If author_w.Length Mod 2 = 0 Then '2的倍數時要+第幾輪參數
                    Dim w_n As Integer = Int(DateDiff("d", start_d(dx), v1) / (author_w.Length * 7 / 2)) '第幾輪
                    i = (duty_d2(start_d(dx), v1) + w_n) Mod author_w.Length
                Else
                    i = (duty_d2(start_d(dx), v1)) Mod author_w.Length
                End If
                i1 = i + 1
                i2 = i1 + 1
                If i = author_w.Length - 1 Then
                    i1 = 0
                    i2 = 1
                Else
                    If i1 = author_w.Length - 1 Then
                        i2 = 0
                    End If
                End If

                Dim author As String = author_w(i)
                Dim author1 As String = author_w(i1)
                Dim author2 As String = author_w(i2)
                c.Controls.Add(New LiteralControl("<br>" + "內-" + author + "<br>" + "外-" + author1 + "<br>" + "外-" + author2))
            Case 15

                Dim author_w() = {"劉建台", "陳進富", "賴易昌", "賴曉蓉", "林紫蓁", "王孝中", "朱志杰", "姜智敏", "楊尚淳", "顏廷諭"}
                If author_w.Length Mod 2 = 0 Then '2的倍數時要+第幾輪參數
                    Dim w_n As Integer = Int(DateDiff("d", start_d(dx), v1) / (author_w.Length * 7 / 2)) '第幾輪
                    i = (duty_d2(start_d(dx), v1) + w_n) Mod author_w.Length
                Else
                    i = (duty_d2(start_d(dx), v1)) Mod author_w.Length
                End If
                i1 = i + 1
                i2 = i1 + 1
                If i = author_w.Length - 1 Then
                    i1 = 0
                    i2 = 1
                Else
                    If i1 = author_w.Length - 1 Then
                        i2 = 0
                    End If
                End If

                Dim author As String = author_w(i)
                Dim author1 As String = author_w(i1)
                Dim author2 As String = author_w(i2)
                c.Controls.Add(New LiteralControl("<br>" + "內-" + author + "<br>" + "外-" + author1 + "<br>" + "外-" + author2))
            Case 16

                Dim author_w() = {"賴易昌", "劉建台", "姜智敏", "林紫蓁", "朱志杰", "王孝中", "楊尚淳", "顏廷諭", "陳進富", "賴曉蓉"}
                If author_w.Length Mod 5 = 0 Then '2的倍數時要+第幾輪參數
                    Dim w_n As Integer = Int(DateDiff("d", start_d(dx), v1) / (author_w.Length * 7 / 5)) '第幾輪
                    i = (duty_d4(start_d(dx), v1) + w_n) Mod author_w.Length
                Else
                    i = (duty_d4(start_d(dx), v1)) Mod author_w.Length
                End If
                i1 = i + 1
                i2 = i1 + 1
                If i = author_w.Length - 1 Then
                    i1 = 0
                    i2 = 1
                Else
                    If i1 = author_w.Length - 1 Then
                        i2 = 0
                    End If
                End If

                Dim author As String = author_w(i)
                Dim author1 As String = author_w(i1)
                Dim author2 As String = author_w(i2)
                c.Controls.Add(New LiteralControl("<br>" + "內-" + author + "<br>" + "外-" + author1 + "<br>" + "外-" + author2))
            Case 17

                Dim author_w() = {"賴易昌", "劉建台", "林紫蓁", "朱志杰", "王孝中", "楊尚淳", "顏廷諭", "陳進富", "賴曉蓉"}
                If author_w.Length Mod 5 = 0 Then '2的倍數時要+第幾輪參數
                    Dim w_n As Integer = Int(DateDiff("d", start_d(dx), v1) / (author_w.Length * 7 / 5)) '第幾輪
                    i = (duty_d4(start_d(dx), v1) + w_n) Mod author_w.Length
                Else
                    i = (duty_d4(start_d(dx), v1)) Mod author_w.Length
                End If
                i1 = i + 1
                i2 = i1 + 1
                If i = author_w.Length - 1 Then
                    i1 = 0
                    i2 = 1
                Else
                    If i1 = author_w.Length - 1 Then
                        i2 = 0
                    End If
                End If

                Dim author As String = author_w(i)
                Dim author1 As String = author_w(i1)
                Dim author2 As String = author_w(i2)
                c.Controls.Add(New LiteralControl("<br>" + "內-" + author + "<br>" + "外-" + author1 + "<br>" + "外-" + author2))
            Case 18

                Dim author_w() = {"賴易昌", "劉建台", "朱志杰", "王孝中", "楊尚淳", "顏廷諭", "陳進富", "賴曉蓉"}
                If author_w.Length Mod 5 = 0 Then '2的倍數時要+第幾輪參數
                    Dim w_n As Integer = Int(DateDiff("d", start_d(dx), v1) / (author_w.Length * 7 / 5)) '第幾輪
                    i = (duty_d5(start_d(dx), v1) + w_n) Mod author_w.Length
                Else
                    i = (duty_d5(start_d(dx), v1)) Mod author_w.Length
                End If
                i1 = i + 1
                i2 = i1 + 1
                If i = author_w.Length - 1 Then
                    i1 = 0
                    i2 = 1
                Else
                    If i1 = author_w.Length - 1 Then
                        i2 = 0
                    End If
                End If

                Dim author As String = author_w(i)
                Dim author1 As String = author_w(i1)
                Dim author2 As String = author_w(i2)
                c.Controls.Add(New LiteralControl("<br>" + "內-" + author + "<br>" + "外-" + author1))
            Case 19

                Dim author_w() = {"賴易昌", "劉建台", "朱志杰", "王孝中", "楊尚淳", "陳進富", "賴曉蓉"}
                If author_w.Length Mod 5 = 0 Then '2的倍數時要+第幾輪參數
                    Dim w_n As Integer = Int(DateDiff("d", start_d(dx), v1) / (author_w.Length * 7 / 5)) '第幾輪
                    i = (duty_d5(start_d(dx), v1) + w_n) Mod author_w.Length
                Else
                    i = (duty_d5(start_d(dx), v1)) Mod author_w.Length
                End If
                i1 = i + 1
                i2 = i1 + 1
                If i = author_w.Length - 1 Then
                    i1 = 0
                    i2 = 1
                Else
                    If i1 = author_w.Length - 1 Then
                        i2 = 0
                    End If
                End If

                Dim author As String = author_w(i)
                Dim author1 As String = author_w(i1)
                Dim author2 As String = author_w(i2)
                c.Controls.Add(New LiteralControl("<br>" + "內-" + author + "<br>" + "外-" + author1))
            Case 20

                Dim author_w() = {"朱志杰", "林敬寅", "陳進富", "賴曉蓉", "賴易昌", "劉建台", "楊尚淳", "王孝中"}
                If author_w.Length Mod 5 = 0 Then '2的倍數時要+第幾輪參數
                    Dim w_n As Integer = Int(DateDiff("d", start_d(dx), v1) / (author_w.Length * 7 / 5)) '第幾輪
                    i = (duty_d5(start_d(dx), v1) + w_n) Mod author_w.Length
                Else
                    i = (duty_d5(start_d(dx), v1)) Mod author_w.Length
                End If
                i1 = i + 1
                i2 = i1 + 1
                If i = author_w.Length - 1 Then
                    i1 = 0
                    i2 = 1
                Else
                    If i1 = author_w.Length - 1 Then
                        i2 = 0
                    End If
                End If

                Dim author As String = author_w(i)
                Dim author1 As String = author_w(i1)
                Dim author2 As String = author_w(i2)
                c.Controls.Add(New LiteralControl("<br>" + "內-" + author + "<br>" + "外-" + author1))
            Case 21

                Dim author_w() = {"朱志杰", "劉建台", "賴曉蓉", "賴易昌", "林敬寅", "楊尚淳", "王孝中"}
                If author_w.Length Mod 5 = 0 Then '2的倍數時要+第幾輪參數
                    Dim w_n As Integer = Int(DateDiff("d", start_d(dx), v1) / (author_w.Length * 7 / 5)) '第幾輪
                    i = (duty_d5(start_d(dx), v1) + w_n) Mod author_w.Length
                Else
                    i = (duty_d5(start_d(dx), v1)) Mod author_w.Length
                End If
                i1 = i + 1
                i2 = i1 + 1
                If i = author_w.Length - 1 Then
                    i1 = 0
                    i2 = 1
                Else
                    If i1 = author_w.Length - 1 Then
                        i2 = 0
                    End If
                End If

                Dim author As String = author_w(i)
                Dim author1 As String = author_w(i1)
                Dim author2 As String = author_w(i2)
                c.Controls.Add(New LiteralControl("<br>" + "內-" + author + "<br>" + "外-" + author1))
            Case 22

                Dim author_w() = {"顏廷諭", "吳俐節", "朱志杰", "劉建台", "賴曉蓉", "賴易昌", "林敬寅", "楊尚淳", "王孝中"}
                If author_w.Length Mod 5 = 0 Then '2的倍數時要+第幾輪參數
                    Dim w_n As Integer = Int(DateDiff("d", start_d(dx), v1) / (author_w.Length * 7 / 5)) '第幾輪
                    i = (duty_d7(start_d(dx), v1) + w_n) Mod author_w.Length
                Else
                    i = (duty_d7(start_d(dx), v1)) Mod author_w.Length
                End If
                
                Dim author_w1() = {"顏廷諭", "朱志杰", "劉建台", "賴曉蓉", "賴易昌", "林敬寅", "楊尚淳", "王孝中"}
                If author_w1.Length Mod 5 = 0 Then '2的倍數時要+第幾輪參數
                    Dim w_n1 As Integer = Int(DateDiff("d", start_d(dx), v1) / (author_w1.Length * 7 / 5)) '第幾輪
                    ii = (duty_d7(start_d(dx), v1) + w_n1) Mod author_w1.Length
                Else
                    ii = (duty_d7(start_d(dx), v1)) Mod author_w1.Length
                End If
                
                i1 = ii + 1
              
                If ii = author_w1.Length - 1 Then
                    i1 = 0
                   
                
                End If
                Dim author As String = author_w(i)
                Dim author1 As String = author_w1(i1)
                'MsgBox("author==" & author)
                'MsgBox("author1==" & author1)
                'MsgBox(String.Compare(author, author1))
               
                If String.Compare(author, author1) = 0 Then
                   
                    i1 = i1 + 1
                    
                    If i1 = 8 Then
                        'MsgBox("i1==" & i1)
                        i1 = 0
                    End If
                    'author1 = author_w1(i1)
                   
                End If
                author1 = author_w1(i1)
                'MsgBox("author==" & author)
                'MsgBox("author1==" & author1)
                ' Dim author2 As String = author_w(i2)
                c.Controls.Add(New LiteralControl("<br>" + "內-" + author + "<br>" + "外-" + author1))
            Case 23

                Dim author_w() = {"顏廷諭", "朱志杰", "劉建台", "賴曉蓉", "賴易昌", "林敬寅", "楊尚淳", "王孝中"}
                If author_w.Length Mod 5 = 0 Then '2的倍數時要+第幾輪參數
                    Dim w_n As Integer = Int(DateDiff("d", start_d(dx), v1) / (author_w.Length * 7 / 5)) '第幾輪
                    i = (duty_d7(start_d(dx), v1) + w_n) Mod author_w.Length
                Else
                    i = (duty_d7(start_d(dx), v1)) Mod author_w.Length
                End If
                
                Dim author_w1() = {"顏廷諭", "朱志杰", "劉建台", "賴曉蓉", "賴易昌", "林敬寅", "楊尚淳", "王孝中"}
                If author_w1.Length Mod 5 = 0 Then '2的倍數時要+第幾輪參數
                    Dim w_n1 As Integer = Int(DateDiff("d", start_d(dx), v1) / (author_w1.Length * 7 / 5)) '第幾輪
                    ii = (duty_d7(start_d(dx), v1) + w_n1) Mod author_w1.Length
                Else
                    ii = (duty_d7(start_d(dx), v1)) Mod author_w1.Length
                End If
                
                i1 = ii + 1
              
                If ii = author_w1.Length - 1 Then
                    i1 = 0
                   
                
                End If
                Dim author As String = author_w(i)
                Dim author1 As String = author_w1(i1)
                'MsgBox("author==" & author)
                'MsgBox("author1==" & author1)
                'MsgBox(String.Compare(author, author1))
               
                If String.Compare(author, author1) = 0 Then
                   
                    i1 = i1 + 1
                    
                    If i1 = 8 Then
                        'MsgBox("i1==" & i1)
                        i1 = 0
                    End If
                    'author1 = author_w1(i1)
                   
                End If
                author1 = author_w1(i1)
                'MsgBox("author==" & author)
                'MsgBox("author1==" & author1)
                ' Dim author2 As String = author_w(i2)
                c.Controls.Add(New LiteralControl("<br>" + "內-" + author + "<br>" + "外-" + author1))
            Case 24

                Dim author_w() = {"賴曉蓉", "朱志杰", "吳俐節", "楊尚淳", "王孝中", "顏廷諭", "劉建台"}
                If author_w.Length Mod 5 = 0 Then '2的倍數時要+第幾輪參數
                    Dim w_n As Integer = Int(DateDiff("d", start_d(dx), v1) / (author_w.Length * 7 / 5)) '第幾輪
                    i = (duty_d7(start_d(dx), v1) + w_n) Mod author_w.Length
                Else
                    i = (duty_d7(start_d(dx), v1)) Mod author_w.Length
                End If
                
                Dim author_w1() = {"賴曉蓉", "朱志杰", "林敬寅", "楊尚淳", "王孝中", "顏廷諭", "劉建台"}
                If author_w1.Length Mod 5 = 0 Then '2的倍數時要+第幾輪參數
                    Dim w_n1 As Integer = Int(DateDiff("d", start_d(dx), v1) / (author_w1.Length * 7 / 5)) '第幾輪
                    ii = (duty_d7(start_d(dx), v1) + w_n1) Mod author_w1.Length
                Else
                    ii = (duty_d7(start_d(dx), v1)) Mod author_w1.Length
                End If
                
                i1 = ii + 1
              
                If ii = author_w1.Length - 1 Then
                    i1 = 0
                   
                
                End If
                Dim author As String = author_w(i)
                Dim author1 As String = author_w1(i1)
                'MsgBox("author==" & author)
                'MsgBox("author1==" & author1)
                'MsgBox(String.Compare(author, author1))
               
                If String.Compare(author, author1) = 0 Then
                   
                    i1 = i1 + 1
                    
                    If i1 = 8 Then
                        'MsgBox("i1==" & i1)
                        i1 = 0
                    End If
                    'author1 = author_w1(i1)
                   
                End If
                author1 = author_w1(i1)
                'MsgBox("author==" & author)
                'MsgBox("author1==" & author1)
                ' Dim author2 As String = author_w(i2)
                c.Controls.Add(New LiteralControl("<br>" + "內-" + author + "<br>" + "外-" + author1))
            Case 25

                Dim author_w() = {"王孝中", "顏廷諭", "朱志杰", "倪翊凱", "楊尚淳", "林敬寅", "賴曉蓉", "劉建台", "吳俐節", "丁振原"}
                If author_w.Length Mod 5 = 0 Then '2的倍數時要+第幾輪參數
                    Dim w_n As Integer = Int(DateDiff("d", start_d(dx), v1) / (author_w.Length * 7 / 5)) '第幾輪
                    i = (duty_d7(start_d(dx), v1) + w_n) Mod author_w.Length
                Else
                    i = (duty_d7(start_d(dx), v1)) Mod author_w.Length
                End If
                
                Dim author_w1() = {"王孝中", "顏廷諭", "朱志杰", "倪翊凱", "楊尚淳", "林敬寅", "賴曉蓉", "劉建台", "丁振原"}
                If author_w1.Length Mod 5 = 0 Then '2的倍數時要+第幾輪參數
                    Dim w_n1 As Integer = Int(DateDiff("d", start_d(dx), v1) / (author_w1.Length * 7 / 5)) '第幾輪
                    ii = (duty_d7(start_d(dx), v1) + w_n1) Mod author_w1.Length
                Else
                    ii = (duty_d7(start_d(dx), v1)) Mod author_w1.Length
                End If
                
                i1 = ii + 1
              
                If ii = author_w1.Length - 1 Then
                    i1 = 0
                   
                
                End If
                Dim author As String = author_w(i)
                Dim author1 As String = author_w1(i1)
                'MsgBox("author==" & author)
                'MsgBox("author1==" & author1)
                'MsgBox(String.Compare(author, author1))
               
                If String.Compare(author, author1) = 0 Then
                   
                    i1 = i1 + 1
                    
                    If i1 = 8 Then
                        'MsgBox("i1==" & i1)
                        i1 = 0
                    End If
                    'author1 = author_w1(i1)
                   
                End If
                author1 = author_w1(i1)
                'MsgBox("author==" & author)
                'MsgBox("author1==" & author1)
                ' Dim author2 As String = author_w(i2)
                c.Controls.Add(New LiteralControl("<br>" + "內-" + author + "<br>" + "外-" + author1))
            Case 26

                Dim author_w() = {"吳俐節", "丁振原", "王孝中", "顏廷諭", "朱志杰", "楊尚淳", "賴曉蓉", "劉建台", "倪翊凱"}
                If author_w.Length Mod 5 = 0 Then '2的倍數時要+第幾輪參數
                    Dim w_n As Integer = Int(DateDiff("d", start_d(dx), v1) / (author_w.Length * 7 / 5)) '第幾輪
                    i = (duty_d7(start_d(dx), v1) + w_n) Mod author_w.Length
                Else
                    i = (duty_d7(start_d(dx), v1)) Mod author_w.Length
                End If
                
                Dim author_w1() = {"丁振原", "王孝中", "顏廷諭", "朱志杰", "楊尚淳", "賴曉蓉", "劉建台", "倪翊凱"}
                If author_w1.Length Mod 5 = 0 Then '2的倍數時要+第幾輪參數
                    Dim w_n1 As Integer = Int(DateDiff("d", start_d(dx), v1) / (author_w1.Length * 7 / 5)) '第幾輪
                    ii = (duty_d7(start_d(dx), v1) + w_n1) Mod author_w1.Length
                Else
                    ii = (duty_d7(start_d(dx), v1)) Mod author_w1.Length
                End If
                
                i1 = ii + 1
              
                If ii = author_w1.Length - 1 Then
                    i1 = 0
                   
                
                End If
                Dim author As String = author_w(i)
                Dim author1 As String = author_w1(i1)
                'MsgBox("author==" & author)
                'MsgBox("author1==" & author1)
                'MsgBox(String.Compare(author, author1))
               
                If String.Compare(author, author1) = 0 Then
                   
                    i1 = i1 + 1
                    
                    If i1 = 8 Then
                        'MsgBox("i1==" & i1)
                        i1 = 0
                    End If
                    'author1 = author_w1(i1)
                   
                End If
                author1 = author_w1(i1)
                'MsgBox("author==" & author)
                'MsgBox("author1==" & author1)
                ' Dim author2 As String = author_w(i2)
                c.Controls.Add(New LiteralControl("<br>" + "內-" + author + "<br>" + "外-" + author1))
            Case 27

                Dim author_w() = {"楊尚淳", "黃憲舜", "賴曉蓉", "劉建台", "倪翊凱", "朱志杰", "吳俐節", "丁振原", "王孝中", "顏廷諭"}
                If author_w.Length Mod 5 = 0 Then '2的倍數時要+第幾輪參數
                    Dim w_n As Integer = Int(DateDiff("d", start_d(dx), v1) / (author_w.Length * 7 / 5)) '第幾輪
                    i = (duty_d7(start_d(dx), v1) + w_n) Mod author_w.Length
                Else
                    i = (duty_d7(start_d(dx), v1)) Mod author_w.Length
                End If
                
                Dim author_w1() = {"黃憲舜", "劉建台", "賴曉蓉", "倪翊凱", "朱志杰", "楊尚淳", "丁振原", "王孝中", "顏廷諭"}
                If author_w1.Length Mod 5 = 0 Then '2的倍數時要+第幾輪參數
                    Dim w_n1 As Integer = Int(DateDiff("d", start_d(dx), v1) / (author_w1.Length * 7 / 5)) '第幾輪
                    ii = (duty_d7(start_d(dx), v1) + w_n1) Mod author_w1.Length
                Else
                    ii = (duty_d7(start_d(dx), v1)) Mod author_w1.Length
                End If
                
                i1 = ii + 1
              
                If ii = author_w1.Length - 1 Then
                    i1 = 0
                   
                
                End If
                Dim author As String = author_w(i)
                Dim author1 As String = author_w1(i1)
                'MsgBox("author==" & author)
                'MsgBox("author1==" & author1)
                'MsgBox(String.Compare(author, author1))
               
                If String.Compare(author, author1) = 0 Then
                   
                    i1 = i1 + 1
                    
                    If i1 = 8 Then
                        'MsgBox("i1==" & i1)
                        i1 = 0
                    End If
                    'author1 = author_w1(i1)
                   
                End If
                author1 = author_w1(i1)
                'MsgBox("author==" & author)
                'MsgBox("author1==" & author1)
                ' Dim author2 As String = author_w(i2)
                c.Controls.Add(New LiteralControl("<br>" + "內-" + author + "<br>" + "外-" + author1))
            Case 28

                Dim author_w() = {"王孝中", "顏廷諭", "楊尚淳", "賴曉蓉", "朱志杰", "劉建台", "吳俐節", "丁振原", "倪翊凱"}
                If author_w.Length Mod 5 = 0 Then '2的倍數時要+第幾輪參數
                    Dim w_n As Integer = Int(DateDiff("d", start_d(dx), v1) / (author_w.Length * 7 / 5)) '第幾輪
                    i = (duty_d7(start_d(dx), v1) + w_n) Mod author_w.Length
                Else
                    i = (duty_d7(start_d(dx), v1)) Mod author_w.Length
                End If
                
                Dim author_w1() = {"王孝中", "顏廷諭", "楊尚淳", "賴曉蓉", "朱志杰", "劉建台", "丁振原", "倪翊凱"}
                If author_w1.Length Mod 5 = 0 Then '2的倍數時要+第幾輪參數
                    Dim w_n1 As Integer = Int(DateDiff("d", start_d(dx), v1) / (author_w1.Length * 7 / 5)) '第幾輪
                    ii = (duty_d7(start_d(dx), v1) + w_n1) Mod author_w1.Length
                Else
                    ii = (duty_d7(start_d(dx), v1)) Mod author_w1.Length
                End If
                
                i1 = ii + 1
              
                If ii = author_w1.Length - 1 Then
                    i1 = 0
                   
                
                End If
                Dim author As String = author_w(i)
                Dim author1 As String = author_w1(i1)
                'MsgBox("author==" & author)
                'MsgBox("author1==" & author1)
                'MsgBox(String.Compare(author, author1))
               
                If String.Compare(author, author1) = 0 Then
                   
                    i1 = i1 + 1
                    
                    If i1 = 8 Then
                        'MsgBox("i1==" & i1)
                        i1 = 0
                    End If
                    'author1 = author_w1(i1)
                   
                End If
                author1 = author_w1(i1)
                'MsgBox("author==" & author)
                'MsgBox("author1==" & author1)
                ' Dim author2 As String = author_w(i2)
                c.Controls.Add(New LiteralControl("<br>" + "內-" + author + "<br>" + "外-" + author1))
            Case 29

                Dim author_w() = {"黃憲舜", "朱志杰", "劉建台", "吳俐節", "丁振原", "楊尚淳", "賴曉蓉", "倪翊凱", "王孝中", "顏廷諭"}
                If author_w.Length Mod 5 = 0 Then '2的倍數時要+第幾輪參數
                    Dim w_n As Integer = Int(DateDiff("d", start_d(dx), v1) / (author_w.Length * 7 / 5)) '第幾輪
                    i = (duty_d7(start_d(dx), v1) + w_n) Mod author_w.Length
                Else
                    i = (duty_d7(start_d(dx), v1)) Mod author_w.Length
                End If
                
                Dim author_w1() = {"黃憲舜", "朱志杰", "劉建台", "丁振原", "楊尚淳", "賴曉蓉", "倪翊凱", "王孝中", "顏廷諭"}
                If author_w1.Length Mod 5 = 0 Then '2的倍數時要+第幾輪參數
                    Dim w_n1 As Integer = Int(DateDiff("d", start_d(dx), v1) / (author_w1.Length * 7 / 5)) '第幾輪
                    ii = (duty_d7(start_d(dx), v1) + w_n1) Mod author_w1.Length
                Else
                    ii = (duty_d7(start_d(dx), v1)) Mod author_w1.Length
                End If
                
                i1 = ii + 1
              
                If ii = author_w1.Length - 1 Then
                    i1 = 0
                   
                
                End If
                Dim author As String = author_w(i)
                Dim author1 As String = author_w1(i1)
                'MsgBox("author==" & author)
                'MsgBox("author1==" & author1)
                'MsgBox(String.Compare(author, author1))
               
                If String.Compare(author, author1) = 0 Then
                   
                    i1 = i1 + 1
                    
                    If i1 = 8 Then
                        'MsgBox("i1==" & i1)
                        i1 = 0
                    End If
                    'author1 = author_w1(i1)
                   
                End If
                author1 = author_w1(i1)
                'MsgBox("author==" & author)
                'MsgBox("author1==" & author1)
                ' Dim author2 As String = author_w(i2)
                c.Controls.Add(New LiteralControl("<br>" + "內-" + author + "<br>" + "外-" + author1))
            Case 30

                Dim author_w() = {"丁振原", "顏廷諭", "林紫蓁", "黃憲舜", "朱志杰", "劉建台", "楊尚淳", "賴曉蓉", "倪翊凱", "王孝中", "吳俐節"}
                If author_w.Length Mod 2 = 0 Then '2的倍數時要+第幾輪參數
                    Dim w_n As Integer = Int(DateDiff("d", start_d(dx), v1) / (author_w.Length * 7 / 2)) '第幾輪
                    i = (duty_d8(start_d(dx), v1) + w_n) Mod author_w.Length
                Else
                    i = (duty_d8(start_d(dx), v1)) Mod author_w.Length
                End If
                i1 = i + 1
                i2 = i1 + 1
                If i = author_w.Length - 1 Then
                    i1 = 0
                    i2 = 1
                Else
                    If i1 = author_w.Length - 1 Then
                        i2 = 0
                    End If
                End If

                Dim author As String = author_w(i)
                Dim author1 As String = author_w(i1)
                Dim author2 As String = author_w(i2)
                c.Controls.Add(New LiteralControl("<br>" + "內-" + author + "<br>" + "外-" + author1))
            Case 31

                Dim author_w() = {"朱志杰", "吳俐節", "高育源", "顏廷諭", "林紫蓁", "黃憲舜", "楊尚淳", "劉建台", "王孝中", "倪翊凱"}
                If author_w.Length Mod 2 = 0 Then '2的倍數時要+第幾輪參數
                    Dim w_n As Integer = Int(DateDiff("d", start_d(dx), v1) / (author_w.Length * 7 / 2)) '第幾輪
                    i = (duty_d8(start_d(dx), v1) + w_n) Mod author_w.Length
                Else
                    i = (duty_d8(start_d(dx), v1)) Mod author_w.Length
                End If
                i1 = i + 1
                i2 = i1 + 1
                If i = author_w.Length - 1 Then
                    i1 = 0
                    i2 = 1
                Else
                    If i1 = author_w.Length - 1 Then
                        i2 = 0
                    End If
                End If

                Dim author As String = author_w(i)
                Dim author1 As String = author_w(i1)
                Dim author2 As String = author_w(i2)
                c.Controls.Add(New LiteralControl("<br>" + "內-" + author + "<br>" + "外-" + author1))
            Case 32

                Dim author_w() = {"朱志杰", "黃憲舜", "楊尚淳", "劉建台", "王孝中", "倪翊凱", "林紫蓁", "吳俐節", "高育源", "顏廷諭", "林冠葦"}
                If author_w.Length Mod 2 = 0 Then '2的倍數時要+第幾輪參數
                    Dim w_n As Integer = Int(DateDiff("d", start_d(dx), v1) / (author_w.Length * 7 / 2)) '第幾輪
                    i = (duty_d8(start_d(dx), v1) + w_n) Mod author_w.Length
                Else
                    i = (duty_d8(start_d(dx), v1)) Mod author_w.Length
                End If
                i1 = i + 1
                i2 = i1 + 1
                If i = author_w.Length - 1 Then
                    i1 = 0
                    i2 = 1
                Else
                    If i1 = author_w.Length - 1 Then
                        i2 = 0
                    End If
                End If

                Dim author As String = author_w(i)
                Dim author1 As String = author_w(i1)
                Dim author2 As String = author_w(i2)
                c.Controls.Add(New LiteralControl("<br>" + "內-" + author + "<br>" + "外-" + author1))
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

function showweek(t)
 if t <> ""
 showweek = weekdayname(datepart("w",t))
 end if
 end function

 function showmessage(vt2,vt3)
 if vt2 <> ""
 showmessage = vt2 & "<br/>" & vt3 & "小時"
 end if
 end function

 function showdetail(vm1)
 if vm1 <> ""
  if len(vm1) > 30
 showdetail = left(vm1,30) & "..."
  else
 showdetail = vm1
 end if 
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
